"""Single seam between the application and pyOpenVBA.

All pyopenvba imports (public and internal) live here. GUI and services
must never import pyopenvba directly (plan sections 2.9, 5.3, 24).

Internals used and why (pinned pyopenvba==3.4.0):
- ``_encoding_for_codepage``: maps the project's dir-stream code page to a
  Python codec for strict save-time validation; no public equivalent exists.
- ``detect_signature`` / ``CFB``: signature detection without mutating the
  host file, by building a CFB from ``host.vba_project_bytes()``.
"""

from __future__ import annotations

import warnings
from pathlib import Path
from typing import Any

from pyopenvba import ExcelFile, PowerPointFile
from pyopenvba._host import VBAHostFile
from pyopenvba.cfb import CFB
from pyopenvba.exceptions import UnsupportedFormatError, VBAProjectError
from pyopenvba.vba import (
    VBAModuleKind,
    _encoding_for_codepage,
    detect_signature,
)

from vba_addin_editor.adapters.source_codec import (
    split_attribute_header,
    to_editor_text,
    to_vba_crlf,
)
from vba_addin_editor.domain.document import (
    DocumentDraft,
    DocumentSnapshot,
    FileFingerprint,
    HostKind,
    ModuleDisplayKind,
    ModuleSnapshot,
    ProjectSafetyInfo,
)
from vba_addin_editor.domain.results import CandidateVerificationResult

SUPPORTED_EXTENSIONS = (".xlam", ".ppam")
_VBA_ENTRY = {".xlam": "xl/vbaProject.bin", ".ppam": "ppt/vbaProject.bin"}
_HOST_PROCESS = {".xlam": "EXCEL.EXE", ".ppam": "POWERPNT.EXE"}


class AdapterError(Exception):
    """Plain-language adapter failure; .details carries technical info."""

    def __init__(self, message: str, details: dict[str, Any] | None = None) -> None:
        super().__init__(message)
        self.details = details or {}


class PPAMPowerPointFile(PowerPointFile):
    """Application-local PPAM compatibility subclass (plan 4.2, 24.2).

    PPAM is a ZIP OOXML package with ``ppt/vbaProject.bin``, same as PPTM;
    only the extension dispatch whitelist differs. Proven against authentic
    Office fixtures by tests/test_ppam_live.py (release gate).
    """

    _zip_formats = frozenset({".pptm", ".potm", ".ppam"})
    _no_vba_hint = "Make sure the PowerPoint add-in contains a VBA project."


def _host_class(path: Path) -> type[VBAHostFile]:
    suffix = path.suffix.lower()
    if suffix == ".xlam":
        return ExcelFile
    if suffix == ".ppam":
        return PPAMPowerPointFile
    raise AdapterError("Only .xlam and .ppam add-ins are supported.")


def vba_entry_for(path: Path) -> str:
    return _VBA_ENTRY[path.suffix.lower()]


def host_process_for(path: Path) -> str:
    return _HOST_PROCESS[path.suffix.lower()]


def host_kind_for(path: Path) -> HostKind:
    return "excel" if path.suffix.lower() == ".xlam" else "powerpoint"


def make_display_kind(pyopenvba_kind: str, *, is_new_app_class: bool = False) -> str:
    if pyopenvba_kind == "standard":
        return ModuleDisplayKind.STANDARD
    return ModuleDisplayKind.CLASS if is_new_app_class else ModuleDisplayKind.AMBIGUOUS


class PyOpenVBAAdapter:
    # -- safety inspection ------------------------------------------------

    def inspect_signature(self, host: VBAHostFile) -> tuple[bool, tuple[str, ...]]:
        cfb = CFB.from_bytes(host.vba_project_bytes())
        info = detect_signature(cfb)
        return info.present, tuple(info.kinds)

    # -- snapshot ---------------------------------------------------------

    def open_snapshot(
        self,
        path: Path,
        fingerprint: FileFingerprint,
        host_process_running: bool = False,
    ) -> DocumentSnapshot:
        path = Path(path)
        suffix = path.suffix.lower()
        if suffix not in SUPPORTED_EXTENSIONS:
            raise AdapterError("Only .xlam and .ppam add-ins are supported.")
        try:
            host = _host_class(path)(path)
        except UnsupportedFormatError as exc:
            raise AdapterError(
                "Only .xlam and .ppam add-ins are supported.", {"exception": repr(exc)}
            ) from exc
        except VBAProjectError as exc:
            raise AdapterError(
                "This add-in does not contain a readable VBA project. The file has not been changed.",
                {"exception": repr(exc)},
            ) from exc
        except Exception as exc:
            raise AdapterError(
                "The Office add-in package appears damaged or is not a valid add-in file.",
                {"exception": repr(exc)},
            ) from exc
        with host:
            try:
                project = host.vba_project()
            except Exception as exc:
                raise AdapterError(
                    "The VBA project inside this add-in could not be read.",
                    {"exception": repr(exc)},
                ) from exc
            sig_present, sig_kinds = self.inspect_signature(host)
            safety = ProjectSafetyInfo(
                password_protected=bool(
                    project.protection is not None and project.protection.has_password
                ),
                signature_present=sig_present,
                signature_kinds=sig_kinds,
                writable=True,  # refined by save service preflight
                host_process_running=host_process_running,
            )
            modules: list[ModuleSnapshot] = []
            for m in project.modules:
                source = to_editor_text(m.source)
                header, body = split_attribute_header(source)
                kind = make_display_kind(m.kind.name)
                modules.append(
                    ModuleSnapshot(
                        id=f"{m.stream_name}:{len(modules)}",
                        original_name=m.name,
                        kind=kind,
                        pyopenvba_kind=m.kind.name,
                        full_source=source,
                        hidden_header=header,
                        body=body,
                        is_read_only=m.is_read_only,
                        is_private=m.is_private,
                        destructive_ops_safe=(m.kind == VBAModuleKind.standard and not m.is_read_only),
                        ends_with_newline=body.endswith("\n"),
                    )
                )
            return DocumentSnapshot(
                path=path,
                host_kind=host_kind_for(path),
                extension=suffix,  # type: ignore[arg-type]
                file_fingerprint=fingerprint,
                project_name=project.name,
                code_page=project.code_page,
                safety=safety,
                modules=tuple(modules),
            )

    def codepage_encoding(self, code_page: int) -> str:
        return _encoding_for_codepage(code_page)

    # -- candidate build (plan 14 Stage G/H/I, replay algorithm 25) -------

    def build_candidate(
        self,
        original_path: Path,
        draft: DocumentDraft,
        candidate_path: Path,
        *,
        allow_signature_removal: bool,
    ) -> None:
        """Open the ORIGINAL fresh, replay the draft onto it, serialize to candidate."""
        host = _host_class(original_path)(original_path)
        with host:
            project = host.vba_project()
            try:
                # 1. Delete permitted original modules gone in final state.
                for mod in draft.deleted_original_modules():
                    project.delete_module(mod.origin_name or mod.current_name)

                # 2. Renames via unique temporary names to avoid collisions.
                rename_items = draft.changed_names()
                temp_map: dict[str, str] = {}
                for item in rename_items:
                    temp = _make_unique_temp_name(project)
                    project.rename_module(item.origin_name or item.current_name, temp)
                    temp_map[item.id] = temp
                for item in rename_items:
                    project.rename_module(temp_map[item.id], item.current_name)

                # 3. Add new modules.
                for item in draft.new_modules():
                    kind = (
                        VBAModuleKind.standard
                        if item.pyopenvba_kind == "standard"
                        else VBAModuleKind.other
                    )
                    project.add_module(item.current_name, to_vba_crlf(item.body), kind=kind)

                # 4. Body changes on surviving originals.
                for item in draft.changed_existing_modules():
                    host.set_module(item.current_name, to_vba_crlf(item.body))

                with warnings.catch_warnings(record=True) as caught:
                    warnings.simplefilter("always")
                    host.save(
                        candidate_path,
                        allow_invalidate_signature=allow_signature_removal,
                    )
                unexpected = [str(w.message) for w in caught if not allow_signature_removal]
                if unexpected:
                    raise AdapterError(
                        "The VBA project signature was dropped without confirmation; save aborted.",
                        {"warnings": unexpected},
                    )
            except ValueError as exc:
                raise AdapterError(str(exc)) from exc

    # -- candidate verification (plan 14 Stage J, 77) ----------------------

    def verify_candidate(
        self,
        original_path: Path,
        candidate_path: Path,
        expected: DocumentDraft,
    ) -> CandidateVerificationResult:
        problems: list[str] = []
        details: dict[str, Any] = {}
        entry = vba_entry_for(original_path)

        import hashlib
        import zipfile
        if not candidate_path.exists() or candidate_path.stat().st_size == 0:
            return CandidateVerificationResult(False, ("Candidate file missing or empty.",))
        if candidate_path.suffix.lower() != original_path.suffix.lower():
            problems.append("Candidate extension does not match original.")
        try:
            with zipfile.ZipFile(candidate_path) as cand, zipfile.ZipFile(original_path) as orig:
                if cand.testzip() is not None:
                    problems.append("Candidate ZIP has a corrupt entry.")
                cand_names = set(cand.namelist())
                orig_names = set(orig.namelist())
                if cand_names != orig_names:
                    problems.append(
                        f"ZIP entry set changed: only-in-original={sorted(orig_names - cand_names)} "
                        f"only-in-candidate={sorted(cand_names - orig_names)}"
                    )
                differing: list[str] = []
                for name in sorted(orig_names & cand_names):
                    if name == entry:
                        continue
                    if hashlib.sha256(orig.read(name)).digest() != hashlib.sha256(
                        cand.read(name)
                    ).digest():
                        differing.append(name)
                if differing:
                    problems.append(f"Non-VBA package payload changed: {differing}")
        except zipfile.BadZipFile as exc:
            return CandidateVerificationResult(
                False, ("Candidate is not a valid Office package.", {"exception": repr(exc)})
            )

        expected_final = {m.current_name.casefold(): m for m in expected.final_module_state()}
        try:
            cand_host = _host_class(candidate_path)(candidate_path)
        except Exception as exc:  # noqa: BLE001 - translated below
            return CandidateVerificationResult(
                False, ("Candidate could not be reopened.", {"exception": repr(exc)})
            )
        with cand_host:
            project = cand_host.vba_project()
            structural = project.validate(cand_host._get_cfb())
            if structural:
                problems.extend(f"Structural validation: {s}" for s in structural)
            if project.code_page != expected.baseline.code_page:
                problems.append(
                    f"Code page changed: {expected.baseline.code_page} -> {project.code_page}"
                )
            actual = {m.name.casefold(): m for m in project.modules}
            for name in sorted(set(actual) - set(expected_final)):
                problems.append(f"Unexpected module in candidate: {name!r}")
            for name, mod in expected_final.items():
                if name not in actual:
                    problems.append(f"Expected module missing in candidate: {mod.current_name!r}")
                    continue
                got = cand_host.get_module(actual[name].name)
                got_body = to_vba_crlf(split_attribute_header(to_editor_text(got))[1])
                want = to_vba_crlf(mod.body)
                if got_body != want:
                    problems.append(f"Body mismatch for module {mod.current_name!r}.")
            sig_present, kinds = self.inspect_signature(cand_host)
            if expected.baseline.safety.signature_present:
                if sig_present and allow_removal_expected(expected):
                    problems.append("Signature still present after approved removal.")
                details["candidate_signature_kinds"] = list(kinds)
            else:
                if sig_present:
                    problems.append("Candidate unexpectedly gained a signature.")
        return CandidateVerificationResult(not problems, tuple(problems), details)

    # -- import/export -----------------------------------------------------

    def export_module_source(self, snapshot: DocumentSnapshot, module_name: str) -> str:
        mod = next(
            (m for m in snapshot.modules if m.original_name.casefold() == module_name.casefold()),
            None,
        )
        if mod is None:
            raise KeyError(module_name)
        return to_vba_crlf(mod.body if mod.kind == ModuleDisplayKind.STANDARD else mod.full_source)

    def parse_import_file(self, path: Path) -> tuple[str, str, bool]:
        """Return (suggested_name, body_or_full_source, is_class_form)."""
        raw = path.read_bytes()
        if len(raw) > 8 * 1024 * 1024:
            raise AdapterError("Import file is unreasonably large.")
        text = to_editor_text(raw.decode("utf-8", errors="strict"))
        header, _ = split_attribute_header(text)
        is_class = "VERSION 1.0 CLASS" in header
        name = path.stem
        return name, text, is_class


def allow_removal_expected(draft: DocumentDraft) -> bool:
    return draft.signed_save_confirmed


def _make_unique_temp_name(project: Any) -> str:
    """ASCII temp component name guaranteed unique and short (plan 14 Stage H)."""
    import uuid

    existing = {m.name.casefold() for m in project.modules}
    for _ in range(100):
        candidate = f"VBAAE_TMP_{uuid.uuid4().hex[:6].upper()}"
        if candidate.casefold() not in existing:
            return candidate
    raise AdapterError("Could not generate a temporary module name for rename.")
