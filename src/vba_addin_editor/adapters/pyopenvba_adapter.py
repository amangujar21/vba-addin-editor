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

import hashlib
import warnings
import zipfile
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
    parse_project_stream,
    parse_projectwm,
    serialize_project_stream,
)

from vba_addin_editor.adapters.ooxml_package_adapter import OoxmlPackageAdapter, PackageError
from vba_addin_editor.adapters.source_codec import (
    split_attribute_header,
    to_editor_text,
    to_vba_crlf,
)
from vba_addin_editor.domain.capabilities import (
    ComponentCapabilities,
    build_project_index,
    classify_component,
    empty_project_index,
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

SUPPORTED_EXTENSIONS = (".xlam", ".ppam", ".pptm")
_VBA_ENTRY = {
    ".xlam": "xl/vbaProject.bin",
    ".ppam": "ppt/vbaProject.bin",
    ".pptm": "ppt/vbaProject.bin",
}
_HOST_PROCESS = {
    ".xlam": "EXCEL.EXE",
    ".ppam": "POWERPNT.EXE",
    ".pptm": "POWERPNT.EXE",
}


class AdapterError(Exception):
    """Plain-language adapter failure; .details carries technical info."""

    def __init__(self, message: str, details: dict[str, Any] | None = None) -> None:
        super().__init__(message)
        self.details = details or {}


def _decrypt_project_data(value: str) -> bytes:
    """Decode an MS-OVBA Data Encryption value from CMG/DPB/GC.

    The encoded record carries its seed, version, project key, ignored-byte
    count, four-byte little-endian data length, and data. Protection state
    must be read from decoded CMG bits; raw DPB length is not a protection
    signal.
    """
    try:
        encoded = bytes.fromhex(value)
    except ValueError as exc:
        raise AdapterError("The VBA project protection metadata is malformed.") from exc
    if len(encoded) < 7:
        raise AdapterError("The VBA project protection metadata is truncated.")

    seed = encoded[0]
    if encoded[1] ^ seed != 2:
        raise AdapterError("The VBA project protection metadata has an unsupported version.")

    ignored_length = (seed & 0x06) // 2
    encrypted_byte_2 = encoded[1]
    encrypted_byte_1 = encoded[2]
    unencrypted_byte_1 = encoded[2] ^ seed
    decoded = bytearray()
    for encrypted_byte in encoded[3:]:
        byte = encrypted_byte ^ ((encrypted_byte_2 + unencrypted_byte_1) & 0xFF)
        decoded.append(byte)
        encrypted_byte_2 = encrypted_byte_1
        encrypted_byte_1 = encrypted_byte
        unencrypted_byte_1 = byte

    length_offset = ignored_length
    data_offset = length_offset + 4
    if len(decoded) < data_offset:
        raise AdapterError("The VBA project protection metadata has no data length.")
    data_length = int.from_bytes(decoded[length_offset:data_offset], "little")
    if len(decoded) != data_offset + data_length:
        raise AdapterError("The VBA project protection metadata has an invalid data length.")
    return bytes(decoded[data_offset:])


def _has_active_project_protection(protection: Any) -> bool:
    """Return whether decoded CMG says user, host, or VBE protection is active.

    pyOpenVBA 3.4.0 derives ``has_password`` from raw DPB string length. DPB
    describes password material, while CMG is the authoritative active
    protection state. Fall back to pyOpenVBA only when CMG is absent; malformed
    or unexpected CMG data fails closed because its protection flags are unknown.
    """
    if protection is None:
        return False
    cmg = getattr(protection, "cmg", "")
    if not cmg:
        return bool(getattr(protection, "has_password", False))
    try:
        state = _decrypt_project_data(cmg)
    except AdapterError:
        return True
    if len(state) != 4:
        return True
    flags = int.from_bytes(state, "little")
    return bool(flags & 0x07)


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
    if suffix == ".pptm":
        return PowerPointFile
    raise AdapterError("Only .xlam, .ppam, and .pptm files are supported.")


def vba_entry_for(path: Path) -> str:
    return _VBA_ENTRY[path.suffix.lower()]


def host_process_for(path: Path) -> str:
    return _HOST_PROCESS[path.suffix.lower()]


def host_kind_for(path: Path) -> HostKind:
    suffix = path.suffix.lower()
    if suffix == ".xlam":
        return "excel"
    if suffix in {".ppam", ".pptm"}:
        return "powerpoint"
    raise AdapterError("Only .xlam, .ppam, and .pptm files are supported.")


def make_display_kind(pyopenvba_kind: str, *, is_new_app_class: bool = False) -> str:
    if pyopenvba_kind == "standard":
        return ModuleDisplayKind.STANDARD
    return ModuleDisplayKind.CLASS if is_new_app_class else ModuleDisplayKind.AMBIGUOUS


def _project_bytes_lossy(raw: bytes, code_page: int) -> tuple[object | None, bool]:
    """Parse PROJECT and report whether decoding required replacement characters."""
    encoding = _encoding_for_codepage(code_page)
    lossy = False
    try:
        raw.decode(encoding, errors="strict")
    except UnicodeDecodeError:
        lossy = True
    try:
        parsed = parse_project_stream(raw, code_page=code_page)
    except Exception:  # noqa: BLE001 - fail closed at the classifier
        return None, True
    return parsed, lossy


def _designer_storage_names(cfb: CFB) -> frozenset[str]:
    try:
        storages = cfb.list_storages()
    except Exception:  # noqa: BLE001
        return frozenset()
    return frozenset(name.casefold() for name in storages if name.upper() != "VBA")


def _index_from_project(parsed: object | None, *, lossy: bool, missing: bool) -> object:
    if parsed is None:
        return empty_project_index(missing=missing, lossy=lossy)
    documents = [name for name, _hid in getattr(parsed, "document_modules", [])]
    return build_project_index(
        standard_modules=list(getattr(parsed, "standard_modules", [])),
        class_modules=list(getattr(parsed, "class_modules", [])),
        document_modules=documents,
        base_classes=list(getattr(parsed, "base_classes", [])),
        lossy=lossy,
        missing=missing,
    )


def classify_host_module(
    module: object,
    *,
    index,
    designer_storages: frozenset[str],
) -> ComponentCapabilities:
    return classify_component(
        logical_name=module.name,
        stream_name=module.stream_name or module.name,
        dir_kind=module.kind.name,
        is_read_only=bool(module.is_read_only),
        index=index,
        designer_storages=designer_storages,
    )


class PyOpenVBAAdapter:
    def __init__(self, package_adapter: OoxmlPackageAdapter | None = None) -> None:
        self.package_adapter = package_adapter or OoxmlPackageAdapter()

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
            raise AdapterError("Only .xlam, .ppam, and .pptm files are supported.")
        try:
            host = _host_class(path)(path)
        except UnsupportedFormatError as exc:
            raise AdapterError(
                "Only .xlam, .ppam, and .pptm files are supported.", {"exception": repr(exc)}
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
            cfb = CFB.from_bytes(host.vba_project_bytes())
            try:
                project_raw = cfb.get_stream("PROJECT")
                parsed_project, lossy = _project_bytes_lossy(project_raw, project.code_page)
                missing_project = parsed_project is None
            except KeyError:
                parsed_project, lossy, missing_project = None, False, True
            type_index = _index_from_project(
                parsed_project, lossy=lossy, missing=missing_project
            )
            designer_storages = _designer_storage_names(cfb)
            safety = ProjectSafetyInfo(
                password_protected=_has_active_project_protection(project.protection),
                signature_present=sig_present,
                signature_kinds=sig_kinds,
                writable=True,  # refined by save service preflight
                host_process_running=host_process_running,
            )
            modules: list[ModuleSnapshot] = []
            for m in project.modules:
                source = to_editor_text(m.source)
                header, body = split_attribute_header(source)
                caps = classify_host_module(
                    m, index=type_index, designer_storages=designer_storages
                )
                modules.append(
                    ModuleSnapshot(
                        id=f"{m.stream_name}:{len(modules)}",
                        original_name=m.name,
                        kind=caps.display_kind,
                        pyopenvba_kind=m.kind.name,
                        full_source=source,
                        hidden_header=header,
                        body=body,
                        is_read_only=m.is_read_only,
                        is_private=m.is_private,
                        destructive_ops_safe=caps.destructive_ops_safe,
                        ends_with_newline=body.endswith("\n"),
                        can_delete=caps.can_delete,
                        can_rename=caps.can_rename,
                        restriction_reason=caps.restriction_reason,
                        stream_name=m.stream_name or m.name,
                        project_item_kind=caps.project_item_kind,
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
        _assert_destructive_ops_against_snapshot(draft)
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
                        # pyOpenVBA 3.4.0 can falsely infer protection from raw
                        # DPB length. Override only when decoded CMG proves the
                        # project is unlocked; active/malformed protection stays
                        # blocked by both the service preflight and the library.
                        allow_protected=not _has_active_project_protection(project.protection),
                    )
                unexpected = [str(w.message) for w in caught if not allow_signature_removal]
                if unexpected:
                    raise AdapterError(
                        "The VBA project signature was dropped without confirmation; save aborted.",
                        {"warnings": unexpected},
                    )
            except ValueError as exc:
                raise AdapterError(str(exc)) from exc
        surviving = {mod.current_name.casefold() for mod in draft.final_module_state()}
        logical_deletes = [
            mod.origin_name or mod.current_name
            for mod in draft.deleted_original_modules()
            if (mod.origin_name or mod.current_name).casefold() not in surviving
        ]
        if logical_deletes:
            _scrub_project_declarations(
                candidate_path,
                logical_deletes,
                draft.baseline.code_page,
                self.package_adapter,
            )

    # -- candidate verification (plan 14 Stage J, 77) ----------------------

    def verify_candidate(
        self,
        reference_path: Path,
        candidate_path: Path,
        expected: DocumentDraft,
        *,
        allowed_non_vba_changes: frozenset[str] = frozenset(),
    ) -> CandidateVerificationResult:
        """Verify the candidate against a reference package.

        `reference_path` is the original for pre-commit verification and the
        backup for post-commit verification (plan 12.12).
        """
        problems: list[str] = []
        details: dict[str, Any] = {}
        entry = vba_entry_for(reference_path)

        if not candidate_path.exists() or candidate_path.stat().st_size == 0:
            return CandidateVerificationResult(False, ("Candidate file missing or empty.",))
        if candidate_path.suffix.lower() != reference_path.suffix.lower():
            problems.append("Candidate extension does not match original.")
        try:
            with zipfile.ZipFile(candidate_path) as cand, zipfile.ZipFile(reference_path) as ref:
                if cand.testzip() is not None:
                    problems.append("Candidate ZIP has a corrupt entry.")
                cand_names = set(cand.namelist())
                ref_names = set(ref.namelist())
                if cand_names != ref_names:
                    problems.append(
                        f"ZIP entry set changed: only-in-original={sorted(ref_names - cand_names)} "
                        f"only-in-candidate={sorted(cand_names - ref_names)}"
                    )
                differing: list[str] = []
                for name in sorted(ref_names & cand_names):
                    if name == entry or name in allowed_non_vba_changes:
                        continue
                    if (
                        hashlib.sha256(ref.read(name)).digest()
                        != hashlib.sha256(cand.read(name)).digest()
                    ):
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
                expected_kind = (
                    "standard" if mod.pyopenvba_kind == "standard" else "other"
                )
                if actual[name].kind.name != expected_kind:
                    problems.append(
                        f"Module kind mismatch for {mod.current_name!r}: "
                        f"{actual[name].kind.name} != {expected_kind}."
                    )
            problems.extend(_verify_project_cleanup(cand_host, expected))
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


def _assert_destructive_ops_against_snapshot(draft: DocumentDraft) -> None:
    trusted = {module.id: module for module in draft.baseline.modules}
    for mod in draft.deleted_original_modules():
        snap = trusted.get(mod.id)
        if snap is None or not snap.can_delete:
            reason = getattr(snap, "restriction_reason", None) if snap is not None else "unknown"
            raise AdapterError(
                f"Deleting {mod.origin_name or mod.current_name!r} is not permitted "
                f"({reason or 'unverified component'}).",
                {"module_id": mod.id, "reason": reason},
            )
    for mod in draft.changed_names():
        snap = trusted.get(mod.id)
        if snap is None or not snap.can_rename:
            reason = getattr(snap, "restriction_reason", None) if snap is not None else "unknown"
            raise AdapterError(
                f"Renaming {mod.origin_name or mod.current_name!r} is not permitted "
                f"({reason or 'unverified component'}).",
                {"module_id": mod.id, "reason": reason},
            )


def _scrub_project_declarations(
    candidate_path: Path,
    logical_names: list[str],
    code_page: int,
    package_adapter: OoxmlPackageAdapter,
) -> None:
    """Shim pyOpenVBA 3.4.0: save() deletes PROJECT lines by stream name.

    serialize_project_stream expects logical names. When they differ, Class=
    declarations leak. Rewrite PROJECT with the logical names after save.
    ZIP rewriting is owned by OoxmlPackageAdapter.
    """
    entry = vba_entry_for(candidate_path)
    with zipfile.ZipFile(candidate_path) as package:
        raw = package.read(entry)
    cfb = CFB.from_bytes(raw)
    try:
        project_raw = cfb.get_stream("PROJECT")
    except KeyError:
        return
    rewritten = serialize_project_stream(
        project_raw,
        {},
        delete_names=set(logical_names),
        code_page=code_page,
    )
    cfb.write_stream("PROJECT", rewritten)
    try:
        package_adapter.replace_members(candidate_path, {entry: cfb.to_bytes()})
    except PackageError as exc:
        raise AdapterError(str(exc), getattr(exc, "details", None)) from exc


def _verify_project_cleanup(host: VBAHostFile, expected: DocumentDraft) -> list[str]:
    problems: list[str] = []
    cfb = CFB.from_bytes(host.vba_project_bytes())
    expected_names = {
        module.current_name.casefold() for module in expected.final_module_state()
    }
    surviving_streams = set(expected_names)
    for module in expected.final_module_state():
        if module.stream_name:
            surviving_streams.add(module.stream_name.casefold())
    deleted_names = {
        (module.origin_name or module.current_name).casefold()
        for module in expected.deleted_original_modules()
        if (module.origin_name or module.current_name).casefold() not in expected_names
    }
    deleted_streams = {
        (module.stream_name or module.origin_name or module.current_name)
        for module in expected.deleted_original_modules()
        if (module.stream_name or module.origin_name or module.current_name).casefold()
        not in surviving_streams
    }
    try:
        parsed, lossy = _project_bytes_lossy(cfb.get_stream("PROJECT"), expected.baseline.code_page)
    except KeyError:
        problems.append("Candidate PROJECT stream is missing.")
        return problems
    if parsed is None:
        problems.append("Candidate PROJECT stream could not be parsed.")
        return problems
    if lossy:
        problems.append("Candidate PROJECT stream decoded with replacement characters.")
    declared = {name.casefold() for name in parsed.standard_modules}
    declared.update(name.casefold() for name in parsed.class_modules)
    declared.update(name.casefold() for name, _hid in parsed.document_modules)
    declared.update(name.casefold() for name in parsed.base_classes)
    for name in sorted(deleted_names):
        if name in declared:
            problems.append(f"Deleted module still declared in PROJECT: {name!r}")
    dir_names = {module.name.casefold() for module in host.vba_project().modules}
    for name in sorted(deleted_names):
        if name in dir_names:
            problems.append(f"Deleted module still listed in dir: {name!r}")
    try:
        vba_streams = {item.casefold() for item in cfb.list_streams_in_storage("VBA")}
    except Exception as exc:  # noqa: BLE001 - inability to verify is failure
        problems.append(f"Could not enumerate VBA streams ({type(exc).__name__}).")
        return problems
    for stream in deleted_streams:
        if stream and stream.casefold() in vba_streams:
            problems.append(f"Deleted module stream still present: {stream!r}")
    try:
        wm_raw = cfb.get_stream("PROJECTwm")
    except KeyError:
        if deleted_names:
            problems.append("Candidate PROJECTwm stream is missing.")
        return problems
    try:
        wm = parse_projectwm(wm_raw, code_page=expected.baseline.code_page)
    except (OSError, UnicodeError, ValueError, TypeError) as exc:
        problems.append(f"Candidate PROJECTwm stream could not be parsed ({type(exc).__name__}).")
        return problems
    wm_names = {item[0].casefold() for item in wm} if wm else set()
    for name in sorted(deleted_names):
        if name in wm_names:
            problems.append(f"Deleted module still listed in PROJECTwm: {name!r}")
    return problems


def _make_unique_temp_name(project: Any) -> str:
    """ASCII temp component name guaranteed unique and short (plan 14 Stage H)."""
    import uuid

    existing = {m.name.casefold() for m in project.modules}
    for _ in range(100):
        candidate = f"VBAAE_TMP_{uuid.uuid4().hex[:6].upper()}"
        if candidate.casefold() not in existing:
            return candidate
    raise AdapterError("Could not generate a temporary module name for rename.")
