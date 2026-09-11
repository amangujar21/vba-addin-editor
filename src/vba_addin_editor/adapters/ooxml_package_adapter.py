"""OOXML package adapter: XML part discovery, signature detection, ZIP rewrite.

Owns all zipfile/XML byte manipulation for editable OOXML packages
(plan sections 5.1, 9, 13). Must not understand VBA module semantics; GUI and
services never touch ZIP entries directly.
"""

from __future__ import annotations

import hashlib
import zipfile
from pathlib import Path

from vba_addin_editor.adapters.xml_codec import (
    XmlCodecError,
    decode_xml,
    encode_xml,
    validate_xml_bytes,
)
from vba_addin_editor.domain.changes import compute_changes
from vba_addin_editor.domain.document import (
    DocumentDraft,
    PackageSafetyInfo,
    XmlPartDraft,
    XmlPartSnapshot,
)
from vba_addin_editor.domain.results import CandidateVerificationResult

XML_EDITABLE_EXTENSIONS = {".xlam", ".ppam", ".pptm"}
XML_PART_SUFFIXES = {".xml", ".rels"}
CONTENT_TYPES_NAME = "[Content_Types].xml"

MAX_XML_PART_BYTES = 32 * 1024 * 1024
MAX_TOTAL_XML_BYTES = 256 * 1024 * 1024
MAX_XML_PART_COUNT = 10_000

_SIGNATURE_DIR_PREFIX = "_xmlsignatures/"
_SIGNATURE_REL_TYPE_SUFFIX = "digital-signature/origin"
_SIGNATURE_CONTENT_TYPE = "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"


class PackageError(Exception):
    """Plain-language package failure; .details carries technical info."""

    def __init__(self, message: str, details: dict | None = None) -> None:
        super().__init__(message)
        self.details = details or {}


def _is_editable_xml_part(name: str) -> bool:
    lowered = name.lower()
    if lowered == CONTENT_TYPES_NAME.lower():
        return True
    return lowered.endswith(tuple(XML_PART_SUFFIXES))


def _check_duplicates(zf: zipfile.ZipFile, path: Path) -> None:
    names = zf.namelist()
    if len(names) != len(set(names)):
        seen: set[str] = set()
        dups = sorted({n for n in names if n in seen or seen.add(n)})
        raise PackageError(
            f"This Office file contains duplicate package entries ({dups}); "
            "it cannot be edited safely.",
            {"duplicates": dups},
        )

class OoxmlPackageAdapter:
    """Standard-library-only OOXML package operations for supported Office files."""
    # -- discovery ---------------------------------------------------------

    def snapshot_xml_parts(self, path: Path) -> tuple[XmlPartSnapshot, ...]:
        path = Path(path)
        parts: list[XmlPartSnapshot] = []
        total = 0
        with zipfile.ZipFile(path) as zf:
            _check_duplicates(zf, path)
            entries = [i for i in zf.infolist() if not i.is_dir()]
            if len(entries) > MAX_XML_PART_COUNT:
                raise PackageError(
                    "This Office file contains too many package entries to edit safely."
                )
            for info in entries:
                if not _is_editable_xml_part(info.filename):
                    continue
                raw = zf.read(info.filename)
                total += len(raw)
                if len(raw) > MAX_XML_PART_BYTES or total > MAX_TOTAL_XML_BYTES:
                    raise PackageError(
                        "An XML part in this Office file exceeds the safe editing "
                        "size limit; XML editing is not available for it."
                    )
                try:
                    decoded = decode_xml(raw)
                except XmlCodecError as exc:
                    decoded = None
                    problems = (f"{info.filename}: {exc}",)
                else:
                    problems = validate_xml_bytes(raw, part_path=info.filename)

                # Existing OOXML packages can contain XML parts that our editor
                # cannot parse even though Office/other tooling tolerates them.
                # Baseline parse failure is per-part state, not an open failure;
                # untouched members are copied and verified byte-for-byte.
                lowered = info.filename.lower()
                parts.append(
                    XmlPartSnapshot(
                        path=info.filename,
                        text=decoded.text if decoded is not None else None,
                        encoding=decoded.encoding if decoded is not None else None,
                        bom=decoded.bom if decoded is not None else b"",
                        newline=decoded.newline if decoded is not None else "\n",
                        original_sha256=hashlib.sha256(raw).hexdigest(),
                        original_size=len(raw),
                        is_relationships_part=lowered.endswith(".rels"),
                        is_content_types_part=lowered == CONTENT_TYPES_NAME.lower(),
                        open_problem="; ".join(problems) if problems else None,
                    )
                )
        parts.sort(key=lambda p: p.path.lower())
        return tuple(parts)

    # -- signature inspection ------------------------------------------------

    def inspect_package_signature(self, path: Path) -> PackageSafetyInfo:
        """Detect OPC/package-level digital signatures (plan 9.5)."""
        path = Path(path)
        names: list[str] = []
        with zipfile.ZipFile(path) as zf:
            for info in zf.infolist():
                if info.is_dir():
                    continue
                name = info.filename
                lowered = name.lower()
                if lowered.startswith(_SIGNATURE_DIR_PREFIX):
                    names.append(name)
                    continue
                if name.lower().endswith(".rels"):
                    try:
                        text = zf.read(name).decode("utf-8", errors="replace")
                    except OSError:  # pragma: no cover - zf.read rarely raises OSError
                        continue
                    if _SIGNATURE_REL_TYPE_SUFFIX in text:
                        names.append(name)
                    continue
                if name.lower() == CONTENT_TYPES_NAME.lower():
                    text = zf.read(name).decode("utf-8", errors="replace")
                    if _SIGNATURE_CONTENT_TYPE in text:
                        names.append(name)
        return PackageSafetyInfo(
            opc_signature_present=bool(names),
            signature_part_names=tuple(names),
        )

    # -- validation ----------------------------------------------------------

    def validate_draft_part(self, part: XmlPartDraft) -> tuple[str, ...]:
        """Validate a changed XML draft part the way it would be written."""
        from vba_addin_editor.adapters.xml_codec import validate_xml_draft

        return validate_xml_draft(part)

    def replace_members(self, path: Path, replacements: dict[str, bytes]) -> None:
        """Rewrite selected ZIP members in place, preserving other entries and metadata."""
        path = Path(path)
        if not replacements:
            raise PackageError("No package members to replace.")
        temp = path.with_name(path.stem + ".vbaae-member" + path.suffix)
        try:
            with zipfile.ZipFile(path) as src:
                _check_duplicates(src, path)
                missing = sorted(set(replacements) - set(src.namelist()))
                if missing:
                    raise PackageError(f"Package members missing from the file: {missing}")
                with zipfile.ZipFile(temp, "w") as dst:
                    if src.comment:
                        dst.comment = src.comment
                    for info in src.infolist():
                        data = replacements.get(info.filename)
                        if data is None:
                            data = src.read(info.filename)
                        new_info = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                        new_info.compress_type = info.compress_type
                        new_info.comment = info.comment
                        new_info.extra = info.extra
                        new_info.internal_attr = info.internal_attr
                        new_info.external_attr = info.external_attr
                        new_info.create_system = info.create_system
                        new_info.create_version = info.create_version
                        dst.writestr(new_info, data)
            with zipfile.ZipFile(temp) as check:
                bad = check.testzip()
                if bad is not None:
                    raise PackageError(
                        f"The edited package failed its integrity check at entry {bad!r}."
                    )
            temp.replace(path)
        except Exception:
            if temp.exists():
                try:
                    temp.unlink()
                except OSError:
                    pass
            raise

    # -- candidate rewrite (plan 9.6, 12.8) -----------------------------------

    def write_xml_candidate(
        self,
        source_path: Path,
        dest_path: Path,
        changed_parts: list,
    ) -> None:
        """Rebuild the package into dest_path with exactly the given XML parts replaced."""
        source_path = Path(source_path)
        dest_path = Path(dest_path)
        changes = {p.path: p for p in changed_parts}
        if not changes:
            raise PackageError("No XML changes to write.")

        # Validate all changed parts first (plan 12.7): no candidate on failure.
        for part in changed_parts:
            raw = encode_xml(
                part.text, encoding=part.encoding, bom=part.bom, newline=part.newline
            )
            problems = validate_xml_bytes(raw, part_path=part.path)
            if problems:
                raise PackageError(
                    "The edited XML could not be validated: " + "; ".join(problems)
                )

        replacement_bytes: dict[str, bytes] = {}
        with zipfile.ZipFile(source_path) as src:
            _check_duplicates(src, source_path)
            src_names = set(src.namelist())
            missing = sorted(set(changes) - src_names)
            if missing:
                raise PackageError(
                    f"Changed XML parts missing from the package: {missing}"
                )
            for name, part in changes.items():
                replacement_bytes[name] = encode_xml(
                    part.text, encoding=part.encoding, bom=part.bom, newline=part.newline
                )
            with zipfile.ZipFile(dest_path, "w") as dst:
                for info in src.infolist():
                    data = replacement_bytes.get(info.filename)
                    if data is None:
                        data = src.read(info.filename)
                    new_info = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                    new_info.compress_type = info.compress_type
                    new_info.comment = info.comment
                    new_info.extra = info.extra
                    new_info.internal_attr = info.internal_attr
                    new_info.external_attr = info.external_attr
                    new_info.create_system = info.create_system
                    new_info.create_version = info.create_version
                    dst.writestr(new_info, data)
                if src.comment:
                    dst.comment = src.comment
        with zipfile.ZipFile(dest_path) as check:
            bad = check.testzip()
            if bad is not None:
                raise PackageError(
                    f"The edited package failed its integrity check at entry {bad!r}."
                )

    # -- candidate verification (plan 13) --------------------------------------

    def verify_candidate_xml(
        self,
        reference_path: Path,
        candidate_path: Path,
        draft: DocumentDraft,
    ) -> CandidateVerificationResult:
        problems: list[str] = []
        try:
            with zipfile.ZipFile(reference_path) as ref, zipfile.ZipFile(
                candidate_path
            ) as cand:
                _check_duplicates(ref, Path(reference_path))
                _check_duplicates(cand, Path(candidate_path))
                if cand.testzip() is not None:
                    problems.append("Candidate ZIP has a corrupt entry.")
                ref_names = set(ref.namelist())
                cand_names = set(cand.namelist())
                if ref_names != cand_names:
                    problems.append(
                        f"ZIP entry set changed: only-in-reference={sorted(ref_names - cand_names)} "
                        f"only-in-candidate={sorted(cand_names - ref_names)}"
                    )
                    return CandidateVerificationResult(False, tuple(problems))

                changed = {p.path for p in draft.changed_xml_parts()}
                for path in sorted(changed):
                    if path not in cand_names:
                        problems.append(f"Changed XML part missing from candidate: {path!r}")
                        continue
                    part = draft.xml_part_by_path(path)
                    if part is None or part.text is None or part.encoding is None:
                        problems.append(
                            f"Changed XML part {path!r} has no editable text metadata."
                        )
                        continue
                    expected = encode_xml(
                        part.text, encoding=part.encoding, bom=part.bom, newline=part.newline
                    )
                    actual = cand.read(path)
                    if actual != expected:
                        problems.append(
                            f"Changed XML part {path!r} does not match the edited draft bytes."
                        )
                    problems.extend(validate_xml_bytes(actual, part_path=path))

                # Do not parse untouched baseline XML here. A malformed
                # pre-existing part is allowed, but it must remain byte-identical.
                # Changed XML is validated separately above.
                for path in sorted(ref_names - changed):
                    if _is_editable_xml_part(path) and ref.read(path) != cand.read(path):
                        problems.append(
                            f"XML part {path!r} changed without being part of the edit set."
                        )
                if not compute_changes(draft).has_vba_changes:
                    vba_entries = [n for n in ref_names if n.lower().endswith("vbaproject.bin")]
                    for entry in vba_entries:
                        if ref.read(entry) != cand.read(entry):
                            problems.append(
                                f"VBA project payload changed on an XML-only save: {entry!r}"
                            )
        except (zipfile.BadZipFile, OSError) as exc:
            problems.append(f"Package comparison failed: {exc!r}")
        return CandidateVerificationResult(not problems, tuple(problems))
