"""OOXML package adapter: XML part discovery, signature detection, ZIP rewrite.

Owns all zipfile/XML byte manipulation for editable OOXML packages
(plan sections 5.1, 9, 13). Must not understand VBA module semantics; GUI and
services never touch ZIP entries directly.
"""

from __future__ import annotations

import hashlib
import io
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


PackageSource = Path | bytes


def open_package(source: PackageSource) -> zipfile.ZipFile:
    """Open a package from a path or from bytes already in memory."""
    if isinstance(source, (bytes, bytearray, memoryview)):
        return zipfile.ZipFile(io.BytesIO(source))
    return zipfile.ZipFile(Path(source))


MAX_ADDED_MEMBER_BYTES = 1024 * 1024
_ADDED_MEMBER_DATE = (1980, 1, 1, 0, 0, 0)  # Office writes this date on every entry


def is_safe_member_name(name: str) -> bool:
    """A ZIP member name this editor may create: relative, '/'-separated, no '..'."""
    if not name or name.startswith("/") or "\\" in name or name.endswith("/"):
        return False
    return all(seg and seg not in (".", "..") and ":" not in seg for seg in name.split("/"))


# The only kind of existing member a draft may drop: an image (an unused ribbon icon).
REMOVABLE_MEMBER_EXTENSIONS = frozenset(
    {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".ico", ".tif", ".tiff", ".emf", ".wmf"}
)


def is_removable_member(name: str) -> bool:
    return is_safe_member_name(name) and Path(name).suffix.lower() in REMOVABLE_MEMBER_EXTENSIONS


def package_member_names(source: PackageSource) -> list[str]:
    with open_package(source) as zf:
        return zf.namelist()


def read_member_from_bytes(package: bytes, name: str, limit: int = MAX_ADDED_MEMBER_BYTES) -> bytes | None:
    """One member of a package held in memory (never a file), or None if absent or too large."""
    try:
        with open_package(package) as zf:
            info = zf.getinfo(name)
            return zf.read(info) if info.file_size <= limit else None
    except (KeyError, OSError, zipfile.BadZipFile):
        return None


def _check_duplicates(zf: zipfile.ZipFile, path: object = None) -> None:
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

    def snapshot_xml_parts(self, source: PackageSource) -> tuple[XmlPartSnapshot, ...]:
        parts: list[XmlPartSnapshot] = []
        total = 0
        with open_package(source) as zf:
            _check_duplicates(zf)
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

    def inspect_package_signature(self, source: PackageSource) -> PackageSafetyInfo:
        """Detect OPC/package-level digital signatures (plan 9.5)."""
        names: list[str] = []
        with open_package(source) as zf:
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

    def build_package(
        self,
        source: PackageSource,
        replacements: dict[str, bytes | None],
        additions: dict[str, bytes] | None = None,
    ) -> bytes:
        """The package with the given members replaced (None drops one), in memory.

        ``additions`` are new members, appended after the existing entries;
        none may exist yet. Every other entry keeps its bytes and ZIP metadata.
        """
        additions = dict(additions or {})
        if not replacements and not additions:
            raise PackageError("No package members to replace.")
        out = io.BytesIO()
        with open_package(source) as src:
            _check_duplicates(src)
            names = src.namelist()
            missing = sorted(set(replacements) - set(names))
            if missing:
                raise PackageError(f"Package members missing from the file: {missing}")
            lowered = {n.lower() for n in names}
            clashing = sorted(n for n in additions if n.lower() in lowered)
            if clashing:
                raise PackageError(f"New package members already exist in the file: {clashing}")
            unsafe = sorted(n for n in additions if not is_safe_member_name(n))
            if unsafe:
                raise PackageError(f"New package member names are not allowed: {unsafe}")
            with zipfile.ZipFile(out, "w") as dst:
                if src.comment:
                    dst.comment = src.comment
                for info in src.infolist():
                    if info.filename in replacements:
                        data = replacements[info.filename]
                        if data is None:
                            continue
                    else:
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
                for name in sorted(additions, key=str.lower):
                    new_info = zipfile.ZipInfo(name, date_time=_ADDED_MEMBER_DATE)
                    new_info.compress_type = zipfile.ZIP_DEFLATED
                    new_info.create_system = 0  # MS-DOS, as Office writes it
                    dst.writestr(new_info, additions[name])
        built = out.getvalue()
        with open_package(built) as check:
            bad = check.testzip()
            if bad is not None:
                raise PackageError(
                    f"The edited package failed its integrity check at entry {bad!r}."
                )
        return built

    def replace_members(self, path: Path, replacements: dict[str, bytes]) -> None:
        """Rewrite selected ZIP members in place, preserving other entries and metadata."""
        path = Path(path)
        built = self.build_package(path, dict(replacements))
        temp = path.with_name(path.stem + ".vbaae-member" + path.suffix)
        try:
            temp.write_bytes(built)
            temp.replace(path)
        except Exception:
            if temp.exists():
                try:
                    temp.unlink()
                except OSError:
                    pass
            raise

    # -- candidate rewrite (plan 9.6, 12.8) -----------------------------------

    def xml_replacements(self, changed_parts: list) -> dict[str, bytes]:
        """Encoded bytes of every changed XML part, validated as they will be written.

        New parts (``is_new``) are additions, not replacements: see ``added_members``.
        """
        return self._encoded_parts([p for p in changed_parts if not getattr(p, "is_new", False)])

    def added_members(self, draft: DocumentDraft) -> dict[str, bytes]:
        """Bytes of every member the draft adds: new XML parts (validated) and images."""
        added = self._encoded_parts(draft.new_xml_parts())
        for binary in draft.binary_parts:
            if binary.path in added:
                raise PackageError(f"Package member {binary.path!r} is added twice.")
            if len(binary.data) > MAX_ADDED_MEMBER_BYTES:
                raise PackageError(
                    f"{binary.path} is {len(binary.data):,} bytes; an added image may be "
                    f"at most {MAX_ADDED_MEMBER_BYTES:,} bytes."
                )
            added[binary.path] = binary.data
        unsafe = sorted(name for name in added if not is_safe_member_name(name))
        if unsafe:
            raise PackageError(f"New package member names are not allowed: {unsafe}")
        return added

    def removed_members(self, draft: DocumentDraft) -> dict[str, None]:
        """Existing members the draft drops, as ``build_package`` replacements (None)."""
        removed = draft.removed_member_paths()
        refused = sorted(name for name in removed if not is_removable_member(name))
        if refused:
            raise PackageError(f"Only unused images may be removed from the package: {refused}")
        clash = sorted(set(removed) & {p.path for p in draft.changed_xml_parts()})
        if clash:
            raise PackageError(f"Package members are both edited and removed: {clash}")
        return {name: None for name in removed}

    def _encoded_parts(self, parts: list) -> dict[str, bytes]:
        replacements: dict[str, bytes] = {}
        # Validate all changed parts first (plan 12.7): no candidate on failure.
        for part in parts:
            raw = encode_xml(
                part.text, encoding=part.encoding, bom=part.bom, newline=part.newline
            )
            problems = validate_xml_bytes(raw, part_path=part.path)
            if problems:
                raise PackageError(
                    "The edited XML could not be validated: " + "; ".join(problems)
                )
            replacements[part.path] = raw
        return replacements

    def write_xml_candidate(
        self,
        source_path: Path,
        dest_path: Path,
        changed_parts: list,
    ) -> None:
        """Rebuild the package into dest_path with exactly the given XML parts replaced."""
        if not changed_parts:
            raise PackageError("No XML changes to write.")
        replacements = self.xml_replacements(changed_parts)
        with open_package(source_path) as src:
            missing = sorted(set(replacements) - set(src.namelist()))
        if missing:
            raise PackageError(f"Changed XML parts missing from the package: {missing}")
        Path(dest_path).write_bytes(self.build_package(source_path, dict(replacements)))

    # -- candidate verification (plan 13) --------------------------------------

    def verify_candidate_xml(
        self,
        reference_path: PackageSource,
        candidate_path: PackageSource,
        draft: DocumentDraft,
        signature_edits: dict[str, bytes | None] | None = None,
    ) -> CandidateVerificationResult:
        """``signature_edits``: package edits of an approved VBA signature drop
        (None = part removed); each must match exactly and nothing else may differ."""
        problems: list[str] = []
        signature_edits = signature_edits or {}
        removed = {path for path, data in signature_edits.items() if data is None}
        try:
            added = self.added_members(draft)
            dropped = set(self.removed_members(draft))
        except PackageError as exc:
            return CandidateVerificationResult(False, (str(exc),))
        try:
            with open_package(reference_path) as ref, open_package(candidate_path) as cand:
                _check_duplicates(ref)
                _check_duplicates(cand)
                if cand.testzip() is not None:
                    problems.append("Candidate ZIP has a corrupt entry.")
                all_ref_names = set(ref.namelist())
                missing = sorted(dropped - all_ref_names)
                if missing:
                    problems.append(f"Members to remove are not in the file: {missing}")
                    return CandidateVerificationResult(False, tuple(problems))
                ref_names = all_ref_names - removed - dropped
                cand_names = set(cand.namelist())
                expected_names = ref_names | set(added)
                if expected_names != cand_names or ref_names & set(added):
                    problems.append(
                        f"ZIP entry set changed: only-in-reference={sorted(expected_names - cand_names)} "
                        f"only-in-candidate={sorted(cand_names - expected_names)}"
                    )
                    return CandidateVerificationResult(False, tuple(problems))

                # Each added member holds exactly the bytes the draft adds.
                for path, data in sorted(added.items()):
                    actual = cand.read(path)
                    if actual != data:
                        problems.append(f"Added package member {path!r} does not match the draft.")
                    elif _is_editable_xml_part(path):
                        problems.extend(validate_xml_bytes(actual, part_path=path))

                changed = {p.path for p in draft.modified_xml_parts()}
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
                for path, data in signature_edits.items():
                    if data is not None and path in cand_names and cand.read(path) != data:
                        problems.append(f"Signature removal changed {path!r} unexpectedly.")
                for path in sorted(ref_names - changed - set(signature_edits)):
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
