"""Core domain models: snapshot (immutable open-time state) and draft (editable)."""

from __future__ import annotations

import uuid
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Literal

HostKind = Literal["excel", "powerpoint"]
Extension = Literal[".xlam", ".ppam", ".pptm"]


class ModuleDisplayKind:
    """UI-facing module classification (plan section 3.6)."""

    STANDARD = "standard"
    CLASS = "class"
    AMBIGUOUS = "ambiguous"  # existing kind=other with unproven subtype


@dataclass(frozen=True)
class FileFingerprint:
    """SHA-256 is the authority; size/mtime are fast pre-checks."""

    sha256: str
    size: int
    mtime_ns: int


@dataclass(frozen=True)
class ProjectSafetyInfo:
    password_protected: bool
    signature_present: bool
    signature_kinds: tuple[str, ...]
    writable: bool
    host_process_running: bool


@dataclass(frozen=True)
class PackageSafetyInfo:
    """OPC/package-level signature state (separate from the VBA signature)."""

    opc_signature_present: bool = False
    signature_part_names: tuple[str, ...] = ()


@dataclass(frozen=True)
class ModuleSnapshot:
    id: str
    original_name: str
    kind: str  # ModuleDisplayKind value
    pyopenvba_kind: str  # "standard" | "other"
    full_source: str
    hidden_header: str
    body: str
    is_read_only: bool
    is_private: bool
    destructive_ops_safe: bool
    ends_with_newline: bool


@dataclass(frozen=True)
class XmlPartSnapshot:
    """Immutable open-time state of one editable XML package part."""

    path: str  # exact ZIP member name, no leading "/"
    text: str  # LF-normalized for the editor
    encoding: str  # canonical codec name for the payload (without BOM)
    bom: bytes
    newline: str  # dominant original newline: "\n" or "\r\n"
    original_sha256: str
    original_size: int
    is_relationships_part: bool
    is_content_types_part: bool


@dataclass(frozen=True)
class DocumentSnapshot:
    path: Path
    host_kind: HostKind
    extension: Extension
    file_fingerprint: FileFingerprint
    project_name: str
    code_page: int
    safety: ProjectSafetyInfo
    modules: tuple[ModuleSnapshot, ...]
    xml_parts: tuple[XmlPartSnapshot, ...] = ()
    package_safety: PackageSafetyInfo = PackageSafetyInfo()


@dataclass
class ModuleDraft:
    id: str
    origin_name: str | None  # None => newly created
    current_name: str
    body: str
    kind: str
    pyopenvba_kind: str  # "standard" | "other"
    is_new: bool
    is_deleted: bool
    destructive_ops_safe: bool
    original_body: str | None = None  # for revert of existing modules


@dataclass
class XmlPartDraft:
    """Editable state of one XML package part."""

    path: str
    text: str
    original_text: str
    encoding: str
    bom: bytes
    newline: str
    is_relationships_part: bool
    is_content_types_part: bool

    def is_dirty(self) -> bool:
        return self.text != self.original_text


@dataclass
class DocumentDraft:
    baseline: DocumentSnapshot
    modules: list[ModuleDraft] = field(default_factory=list)
    xml_parts: list[XmlPartDraft] = field(default_factory=list)
    signed_save_confirmed: bool = False

    # -- queries -------------------------------------------------------

    def module_by_id(self, module_id: str) -> ModuleDraft | None:
        return next((m for m in self.modules if m.id == module_id), None)

    def find_current(self, name: str) -> ModuleDraft | None:
        needle = name.casefold()
        return next(
            (m for m in self.modules if not m.is_deleted and m.current_name.casefold() == needle),
            None,
        )

    def deleted_original_modules(self) -> list[ModuleDraft]:
        return [m for m in self.modules if m.is_deleted and not m.is_new]

    def surviving_originals(self) -> list[ModuleDraft]:
        return [m for m in self.modules if not m.is_deleted and not m.is_new]

    def changed_names(self) -> list[ModuleDraft]:
        return [
            m
            for m in self.surviving_originals()
            if m.origin_name is not None and m.current_name.casefold() != m.origin_name.casefold()
        ]

    def new_modules(self) -> list[ModuleDraft]:
        return [m for m in self.modules if m.is_new and not m.is_deleted]

    def changed_existing_modules(self) -> list[ModuleDraft]:
        return [m for m in self.surviving_originals() if m.body != m.original_body]

    def xml_part_by_path(self, path: str) -> XmlPartDraft | None:
        return next((p for p in self.xml_parts if p.path == path), None)

    def changed_xml_parts(self) -> list[XmlPartDraft]:
        return [p for p in self.xml_parts if p.is_dirty()]

    def final_module_state(self) -> list[ModuleDraft]:
        return [
            m for m in self.modules if not m.is_deleted
        ]

    def is_dirty(self) -> bool:
        return bool(
            self.deleted_original_modules()
            or self.changed_names()
            or self.new_modules()
            or self.changed_existing_modules()
            or self.changed_xml_parts()
        )


def draft_from_snapshot(snapshot: DocumentSnapshot) -> DocumentDraft:
    modules = [
        ModuleDraft(
            id=m.id,
            origin_name=m.original_name,
            current_name=m.original_name,
            body=m.body,
            kind=m.kind,
            pyopenvba_kind=m.pyopenvba_kind,
            is_new=False,
            is_deleted=False,
            destructive_ops_safe=m.destructive_ops_safe,
            original_body=m.body,
        )
        for m in snapshot.modules
    ]
    xml_parts = [
        XmlPartDraft(
            path=p.path,
            text=p.text,
            original_text=p.text,
            encoding=p.encoding,
            bom=p.bom,
            newline=p.newline,
            is_relationships_part=p.is_relationships_part,
            is_content_types_part=p.is_content_types_part,
        )
        for p in snapshot.xml_parts
    ]
    return DocumentDraft(baseline=snapshot, modules=modules, xml_parts=xml_parts)


def new_module_id() -> str:
    return uuid.uuid4().hex[:12]


def revert_all(draft: DocumentDraft) -> DocumentDraft:
    return draft_from_snapshot(draft.baseline)


def snapshot_with_fingerprint(snapshot: DocumentSnapshot, fp: FileFingerprint) -> DocumentSnapshot:
    return replace(snapshot, file_fingerprint=fp)
