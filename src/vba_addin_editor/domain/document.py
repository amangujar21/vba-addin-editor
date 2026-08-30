"""Core domain models: snapshot (immutable open-time state) and draft (editable)."""

from __future__ import annotations

import uuid
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Literal

HostKind = Literal["excel", "powerpoint"]
Extension = Literal[".xlam", ".ppam"]


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
class DocumentSnapshot:
    path: Path
    host_kind: HostKind
    extension: Extension
    file_fingerprint: FileFingerprint
    project_name: str
    code_page: int
    safety: ProjectSafetyInfo
    modules: tuple[ModuleSnapshot, ...]


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
class DocumentDraft:
    baseline: DocumentSnapshot
    modules: list[ModuleDraft] = field(default_factory=list)
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
    return DocumentDraft(baseline=snapshot, modules=modules)


def new_module_id() -> str:
    return uuid.uuid4().hex[:12]


def revert_all(draft: DocumentDraft) -> DocumentDraft:
    return draft_from_snapshot(draft.baseline)


def snapshot_with_fingerprint(snapshot: DocumentSnapshot, fp: FileFingerprint) -> DocumentSnapshot:
    return replace(snapshot, file_fingerprint=fp)
