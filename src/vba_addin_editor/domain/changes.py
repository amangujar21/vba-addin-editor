"""Change set: single source of truth for dirty state, review UI, replay, tests."""

from __future__ import annotations

from dataclasses import dataclass

from vba_addin_editor.domain.document import DocumentDraft


@dataclass(frozen=True)
class ModuleChange:
    name: str
    module_id: str


@dataclass(frozen=True)
class ModuleRename:
    old_name: str
    new_name: str
    module_id: str


@dataclass(frozen=True)
class ChangeSet:
    modified: tuple[ModuleChange, ...] = ()
    added: tuple[ModuleChange, ...] = ()
    renamed: tuple[ModuleRename, ...] = ()
    deleted: tuple[ModuleChange, ...] = ()

    @property
    def is_empty(self) -> bool:
        return not (self.modified or self.added or self.renamed or self.deleted)

    def summary_lines(self) -> list[str]:
        lines: list[str] = []
        if self.modified:
            lines.append("Modified")
            lines.extend(f"  • {c.name}" for c in self.modified)
        if self.added:
            lines.append("Added")
            lines.extend(f"  • {c.name}" for c in self.added)
        if self.renamed:
            lines.append("Renamed")
            lines.extend(f"  • {r.old_name} → {r.new_name}" for r in self.renamed)
        if self.deleted:
            lines.append("Deleted")
            lines.extend(f"  • {c.name}" for c in self.deleted)
        return lines


def compute_changes(draft: DocumentDraft) -> ChangeSet:
    modified = tuple(
        ModuleChange(name=m.current_name, module_id=m.id)
        for m in draft.changed_existing_modules()
    )
    added = tuple(
        ModuleChange(name=m.current_name, module_id=m.id) for m in draft.new_modules()
    )
    renamed = tuple(
        ModuleRename(old_name=m.origin_name or "", new_name=m.current_name, module_id=m.id)
        for m in draft.changed_names()
    )
    deleted = tuple(
        ModuleChange(name=m.origin_name or m.current_name, module_id=m.id)
        for m in draft.deleted_original_modules()
    )
    return ChangeSet(modified=modified, added=added, renamed=renamed, deleted=deleted)


def dirty_count(draft: DocumentDraft) -> int:
    cs = compute_changes(draft)
    return len(cs.modified) + len(cs.added) + len(cs.renamed) + len(cs.deleted)


def revert_module(draft: DocumentDraft, module_id: str) -> None:
    """Revert one module in place: restore original state or remove new/delete marks."""
    baseline = {m.id: m for m in draft.baseline.modules}
    target = draft.module_by_id(module_id)
    if target is None:
        return
    if module_id in baseline:
        b = baseline[module_id]
        target.is_deleted = False
        target.is_new = False
        target.current_name = b.original_name
        target.body = b.body
    else:
        draft.modules.remove(target)
