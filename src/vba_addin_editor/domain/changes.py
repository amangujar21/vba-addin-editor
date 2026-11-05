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
class XmlPartChange:
    path: str


@dataclass(frozen=True)
class ChangeSet:
    modified: tuple[ModuleChange, ...] = ()
    added: tuple[ModuleChange, ...] = ()
    renamed: tuple[ModuleRename, ...] = ()
    deleted: tuple[ModuleChange, ...] = ()
    xml_modified: tuple[XmlPartChange, ...] = ()
    # New package members (XML parts and images), not in the file yet.
    parts_added: tuple[XmlPartChange, ...] = ()

    @property
    def is_empty(self) -> bool:
        return not (
            self.modified
            or self.added
            or self.renamed
            or self.deleted
            or self.xml_modified
            or self.parts_added
        )

    @property
    def has_vba_changes(self) -> bool:
        return bool(self.modified or self.added or self.renamed or self.deleted)

    @property
    def has_xml_changes(self) -> bool:
        """Any package-level (non-VBA) change: edited or added members."""
        return bool(self.xml_modified or self.parts_added)

    def summary_lines(self) -> list[str]:
        lines: list[str] = []
        if self.modified:
            lines.append("VBA — Modified")
            lines.extend(f"  • {c.name}" for c in self.modified)
        if self.added:
            lines.append("VBA — Added")
            lines.extend(f"  • {c.name}" for c in self.added)
        if self.renamed:
            lines.append("VBA — Renamed")
            lines.extend(f"  • {r.old_name} → {r.new_name}" for r in self.renamed)
        if self.deleted:
            lines.append("VBA — Deleted")
            lines.extend(f"  • {c.name}" for c in self.deleted)
        if self.xml_modified:
            lines.append("XML — Modified")
            lines.extend(f"  • {c.path}" for c in self.xml_modified)
        if self.parts_added:
            lines.append("Package — Added")
            lines.extend(f"  • {c.path}" for c in self.parts_added)
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
    xml_modified = tuple(
        XmlPartChange(path=p.path) for p in draft.modified_xml_parts()
    )
    parts_added = tuple(XmlPartChange(path=path) for path in draft.added_member_paths())
    return ChangeSet(
        modified=modified, added=added, renamed=renamed, deleted=deleted,
        xml_modified=xml_modified, parts_added=parts_added,
    )


def dirty_count(draft: DocumentDraft) -> int:
    cs = compute_changes(draft)
    return (
        len(cs.modified)
        + len(cs.added)
        + len(cs.renamed)
        + len(cs.deleted)
        + len(cs.xml_modified)
        + len(cs.parts_added)
    )


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


def revert_xml_part(draft: DocumentDraft, path: str) -> None:
    """Revert one XML part in place to its snapshot text."""
    part = draft.xml_part_by_path(path)
    if part is None:
        return
    if part.is_new:
        draft.xml_parts.remove(part)
        return
    part.text = part.original_text
