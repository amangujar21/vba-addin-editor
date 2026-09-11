"""Component capabilities derived from trusted PROJECT/dir metadata.

JSON flags and draft booleans never authorize destructive operations; callers
must resolve original modules against a trusted snapshot.
"""

from __future__ import annotations

from dataclasses import dataclass

from vba_addin_editor.domain.document import ModuleDisplayKind

# Stable restriction tokens. UI maps these to explanations.
READ_ONLY = "read_only"
DOCUMENT = "document"
DESIGNER = "designer"
UNKNOWN = "unknown"
CONFLICTING_METADATA = "conflicting_metadata"
DECODING_LOSS = "decoding_loss"
MISSING_PROJECT = "missing_project"
PASSWORD_PROTECTED = "password_protected"

RESTRICTION_MESSAGES = {
    READ_ONLY: "This module is marked read-only, so renaming or deleting it is disabled.",
    DOCUMENT: (
        "This is a host document module (workbook, sheet, or presentation "
        "code-behind). Renaming or deleting it is disabled."
    ),
    DESIGNER: (
        "This is a designer component (for example a UserForm). Renaming or "
        "deleting it is disabled."
    ),
    UNKNOWN: (
        "This component's type could not be verified from project metadata, "
        "so renaming or deleting it is disabled."
    ),
    CONFLICTING_METADATA: (
        "Project metadata for this component is duplicated or conflicting, "
        "so renaming or deleting it is disabled."
    ),
    DECODING_LOSS: (
        "Project metadata could not be decoded reliably, so renaming or "
        "deleting modules is disabled."
    ),
    MISSING_PROJECT: (
        "The VBA PROJECT stream is missing or unreadable, so renaming or "
        "deleting modules is disabled."
    ),
    PASSWORD_PROTECTED: "This VBA project is password-protected and cannot be modified safely.",
}


@dataclass(frozen=True)
class ProjectTypeIndex:
    """Case-folded logical names declared in the PROJECT stream."""

    standard: frozenset[str]
    classes: frozenset[str]
    documents: frozenset[str]
    base_classes: frozenset[str]
    duplicates: frozenset[str]
    lossy: bool
    missing: bool


@dataclass(frozen=True)
class ComponentCapabilities:
    display_kind: str
    project_item_kind: str
    can_delete: bool
    can_rename: bool
    restriction_reason: str | None

    @property
    def destructive_ops_safe(self) -> bool:
        return self.can_delete and self.can_rename

    @property
    def restriction_message(self) -> str | None:
        if self.restriction_reason is None:
            return None
        return RESTRICTION_MESSAGES.get(self.restriction_reason, self.restriction_reason)


def empty_project_index(*, missing: bool = False, lossy: bool = False) -> ProjectTypeIndex:
    return ProjectTypeIndex(
        standard=frozenset(),
        classes=frozenset(),
        documents=frozenset(),
        base_classes=frozenset(),
        duplicates=frozenset(),
        lossy=lossy,
        missing=missing,
    )


def build_project_index(
    *,
    standard_modules: list[str],
    class_modules: list[str],
    document_modules: list[str],
    base_classes: list[str],
    lossy: bool,
    missing: bool = False,
) -> ProjectTypeIndex:
    bags = {
        "standard": list(standard_modules),
        "classes": list(class_modules),
        "documents": list(document_modules),
        "base_classes": list(base_classes),
    }
    folded: dict[str, set[str]] = {key: set() for key in bags}
    duplicates: set[str] = set()
    seen_any: set[str] = set()
    for key, names in bags.items():
        for name in names:
            item = name.casefold()
            if item in folded[key] or item in seen_any:
                duplicates.add(item)
            folded[key].add(item)
            seen_any.add(item)
    return ProjectTypeIndex(
        standard=frozenset(folded["standard"]),
        classes=frozenset(folded["classes"]),
        documents=frozenset(folded["documents"]),
        base_classes=frozenset(folded["base_classes"]),
        duplicates=frozenset(duplicates),
        lossy=lossy,
        missing=missing,
    )


def classify_component(
    *,
    logical_name: str,
    stream_name: str,
    dir_kind: str,
    is_read_only: bool,
    index: ProjectTypeIndex,
    designer_storages: frozenset[str],
) -> ComponentCapabilities:
    """Classify one dir-stream module using PROJECT declarations.

    Never infers subtype from the logical name (Class1, ThisWorkbook, …)
    or from VB_PredeclaredId. ``dir_kind`` is the dir-stream kind name
    (``standard`` or ``other``).
    """
    if index.missing:
        return _restricted(ModuleDisplayKind.AMBIGUOUS, "unknown", MISSING_PROJECT)
    if index.lossy:
        display = ModuleDisplayKind.STANDARD if dir_kind == "standard" else ModuleDisplayKind.AMBIGUOUS
        kind = "standard" if dir_kind == "standard" else "unknown"
        return _restricted(display, kind, DECODING_LOSS)

    folded = logical_name.casefold()
    stream_folded = stream_name.casefold()
    in_std = folded in index.standard
    in_class = folded in index.classes
    in_doc = folded in index.documents
    in_base = folded in index.base_classes
    in_designer = stream_folded in designer_storages or folded in designer_storages
    category_hits = sum((in_std, in_class, in_doc, in_base))

    if folded in index.duplicates or category_hits > 1:
        return _restricted(ModuleDisplayKind.AMBIGUOUS, "unknown", CONFLICTING_METADATA)

    if in_doc:
        return _restricted(ModuleDisplayKind.AMBIGUOUS, "document", DOCUMENT)
    if in_base or in_designer:
        return _restricted(ModuleDisplayKind.AMBIGUOUS, "designer", DESIGNER)

    if in_class and dir_kind == "other":
        if is_read_only:
            return _restricted(ModuleDisplayKind.CLASS, "class", READ_ONLY)
        return ComponentCapabilities(
            display_kind=ModuleDisplayKind.CLASS,
            project_item_kind="class",
            can_delete=True,
            can_rename=True,
            restriction_reason=None,
        )

    if in_std and dir_kind == "standard":
        if is_read_only:
            return _restricted(ModuleDisplayKind.STANDARD, "standard", READ_ONLY)
        return ComponentCapabilities(
            display_kind=ModuleDisplayKind.STANDARD,
            project_item_kind="standard",
            can_delete=True,
            can_rename=True,
            restriction_reason=None,
        )

    # PROJECT/dir mismatch or undeclared other component.
    display = ModuleDisplayKind.STANDARD if dir_kind == "standard" else ModuleDisplayKind.AMBIGUOUS
    kind = "standard" if dir_kind == "standard" else "unknown"
    return _restricted(display, kind, UNKNOWN)


def _restricted(display_kind: str, project_item_kind: str, reason: str) -> ComponentCapabilities:
    return ComponentCapabilities(
        display_kind=display_kind,
        project_item_kind=project_item_kind,
        can_delete=False,
        can_rename=False,
        restriction_reason=reason,
    )


def capabilities_from_snapshot_module(module) -> ComponentCapabilities:
    return ComponentCapabilities(
        display_kind=module.kind,
        project_item_kind=getattr(module, "project_item_kind", "unknown"),
        can_delete=getattr(module, "can_delete", False),
        can_rename=getattr(module, "can_rename", False),
        restriction_reason=getattr(module, "restriction_reason", None),
    )
