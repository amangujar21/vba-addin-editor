"""Revision-bound pre-save change review (IMP-05)."""

from __future__ import annotations

import difflib
from dataclasses import dataclass

from vba_addin_editor.adapters.source_codec import compose_module_source
from vba_addin_editor.domain.changes import compute_changes
from vba_addin_editor.domain.document import DocumentDraft
from vba_addin_editor.domain.session import DocumentSession, ProposalBinding


@dataclass(frozen=True)
class ReviewItem:
    target_id: str
    category: str
    operations: tuple[str, ...]
    old_name: str | None
    new_name: str | None
    old_text: str | None
    new_text: str | None
    unified_diff: str


@dataclass(frozen=True)
class ReviewModel:
    binding: ProposalBinding
    destination: str
    file_type: str
    items: tuple[ReviewItem, ...]
    vba_count: int
    xml_count: int
    encoding_warnings: tuple[str, ...]
    signature_will_be_removed: bool
    backup_policy: str
    operation_type: str


def _diff(old: str | None, new: str | None, name: str) -> str:
    old_lines = (old or "").splitlines(keepends=True)
    new_lines = (new or "").splitlines(keepends=True)
    return "".join(
        difflib.unified_diff(old_lines, new_lines, fromfile=f"a/{name}", tofile=f"b/{name}", lineterm="")
    )


def _composed(draft: DocumentDraft, module_id: str, body: str, name: str) -> str:
    mod = draft.module_by_id(module_id)
    baseline = next((item for item in draft.baseline.modules if item.id == module_id), None)
    header = (mod.hidden_header if mod is not None else "") or (
        baseline.hidden_header if baseline is not None else ""
    )
    kind = mod.kind if mod is not None else "standard"
    py_kind = mod.pyopenvba_kind if mod is not None else "standard"
    return compose_module_source(
        body=body, hidden_header=header, current_name=name, kind=kind, pyopenvba_kind=py_kind
    )


class ReviewService:
    def build(
        self,
        session: DocumentSession,
        *,
        operation_type: str = "save",
        destination: str | None = None,
    ) -> ReviewModel:
        draft = session.draft
        changes = compute_changes(draft)
        grouped: dict[str, list[str]] = {}
        names: dict[str, tuple[str | None, str | None]] = {}
        for item in changes.modified:
            grouped.setdefault(item.module_id, []).append("edit")
            names[item.module_id] = (item.name, item.name)
        for item in changes.added:
            grouped.setdefault(item.module_id, []).append("add")
            names[item.module_id] = (None, item.name)
        for item in changes.renamed:
            grouped.setdefault(item.module_id, []).append("rename")
            names[item.module_id] = (item.old_name, item.new_name)
        for item in changes.deleted:
            grouped.setdefault(item.module_id, []).append("delete")
            names[item.module_id] = (item.name, None)
        review_items: list[ReviewItem] = []
        for module_id, ops in grouped.items():
            mod = draft.module_by_id(module_id)
            baseline = next((item for item in draft.baseline.modules if item.id == module_id), None)
            old_name, new_name = names.get(module_id, (None, None))
            old_text = baseline.full_source if baseline is not None else None
            if "delete" in ops:
                new_text = None
            elif mod is not None:
                new_text = _composed(draft, module_id, mod.body, mod.current_name)
            else:
                new_text = None
            review_items.append(
                ReviewItem(
                    target_id=module_id,
                    category="VBA",
                    operations=tuple(ops),
                    old_name=old_name,
                    new_name=new_name,
                    old_text=old_text,
                    new_text=new_text,
                    unified_diff=_diff(old_text, new_text, new_name or old_name or module_id),
                )
            )
        for xml in changes.xml_modified:
            part = draft.xml_part_by_path(xml.path)
            old_text = part.original_text if part is not None else None
            new_text = part.text if part is not None else None
            review_items.append(
                ReviewItem(
                    target_id=xml.path,
                    category="XML",
                    operations=("edit",),
                    old_name=xml.path,
                    new_name=xml.path,
                    old_text=old_text,
                    new_text=new_text,
                    unified_diff=_diff(old_text, new_text, xml.path),
                )
            )
        vba_count = sum(1 for item in review_items if item.category == "VBA")
        xml_count = sum(1 for item in review_items if item.category == "XML")
        signature = bool(
            changes.has_vba_changes and draft.baseline.safety.signature_present
        )
        return ReviewModel(
            binding=session.binding(),
            destination=destination or str(draft.baseline.path),
            file_type=draft.baseline.extension,
            items=tuple(review_items),
            vba_count=vba_count,
            xml_count=xml_count,
            encoding_warnings=(),
            signature_will_be_removed=signature,
            backup_policy="A same-folder backup is created before the original is replaced."
            if operation_type == "save"
            else "Save a Copy does not replace the original. An existing destination is backed up.",
            operation_type=operation_type,
        )
