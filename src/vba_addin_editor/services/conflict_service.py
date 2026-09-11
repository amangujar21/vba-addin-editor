"""External-change comparison and three-way resolution (IMP-03)."""

from __future__ import annotations

from dataclasses import dataclass, field

from vba_addin_editor.domain.document import (
    DocumentDraft,
    DocumentSnapshot,
    ModuleDraft,
)
from vba_addin_editor.domain.session import DocumentSession, ProposalBinding
from vba_addin_editor.services.document_service import DocumentService


@dataclass
class ConflictItem:
    target_id: str
    category: str
    property_name: str
    baseline: str | None
    draft: str | None
    external: str | None
    resolution: str | None = None  # keep_draft | keep_disk | merged
    merged_text: str | None = None
    requires_choice: bool = True
    structural: str | None = None


@dataclass
class ConflictProposal:
    binding: ProposalBinding
    external_sha256: str
    items: list[ConflictItem] = field(default_factory=list)
    uneditable_from_external: bool = False
    parse_failed: bool = False


class ConflictService:
    def __init__(self, document_service: DocumentService | None = None) -> None:
        self.document_service = document_service or DocumentService()

    def compare(self, session: DocumentSession, external: DocumentSnapshot) -> ConflictProposal:
        draft = session.draft
        baseline = draft.baseline
        items: list[ConflictItem] = []
        base_by_stream = {m.stream_name or m.original_name: m for m in baseline.modules}
        draft_by_id = {m.id: m for m in draft.modules}
        ext_by_stream = {m.stream_name or m.original_name: m for m in external.modules}
        seen_streams: set[str] = set()
        for stream, base_mod in base_by_stream.items():
            seen_streams.add(stream)
            dmod = draft_by_id.get(base_mod.id)
            emod = ext_by_stream.get(stream)
            items.extend(self._module_conflicts(base_mod, dmod, emod))
        for stream, emod in ext_by_stream.items():
            if stream in seen_streams:
                continue
            items.append(
                ConflictItem(
                    target_id=emod.id,
                    category="VBA",
                    property_name="module",
                    baseline=None,
                    draft=None,
                    external=emod.original_name,
                    requires_choice=True,
                    structural="external_add",
                )
            )
        base_xml = {part.path: part for part in baseline.xml_parts}
        draft_xml = {part.path: part for part in draft.xml_parts}
        ext_xml = {part.path: part for part in external.xml_parts}
        for path, base_part in base_xml.items():
            dpart = draft_xml.get(path)
            epart = ext_xml.get(path)
            btxt = base_part.text
            dtxt = dpart.text if dpart is not None else None
            etxt = epart.text if epart is not None else None
            items.append(self._property_conflict(path, "XML", "text", btxt, dtxt, etxt))
        auto_resolved = []
        for item in items:
            if not item.requires_choice:
                auto_resolved.append(item)
        return ConflictProposal(
            binding=session.binding(),
            external_sha256=external.file_fingerprint.sha256,
            items=items,
        )

    def _module_conflicts(self, base, dmod: ModuleDraft | None, emod) -> list[ConflictItem]:
        items: list[ConflictItem] = []
        d_deleted = dmod is None or dmod.is_deleted
        e_deleted = emod is None
        if d_deleted and e_deleted:
            return items
        if d_deleted and emod is not None and emod.body != base.body:
            items.append(
                ConflictItem(
                    target_id=base.id,
                    category="VBA",
                    property_name="module",
                    baseline=base.original_name,
                    draft="deleted",
                    external=emod.body,
                    requires_choice=True,
                    structural="delete_versus_edit",
                )
            )
            return items
        if e_deleted and dmod is not None and not dmod.is_deleted and dmod.body != base.body:
            items.append(
                ConflictItem(
                    target_id=base.id,
                    category="VBA",
                    property_name="module",
                    baseline=base.original_name,
                    draft=dmod.body,
                    external="deleted",
                    requires_choice=True,
                    structural="delete_versus_edit",
                )
            )
            return items
        if dmod is None or emod is None:
            return items
        items.append(
            self._property_conflict(base.id, "VBA", "body", base.body, dmod.body, emod.body)
        )
        items.append(
            self._property_conflict(
                base.id, "VBA", "name", base.original_name, dmod.current_name, emod.original_name
            )
        )
        return items

    def _property_conflict(
        self,
        target_id: str,
        category: str,
        property_name: str,
        baseline: str | None,
        draft: str | None,
        external: str | None,
    ) -> ConflictItem:
        if draft == baseline:
            return ConflictItem(
                target_id=target_id,
                category=category,
                property_name=property_name,
                baseline=baseline,
                draft=draft,
                external=external,
                resolution="keep_disk",
                requires_choice=False,
            )
        if external == baseline:
            return ConflictItem(
                target_id=target_id,
                category=category,
                property_name=property_name,
                baseline=baseline,
                draft=draft,
                external=external,
                resolution="keep_draft",
                requires_choice=False,
            )
        if draft == external:
            return ConflictItem(
                target_id=target_id,
                category=category,
                property_name=property_name,
                baseline=baseline,
                draft=draft,
                external=external,
                resolution="keep_draft",
                requires_choice=False,
            )
        return ConflictItem(
            target_id=target_id,
            category=category,
            property_name=property_name,
            baseline=baseline,
            draft=draft,
            external=external,
            requires_choice=True,
        )

    def all_resolved(self, proposal: ConflictProposal) -> bool:
        return all((not item.requires_choice) or item.resolution for item in proposal.items)

    def apply(
        self,
        session: DocumentSession,
        proposal: ConflictProposal,
        external: DocumentSnapshot,
    ) -> DocumentDraft:
        if not session.matches_proposal(proposal.binding):
            raise ValueError("stale_proposal")
        if not self.all_resolved(proposal):
            raise ValueError("unresolved_conflicts")
        new_draft = draft_from_external(external)
        resolutions = {(item.target_id, item.property_name): item for item in proposal.items}
        for module in new_draft.modules:
            body_item = resolutions.get((module.id, "body")) or resolutions.get((module.origin_name or "", "body"))
            # Match by stream via baseline id on the old session draft.
            old = session.draft.module_by_id(module.id)
            if old is None:
                old = next(
                    (
                        item
                        for item in session.draft.modules
                        if item.stream_name == module.stream_name
                    ),
                    None,
                )
            name_item = resolutions.get((old.id if old else module.id, "name"))
            body_item = resolutions.get((old.id if old else module.id, "body"))
            if body_item and body_item.resolution == "keep_draft" and old is not None:
                module.body = old.body
            if name_item and name_item.resolution == "keep_draft" and old is not None:
                module.current_name = old.current_name
        for part in new_draft.xml_parts:
            item = resolutions.get((part.path, "text"))
            old_part = session.draft.xml_part_by_path(part.path)
            if item and item.resolution == "keep_draft" and old_part is not None:
                part.text = old_part.text
            if item and item.resolution == "merged":
                part.text = item.merged_text
        new_draft.signed_save_confirmed = False
        return new_draft


def draft_from_external(snapshot: DocumentSnapshot):
    from vba_addin_editor.domain.document import draft_from_snapshot

    return draft_from_snapshot(snapshot)
