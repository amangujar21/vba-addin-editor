"""External-change comparison and three-way resolution (IMP-03)."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from vba_addin_editor.adapters.pyopenvba_adapter import AdapterError
from vba_addin_editor.domain.document import (
    DocumentDraft,
    DocumentSnapshot,
    ModuleDraft,
    XmlPartDraft,
    clone_draft,
    draft_from_snapshot,
)
from vba_addin_editor.domain.results import SaveResult
from vba_addin_editor.domain.session import DocumentSession, ProposalBinding
from vba_addin_editor.platform import paths
from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.services.validation_service import validate_draft


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
    stream_key: str = ""


@dataclass
class ConflictProposal:
    binding: ProposalBinding
    external_sha256: str
    items: list[ConflictItem] = field(default_factory=list)
    uneditable_from_external: bool = False
    parse_failed: bool = False
    parse_problem: str | None = None


def _stream_key(module) -> str:
    stream = (getattr(module, "stream_name", None) or "").strip()
    if stream:
        return "s:" + stream.casefold()
    origin = (
        getattr(module, "origin_name", None)
        or getattr(module, "original_name", None)
        or ""
    )
    if origin:
        return "n:" + origin.casefold()
    current = getattr(module, "current_name", "") or ""
    if getattr(module, "is_new", False):
        return "new:" + getattr(module, "id", current)
    return "n:" + current.casefold()


class ConflictService:
    def __init__(self, document_service: DocumentService | None = None) -> None:
        self.document_service = document_service or DocumentService()

    def snapshot_external(self, path: Path) -> tuple[DocumentSnapshot | None, str | None, str | None]:
        """Open a stable external snapshot. Returns (snapshot, sha256, error)."""
        path = Path(path)
        try:
            before = paths.fingerprint(path)
        except FileNotFoundError:
            return None, None, "source_missing"
        except paths.UnstableSourceError:
            return None, None, "source_unstable"
        except OSError as exc:
            return None, None, type(exc).__name__
        try:
            draft = self.document_service.open(path)
        except (AdapterError, OSError, ValueError) as exc:
            return None, before.sha256, str(exc)
        try:
            after = paths.fingerprint(path)
        except OSError as exc:
            return None, before.sha256, type(exc).__name__
        if not after.content_equal(before):
            return None, before.sha256, "source_unstable"
        return draft.baseline, before.sha256, None

    def compare(self, session: DocumentSession, external: DocumentSnapshot) -> ConflictProposal:
        draft = session.draft
        baseline = draft.baseline
        items: list[ConflictItem] = []
        base_by_key = {_stream_key(m): m for m in baseline.modules}
        draft_by_key = {_stream_key(m): m for m in draft.modules if not m.is_new}
        draft_new = [m for m in draft.modules if m.is_new and not m.is_deleted]
        ext_by_key = {_stream_key(m): m for m in external.modules}

        keys = set(base_by_key) | set(draft_by_key) | set(ext_by_key)
        for key in sorted(keys):
            items.extend(
                self._module_conflicts(
                    key,
                    base_by_key.get(key),
                    draft_by_key.get(key),
                    ext_by_key.get(key),
                )
            )

        ext_names = {m.original_name.casefold() for m in external.modules}
        for added in draft_new:
            collision = added.current_name.casefold() in ext_names
            items.append(
                ConflictItem(
                    target_id=added.id,
                    category="VBA",
                    property_name="module",
                    baseline=None,
                    draft=added.current_name,
                    external="added" if collision else None,
                    resolution=None if collision else "keep_draft",
                    requires_choice=collision,
                    structural="name_collision" if collision else "draft_add",
                    stream_key=_stream_key(added),
                )
            )

        items.extend(self._xml_conflicts(baseline, draft, external))
        return ConflictProposal(
            binding=session.binding(),
            external_sha256=external.file_fingerprint.sha256,
            items=items,
        )

    def _module_conflicts(self, key: str, base, dmod: ModuleDraft | None, emod) -> list[ConflictItem]:
        items: list[ConflictItem] = []
        target = (
            (dmod.id if dmod is not None else None)
            or (base.id if base is not None else None)
            or (emod.id if emod is not None else key)
        )
        d_deleted = dmod is not None and dmod.is_deleted
        e_deleted = emod is None
        b_present = base is not None

        if not b_present:
            if dmod is not None and emod is None:
                return items
            if dmod is None and emod is not None:
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
                        stream_key=key,
                    )
                )
            return items

        if d_deleted and e_deleted:
            return items

        if d_deleted and emod is not None:
            if emod.body != base.body or emod.original_name.casefold() != base.original_name.casefold():
                items.append(
                    ConflictItem(
                        target_id=target,
                        category="VBA",
                        property_name="module",
                        baseline=base.original_name,
                        draft="deleted",
                        external=emod.body,
                        requires_choice=True,
                        structural="delete_versus_edit",
                        stream_key=key,
                    )
                )
            else:
                items.append(
                    ConflictItem(
                        target_id=target,
                        category="VBA",
                        property_name="module",
                        baseline=base.original_name,
                        draft="deleted",
                        external=emod.original_name,
                        resolution="keep_draft",
                        requires_choice=False,
                        structural="draft_delete",
                        stream_key=key,
                    )
                )
            return items

        if e_deleted and dmod is not None and not dmod.is_deleted:
            if dmod.body != base.body or dmod.current_name.casefold() != base.original_name.casefold():
                items.append(
                    ConflictItem(
                        target_id=target,
                        category="VBA",
                        property_name="module",
                        baseline=base.original_name,
                        draft=dmod.body,
                        external="deleted",
                        requires_choice=True,
                        structural="delete_versus_edit",
                        stream_key=key,
                    )
                )
            else:
                items.append(
                    ConflictItem(
                        target_id=target,
                        category="VBA",
                        property_name="module",
                        baseline=base.original_name,
                        draft=dmod.current_name,
                        external="deleted",
                        resolution="keep_disk",
                        requires_choice=False,
                        structural="external_delete",
                        stream_key=key,
                    )
                )
            return items

        if dmod is None or emod is None:
            return items

        d_name = dmod.current_name
        e_name = emod.original_name
        b_name = base.original_name
        if d_name.casefold() != b_name.casefold() and e_name.casefold() != b_name.casefold() and d_name.casefold() != e_name.casefold():
            items.append(
                ConflictItem(
                    target_id=target,
                    category="VBA",
                    property_name="name",
                    baseline=b_name,
                    draft=d_name,
                    external=e_name,
                    requires_choice=True,
                    structural="rename_versus_rename",
                    stream_key=key,
                )
            )
        else:
            items.append(self._property_conflict(target, "VBA", "name", b_name, d_name, e_name, key))

        items.append(self._property_conflict(target, "VBA", "body", base.body, dmod.body, emod.body, key))

        b_kind = getattr(base, "project_item_kind", "") or ""
        d_kind = dmod.project_item_kind or dmod.kind
        e_kind = getattr(emod, "project_item_kind", "") or getattr(emod, "kind", "")
        if d_kind != b_kind or e_kind != b_kind:
            items.append(
                self._property_conflict(target, "VBA", "kind", b_kind, d_kind, e_kind, key)
            )
            if items[-1].requires_choice:
                items[-1].structural = "kind_change"
        return items

    def _xml_conflicts(self, baseline, draft: DocumentDraft, external: DocumentSnapshot) -> list[ConflictItem]:
        items: list[ConflictItem] = []
        base_xml = {part.path: part for part in baseline.xml_parts}
        draft_xml = {part.path: part for part in draft.xml_parts}
        ext_xml = {part.path: part for part in external.xml_parts}
        paths_union = set(base_xml) | set(draft_xml) | set(ext_xml)
        for path in sorted(paths_union):
            base_part = base_xml.get(path)
            dpart = draft_xml.get(path)
            epart = ext_xml.get(path)
            btxt = base_part.text if base_part is not None else None
            dtxt = dpart.text if dpart is not None else None
            etxt = epart.text if epart is not None else None
            if base_part is None and dpart is None and epart is not None:
                items.append(
                    ConflictItem(
                        target_id=path,
                        category="XML",
                        property_name="text",
                        baseline=None,
                        draft=None,
                        external=etxt,
                        requires_choice=True,
                        structural="external_add",
                    )
                )
                continue
            if base_part is not None and epart is None:
                items.append(
                    ConflictItem(
                        target_id=path,
                        category="XML",
                        property_name="text",
                        baseline=btxt,
                        draft=dtxt,
                        external=None,
                        requires_choice=True,
                        structural="external_delete",
                    )
                )
                continue
            items.append(self._property_conflict(path, "XML", "text", btxt, dtxt, etxt, path))
        return items

    def _property_conflict(
        self,
        target_id: str,
        category: str,
        property_name: str,
        baseline: str | None,
        draft: str | None,
        external: str | None,
        stream_key: str = "",
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
                stream_key=stream_key,
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
                stream_key=stream_key,
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
                stream_key=stream_key,
            )
        return ConflictItem(
            target_id=target_id,
            category=category,
            property_name=property_name,
            baseline=baseline,
            draft=draft,
            external=external,
            requires_choice=True,
            stream_key=stream_key,
        )

    def all_resolved(self, proposal: ConflictProposal) -> bool:
        if proposal.parse_failed:
            return False
        return all((not item.requires_choice) or item.resolution for item in proposal.items)

    def apply(
        self,
        session: DocumentSession,
        proposal: ConflictProposal,
        external: DocumentSnapshot,
        *,
        live_external_hash: str | None = None,
    ) -> DocumentDraft:
        if proposal.parse_failed:
            raise ValueError("unparseable_external")
        if not session.matches_proposal(proposal.binding):
            raise ValueError("stale_proposal")
        if live_external_hash is not None and live_external_hash != proposal.external_sha256:
            raise ValueError("stale_external")
        if external.file_fingerprint.sha256 != proposal.external_sha256:
            raise ValueError("stale_external")
        if not self.all_resolved(proposal):
            raise ValueError("unresolved_conflicts")
        for item in proposal.items:
            if item.resolution == "merged" and not (item.merged_text or "").strip() and item.merged_text is None:
                raise ValueError("unresolved_conflicts")

        new_draft = draft_from_snapshot(external)
        old = session.draft
        old_by_key = {_stream_key(m): m for m in old.modules}
        new_by_key = {_stream_key(m): m for m in new_draft.modules}

        for item in proposal.items:
            if item.category != "VBA":
                continue
            self._apply_vba_item(item, old, new_draft, old_by_key, new_by_key)

        for item in proposal.items:
            if item.category != "XML":
                continue
            self._apply_xml_item(item, old, new_draft)

        new_draft.signed_save_confirmed = False
        problems = validate_draft(new_draft)
        if problems:
            raise ValueError("invalid_draft:" + "; ".join(problems))
        return new_draft

    def _apply_vba_item(
        self,
        item: ConflictItem,
        old: DocumentDraft,
        new_draft: DocumentDraft,
        old_by_key: dict[str, ModuleDraft],
        new_by_key: dict[str, ModuleDraft],
    ) -> None:
        key = item.stream_key or item.target_id
        old_mod = old.module_by_id(item.target_id) or old_by_key.get(key)
        new_mod = new_by_key.get(key)
        if old_mod is None and item.structural == "draft_add":
            old_mod = old.module_by_id(item.target_id)

        if item.structural in {"draft_delete", "delete_versus_edit"} and item.resolution == "keep_draft":
            if item.draft == "deleted" or (old_mod is not None and old_mod.is_deleted):
                if new_mod is not None:
                    new_mod.is_deleted = True
                return
            if item.external == "deleted" and old_mod is not None:
                self._ensure_draft_module(new_draft, new_by_key, old_mod)
            return

        if item.structural == "external_add" and item.resolution == "keep_draft":
            if new_mod is not None:
                new_draft.modules.remove(new_mod)
                new_by_key.pop(key, None)
            return

        if item.structural == "draft_add" and item.resolution == "keep_draft" and old_mod is not None:
            self._ensure_draft_module(new_draft, new_by_key, old_mod)
            return

        if item.structural == "name_collision" and item.resolution == "keep_draft" and old_mod is not None:
            self._ensure_draft_module(new_draft, new_by_key, old_mod)
            return

        if new_mod is None:
            if item.resolution == "keep_draft" and old_mod is not None and not old_mod.is_deleted:
                self._ensure_draft_module(new_draft, new_by_key, old_mod)
            return

        if item.property_name == "body":
            if item.resolution == "keep_draft" and old_mod is not None:
                new_mod.body = old_mod.body
            elif item.resolution == "merged":
                new_mod.body = item.merged_text or ""
        elif item.property_name == "name":
            if item.resolution == "keep_draft" and old_mod is not None:
                new_mod.current_name = old_mod.current_name
            elif item.resolution == "merged" and item.merged_text:
                new_mod.current_name = item.merged_text
        elif item.property_name == "kind" and item.resolution == "keep_draft" and old_mod is not None:
            new_mod.kind = old_mod.kind
            new_mod.project_item_kind = old_mod.project_item_kind
            new_mod.pyopenvba_kind = old_mod.pyopenvba_kind

    def _ensure_draft_module(
        self, new_draft: DocumentDraft, new_by_key: dict[str, ModuleDraft], old_mod: ModuleDraft
    ) -> None:
        key = _stream_key(old_mod)
        if any(m.id == old_mod.id for m in new_draft.modules):
            return
        clone = clone_draft(DocumentDraft(baseline=new_draft.baseline, modules=[old_mod])).modules[0]
        clone.is_new = True if not old_mod.stream_name else old_mod.is_new
        new_draft.modules.append(clone)
        new_by_key[key] = clone

    def _apply_xml_item(self, item: ConflictItem, old: DocumentDraft, new_draft: DocumentDraft) -> None:
        path = item.target_id
        old_part = old.xml_part_by_path(path)
        new_part = new_draft.xml_part_by_path(path)
        if item.structural == "external_add" and item.resolution == "keep_draft":
            if new_part is not None:
                new_draft.xml_parts.remove(new_part)
            return
        if item.structural == "external_delete" and item.resolution == "keep_draft" and old_part is not None:
            if new_part is None:
                new_draft.xml_parts.append(
                    XmlPartDraft(
                        path=old_part.path,
                        text=old_part.text,
                        original_text=old_part.original_text,
                        encoding=old_part.encoding,
                        bom=old_part.bom,
                        newline=old_part.newline,
                        is_relationships_part=old_part.is_relationships_part,
                        is_content_types_part=old_part.is_content_types_part,
                        open_problem=old_part.open_problem,
                    )
                )
            return
        if new_part is None:
            return
        if item.resolution == "keep_draft" and old_part is not None:
            new_part.text = old_part.text
        elif item.resolution == "merged":
            new_part.text = item.merged_text

    def commit(
        self,
        session: DocumentSession,
        proposal: ConflictProposal,
        external: DocumentSnapshot,
        *,
        live_external_hash: str,
        external_path: Path,
        session_service,
        recovery_service,
    ) -> SaveResult:
        try:
            new_draft = self.apply(
                session, proposal, external, live_external_hash=live_external_hash
            )
        except ValueError as exc:
            token = str(exc)
            if token.startswith("invalid_draft:"):
                return SaveResult.blocked(
                    "invalid_draft",
                    token.split(":", 1)[1],
                    operation_type="conflict",
                )
            reason = {
                "stale_proposal": "stale_proposal",
                "stale_external": "external_change",
                "unresolved_conflicts": "invalid_draft",
                "unparseable_external": "external_change",
            }.get(token, "invalid_draft")
            return SaveResult.blocked(reason, operation_type="conflict")
        recovery_service.retain_named_checkpoint(session, "pre_resolution")
        session.draft = new_draft
        published = session_service.publish_verified_baseline(session, Path(external_path))
        if published is not None:
            return published
        session.conflicts_pending = False
        failed = recovery_service.checkpoint(session)
        if failed is not None:
            return failed
        return SaveResult.success(None, Path(external_path), operation_type="conflict")


def draft_from_external(snapshot: DocumentSnapshot):
    return draft_from_snapshot(snapshot)
