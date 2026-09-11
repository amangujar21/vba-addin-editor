"""Apply undo/redo commands to a draft (IMP-04)."""

from __future__ import annotations

from vba_addin_editor.domain.document import DocumentDraft, ModuleDraft
from vba_addin_editor.domain.history import HistoryCommand
from vba_addin_editor.domain.session import DocumentSession


class HistoryService:
    def record_text(
        self,
        session: DocumentSession,
        *,
        target_id: str,
        before: str,
        after: str,
        op: str = "text",
        now_ms: int = 0,
        coalesce: bool = True,
    ) -> str | None:
        if before == after:
            return None
        command = HistoryCommand(
            op=op,  # type: ignore[arg-type]
            target_id=target_id,
            before=before,
            after=after,
            timestamp_ms=now_ms,
        )
        warning = session.history.record(command, coalesce=coalesce, now_ms=now_ms)
        session.bump()
        return warning

    def record_structural(self, session: DocumentSession, command: HistoryCommand) -> None:
        session.history.record(command, coalesce=False, now_ms=command.timestamp_ms)
        session.bump()

    def undo(self, session: DocumentSession) -> HistoryCommand | None:
        command = session.history.undo()
        if command is None:
            return None
        self._apply(session.draft, command, reverse=True)
        session.bump()
        return command

    def redo(self, session: DocumentSession) -> HistoryCommand | None:
        command = session.history.redo()
        if command is None:
            return None
        self._apply(session.draft, command, reverse=False)
        session.bump()
        return command

    def _apply(self, draft: DocumentDraft, command: HistoryCommand, *, reverse: bool) -> None:
        state = command.before if reverse else command.after
        if command.op in {"text", "paste", "replace_all"}:
            module = draft.module_by_id(command.target_id)
            if module is not None:
                module.body = state
            return
        if command.op == "xml_text":
            part = draft.xml_part_by_path(command.target_id)
            if part is not None:
                part.text = state
            return
        if command.op == "rename":
            module = draft.module_by_id(command.target_id)
            if module is not None:
                module.current_name = state
            return
        if command.op == "delete":
            module = draft.module_by_id(command.target_id)
            if module is None:
                return
            module.is_deleted = bool(state) if not reverse else bool(command.before)
            return
        if command.op == "add":
            if reverse:
                module = draft.module_by_id(command.target_id)
                if module is not None:
                    draft.modules.remove(module)
            else:
                if draft.module_by_id(command.target_id) is None and isinstance(command.after, dict):
                    draft.modules.append(_module_from_snapshot(command.after))
            return
        if command.op in {"import_batch", "folder_sync", "revert"} and isinstance(state, list):
            draft.modules[:] = [_module_from_snapshot(item) for item in state]


def _module_from_snapshot(item: dict) -> ModuleDraft:
    return ModuleDraft(
        id=item["id"],
        origin_name=item.get("origin_name"),
        current_name=item["current_name"],
        body=item.get("body") or "",
        kind=item.get("kind") or "standard",
        pyopenvba_kind=item.get("pyopenvba_kind") or "standard",
        is_new=bool(item.get("is_new")),
        is_deleted=bool(item.get("is_deleted")),
        destructive_ops_safe=bool(item.get("destructive_ops_safe", item.get("can_delete"))),
        original_body=item.get("original_body"),
        can_delete=bool(item.get("can_delete")),
        can_rename=bool(item.get("can_rename")),
        restriction_reason=item.get("restriction_reason"),
        stream_name=item.get("stream_name") or "",
        hidden_header=item.get("hidden_header") or "",
        project_item_kind=item.get("project_item_kind") or "unknown",
    )


def snapshot_modules(draft: DocumentDraft) -> list[dict]:
    return [
        {
            "id": module.id,
            "origin_name": module.origin_name,
            "current_name": module.current_name,
            "body": module.body,
            "kind": module.kind,
            "pyopenvba_kind": module.pyopenvba_kind,
            "is_new": module.is_new,
            "is_deleted": module.is_deleted,
            "destructive_ops_safe": module.destructive_ops_safe,
            "original_body": module.original_body,
            "can_delete": module.can_delete,
            "can_rename": module.can_rename,
            "restriction_reason": module.restriction_reason,
            "stream_name": module.stream_name,
            "hidden_header": module.hidden_header,
            "project_item_kind": module.project_item_kind,
        }
        for module in draft.modules
    ]
