"""Apply undo/redo commands to a draft (IMP-04)."""

from __future__ import annotations

from vba_addin_editor.domain.document import (
    BinaryPartDraft,
    DocumentDraft,
    ModuleDraft,
    XmlPartDraft,
)
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
        if command.op in {"replace_all", "folder_sync", "import_batch", "revert"} and isinstance(state, dict):
            modules = state.get("modules")
            if isinstance(modules, list):
                draft.modules[:] = [_module_from_snapshot(item) for item in modules]
            package = state.get("package")
            if isinstance(package, dict):
                restore_package(draft, package)
            xml_state = state.get("xml")
            if isinstance(xml_state, list):
                by_path = {item.get("path"): item for item in xml_state if isinstance(item, dict)}
                for part in draft.xml_parts:
                    payload = by_path.get(part.path)
                    if payload is not None:
                        part.text = payload.get("text")
            return
        if command.op in {"text", "paste"}:
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


def snapshot_xml(draft: DocumentDraft) -> list[dict]:
    return [{"path": part.path, "text": part.text} for part in draft.xml_parts]


def snapshot_package(draft: DocumentDraft) -> dict:
    """The members a draft adds (new XML parts, images) or drops, for undo of package edits."""
    return {
        "new_xml": [
            {
                "path": part.path,
                "text": part.text,
                "encoding": part.encoding,
                "bom": part.bom,
                "newline": part.newline,
                "is_relationships_part": part.is_relationships_part,
                "is_content_types_part": part.is_content_types_part,
            }
            for part in draft.new_xml_parts()
        ],
        "binary": [
            {"path": b.path, "data": b.data, "content_type": b.content_type}
            for b in draft.binary_parts
        ],
        "removed": list(draft.removed_parts),
    }


def restore_package(draft: DocumentDraft, package: dict) -> None:
    draft.xml_parts[:] = [p for p in draft.xml_parts if not p.is_new] + [
        XmlPartDraft(
            path=item["path"],
            text=item.get("text"),
            original_text=None,
            encoding=item.get("encoding") or "utf-8",
            bom=item.get("bom") or b"",
            newline=item.get("newline") or "\n",
            is_relationships_part=bool(item.get("is_relationships_part")),
            is_content_types_part=bool(item.get("is_content_types_part")),
            is_new=True,
        )
        for item in package.get("new_xml", [])
    ]
    draft.binary_parts[:] = [
        BinaryPartDraft(item["path"], bytes(item["data"]), item["content_type"])
        for item in package.get("binary", [])
    ]
    draft.removed_parts[:] = list(package.get("removed", []))


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
