"""Unified undo/redo history for one document session (IMP-04)."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Literal

CommandOp = Literal[
    "text",
    "xml_text",
    "add",
    "delete",
    "rename",
    "import_batch",
    "folder_sync",
    "revert",
    "replace_all",
    "paste",
]

TYPING_COALESCE_MS = 750
MAX_COMMANDS = 1000
MAX_PAYLOAD_BYTES = 64 * 1024 * 1024


def _payload_size(value: Any) -> int:
    if value is None:
        return 0
    if isinstance(value, str):
        return len(value.encode("utf-8"))
    if isinstance(value, dict):
        return sum(_payload_size(v) + _payload_size(k) for k, v in value.items())
    if isinstance(value, (list, tuple)):
        return sum(_payload_size(v) for v in value)
    return 64


@dataclass
class HistoryCommand:
    op: CommandOp
    target_id: str
    before: Any
    after: Any
    selection_before: str | None = None
    selection_after: str | None = None
    timestamp_ms: int = 0
    label: str = ""

    @property
    def payload_bytes(self) -> int:
        return _payload_size(self.before) + _payload_size(self.after)


@dataclass
class UndoHistory:
    undo_stack: list[HistoryCommand] = field(default_factory=list)
    redo_stack: list[HistoryCommand] = field(default_factory=list)
    payload_bytes: int = 0
    warned_overflow: bool = False

    def clear(self) -> None:
        self.undo_stack.clear()
        self.redo_stack.clear()
        self.payload_bytes = 0
        self.warned_overflow = False

    def can_undo(self) -> bool:
        return bool(self.undo_stack)

    def can_redo(self) -> bool:
        return bool(self.redo_stack)

    def record(
        self,
        command: HistoryCommand,
        *,
        coalesce: bool = False,
        now_ms: int = 0,
    ) -> str | None:
        """Push a command. Returns a warning if older history was dropped."""
        warning = None
        if (
            coalesce
            and command.op in {"text", "xml_text"}
            and self.undo_stack
            and not self.redo_stack
        ):
            last = self.undo_stack[-1]
            if (
                last.op == command.op
                and last.target_id == command.target_id
                and now_ms - last.timestamp_ms <= TYPING_COALESCE_MS
            ):
                self.payload_bytes -= last.payload_bytes
                last.after = command.after
                last.selection_after = command.selection_after
                last.timestamp_ms = now_ms
                self.payload_bytes += last.payload_bytes
                return None
        self.redo_stack.clear()
        self.undo_stack.append(command)
        self.payload_bytes += command.payload_bytes
        warning = self._evict()
        return warning

    def undo(self) -> HistoryCommand | None:
        if not self.undo_stack:
            return None
        command = self.undo_stack.pop()
        self.redo_stack.append(command)
        return command

    def redo(self) -> HistoryCommand | None:
        if not self.redo_stack:
            return None
        command = self.redo_stack.pop()
        self.undo_stack.append(command)
        return command

    def _evict(self) -> str | None:
        warning = None
        while len(self.undo_stack) > 1 and (
            len(self.undo_stack) > MAX_COMMANDS or self.payload_bytes > MAX_PAYLOAD_BYTES
        ):
            dropped = self.undo_stack.pop(0)
            self.payload_bytes -= dropped.payload_bytes
            warning = "Older undo history was cleared to stay within memory limits."
        if len(self.undo_stack) == 1 and self.payload_bytes > MAX_PAYLOAD_BYTES:
            warning = "Older undo history was cleared to stay within memory limits."
        if warning:
            self.warned_overflow = True
        return warning
