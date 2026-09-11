"""Map operation results to user actions and copyable reports (IMP-07)."""

from __future__ import annotations

from vba_addin_editor.domain.diagnostics import (
    build_diagnostic_report,
    explanation_for,
)
from vba_addin_editor.domain.results import SaveResult


def actions_for(result: SaveResult) -> list[str]:
    reason = result.reason
    if result.retryable:
        actions = ["Retry"]
    else:
        actions = []
    if reason == "external_change":
        actions.append("Compare Changes")
    if reason in {"source_missing", "destination_exists", "same_path", "reparse_point"}:
        actions.append("Choose Destination")
    if reason == "recovery_required":
        actions.append("Restore Backup")
    if reason == "recovery_write_failed":
        actions.append("Save Recovery Now")
    if reason == "stale_proposal":
        actions.append("Rebuild Review")
    actions.append("Copy Diagnostic Report")
    return actions


def report_for(result: SaveResult, *, include_full_paths: bool = False, extension: str | None = None) -> str:
    return build_diagnostic_report(
        envelope=result.envelope(),
        extension=extension,
        include_full_paths=include_full_paths,
    )


def dialog_text(result: SaveResult) -> str:
    explanation, action = explanation_for(result.reason)
    message = result.message or explanation
    return f"{message}\n\n{action}"
