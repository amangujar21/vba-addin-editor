"""Structured results for save pipeline and candidate verification."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

BlockedReason = str  # "office_running" | "external_change" | "password_protected" | ...


@dataclass(frozen=True)
class CandidateVerificationResult:
    ok: bool
    problems: tuple[str, ...] = ()
    details: dict[str, Any] = field(default_factory=dict)


@dataclass
class SaveResult:
    """One save attempt outcome. UI maps `reason` to plain-language dialogs."""

    kind: str  # "success" | "no_changes" | "blocked" | "needs_signature_confirmation" \
    #          # | "candidate_failed" | "commit_failed" | "recovery_required" | "error"
    reason: str | None = None
    message: str | None = None
    backup_path: Path | None = None
    candidate_path: Path | None = None
    problems: tuple[str, ...] = ()
    details: dict[str, Any] = field(default_factory=dict)

    @classmethod
    def success(cls, backup: Path | None, snapshot_path: Path) -> SaveResult:
        return cls(kind="success", backup_path=backup, details={"saved": str(snapshot_path)})

    @classmethod
    def no_changes(cls) -> SaveResult:
        return cls(kind="no_changes", message="No changes to save.")

    @classmethod
    def blocked(cls, reason: str, message: str | None = None) -> SaveResult:
        return cls(kind="blocked", reason=reason, message=message)

    @classmethod
    def needs_signature_confirmation(cls) -> SaveResult:
        return cls(kind="needs_signature_confirmation")

    @classmethod
    def candidate_failed(cls, problems: tuple[str, ...], candidate: Path | None) -> SaveResult:
        return cls(
            kind="candidate_failed",
            problems=problems,
            candidate_path=candidate,
            message="The edited copy failed verification. Your original add-in was not changed.",
        )

    @classmethod
    def recovery_required(cls, backup: Path | None, problems: tuple[str, ...]) -> SaveResult:
        return cls(
            kind="recovery_required",
            backup_path=backup,
            problems=problems,
            message="The add-in was replaced but final verification failed. Use Restore Backup.",
        )

    @classmethod
    def error(cls, message: str, reason: str | None = None) -> SaveResult:
        return cls(kind="error", reason=reason, message=message)
