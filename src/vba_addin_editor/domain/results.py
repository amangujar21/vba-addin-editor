"""Structured results for save pipeline and candidate verification."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from vba_addin_editor.domain.diagnostics import (
    OperationEnvelope,
    explanation_for,
    is_retryable,
    new_operation_id,
    utc_now,
)

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
    operation_id: str = field(default_factory=new_operation_id)
    operation_type: str = "save"
    stage: str | None = None
    retryable: bool = False
    win32_error: int | None = None
    win32_name: str | None = None
    revision: int | None = None
    utc_time: str = field(default_factory=utc_now)

    def envelope(self) -> OperationEnvelope:
        artifacts = {}
        existence = {}
        if self.backup_path is not None:
            artifacts["backup"] = str(self.backup_path)
            existence["backup"] = self.backup_path.exists()
        if self.candidate_path is not None:
            artifacts["candidate"] = str(self.candidate_path)
            existence["candidate"] = self.candidate_path.exists()
        saved = self.details.get("saved")
        if saved:
            artifacts["original"] = str(saved)
            existence["original"] = Path(str(saved)).exists()
        hashes = {
            key[5:]: value
            for key, value in self.details.items()
            if key.startswith("hash_") and isinstance(value, str)
        }
        return OperationEnvelope(
            operation_id=self.operation_id,
            operation_type=self.operation_type,
            stage=self.stage,
            reason=self.reason,
            utc_time=self.utc_time,
            retryable=self.retryable,
            win32_error=self.win32_error,
            win32_name=self.win32_name,
            revision=self.revision,
            hashes=hashes,
            artifact_paths=artifacts,
            artifact_existence=existence,
            details=dict(self.details),
        )

    @classmethod
    def success(
        cls,
        backup: Path | None,
        snapshot_path: Path,
        *,
        operation_type: str = "save",
        stage: str = "complete",
        revision: int | None = None,
    ) -> SaveResult:
        return cls(
            kind="success",
            backup_path=backup,
            details={"saved": str(snapshot_path)},
            operation_type=operation_type,
            stage=stage,
            revision=revision,
        )

    @classmethod
    def no_changes(cls) -> SaveResult:
        return cls(
            kind="no_changes",
            reason="no_changes",
            message="No changes to save.",
            stage="preflight",
        )

    @classmethod
    def blocked(
        cls,
        reason: str,
        message: str | None = None,
        *,
        stage: str | None = None,
        win32_error: int | None = None,
        win32_name: str | None = None,
        details: dict[str, Any] | None = None,
        candidate_path: Path | None = None,
        backup_path: Path | None = None,
        revision: int | None = None,
        operation_type: str = "save",
    ) -> SaveResult:
        text = message or explanation_for(reason)[0]
        return cls(
            kind="blocked",
            reason=reason,
            message=text,
            stage=stage or "preflight",
            retryable=is_retryable(reason),
            win32_error=win32_error,
            win32_name=win32_name,
            details=details or {},
            candidate_path=candidate_path,
            backup_path=backup_path,
            revision=revision,
            operation_type=operation_type,
        )

    @classmethod
    def needs_signature_confirmation(cls) -> SaveResult:
        return cls(kind="needs_signature_confirmation", stage="preflight")

    @classmethod
    def candidate_failed(cls, problems: tuple[str, ...], candidate: Path | None) -> SaveResult:
        return cls(
            kind="candidate_failed",
            reason="candidate_failed",
            problems=problems,
            candidate_path=candidate,
            message="The edited copy failed verification. Your original add-in was not changed.",
            stage="verify",
            retryable=False,
        )

    @classmethod
    def recovery_required(cls, backup: Path | None, problems: tuple[str, ...]) -> SaveResult:
        return cls(
            kind="recovery_required",
            reason="recovery_required",
            backup_path=backup,
            problems=problems,
            message="The add-in was replaced but final verification failed. Use Restore Backup.",
            stage="post_commit",
            retryable=False,
        )

    @classmethod
    def error(
        cls,
        message: str,
        reason: str | None = None,
        *,
        stage: str | None = None,
        win32_error: int | None = None,
        win32_name: str | None = None,
        details: dict[str, Any] | None = None,
        candidate_path: Path | None = None,
        backup_path: Path | None = None,
        retryable: bool = False,
        operation_type: str = "save",
    ) -> SaveResult:
        token = reason or "unknown"
        return cls(
            kind="error",
            reason=token,
            message=message,
            stage=stage,
            retryable=retryable if reason is None else (retryable or is_retryable(token)),
            win32_error=win32_error,
            win32_name=win32_name,
            details=details or {},
            candidate_path=candidate_path,
            backup_path=backup_path,
            operation_type=operation_type,
        )
