"""Shared operation result envelope and sanitized diagnostic reports (IMP-07)."""

from __future__ import annotations

import os
import platform
import sys
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from vba_addin_editor.version import APP_NAME, PYOPENVBA_PIN, VERSION, build_identity

# Stable reason tokens. Unknown errors stay unknown; never relabel them as locks.
OFFICE_RUNNING = "office_running"
PROCESS_PROBE_FAILED = "process_probe_failed"
FILE_IN_USE = "file_in_use"
ACCESS_DENIED = "access_denied"
READ_ONLY = "read_only"
SOURCE_MISSING = "source_missing"
SOURCE_UNSTABLE = "source_unstable"
EXTERNAL_CHANGE = "external_change"
DESTINATION_CHANGED = "destination_changed"
DESTINATION_EXISTS = "destination_exists"
INVALID_DRAFT = "invalid_draft"
ENCODING = "encoding"
INVALID_XML = "invalid_xml"
UNSUPPORTED_COMPONENT_OPERATION = "unsupported_component_operation"
CANDIDATE_FAILED = "candidate_failed"
COMMIT_FAILED = "commit_failed"
RECOVERY_REQUIRED = "recovery_required"
RECOVERY_WRITE_FAILED = "recovery_write_failed"
STALE_PROPOSAL = "stale_proposal"
NO_CHANGES = "no_changes"
PASSWORD_PROTECTED = "password_protected"
PACKAGE_SIGNED = "package_signed"
UNSUPPORTED_EXTENSION = "unsupported_extension"
REPARSE_POINT = "reparse_point"
SAME_PATH = "same_path"
UNKNOWN = "unknown"

RETRYABLE_REASONS = frozenset(
    {
        OFFICE_RUNNING,
        PROCESS_PROBE_FAILED,
        FILE_IN_USE,
        ACCESS_DENIED,
        SOURCE_UNSTABLE,
    }
)

REASON_EXPLANATIONS: dict[str, tuple[str, str]] = {
    OFFICE_RUNNING: (
        "The corresponding Office application is running. Close it before saving this add-in.",
        "Close the host application, then Retry.",
    ),
    PROCESS_PROBE_FAILED: (
        "Windows could not list running programs, so the editor cannot confirm Office is closed.",
        "Retry after confirming Excel and PowerPoint are closed.",
    ),
    FILE_IN_USE: (
        "Another program is using this file, so it cannot be replaced yet.",
        "Close the program that has the file open, then Retry.",
    ),
    ACCESS_DENIED: (
        "Windows denied access to this file.",
        "Check permissions, then Retry or Choose Destination.",
    ),
    READ_ONLY: (
        "The file is read-only, so it cannot be replaced.",
        "Clear the read-only flag, then Retry.",
    ),
    SOURCE_MISSING: (
        "The original add-in file is missing.",
        "Use Save a Copy to keep the draft, or restore the original path.",
    ),
    SOURCE_UNSTABLE: (
        "The file changed while it was being read.",
        "Retry once the file is idle, or Compare Changes.",
    ),
    EXTERNAL_CHANGE: (
        "The file on disk is not the copy this draft was opened from.",
        "Compare Changes. Do not overwrite blindly.",
    ),
    DESTINATION_CHANGED: (
        "The destination file changed after it was reviewed.",
        "Review the destination again, then Retry.",
    ),
    DESTINATION_EXISTS: (
        "The destination already exists.",
        "Choose another destination or confirm overwrite.",
    ),
    INVALID_DRAFT: (
        "The draft has validation problems and was not written.",
        "Fix the named modules, then Retry.",
    ),
    ENCODING: (
        "The draft contains characters the VBA project's code page cannot store.",
        "Replace the named characters, then Retry.",
    ),
    INVALID_XML: (
        "Edited XML is not safe to write.",
        "Fix the XML problems, then Retry.",
    ),
    UNSUPPORTED_COMPONENT_OPERATION: (
        "A rename or delete is not allowed for this component type.",
        "Remove the unsupported operation, then Retry.",
    ),
    CANDIDATE_FAILED: (
        "The edited copy failed verification. The original add-in was not changed.",
        "Inspect the candidate artifact, then fix the draft.",
    ),
    COMMIT_FAILED: (
        "Windows could not safely replace the file.",
        "Inspect the candidate and original paths. Do not retry blindly.",
    ),
    RECOVERY_REQUIRED: (
        "The add-in was replaced but final verification failed.",
        "Use Restore Backup or inspect the listed artifacts.",
    ),
    RECOVERY_WRITE_FAILED: (
        "Draft recovery could not be written.",
        "Use Save Recovery Now or free disk space. Editing stays enabled.",
    ),
    STALE_PROPOSAL: (
        "This review is out of date because the draft changed.",
        "Rebuild the review, then continue.",
    ),
    NO_CHANGES: ("There are no changes to save.", "Continue editing, or Save a Copy."),
    PASSWORD_PROTECTED: (
        "This VBA project is password-protected and cannot be modified safely.",
        "XML-only package edits remain possible when they do not touch VBA.",
    ),
    PACKAGE_SIGNED: (
        "This Office package has an OPC digital signature, so XML editing is blocked.",
        "Remove the package signature outside this editor, or edit VBA only.",
    ),
    UNSUPPORTED_EXTENSION: (
        "The destination must keep the original file type.",
        "Choose a destination with the same extension.",
    ),
    REPARSE_POINT: (
        "The chosen path is a link or reparse point, which this version does not write to.",
        "Choose a regular folder.",
    ),
    SAME_PATH: (
        "The destination is the same file as the original.",
        "Choose a different destination.",
    ),
    UNKNOWN: (
        "The operation failed for an unexpected reason.",
        "Copy the diagnostic report and inspect artifact paths.",
    ),
}


def utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def new_operation_id() -> str:
    return str(uuid.uuid4())


@dataclass
class OperationEnvelope:
    """Common result envelope carried by save/restore/copy/recovery results."""

    operation_id: str = field(default_factory=new_operation_id)
    operation_type: str = "save"
    stage: str | None = None
    reason: str | None = None
    utc_time: str = field(default_factory=utc_now)
    retryable: bool = False
    win32_error: int | None = None
    win32_name: str | None = None
    revision: int | None = None
    hashes: dict[str, str] = field(default_factory=dict)
    artifact_paths: dict[str, str] = field(default_factory=dict)
    artifact_existence: dict[str, bool] = field(default_factory=dict)
    details: dict[str, Any] = field(default_factory=dict)

    def as_dict(self) -> dict[str, Any]:
        return {
            "operation_id": self.operation_id,
            "operation_type": self.operation_type,
            "stage": self.stage,
            "reason": self.reason,
            "utc_time": self.utc_time,
            "retryable": self.retryable,
            "win32_error": self.win32_error,
            "win32_name": self.win32_name,
            "revision": self.revision,
            "hashes": dict(self.hashes),
            "artifact_paths": dict(self.artifact_paths),
            "artifact_existence": dict(self.artifact_existence),
        }


def explanation_for(reason: str | None) -> tuple[str, str]:
    if not reason:
        return REASON_EXPLANATIONS[UNKNOWN]
    return REASON_EXPLANATIONS.get(reason, REASON_EXPLANATIONS[UNKNOWN])


def is_retryable(reason: str | None) -> bool:
    return reason in RETRYABLE_REASONS


def redact_path(path: str | Path | None) -> str | None:
    if path is None:
        return None
    text = str(path)
    name = Path(text).name
    parent = Path(text).parent.name
    return f"<redacted>/{parent}/{name}" if parent else f"<redacted>/{name}"


def sanitize_text(value: str, *, sentinels: tuple[str, ...] = ()) -> str:
    """Drop exception strings that may contain user source; keep reason tokens."""
    lowered = value.lower()
    if any(token.lower() in lowered for token in sentinels if token):
        return "<redacted user text>"
    return value


def build_diagnostic_report(
    *,
    envelope: OperationEnvelope,
    extension: str | None = None,
    include_full_paths: bool = False,
    extra: dict[str, Any] | None = None,
) -> str:
    """Machine-oriented report with no VBA/XML payloads."""
    identity = build_identity()
    lines = [
        f"{APP_NAME} diagnostic report",
        f"generated_utc: {utc_now()}",
        f"app_version: {identity.get('version', VERSION)}",
        f"source_commit: {identity.get('source_commit', 'unbuilt')}",
        f"dirty_tree: {identity.get('dirty_tree')}",
        f"build_time_utc: {identity.get('build_time_utc')}",
        f"packaged_mode: {identity.get('packaged_mode', 'development')}",
        f"python: {sys.version.split()[0]} {platform.architecture()[0]}",
        f"windows: {platform.version()}",
        f"pyopenvba_pin: {PYOPENVBA_PIN}",
        f"pyopenvba_installed: {identity.get('pyopenvba_installed')}",
        f"extension: {extension or ''}",
        f"operation_id: {envelope.operation_id}",
        f"operation_type: {envelope.operation_type}",
        f"stage: {envelope.stage or ''}",
        f"reason: {envelope.reason or ''}",
        f"retryable: {envelope.retryable}",
        f"win32_error: {envelope.win32_error if envelope.win32_error is not None else ''}",
        f"win32_name: {envelope.win32_name or ''}",
        f"revision: {envelope.revision if envelope.revision is not None else ''}",
    ]
    for key, digest in sorted(envelope.hashes.items()):
        lines.append(f"hash.{key}: {digest}")
    for key, raw in sorted(envelope.artifact_paths.items()):
        shown = raw if include_full_paths else redact_path(raw)
        exists = envelope.artifact_existence.get(key)
        lines.append(f"artifact.{key}: {shown} exists={exists}")
    if extra:
        for key, value in extra.items():
            if key in {"vba", "xml", "source", "body", "exception"}:
                continue
            lines.append(f"{key}: {value}")
    lines.append(f"process: {os.getpid()}")
    return "\n".join(lines) + "\n"
