"""Open workflow: preflight, snapshot, fingerprint (plan 7.3)."""

from __future__ import annotations

from pathlib import Path

from vba_addin_editor.adapters.pyopenvba_adapter import (
    SUPPORTED_EXTENSIONS,
    AdapterError,
    PyOpenVBAAdapter,
)
from vba_addin_editor.domain.document import DocumentDraft, DocumentSnapshot, draft_from_snapshot
from vba_addin_editor.platform import paths
from vba_addin_editor.platform import windows_processes as wp


class DocumentService:
    def __init__(self, adapter: PyOpenVBAAdapter | None = None) -> None:
        self.adapter = adapter or PyOpenVBAAdapter()

    def open(self, path: Path) -> DocumentDraft:
        path = Path(path).resolve()
        if not path.is_file():
            raise AdapterError("The selected add-in file could not be found.")
        if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            raise AdapterError("Only .xlam and .ppam add-ins are supported.")
        fp = paths.fingerprint(path)
        snapshot = self.adapter.open_snapshot(
            path,
            fp,
            host_process_running=wp.host_process_running(path),
        )
        return draft_from_snapshot(snapshot)

    def reload(self, draft: DocumentDraft) -> DocumentDraft:
        return self.open(draft.baseline.path)

    def is_current_on_disk(self, draft: DocumentDraft) -> bool:
        return paths.fingerprint_matches(draft.baseline.path, draft.baseline.file_fingerprint)


def snapshot_report(snapshot: DocumentSnapshot) -> str:
    """Human-readable one-line status used by banners."""
    bits = [f"{snapshot.extension} ", snapshot.project_name or "(unnamed project)"]
    flags: list[str] = []
    if snapshot.safety.password_protected:
        flags.append("password-protected")
    if snapshot.safety.signature_present:
        flags.append("digitally signed")
    if snapshot.safety.host_process_running:
        flags.append("Office is running")
    return "".join(bits) + (" — " + ", ".join(flags) if flags else "")
