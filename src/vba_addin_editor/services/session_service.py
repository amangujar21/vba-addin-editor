"""Open/capture sessions and apply draft mutations through one boundary."""

from __future__ import annotations

import shutil
from dataclasses import replace
from pathlib import Path

from vba_addin_editor.adapters.ooxml_package_adapter import (
    XML_EDITABLE_EXTENSIONS,
    OoxmlPackageAdapter,
)
from vba_addin_editor.adapters.pyopenvba_adapter import AdapterError, PyOpenVBAAdapter
from vba_addin_editor.domain.document import DocumentDraft, draft_from_snapshot
from vba_addin_editor.domain.results import SaveResult
from vba_addin_editor.domain.session import DocumentSession, ProposalBinding, new_session_id
from vba_addin_editor.platform import paths
from vba_addin_editor.platform.session_store import (
    SessionLock,
    acquire_lock,
    baseline_package_path,
    create_session_dir,
    session_dir_for,
    write_json_atomic,
)
from vba_addin_editor.services.document_service import DocumentService


class SessionService:
    def __init__(
        self,
        *,
        document_service: DocumentService | None = None,
        adapter: PyOpenVBAAdapter | None = None,
        package_adapter: OoxmlPackageAdapter | None = None,
        session_root: Path | None = None,
    ) -> None:
        self.document_service = document_service or DocumentService(
            adapter=adapter, package_adapter=package_adapter
        )
        self.adapter = self.document_service.adapter
        self.package_adapter = self.document_service.package_adapter
        self.session_root = session_root
        self._locks: dict[str, SessionLock] = {}

    def open(self, path: Path) -> DocumentSession:
        path = Path(path).resolve()
        last_error: Exception | None = None
        for _ in range(3):
            try:
                return self._open_once(path)
            except paths.UnstableSourceError as exc:
                last_error = exc
        raise AdapterError(
            "The add-in file changed while it was being opened. Try again when it is idle.",
            {"reason": "source_unstable", "exception": repr(last_error)},
        )

    def _open_once(self, path: Path) -> DocumentSession:
        session_id = new_session_id()
        directory = create_session_dir(session_id, root=self.session_root)
        lock = acquire_lock(directory)
        self._locks[session_id] = lock
        captured = baseline_package_path(directory, path.suffix.lower())
        shutil.copy2(path, captured)
        captured_fp = paths.fingerprint(captured)
        original_fp = paths.fingerprint(path)
        if not captured_fp.content_equal(original_fp):
            raise paths.UnstableSourceError("Captured bytes do not match the original file.")
        snapshot = self.adapter.open_snapshot(captured, captured_fp, host_process_running=False)
        snapshot = replace(snapshot, path=path, file_fingerprint=captured_fp)
        if path.suffix.lower() in XML_EDITABLE_EXTENSIONS:
            snapshot = self.document_service._with_package_state(captured, snapshot)
            snapshot = replace(snapshot, path=path, file_fingerprint=captured_fp)
        draft = draft_from_snapshot(snapshot)
        write_json_atomic(
            directory / "session.json",
            {
                "schema_version": 1,
                "session_id": session_id,
                "source_path": str(path),
                "baseline_sha256": captured_fp.sha256,
                "baseline_size": captured_fp.size,
                "extension": path.suffix.lower(),
            },
        )
        return DocumentSession(
            session_id=session_id,
            draft=draft,
            captured_path=captured,
            captured_fingerprint=captured_fp,
            session_dir=directory,
            revision=1,
        )

    def close(self, session: DocumentSession) -> None:
        lock = self._locks.pop(session.session_id, None)
        if lock is not None:
            lock.release()

    def apply(self, session: DocumentSession, mutator) -> int:
        """Apply one mutation or atomic batch, then increment revision."""
        mutator(session.draft)
        return session.bump()

    def replace_baseline(self, session: DocumentSession, draft: DocumentDraft) -> int:
        session.draft = draft
        session.captured_fingerprint = draft.baseline.file_fingerprint
        session.history.clear()
        session.draft.signed_save_confirmed = False
        return session.bump()

    def require_binding(self, session: DocumentSession, binding: ProposalBinding) -> SaveResult | None:
        if session.matches_proposal(binding):
            return None
        return SaveResult.blocked(
            "stale_proposal",
            "This review is out of date because the draft changed. Rebuild it and try again.",
        )

    def session_dir(self, session_id: str) -> Path:
        return session_dir_for(session_id, root=self.session_root)
