"""Open/capture sessions and apply draft mutations through one boundary."""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path

from vba_addin_editor.adapters.ooxml_package_adapter import (
    OoxmlPackageAdapter,
)
from vba_addin_editor.adapters.pyopenvba_adapter import AdapterError, PyOpenVBAAdapter
from vba_addin_editor.domain.document import DocumentDraft, draft_from_snapshot
from vba_addin_editor.domain.results import SaveResult
from vba_addin_editor.domain.session import DocumentSession, ProposalBinding, new_session_id
from vba_addin_editor.platform import paths
from vba_addin_editor.platform.session_store import (
    SessionLock,
    SessionStoreError,
    acquire_lock,
    baseline_package_path,
    create_session_dir,
    read_json,
    session_dir_for,
    write_bytes_atomic,
    write_json_atomic,
)
from vba_addin_editor.services.document_service import DocumentService


def _publication_failure(
    message: str,
    *,
    committed: Path,
    captured: Path,
    session_dir: Path,
    exc: BaseException | None = None,
    extra: dict | None = None,
) -> SaveResult:
    details = {
        "committed": str(committed),
        "captured": str(captured),
        "session_dir": str(session_dir),
        "committed_exists": committed.exists(),
        "captured_exists": captured.exists(),
    }
    if extra:
        details.update(extra)
    if exc is not None:
        details["exception_type"] = type(exc).__name__
    return SaveResult.error(
        message,
        reason="baseline_publication_failed",
        stage="baseline_publish",
        details=details,
        operation_type="session",
        retryable=False,
    )


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
        # One stable read serves the fingerprint, the captured copy and the
        # parse. The captured copy is verified as it is written, never reopened.
        data, original_fp = paths.read_stable(path)
        write_bytes_atomic(captured, data)
        captured_fp = replace(original_fp, mtime_ns=captured.stat().st_mtime_ns)
        snapshot = self._snapshot_from_captured(data, captured_fp, path)
        draft = draft_from_snapshot(snapshot)
        self._write_session_meta(
            directory,
            session_id=session_id,
            source_path=path,
            fingerprint=captured_fp,
            extension=path.suffix.lower(),
            revision=1,
            baseline_generation=1,
        )
        return DocumentSession(
            session_id=session_id,
            draft=draft,
            captured_path=captured,
            captured_fingerprint=captured_fp,
            session_dir=directory,
            revision=1,
            original_exists=True,
            baseline_generation=1,
        )

    def attach_captured(
        self,
        *,
        session_id: str,
        directory: Path,
        source_path: Path,
        extension: str,
        revision: int,
        original_exists: bool,
        baseline_generation: int = 1,
    ) -> DocumentSession:
        """Reopen a session from its captured baseline package, not current disk bytes."""
        directory = Path(directory)
        lock = acquire_lock(directory)
        self._locks[session_id] = lock
        captured = baseline_package_path(directory, extension)
        data, captured_fp = paths.read_stable(captured)
        snapshot = self._snapshot_from_captured(data, captured_fp, Path(source_path))
        draft = draft_from_snapshot(snapshot)
        return DocumentSession(
            session_id=session_id,
            draft=draft,
            captured_path=captured,
            captured_fingerprint=captured_fp,
            session_dir=directory,
            revision=revision,
            original_exists=original_exists,
            baseline_generation=baseline_generation,
        )

    def _snapshot_from_captured(self, data: bytes, captured_fp, user_path):
        """Snapshot of the captured baseline bytes, addressed by the user's path."""
        snapshot = self.adapter.open_snapshot(
            Path(user_path), captured_fp, host_process_running=False, data=data
        )
        return self.document_service._with_package_state(data, snapshot)

    def close(self, session: DocumentSession) -> None:
        lock = self._locks.pop(session.session_id, None)
        if lock is not None:
            lock.release()

    def apply(self, session: DocumentSession, mutator) -> int:
        """Apply one mutation or atomic batch, then increment revision."""
        mutator(session.draft)
        return session.bump()

    def replace_baseline(self, session: DocumentSession, draft: DocumentDraft) -> int:
        """In-memory baseline swap only. Prefer publish_verified_baseline."""
        session.draft = draft
        session.captured_fingerprint = draft.baseline.file_fingerprint
        session.history.clear()
        session.draft.signed_save_confirmed = False
        return session.bump()

    def publish_verified_baseline(
        self,
        session: DocumentSession,
        source_path: Path,
        data: bytes | None = None,
    ) -> SaveResult | None:
        """Atomically replace the captured package from a verified source of truth.

        ``data`` is the committed file's bytes when the caller already holds
        them (the save pipeline does); otherwise the committed file is read.

        On failure after the add-in was already committed, return a structured
        error that names the committed file and session artifacts. Never describe
        the original as unchanged.
        """
        committed = Path(source_path)
        captured = session.captured_path
        fail_msg = (
            "The add-in file was written, but the editor could not publish a new "
            "session baseline. The committed file is the current source of truth. "
            "Do not treat the original as unchanged. Use the committed path and "
            "session recovery artifacts below."
        )
        try:
            try:
                if data is None:
                    data, source_fp = paths.read_stable(committed)
                else:
                    source_fp = paths.written_fingerprint(data, committed)
            except OSError as exc:
                return _publication_failure(
                    fail_msg, committed=committed, captured=captured,
                    session_dir=session.session_dir, exc=exc,
                )
            try:
                write_bytes_atomic(captured, data)
            except SessionStoreError:
                return _publication_failure(
                    fail_msg, committed=committed, captured=captured,
                    session_dir=session.session_dir,
                    extra={"reason": "temp_copy_mismatch"},
                )
            stat = captured.stat()
            captured_fp = replace(
                source_fp, size=stat.st_size, mtime_ns=stat.st_mtime_ns, written_identity=None
            )
            if not captured_fp.content_equal(source_fp):
                return _publication_failure(
                    fail_msg, committed=committed, captured=captured,
                    session_dir=session.session_dir,
                    extra={"reason": "captured_mismatch"},
                )
            next_revision = session.revision + 1
            next_generation = session.baseline_generation + 1
            self._write_session_meta(
                session.session_dir,
                session_id=session.session_id,
                source_path=session.original_path,
                fingerprint=captured_fp,
                extension=session.draft.baseline.extension,
                revision=next_revision,
                baseline_generation=next_generation,
            )
        except OSError as exc:
            return _publication_failure(
                fail_msg, committed=committed, captured=captured,
                session_dir=session.session_dir, exc=exc,
            )

        session.captured_path = captured
        session.captured_fingerprint = captured_fp
        if session.draft.baseline.file_fingerprint.sha256 != captured_fp.sha256:
            session.draft.baseline = replace(
                session.draft.baseline, file_fingerprint=captured_fp
            )
        session.history.clear()
        session.draft.signed_save_confirmed = False
        session.baseline_generation = next_generation
        session.conflicts_pending = False
        session.bump()
        return None

    def require_binding(self, session: DocumentSession, binding: ProposalBinding) -> SaveResult | None:
        if session.matches_proposal(binding):
            return None
        return SaveResult.blocked(
            "stale_proposal",
            "This review is out of date because the draft changed. Rebuild it and try again.",
        )

    def session_dir(self, session_id: str) -> Path:
        return session_dir_for(session_id, root=self.session_root)

    def _write_session_meta(
        self,
        directory: Path,
        *,
        session_id: str,
        source_path: Path,
        fingerprint,
        extension: str,
        revision: int,
        baseline_generation: int,
    ) -> None:
        write_json_atomic(
            directory / "session.json",
            {
                "schema_version": 1,
                "session_id": session_id,
                "source_path": str(source_path),
                "baseline_sha256": fingerprint.sha256,
                "baseline_size": fingerprint.size,
                "extension": extension,
                "revision": revision,
                "baseline_generation": baseline_generation,
            },
        )

    def read_meta(self, directory: Path) -> dict:
        return read_json(Path(directory) / "session.json")
