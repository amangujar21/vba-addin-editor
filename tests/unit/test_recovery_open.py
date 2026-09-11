from __future__ import annotations

import json
from pathlib import Path

from helpers import edit_module1, make_service

from vba_addin_editor.services.recovery_service import RecoveryOpenService, RecoveryService
from vba_addin_editor.services.session_service import SessionService


def _prepare_dirty(work_xlam: Path, tmp_path: Path):
    root = tmp_path / "sessions"
    sessions = SessionService(session_root=root)
    recovery = RecoveryService(session_root=root)
    session = sessions.open(work_xlam)
    edit_module1(session.draft)
    session.bump()
    assert recovery.checkpoint(session) is None
    directory = session.session_dir
    sessions.close(session)
    return root, sessions, recovery, directory


def test_recover_when_original_unchanged(work_xlam: Path, tmp_path: Path):
    _root, sessions, recovery, directory = _prepare_dirty(work_xlam, tmp_path)
    opened = RecoveryOpenService(recovery, sessions).open_recoverable(directory)
    assert opened.status == "clean"
    assert opened.session is not None
    m1 = next(m for m in opened.session.draft.modules if m.current_name == "Module1")
    assert "V2" in m1.body
    assert opened.session.original_exists
    assert not opened.session.conflicts_pending


def test_recover_when_original_changed_does_not_rebase_onto_disk(work_xlam: Path, tmp_path: Path):
    _root, sessions, recovery, directory = _prepare_dirty(work_xlam, tmp_path)
    other = __import__("vba_addin_editor.services.document_service", fromlist=["DocumentService"]).DocumentService().open(work_xlam)
    m1 = next(m for m in other.modules if m.current_name == "Module1")
    m1.body = m1.body.replace("ORIGINAL", "DISK")
    assert make_service().save_addin(other).kind == "success"
    opened = RecoveryOpenService(recovery, sessions).open_recoverable(directory)
    assert opened.status == "needs_conflict"
    assert opened.session is not None
    assert opened.session.conflicts_pending
    recovered = next(m for m in opened.session.draft.modules if m.current_name == "Module1")
    assert "V2" in recovered.body
    assert "DISK" not in recovered.body
    assert opened.session.captured_path.read_bytes() != work_xlam.read_bytes()


def test_recover_when_original_missing(work_xlam: Path, tmp_path: Path):
    _root, sessions, recovery, directory = _prepare_dirty(work_xlam, tmp_path)
    work_xlam.unlink()
    opened = RecoveryOpenService(recovery, sessions).open_recoverable(directory)
    assert opened.status == "missing_source"
    assert opened.session is not None
    assert opened.session.original_exists is False
    dest = tmp_path / "recovered.xlam"
    result = make_service().save_copy(
        opened.session.draft,
        dest,
        source_path=opened.session.captured_path,
        recovered_copy=True,
        session_dir=opened.session.session_dir,
    )
    assert result.kind == "success", result
    assert dest.exists()
    assert not work_xlam.exists()


def test_corrupt_newest_generation_falls_back(work_xlam: Path, tmp_path: Path):
    root = tmp_path / "sessions"
    sessions = SessionService(session_root=root)
    recovery = RecoveryService(session_root=root)
    session = sessions.open(work_xlam)
    edit_module1(session.draft)
    session.bump()
    assert recovery.checkpoint(session) is None
    first = json.loads((session.session_dir / "current.json").read_text(encoding="utf-8"))["generation"]
    sheet = next(m for m in session.draft.modules if m.current_name == "Sheet1")
    sheet.body = sheet.body + "\n' gen2\n"
    session.bump()
    assert recovery.checkpoint(session) is None
    second = json.loads((session.session_dir / "current.json").read_text(encoding="utf-8"))["generation"]
    assert second != first
    (session.session_dir / second / "checkpoint.json").write_text("{not-json", encoding="utf-8")
    directory = session.session_dir
    sessions.close(session)
    opened = RecoveryOpenService(recovery, sessions).open_recoverable(directory)
    assert opened.status == "clean"
    assert opened.used_fallback
    assert opened.session is not None
    m1 = next(m for m in opened.session.draft.modules if m.current_name == "Module1")
    assert "V2" in m1.body
    sheet = next(m for m in opened.session.draft.modules if m.current_name == "Sheet1")
    assert "gen2" not in sheet.body


def test_rejects_unknown_schema_and_duplicate_ids(work_xlam: Path, tmp_path: Path):
    _root, sessions, recovery, directory = _prepare_dirty(work_xlam, tmp_path)
    pointer = json.loads((directory / "current.json").read_text(encoding="utf-8"))
    path = directory / pointer["generation"] / "checkpoint.json"
    payload = json.loads(path.read_text(encoding="utf-8"))
    payload["schema_version"] = 99
    path.write_text(json.dumps(payload), encoding="utf-8")
    opened = RecoveryOpenService(recovery, sessions).open_recoverable(directory)
    assert opened.status == "invalid"

    root2 = tmp_path / "sessions2"
    sessions2 = SessionService(session_root=root2)
    recovery2 = RecoveryService(session_root=root2)
    session = sessions2.open(work_xlam)
    recovery2.checkpoint(session)
    pointer = json.loads((session.session_dir / "current.json").read_text(encoding="utf-8"))
    path = session.session_dir / pointer["generation"] / "checkpoint.json"
    payload = json.loads(path.read_text(encoding="utf-8"))
    payload["draft"]["modules"].append(dict(payload["draft"]["modules"][0]))
    path.write_text(json.dumps(payload), encoding="utf-8")
    sessions2.close(session)
    opened = RecoveryOpenService(recovery2, sessions2).open_recoverable(session.session_dir)
    assert opened.status == "invalid"


def test_mismatched_captured_package_is_invalid(work_xlam: Path, tmp_path: Path):
    _root, sessions, recovery, directory = _prepare_dirty(work_xlam, tmp_path)
    captured = next((directory / "baseline").glob("package.*"))
    captured.write_bytes(captured.read_bytes() + b"tamper")
    opened = RecoveryOpenService(recovery, sessions).open_recoverable(directory)
    assert opened.status == "invalid"
    assert opened.session is None
