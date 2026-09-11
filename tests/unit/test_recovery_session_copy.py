from __future__ import annotations

from pathlib import Path

from helpers import edit_module1, make_service

from vba_addin_editor.domain.history import HistoryCommand
from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.services.history_service import HistoryService
from vba_addin_editor.services.recovery_service import RecoveryService
from vba_addin_editor.services.session_service import SessionService


def test_session_open_captures_baseline_without_writing_original(work_xlam: Path, tmp_path: Path):
    before = work_xlam.read_bytes()
    service = SessionService(session_root=tmp_path / "sessions")
    session = service.open(work_xlam)
    assert session.captured_path.exists()
    assert session.captured_path.read_bytes() == before
    assert work_xlam.read_bytes() == before
    assert session.draft.baseline.path == work_xlam.resolve()
    service.close(session)


def test_recovery_roundtrip(work_xlam: Path, tmp_path: Path):
    sessions = tmp_path / "sessions"
    session = SessionService(session_root=sessions).open(work_xlam)
    edit_module1(session.draft)
    session.bump()
    recovery = RecoveryService(session_root=sessions)
    assert recovery.checkpoint(session) is None
    data = recovery.load_checkpoint(session.session_dir)
    session.draft.modules[0].body = "wiped"
    recovery.apply_checkpoint(session, data)
    recovery.revalidate_capabilities(session.draft)
    m1 = next(m for m in session.draft.modules if m.current_name == "Module1")
    assert "V2" in m1.body
    assert work_xlam.read_bytes() == session.captured_path.read_bytes()


def test_recovery_rejects_unknown_schema(work_xlam: Path, tmp_path: Path):
    sessions = tmp_path / "sessions"
    session = SessionService(session_root=sessions).open(work_xlam)
    recovery = RecoveryService(session_root=sessions)
    recovery.checkpoint(session)
    gen = (session.session_dir / "current.json")
    import json

    pointer = json.loads(gen.read_text(encoding="utf-8"))
    checkpoint = session.session_dir / pointer["generation"] / "checkpoint.json"
    payload = json.loads(checkpoint.read_text(encoding="utf-8"))
    payload["schema_version"] = 99
    checkpoint.write_text(json.dumps(payload), encoding="utf-8")
    try:
        recovery.load_checkpoint(session.session_dir)
        raise AssertionError("expected unknown schema")
    except ValueError:
        pass


def test_clean_save_copy_is_not_no_changes(work_xlam: Path, tmp_path: Path):
    draft = DocumentService().open(work_xlam)
    dest = tmp_path / "copy.xlam"
    result = make_service().save_copy(draft, dest)
    assert result.kind == "success", result
    assert dest.exists()
    assert dest.read_bytes() == work_xlam.read_bytes()
    assert work_xlam.read_bytes() == dest.read_bytes()


def test_save_copy_rejects_same_path(work_xlam: Path):
    draft = DocumentService().open(work_xlam)
    edit_module1(draft)
    result = make_service().save_copy(draft, work_xlam)
    assert result.kind == "blocked"
    assert result.reason == "same_path"


def test_undo_text_and_rename(work_xlam: Path, tmp_path: Path):
    session = SessionService(session_root=tmp_path / "sessions").open(work_xlam)
    history = HistoryService()
    m1 = next(m for m in session.draft.modules if m.current_name == "Module1")
    history.record_text(session, target_id=m1.id, before=m1.body, after=m1.body + "X", coalesce=False)
    m1.body = m1.body + "X"
    history.record_structural(
        session, HistoryCommand(op="rename", target_id=m1.id, before="Module1", after="Renamed")
    )
    m1.current_name = "Renamed"
    history.undo(session)
    assert m1.current_name == "Module1"
    history.undo(session)
    assert m1.body.endswith("X") is False or not m1.body.endswith("X")
