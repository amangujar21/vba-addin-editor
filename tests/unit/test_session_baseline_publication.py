from __future__ import annotations

import os
from pathlib import Path

from helpers import edit_module1, make_service

from vba_addin_editor.services.recovery_service import RecoveryService
from vba_addin_editor.services.session_service import SessionService


def _open(work_xlam: Path, tmp_path: Path):
    return SessionService(session_root=tmp_path / "sessions").open(work_xlam)


def test_save_updates_revision_and_captured_hash(work_xlam: Path, tmp_path: Path):
    session = _open(work_xlam, tmp_path)
    old_rev = session.revision
    old_hash = session.baseline_sha256
    edit_module1(session.draft)
    result = make_service().save_addin(session.draft)
    assert result.kind == "success", result
    pub = SessionService(session_root=tmp_path / "sessions").publish_verified_baseline(session, work_xlam)
    assert pub is None, pub
    assert session.revision > old_rev
    assert session.baseline_sha256 != old_hash
    assert session.captured_path.read_bytes() == work_xlam.read_bytes()


def test_pre_save_proposal_is_rejected_after_save(work_xlam: Path, tmp_path: Path):
    session = _open(work_xlam, tmp_path)
    edit_module1(session.draft)
    session.bump()
    binding = session.binding()
    result = make_service().save_addin(session.draft)
    assert result.kind == "success", result
    SessionService(session_root=tmp_path / "sessions").publish_verified_baseline(session, work_xlam)
    assert not session.matches_proposal(binding)
    assert binding.session_id == session.session_id


def test_clean_save_copy_after_inplace_save_matches_new_baseline(work_xlam: Path, tmp_path: Path):
    sessions = SessionService(session_root=tmp_path / "sessions")
    session = sessions.open(work_xlam)
    edit_module1(session.draft)
    assert make_service().save_addin(session.draft).kind == "success"
    assert sessions.publish_verified_baseline(session, work_xlam) is None
    dest = tmp_path / "copy.xlam"
    result = make_service().save_copy(session.draft, dest, source_path=session.captured_path)
    assert result.kind == "success", result
    assert dest.read_bytes() == session.captured_path.read_bytes()
    assert dest.read_bytes() == work_xlam.read_bytes()


def test_save_then_edit_other_module_copy_keeps_both(work_xlam: Path, tmp_path: Path):
    sessions = SessionService(session_root=tmp_path / "sessions")
    session = sessions.open(work_xlam)
    edit_module1(session.draft)
    assert make_service().save_addin(session.draft).kind == "success"
    assert sessions.publish_verified_baseline(session, work_xlam) is None
    from vba_addin_editor.domain.document import ModuleDisplayKind, ModuleDraft, new_module_id

    session.draft.modules.append(
        ModuleDraft(
            id=new_module_id(),
            origin_name=None,
            current_name="SecondEdit",
            body="Public Sub Second()\r\nEnd Sub\r\n",
            kind=ModuleDisplayKind.STANDARD,
            pyopenvba_kind="standard",
            is_new=True,
            is_deleted=False,
            destructive_ops_safe=True,
            can_delete=True,
            can_rename=True,
            project_item_kind="standard",
        )
    )
    dest = tmp_path / "both.xlam"
    result = make_service().save_copy(session.draft, dest, source_path=session.captured_path)
    assert result.kind == "success", result.problems or result
    from vba_addin_editor.services.document_service import DocumentService

    copied = DocumentService().open(dest)
    m1 = next(m for m in copied.modules if m.current_name == "Module1")
    m2 = next(m for m in copied.modules if m.current_name == "SecondEdit")
    assert "V2" in m1.body
    assert "Second" in m2.body


def test_folder_export_after_save_uses_new_baseline_hash(work_xlam: Path, tmp_path: Path):
    import json

    from vba_addin_editor.services.folder_sync_service import MANIFEST_NAME, FolderSyncService

    sessions = SessionService(session_root=tmp_path / "sessions")
    session = sessions.open(work_xlam)
    edit_module1(session.draft)
    assert make_service().save_addin(session.draft).kind == "success"
    assert sessions.publish_verified_baseline(session, work_xlam) is None
    dest = tmp_path / "export"
    FolderSyncService().export_folder(session, dest)
    manifest = json.loads((dest / MANIFEST_NAME).read_text(encoding="utf-8"))
    assert manifest["baseline_sha256"] == session.baseline_sha256


def test_post_save_edit_is_recoverable_after_crash(work_xlam: Path, tmp_path: Path):
    root = tmp_path / "sessions"
    sessions = SessionService(session_root=root)
    recovery = RecoveryService(session_root=root)
    session = sessions.open(work_xlam)
    edit_module1(session.draft)
    assert make_service().save_addin(session.draft).kind == "success"
    assert sessions.publish_verified_baseline(session, work_xlam) is None
    recovery.mark_complete(session)
    sheet = next(m for m in session.draft.modules if m.current_name == "Sheet1")
    sheet.body = sheet.body + "\n' after-save\n"
    session.bump()
    assert recovery.checkpoint(session) is None
    assert not (session.session_dir / "complete.marker").exists()
    directory = session.session_dir
    sessions.close(session)
    listings = recovery.list_recoverable()
    assert any(item.directory == directory for item in listings)


def test_publication_failure_reports_committed_file(work_xlam: Path, tmp_path: Path, monkeypatch):
    sessions = SessionService(session_root=tmp_path / "sessions")
    session = sessions.open(work_xlam)
    edit_module1(session.draft)
    assert make_service().save_addin(session.draft).kind == "success"
    committed = work_xlam.read_bytes()
    real_replace = os.replace

    def boom(src, dst):
        dst_name = Path(dst).name
        if dst_name.startswith("package") or "pubtmp" in Path(src).name:
            raise OSError("injected publication failure")
        return real_replace(src, dst)

    monkeypatch.setattr(os, "replace", boom)
    pub = sessions.publish_verified_baseline(session, work_xlam)
    assert pub is not None
    assert pub.reason == "baseline_publication_failed"
    assert "committed" in (pub.details or {})
    assert work_xlam.read_bytes() == committed
    assert "unchanged" not in (pub.message or "").lower() or "not" in (pub.message or "").lower()
