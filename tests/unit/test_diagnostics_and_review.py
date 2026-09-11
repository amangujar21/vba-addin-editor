from __future__ import annotations

from pathlib import Path

from helpers import edit_module1

from vba_addin_editor.domain.results import SaveResult
from vba_addin_editor.services.diagnostics_service import actions_for, report_for
from vba_addin_editor.services.review_service import ReviewService
from vba_addin_editor.services.session_service import SessionService


def test_reasons_map_to_retry_and_compare():
    blocked = SaveResult.blocked("office_running", stage="preflight")
    assert blocked.retryable
    assert "Retry" in actions_for(blocked)
    changed = SaveResult.blocked("external_change")
    assert "Compare Changes" in actions_for(changed)
    assert not changed.retryable


def test_diagnostic_report_redacts_paths_and_source(tmp_path: Path):
    secret = "SECRET_VBA_BODY"
    result = SaveResult.error(
        "failed",
        reason="unknown",
        details={"saved": str(tmp_path / "AddIn.xlam")},
    )
    result.details["hash_baseline"] = "abc"
    report = report_for(result, include_full_paths=False)
    assert "SECRET_VBA_BODY" not in report
    assert str(tmp_path) not in report or "<redacted>" in report
    full = report_for(result, include_full_paths=True)
    assert "AddIn.xlam" in full
    assert secret not in full


def test_review_is_revision_bound(work_xlam: Path, tmp_path: Path):
    session = SessionService(session_root=tmp_path / "sessions").open(work_xlam)
    edit_module1(session.draft)
    session.bump()
    model = ReviewService().build(session)
    assert model.binding.revision == session.revision
    session.bump()
    assert not session.matches_proposal(model.binding)


def test_folder_export_then_preview_noop(work_xlam_with_class: Path, tmp_path: Path):
    from vba_addin_editor.services.folder_sync_service import FolderSyncService

    session = SessionService(session_root=tmp_path / "sessions").open(work_xlam_with_class)
    dest = tmp_path / "export"
    FolderSyncService().export_folder(session, dest)
    preview = FolderSyncService().preview(session, dest)
    assert preview.problems == []
    assert all(change.operation in {"noop"} for change in preview.changes)
