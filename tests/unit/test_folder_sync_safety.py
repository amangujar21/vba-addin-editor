from __future__ import annotations

import json
from pathlib import Path

from vba_addin_editor.adapters.pyopenvba_adapter import AdapterError
from vba_addin_editor.services.folder_sync_service import MANIFEST_NAME, FolderSyncService
from vba_addin_editor.services.history_service import HistoryService
from vba_addin_editor.services.session_service import SessionService


def test_export_refuses_externally_changed_managed_file(work_xlam_with_class: Path, tmp_path: Path):
    session = SessionService(session_root=tmp_path / "sessions").open(work_xlam_with_class)
    dest = tmp_path / "export"
    FolderSyncService().export_folder(session, dest)
    managed = next(dest.glob("*.bas"))
    managed.write_text(managed.read_text(encoding="utf-8") + "\n' external\n", encoding="utf-8")
    result = FolderSyncService().export_folder_result(session, dest)
    assert not result.ok
    assert result.conflicts


def test_malicious_manifest_is_structured_problem(work_xlam: Path, tmp_path: Path):
    session = SessionService(session_root=tmp_path / "sessions").open(work_xlam)
    dest = tmp_path / "evil"
    dest.mkdir()
    (dest / MANIFEST_NAME).write_text(
        json.dumps(
            {
                "schema_version": 1,
                "baseline_sha256": session.baseline_sha256,
                "entries": [
                    {
                        "export_id": "abc",
                        "logical_name": "Evil",
                        "kind": "standard",
                        "relative_path": "../escape.bas",
                        "module_id": "x",
                    }
                ],
            }
        ),
        encoding="utf-8",
    )
    preview = FolderSyncService().preview(session, dest)
    assert preview.problems
    assert all("Traceback" not in item for item in preview.problems)


def test_apply_failure_leaves_draft_unchanged(work_xlam: Path, tmp_path: Path):
    from vba_addin_editor.services.folder_sync_service import FolderChange

    session = SessionService(session_root=tmp_path / "sessions").open(work_xlam)
    dest = tmp_path / "export"
    FolderSyncService().export_folder(session, dest)
    preview = FolderSyncService().preview(session, dest)
    preview.changes.append(
        FolderChange(
            export_id="newbad",
            operation="add",
            logical_name="1bad",
            selected=True,
            folder_text="Public Sub X()\nEnd Sub\n",
            kind="standard",
        )
    )
    before = [(m.current_name, m.body, m.is_deleted) for m in session.draft.modules]
    rev = session.revision
    try:
        FolderSyncService().apply(session, preview)
        raise AssertionError("expected validation failure")
    except AdapterError:
        pass
    after = [(m.current_name, m.body, m.is_deleted) for m in session.draft.modules]
    assert after == before
    assert session.revision == rev


def test_apply_batch_is_undoable(work_xlam: Path, tmp_path: Path):
    session = SessionService(session_root=tmp_path / "sessions")
    opened = session.open(work_xlam)
    dest = tmp_path / "export"
    FolderSyncService().export_folder(opened, dest)
    module_file = next(dest.glob("Module1*.bas"))
    text = module_file.read_text(encoding="utf-8").replace("ORIGINAL", "FOLDER")
    module_file.write_bytes(text.encode("utf-8"))
    preview = FolderSyncService().preview(opened, dest)
    edited = [c for c in preview.changes if c.operation == "edit"]
    assert edited
    history = HistoryService()
    FolderSyncService().apply(opened, preview, history_service=history)
    m1 = next(m for m in opened.draft.modules if m.current_name == "Module1")
    assert "FOLDER" in m1.body
    history.undo(opened)
    m1 = next(m for m in opened.draft.modules if m.current_name == "Module1")
    assert "ORIGINAL" in m1.body
