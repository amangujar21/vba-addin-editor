from __future__ import annotations

from vba_addin_editor.domain.document import ModuleDisplayKind, ModuleDraft, new_module_id
from vba_addin_editor.platform import paths
from vba_addin_editor.platform import windows_file_ops as wfo
from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.services.save_service import SaveService


def _adapter():
    from vba_addin_editor.adapters.pyopenvba_adapter import PyOpenVBAAdapter

    return PyOpenVBAAdapter()


def make_service(**overrides):
    base = {
        "adapter": _adapter(),
        "process_probe": lambda p: False,
        "exclusive_probe": lambda p: None,
        "commit": None,
        "progress": None,
    }
    base.update(overrides)
    return SaveService(**base)


def edit_module1(draft) -> None:
    m1 = next(m for m in draft.modules if m.current_name == "Module1")
    m1.body = m1.body.replace("ORIGINAL", "V2")


def test_no_change_save_does_not_write(work_xlam):
    draft = DocumentService().open(work_xlam)
    before = paths.fingerprint(work_xlam)
    result = make_service().save_addin(draft)
    assert result.kind == "no_changes"
    assert paths.fingerprint(work_xlam) == before


def test_edit_save_replaces_in_place_with_backup(work_xlam):
    draft = DocumentService().open(work_xlam)
    before = work_xlam.read_bytes()
    edit_module1(draft)
    result = make_service().save_addin(draft)
    assert result.kind == "success", result
    assert result.backup_path is not None and result.backup_path.exists()
    assert result.backup_path.read_bytes() == before
    # Same path now holds the edit.
    reopened = DocumentService().open(work_xlam)
    m1 = next(m for m in reopened.modules if m.current_name == "Module1")
    assert "V2" in m1.body and "ORIGINAL" not in m1.body
    # No leftover candidates.
    assert not list(work_xlam.parent.glob(".*candidate*"))
    # Non-VBA entries preserved implicitly by verifier; assert clean draft state.
    assert not draft.is_dirty()


def test_save_blocked_when_office_running(work_xlam):
    draft = DocumentService().open(work_xlam)
    edit_module1(draft)
    before = work_xlam.read_bytes()
    result = make_service(process_probe=lambda p: True).save_addin(draft)
    assert result.kind == "blocked" and result.reason == "office_running"
    assert work_xlam.read_bytes() == before


def test_save_blocked_when_locked(work_xlam):
    draft = DocumentService().open(work_xlam)
    edit_module1(draft)
    result = make_service(exclusive_probe=lambda p: "in use").save_addin(draft)
    assert result.kind == "blocked" and result.reason == "locked"


def test_save_blocked_on_external_change(work_xlam):
    draft = DocumentService().open(work_xlam)
    edit_module1(draft)
    work_xlam.write_bytes(work_xlam.read_bytes() + b"\x00")
    result = make_service().save_addin(draft)
    assert result.kind == "blocked" and result.reason == "external_change"


def test_candidate_failure_leaves_original_untouched(work_xlam):
    draft = DocumentService().open(work_xlam)
    before = work_xlam.read_bytes()
    draft.modules.append(
        ModuleDraft(
            id=new_module_id(),
            origin_name=None,
            current_name="Bad Name!",  # adapter raises during replay
            body="Sub X()\r\nEnd Sub\r\n",
            kind=ModuleDisplayKind.STANDARD,
            pyopenvba_kind="standard",
            is_new=True,
            is_deleted=False,
            destructive_ops_safe=True,
        )
    )
    result = make_service().save_addin(draft)
    assert result.kind == "error"
    assert work_xlam.read_bytes() == before
    assert list(work_xlam.parent.glob(".*candidate*")) == []


def test_commit_failure_reports_error_and_keeps_original(work_xlam):
    def failing_commit(replaced, replacement, backup):
        raise wfo.FileOpsError("boom", win32_error=5)

    draft = DocumentService().open(work_xlam)
    edit_module1(draft)
    before = work_xlam.read_bytes()
    result = make_service(commit=failing_commit).save_addin(draft)
    assert result.kind == "error" and result.reason == "commit_failed"
    assert work_xlam.read_bytes() == before
    assert result.backup_path is None


def test_add_rename_delete_replay(work_xlam):
    draft = DocumentService().open(work_xlam)
    edit_module1(draft)
    m1 = next(m for m in draft.modules if m.current_name == "Module1")
    m1.current_name = "Renamed1"
    new = ModuleDraft(
        id=new_module_id(),
        origin_name=None,
        current_name="AddedModule",
        body="Public Sub FromNew()\r\nEnd Sub\r\n",
        kind=ModuleDisplayKind.STANDARD,
        pyopenvba_kind="standard",
        is_new=True,
        is_deleted=False,
        destructive_ops_safe=True,
    )
    draft.modules.append(new)
    tw = next(m for m in draft.modules if m.current_name == "Sheet1")
    tw.is_deleted = True
    result = make_service().save_addin(draft)
    assert result.kind == "success", result
    reopened = DocumentService().open(work_xlam)
    names = {m.current_name for m in reopened.modules}
    assert "Renamed1" in names and "AddedModule" in names and "Sheet1" not in names
    assert "Module1" not in names
    m = next(m for m in reopened.modules if m.current_name == "Renamed1")
    assert "V2" in m.body


def test_rename_chain_a_b_c(work_xlam):
    draft = DocumentService().open(work_xlam)
    m1 = next(m for m in draft.modules if m.current_name == "Module1")
    m1.current_name = "StepB"
    # Now rename again in the same session: StepB -> StepC (pending chain).
    m1.current_name = "StepC"
    result = make_service().save_addin(draft)
    assert result.kind == "success", result
    reopened = DocumentService().open(work_xlam)
    assert {m.current_name for m in reopened.modules} >= {"StepC"}
    assert "Module1" not in {m.current_name for m in reopened.modules}


def test_delete_then_readd_same_name(work_xlam):
    draft = DocumentService().open(work_xlam)
    m1 = next(m for m in draft.modules if m.current_name == "Module1")
    m1.is_deleted = True
    draft.modules.append(
        ModuleDraft(
            id=new_module_id(),
            origin_name=None,
            current_name="Module1",
            body="Public Const TEST_BUILD As String = \"REBORN\"\r\n",
            kind=ModuleDisplayKind.STANDARD,
            pyopenvba_kind="standard",
            is_new=True,
            is_deleted=False,
            destructive_ops_safe=True,
        )
    )
    result = make_service().save_addin(draft)
    assert result.kind == "success", result
    reopened = DocumentService().open(work_xlam)
    m = next(m for m in reopened.modules if m.current_name == "Module1")
    assert "REBORN" in m.body


def test_signature_confirmation_required(work_xlam):
    draft = DocumentService().open(work_xlam)
    edit_module1(draft)
    from dataclasses import replace

    draft.baseline = replace(
        draft.baseline,
        safety=replace(draft.baseline.safety, signature_present=True, signature_kinds=("legacy",)),
    )
    edit_module1(draft)
    result = make_service().save_addin(draft)
    assert result.kind == "needs_signature_confirmation"
    draft.signed_save_confirmed = True
    result = make_service().save_addin(draft)
    assert result.kind == "success", result


def test_save_copy_does_not_touch_original(work_xlam, tmp_path):
    draft = DocumentService().open(work_xlam)
    edit_module1(draft)
    dest = tmp_path / "copy.xlam"
    result = make_service().save_copy(draft, dest)
    assert result.kind == "success", result
    reopened = DocumentService().open(dest)
    m1 = next(m for m in reopened.modules if m.current_name == "Module1")
    assert "V2" in m1.body
