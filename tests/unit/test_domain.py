from __future__ import annotations

from vba_addin_editor.domain.changes import compute_changes, dirty_count, revert_module
from vba_addin_editor.domain.document import revert_all
from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.services.validation_service import validate_draft, validate_module_name


def _draft(work_xlam):
    return DocumentService().open(work_xlam)


def test_clean_draft_is_not_dirty(work_xlam):
    draft = _draft(work_xlam)
    assert not draft.is_dirty()
    assert compute_changes(draft).is_empty
    assert dirty_count(draft) == 0


def test_body_edit_detected(work_xlam):
    draft = _draft(work_xlam)
    m1 = next(m for m in draft.modules if m.current_name == "Module1")
    m1.body = m1.body.replace("ORIGINAL", "V2")
    cs = compute_changes(draft)
    assert len(cs.modified) == 1 and cs.modified[0].name == "Module1"
    assert draft.is_dirty()


def test_rename_and_revert_module(work_xlam):
    draft = _draft(work_xlam)
    m1 = next(m for m in draft.modules if m.current_name == "Module1")
    m1.current_name = "Renamed1"
    cs = compute_changes(draft)
    assert cs.renamed[0].old_name == "Module1" and cs.renamed[0].new_name == "Renamed1"
    revert_module(draft, m1.id)
    assert compute_changes(draft).is_empty

def test_new_and_delete(work_xlam):
    draft = _draft(work_xlam)
    m1 = next(m for m in draft.modules if m.current_name == "Module1")
    m1.is_deleted = True
    tw = next(m for m in draft.modules if m.current_name == "ThisWorkbook")
    tw.is_deleted = True
    assert len(compute_changes(draft).deleted) == 2
    draft = revert_all(draft)
    assert compute_changes(draft).is_empty


def test_name_validation(work_xlam):
    draft = _draft(work_xlam)
    assert validate_module_name("Good_Name1", draft) is None
    assert validate_module_name("", draft) is not None
    assert validate_module_name("1Bad", draft) is not None
    assert validate_module_name("Bäd", draft) is not None
    assert validate_module_name("Sub", draft) is not None
    assert validate_module_name("module1", draft) is not None  # case-insensitive dup
    assert validate_module_name("a" * 65, draft) is not None
    assert validate_draft(draft) == []
