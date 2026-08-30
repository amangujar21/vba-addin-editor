from __future__ import annotations

from vba_addin_editor.domain.changes import (
    compute_changes,
    dirty_count,
    revert_module,
    revert_xml_part,
)
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


def _xml_draft(work_pptm):
    return DocumentService().open(work_pptm)


def test_pptm_draft_has_xml_parts(work_pptm):
    draft = _xml_draft(work_pptm)
    assert draft.xml_parts, "synthetic PPTM should expose XML parts"
    paths = {p.path for p in draft.xml_parts}
    assert "[Content_Types].xml" in paths
    assert any(p.path.endswith(".rels") for p in draft.xml_parts)


def test_xml_text_change_makes_draft_dirty(work_pptm):
    draft = _xml_draft(work_pptm)
    part = draft.xml_parts[0]
    part.text = part.text + "\n<!-- marker -->"
    assert draft.is_dirty()
    assert dirty_count(draft) == 1
    changes = compute_changes(draft)
    assert changes.has_xml_changes and not changes.has_vba_changes
    assert [c.path for c in changes.xml_modified] == [part.path]


def test_read_only_baseline_xml_cannot_become_dirty(work_pptm):
    draft = _xml_draft(work_pptm)
    part = draft.xml_parts[0]
    part.editable = False
    part.well_formed_on_open = False
    part.open_problem = "pre-existing malformed XML"

    part.text = (part.text or "") + "changed despite UI lock"

    assert not part.is_dirty()
    assert part not in draft.changed_xml_parts()
    assert not draft.is_dirty()


def test_xml_return_to_original_is_clean(work_pptm):
    draft = _xml_draft(work_pptm)
    part = draft.xml_parts[0]
    original = part.text
    part.text = original + "x"
    part.text = original
    assert not draft.is_dirty()


def test_revert_all_reverts_xml(work_pptm):
    draft = _xml_draft(work_pptm)
    draft.xml_parts[0].text += "x"
    draft = revert_all(draft)
    assert not draft.is_dirty()


def test_revert_xml_part_reverts_only_selected(work_pptm):
    draft = _xml_draft(work_pptm)
    assert len(draft.xml_parts) >= 2
    a, b = draft.xml_parts[0], draft.xml_parts[1]
    a.text += "x"
    b.text += "y"
    revert_xml_part(draft, a.path)
    assert a.text == a.original_text
    assert b.text == b.original_text + "y"
    assert draft.is_dirty()
