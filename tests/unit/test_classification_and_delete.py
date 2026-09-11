from __future__ import annotations

from pathlib import Path

from conftest import build_xlam_with_class
from helpers import make_service
from pyopenvba.cfb import CFB
from pyopenvba.vba import parse_project_stream

from vba_addin_editor.adapters.pyopenvba_adapter import vba_entry_for
from vba_addin_editor.domain.capabilities import (
    DOCUMENT,
    UNKNOWN,
    build_project_index,
    classify_component,
    empty_project_index,
)
from vba_addin_editor.domain.document import ModuleDisplayKind
from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.services.validation_service import validate_draft


def test_ordinary_class_is_deletable_on_open(work_xlam_with_class: Path):
    draft = DocumentService().open(work_xlam_with_class)
    cls = next(m for m in draft.modules if m.current_name == "OrdinaryClass")
    assert cls.kind == ModuleDisplayKind.CLASS
    assert cls.can_delete and cls.can_rename
    assert cls.project_item_kind == "class"
    sheet = next(m for m in draft.modules if m.current_name == "Sheet1")
    assert not sheet.can_delete and not sheet.can_rename
    assert sheet.restriction_reason == DOCUMENT
    book = next(m for m in draft.modules if m.current_name == "ThisWorkbook")
    assert not book.can_delete


def test_delete_ordinary_class_survives_reopen(work_xlam_with_class: Path):
    draft = DocumentService().open(work_xlam_with_class)
    cls = next(m for m in draft.modules if m.current_name == "OrdinaryClass")
    cls.is_deleted = True
    result = make_service().save_addin(draft)
    assert result.kind == "success", result
    reopened = DocumentService().open(work_xlam_with_class)
    names = {m.current_name for m in reopened.modules}
    assert "OrdinaryClass" not in names
    assert "Module1" in names
    raw = _project_stream(work_xlam_with_class)
    parsed = parse_project_stream(raw, code_page=reopened.baseline.code_page)
    assert "OrdinaryClass" not in parsed.class_modules


def test_delete_class_with_stream_name_mismatch(tmp_path: Path):
    path = build_xlam_with_class(tmp_path / "mismatch.xlam", stream_name="StreamClass")
    draft = DocumentService().open(path)
    cls = next(m for m in draft.modules if m.current_name == "OrdinaryClass")
    assert cls.stream_name == "StreamClass"
    cls.is_deleted = True
    result = make_service().save_addin(draft)
    assert result.kind == "success", result
    reopened = DocumentService().open(path)
    assert "OrdinaryClass" not in {m.current_name for m in reopened.modules}
    parsed = parse_project_stream(_project_stream(path), code_page=reopened.baseline.code_page)
    assert "OrdinaryClass" not in parsed.class_modules


def test_document_delete_blocked_below_gui(work_xlam: Path):
    draft = DocumentService().open(work_xlam)
    sheet = next(m for m in draft.modules if m.current_name == "Sheet1")
    sheet.is_deleted = True
    sheet.can_delete = True
    sheet.destructive_ops_safe = True
    problems = validate_draft(draft)
    assert any("Cannot delete" in item for item in problems)
    before = work_xlam.read_bytes()
    result = make_service().save_addin(draft)
    assert result.kind == "error"
    assert result.reason == "unsupported_component_operation"
    assert work_xlam.read_bytes() == before


def test_new_class_then_reopen_then_delete(work_xlam: Path):
    draft = DocumentService().open(work_xlam)
    from vba_addin_editor.domain.document import ModuleDraft, new_module_id

    draft.modules.append(
        ModuleDraft(
            id=new_module_id(),
            origin_name=None,
            current_name="SavedClass",
            body="Public Z As Long\r\n",
            kind=ModuleDisplayKind.CLASS,
            pyopenvba_kind="other",
            is_new=True,
            is_deleted=False,
            destructive_ops_safe=True,
            can_delete=True,
            can_rename=True,
            project_item_kind="class",
        )
    )
    assert make_service().save_addin(draft).kind == "success"
    reopened = DocumentService().open(work_xlam)
    saved = next(m for m in reopened.modules if m.current_name == "SavedClass")
    assert saved.kind == ModuleDisplayKind.CLASS
    assert saved.can_delete
    saved.is_deleted = True
    assert make_service().save_addin(reopened).kind == "success"
    final = DocumentService().open(work_xlam)
    assert "SavedClass" not in {m.current_name for m in final.modules}


def test_pptm_and_ppam_class_delete(work_pptm_with_class: Path, work_ppam_with_class: Path):
    for path in (work_pptm_with_class, work_ppam_with_class):
        draft = DocumentService().open(path)
        cls = next(m for m in draft.modules if m.current_name == "OrdinaryClass")
        assert cls.can_delete
        cls.is_deleted = True
        result = make_service().save_addin(draft)
        assert result.kind == "success", (path, result)
        names = {m.current_name for m in DocumentService().open(path).modules}
        assert "OrdinaryClass" not in names


def test_classifier_does_not_infer_from_thisworkbook_name():
    index = build_project_index(
        standard_modules=["Module1"],
        class_modules=["ThisWorkbook"],
        document_modules=[],
        base_classes=[],
        lossy=False,
    )
    caps = classify_component(
        logical_name="ThisWorkbook",
        stream_name="ThisWorkbook",
        dir_kind="other",
        is_read_only=False,
        index=index,
        designer_storages=frozenset(),
    )
    assert caps.display_kind == ModuleDisplayKind.CLASS
    assert caps.can_delete


def test_lossy_project_fails_closed():
    index = empty_project_index(lossy=True)
    caps = classify_component(
        logical_name="OrdinaryClass",
        stream_name="OrdinaryClass",
        dir_kind="other",
        is_read_only=False,
        index=index,
        designer_storages=frozenset(),
    )
    assert not caps.can_delete
    assert caps.restriction_reason == "decoding_loss"


def test_unknown_other_module_is_locked():
    index = build_project_index(
        standard_modules=["Module1"],
        class_modules=[],
        document_modules=[],
        base_classes=[],
        lossy=False,
    )
    caps = classify_component(
        logical_name="Mystery",
        stream_name="Mystery",
        dir_kind="other",
        is_read_only=False,
        index=index,
        designer_storages=frozenset(),
    )
    assert not caps.can_delete
    assert caps.restriction_reason == UNKNOWN


def _project_stream(path: Path) -> bytes:
    entry = vba_entry_for(path)
    import zipfile

    with zipfile.ZipFile(path) as package:
        cfb = CFB.from_bytes(package.read(entry))
    return cfb.get_stream("PROJECT")
