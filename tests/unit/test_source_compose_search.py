from __future__ import annotations

from pathlib import Path

from vba_addin_editor.adapters.source_codec import compose_module_source, split_attribute_header
from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.services.import_export_service import ImportExportService
from vba_addin_editor.services.search_service import SearchService, parse_procedures


def test_compose_keeps_procedure_level_attributes():
    body = (
        "Public Sub X()\nEnd Sub\n"
        "Attribute VB_Name = \"should_stay_in_body\"\n"
    )
    source = compose_module_source(
        body=body,
        hidden_header='Attribute VB_Name = "Mod1"',
        current_name="Mod1",
        kind="standard",
    )
    header, got_body = split_attribute_header(source)
    assert "Attribute VB_Name = \"Mod1\"" in header
    assert "should_stay_in_body" in got_body


def test_class_export_includes_current_body(work_xlam_with_class: Path, tmp_path: Path):
    draft = DocumentService().open(work_xlam_with_class)
    cls = next(m for m in draft.modules if m.current_name == "OrdinaryClass")
    cls.body = cls.body + "Public Sub Edited()\r\nEnd Sub\r\n"
    dest = tmp_path / "OrdinaryClass.cls"
    ImportExportService().export_module(draft, cls.id, dest)
    text = dest.read_text(encoding="utf-8")
    assert "Edited" in text
    assert "VERSION 1.0 CLASS" in text
    assert "OrdinaryClass" in text


def test_search_skips_deleted_and_supports_whole_word(work_xlam: Path):
    draft = DocumentService().open(work_xlam)
    m1 = next(m for m in draft.modules if m.current_name == "Module1")
    m1.body = "foo bar foo2 foo\n"
    doomed = next(m for m in draft.modules if m.current_name == "Sheet1")
    doomed.body = "foo hidden"
    doomed.is_deleted = True
    results = SearchService().search(draft, "foo", whole_word=True)
    assert results.total == 2
    assert all(hit.path == "Module1" for hit in results.hits)


def test_procedure_outline_skips_comments_and_joins_continuation():
    body = (
        "' Sub Hidden()\n"
        "Public Sub Visible( _\n"
        "    x As Long)\nEnd Sub\n"
        'MsgBox "Sub NotAProc"\n'
        "Private Property Get Name() As String\nEnd Property\n"
    )
    found = parse_procedures(body, "m1")
    names = {item.name for item in found}
    assert "Visible" in names
    assert "Name" in names
    assert "Hidden" not in names
    assert "NotAProc" not in names
