"""Ribbon callback navigation, stub creation and Add Ribbon Button in the GUI."""

from __future__ import annotations

import tkinter as tk
import zipfile
from pathlib import Path

import pytest
from conftest import build_xlam_with_ribbon
from helpers import make_service

from vba_addin_editor.services import ribbon_service as rs
from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.ui import main_window as mw
from vba_addin_editor.ui.ribbon_dialogs import ButtonSpec

PART = "customUI/customUI14.xml"


@pytest.fixture()
def window(tmp_path: Path, monkeypatch):
    monkeypatch.setenv("VBAAE_SESSION_ROOT", str(tmp_path / "sessions"))
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    root.withdraw()
    path = build_xlam_with_ribbon(tmp_path / "work" / "Ribbon.xlam")
    win = mw.MainWindow(root)
    win.load_path(path)
    root.update()
    try:
        yield win
    finally:
        root.destroy()


def _cursor_on(win, needle: str) -> None:
    win._open_xml_at(PART, 1)
    offset = win.xml_editor.get_text().index(needle) + 2
    win.xml_editor.text.mark_set("insert", f"1.0+{offset}c")


def test_go_to_macro_from_xml(window):
    _cursor_on(window, 'onAction="Hello"')
    window.goto_callback_under_cursor()
    window.root.update()
    module = window.draft.module_by_id(window.current_module_id)
    assert module.current_name == "Module1"
    line = int(window.editor.text.index("insert").split(".")[0])
    assert window.editor.get_text().split("\n")[line - 1].startswith("Public Sub Hello")
    assert window.editor.text.get("sel.first", "sel.last") == "Hello"
    # And back: F12 in the VBA editor jumps to the single button using it.
    window.goto_callback_under_cursor()
    offset = len(window.xml_editor.text.get("1.0", "insert"))
    assert rs.callback_at(window.xml_editor.get_text(), offset).control_id == "btnHello"


def test_create_missing_callback_is_one_undo_step(window, monkeypatch):
    monkeypatch.setattr(mw.messagebox, "askyesno", lambda *a, **k: True)
    monkeypatch.setattr(mw, "choose", lambda *a, **k: 0)
    _cursor_on(window, 'onAction="NoSuchMacro"')
    before = window.draft.find_current("Module1").body
    window.goto_callback_under_cursor()
    window.root.update()
    body = window.draft.find_current("Module1").body
    assert "Public Sub NoSuchMacro(control As IRibbonControl)" in body
    assert window.editor.get_text() == body
    window.undo()
    assert window.draft.find_current("Module1").body == before


def test_add_button_with_wrapper_saves_and_undoes(window, tmp_path: Path):
    m1 = window.draft.find_current("Module1")
    m1.body += "\nPublic Sub TidySheets()\nEnd Sub\n"
    window._open_module_at(m1.id, 1)
    window.editor.set_text(m1.body)
    proc = rs.procedure_at(m1.body, m1.id, m1.body.split("\n").index("Public Sub TidySheets()") + 1)
    callback_name, wrapper, note = window._plan_button_callback(m1, proc)
    assert callback_name == "TidySheetsCallback" and wrapper and "TidySheets" in note
    part = window.draft.xml_part_by_path(PART)
    xml_before, body_before = part.text, window.draft.find_current("Module1").body
    group = next(c for c in rs.find_containers(part.text, PART) if c.control_id == "grpMain")
    spec = ButtonSpec(group, "btnTidy", "Tidy Sheets", "Copy", "large", "", "Tidy up")
    window.apply_ribbon_button(m1.id, callback_name, wrapper, spec)
    window.root.update()

    assert 'onAction="TidySheetsCallback"' in part.text
    assert window.xml_editor.text.get("sel.first", "sel.last").startswith('<button id="btnTidy"')
    body = window.draft.find_current("Module1").body
    assert "Public Sub TidySheetsCallback(control As IRibbonControl)\n    TidySheets\nEnd Sub" in body
    assert rs.new_issues(window.draft) == []

    window.undo()
    assert part.text == xml_before
    assert window.draft.find_current("Module1").body == body_before
    window.redo()
    assert 'id="btnTidy"' in window.draft.xml_part_by_path(PART).text

    result = make_service().save_addin(window.draft)
    assert result.kind == "success", result
    reopened = DocumentService().open(window.draft.baseline.path)
    assert "TidySheetsCallback" in reopened.find_current("Module1").body
    with zipfile.ZipFile(window.draft.baseline.path) as package:
        assert 'id="btnTidy"' in package.read(PART).decode("utf-8")


def test_existing_callback_wrapper_is_reused(window):
    m1 = window.draft.find_current("Module1")
    m1.body += "\nSub Go()\nEnd Sub\n\nSub GoCallback(control As IRibbonControl)\n    Go\nEnd Sub\n"
    proc = next(p for p in rs.parse_procedures(m1.body, m1.id) if p.name == "Go")
    name, wrapper, _note = window._plan_button_callback(m1, proc)
    assert (name, wrapper) == ("GoCallback", None)


def test_save_warns_only_for_new_ribbon_problems(window, monkeypatch):
    asked: list[str] = []
    monkeypatch.setattr(mw.messagebox, "askyesno", lambda _t, msg, **k: asked.append(msg) or False)
    assert window._confirm_ribbon_issues()  # NoSuchMacro/OnToggle pre-exist: no prompt
    assert not asked
    part = window.draft.xml_part_by_path(PART)
    part.text = part.text.replace('onAction="Hello"', 'onAction="Helo"')
    assert not window._confirm_ribbon_issues()
    assert "Helo" in asked[0]


def test_context_menu_items_follow_cursor(window):
    xml_menu = window.xml_editor.text._vbaae_context_menu
    labels = [xml_menu.menu.entrycget(i, "label") for i in range(xml_menu.menu.index("end") + 1)
              if xml_menu.menu.type(i) == "command"]
    assert "Go to Macro" in labels
    _cursor_on(window, 'onAction="Hello"')
    assert window._xml_cursor_callback() is not None
    window.xml_editor.text.mark_set("insert", "1.0")
    assert window._xml_cursor_callback() is None


def test_ribbon_buttons_dialog_lists_problems(window):
    window.show_ribbon_buttons(problems_only=True)
    dialog = window._ribbon_dialog
    names = {entry.callback.name for entry in dialog.visible}
    assert names == {"NoSuchMacro", "OnToggle", "GetToggle"}
    dialog.problems_only.set(False)
    dialog.filter_var.set("hello")
    assert [e.callback.name for e in dialog.visible] == ["Hello"]
    dialog.goto_macro()
    assert window.draft.module_by_id(window.current_module_id).current_name == "Module1"

