"""Right-click Cut/Copy/Paste behavior for both editors (XML plan 23.8)."""

from __future__ import annotations

import tkinter as tk

import pytest

from vba_addin_editor.ui.code_editor import CodeEditor
from vba_addin_editor.ui.text_context_menu import TextContextMenu
from vba_addin_editor.ui.xml_editor import XmlEditor

TEXT = "line one\nline two\nline three\n"


@pytest.fixture(params=[CodeEditor, XmlEditor], ids=["code-editor", "xml-editor"])
def editor(request):
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    root.withdraw()
    widget = request.param(root)
    widget.text.insert("1.0", TEXT)
    widget.update()
    try:
        yield widget
    finally:
        root.destroy()


def _menu(editor) -> TextContextMenu:
    return editor.text._vbaae_context_menu


class FakeEvent:
    def __init__(self, x: int, y: int) -> None:
        self.x = x
        self.y = y
        self.x_root = 2
        self.y_root = 2


def test_menu_created(editor):
    assert isinstance(_menu(editor), TextContextMenu)

def test_cut_enabled_with_selection_and_editable(editor):
    editor.text.clipboard_clear()
    editor.text.clipboard_append("X")
    editor.update()
    editor.text.tag_add("sel", "1.0", "1.4")
    cut, _copy, paste = _menu(editor).menu_states()
    assert cut == "normal" and paste == "normal"



def test_cut_disabled_without_selection(editor):
    cut, _copy, _paste = _menu(editor).menu_states()
    assert cut == "disabled"


def test_cut_removes_selected_text(editor):
    editor.text.tag_add("sel", "1.0", "1.5")
    _menu(editor)._cut()
    editor.update()
    assert editor.get_text().startswith("one\n")


def test_copy_preserves_text(editor):
    editor.text.tag_add("sel", "1.0", "1.4")
    _menu(editor)._copy()
    editor.update()
    assert editor.get_text() == TEXT
    assert editor.text.clipboard_get() == "line"


def test_paste_inserts_clipboard(editor):
    editor.text.clipboard_clear()
    editor.text.clipboard_append("XYZ")
    editor.update()
    editor.text.mark_set("insert", "1.0")
    _menu(editor)._paste()
    editor.update()
    assert editor.get_text().startswith("XYZline one")


def test_paste_replaces_selection(editor):
    editor.text.clipboard_clear()
    editor.text.clipboard_append("XYZ")
    editor.update()
    editor.text.tag_add("sel", "1.0", "1.4")
    _menu(editor)._paste()
    editor.update()
    assert editor.get_text().startswith("XYZ one\n")


def test_paste_disabled_without_clipboard_text(editor):
    editor.text.clipboard_clear()
    _cut, _copy, paste = _menu(editor).menu_states()
    assert paste == "disabled"

def test_right_click_inside_selection_preserves_it(editor):
    editor.text.tag_add("sel", "1.0", "1.8")
    clicked = editor.text.index("@2,1")
    if not (
        editor.text.compare("sel.first", "<=", clicked)
        and editor.text.compare(clicked, "<=", "sel.last")
    ):
        pytest.skip("font geometry puts click outside selection")
    try:
        _menu(editor).prepare_menu(FakeEvent(2, 1))
        editor.text.index("sel.first")
        has_sel = True
    except tk.TclError:
        has_sel = False
    assert has_sel

def test_right_click_outside_selection_clears_it(editor):
    editor.text.tag_add("sel", "2.0", "2.8")
    try:
        _menu(editor).prepare_menu(FakeEvent(2, 1))
    except tk.TclError:
        pass
    try:
        editor.text.index("sel.first")
        pytest.fail("selection should have been cleared")
    except tk.TclError:
        pass
    assert editor.text.index("insert") == editor.text.index("@2,1")


def test_disabled_editor_blocks_cut_and_paste(editor):
    editor.text.config(state="disabled")
    cut, copy, paste = _menu(editor).menu_states()
    assert cut == "disabled" and paste == "disabled"
    assert copy == "disabled"  # no selection in this state


def test_disabled_editor_allows_copy_with_selection(editor):
    editor.text.config(state="disabled")
    editor.text.tag_add("sel", "1.0", "1.4")
    cut, copy, paste = _menu(editor).menu_states()
    assert cut == "disabled" and paste == "disabled"
    assert copy == "normal"


def test_empty_clipboard_does_not_raise(editor):
    editor.text.clipboard_clear()
    _menu(editor).menu_states()  # must not raise


def test_context_menu_operations_participate_in_undo(editor):
    editor.text.tag_add("sel", "1.0", "1.5")
    _menu(editor)._cut()
    editor.update()
    after_cut = editor.get_text()
    editor.text.mark_set("insert", "1.0")
    _menu(editor)._paste()  # paste back at the cursor
    editor.update()
    editor.text.edit_undo()
    assert editor.get_text() == after_cut or editor.get_text() == TEXT
    editor.text.edit_undo()
    assert editor.get_text() == TEXT
