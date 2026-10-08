"""Ctrl+F / Ctrl+H: one inline find bar per editor, never a second window."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path

import pytest

from vba_addin_editor.ui import main_window as mw

BODY = "Sub One()\n    total = total + 1\nEnd Sub\n\nSub Two()\n    Total = subtotal\nEnd Sub\n"


@pytest.fixture()
def window(work_xlam: Path, tmp_path: Path, monkeypatch):
    monkeypatch.setenv("VBAAE_SESSION_ROOT", str(tmp_path / "sessions"))
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    root.withdraw()
    win = mw.MainWindow(root)
    win.load_path(work_xlam)
    root.update()
    win._open_module_at(win.draft.find_current("Module1").id, 1)
    win.editor.set_text(BODY)
    root.update()
    try:
        yield win
    finally:
        root.destroy()


def _toplevels(win) -> list:
    return [w for w in win.root.winfo_children() if isinstance(w, tk.Toplevel)]


def _selection(editor) -> tuple[str, str]:
    return editor.text.index("sel.first"), editor.text.get("sel.first", "sel.last")


def test_ctrl_f_shows_one_bar_and_no_window(window):
    editor = window.editor
    assert not editor.find_bar.winfo_manager()
    # The text widget's own binding stops the key, so the window-level one never
    # opens a second find box (and Tk's Ctrl+H "delete a character" never runs).
    for sequence in ("<<Find>>", "<<Replace>>"):
        script = editor.text.bind(sequence)
        assert script and "break" in script
    window.find_in_editor()
    window.find_in_editor()
    window.root.update()
    assert editor.find_bar.winfo_manager() == "grid"
    assert _toplevels(window) == []
    assert not editor.replace_entry.winfo_manager()
    window.find_in_editor(replace=True)
    assert editor.replace_entry.winfo_manager() == "grid"
    editor.hide_find()
    assert not editor.find_bar.winfo_manager()


def test_find_next_and_previous_wrap(window):
    editor = window.editor
    editor.show_find()
    editor.text.mark_set("insert", "1.0")
    assert editor.find_next("total")
    assert _selection(editor) == ("2.4", "total")
    assert editor.find_status.cget("text") == "1 of 4"
    for expected in ("2.12", "6.4", "6.15"):
        assert editor.find_next()
        assert _selection(editor)[0] == expected
    assert editor.find_next()  # wraps to the first match
    assert _selection(editor)[0] == "2.4"
    assert editor.find_next(backwards=True)  # and back around to the last
    assert _selection(editor)[0] == "6.15"
    assert not editor.find_next("no such text")
    assert editor.find_status.cget("text") == "No matches"


def test_match_case_and_whole_word(window):
    editor = window.editor
    editor.show_find()
    editor.find_var.set("total")
    assert len(editor.text.tag_ranges("find_match")) // 2 == 4
    editor.whole_word.set(True)
    assert len(editor._highlight_matches()) == 3  # not "subtotal"
    editor.match_case.set(True)
    assert len(editor._highlight_matches()) == 2  # not "Total"


def test_find_seeds_from_selection(window):
    editor = window.editor
    editor.text.tag_add("sel", "1.4", "1.7")
    editor.show_find()
    assert editor.find_var.get() == "One"


def test_replace_current_then_all_reaches_the_draft(window):
    editor = window.editor
    editor.show_find(replace=True)
    editor.find_var.set("total")
    editor.replace_var.set("sum")
    editor.text.mark_set("insert", "1.0")
    editor.replace_current()  # selects the first match
    editor.replace_current()  # replaces it and moves on
    assert editor.get_text().split("\n")[1] == "    sum = total + 1"
    assert _selection(editor) == ("2.10", "total")
    assert editor.replace_all() == 3
    assert editor.get_text() == BODY.replace("subtotal", "subsum").replace("total", "sum").replace(
        "Total", "sum"
    )
    window._flush_all_editors()
    module = window.draft.module_by_id(window.current_module_id)
    assert module.body == editor.get_text()
    window.undo()  # typing and replacing since the module was opened is one coalesced step
    assert "sum" not in window.draft.module_by_id(window.current_module_id).body
    assert window.editor.get_text() == window.draft.module_by_id(window.current_module_id).body


def test_replace_leaves_a_read_only_editor_alone(window):
    editor = window.editor
    editor.text.config(state="disabled")
    assert editor.replace_all("total", "sum") == 0
    assert not editor.replace_current()
    editor.text.config(state="normal")
    assert editor.get_text() == BODY


def test_xml_editor_has_its_own_bar(window):
    window.editor_notebook.select(window.xml_tab)
    window.root.update()
    window.find_in_editor()
    assert window.xml_editor.find_bar.winfo_manager() == "grid"
    assert not window.editor.find_bar.winfo_manager()
