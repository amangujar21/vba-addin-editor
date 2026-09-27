"""GUI: UserForms group, read-only Forms tab, and form deletion prompt."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path

import pytest

from vba_addin_editor.ui import main_window as mw
from vba_addin_editor.ui.main_window import MainWindow


@pytest.fixture
def window(tmp_path, monkeypatch):
    monkeypatch.setenv("VBAAE_SESSION_ROOT", str(tmp_path / "sessions"))
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    root.withdraw()
    try:
        yield MainWindow(root)
    finally:
        root.destroy()


def _group_children(win: MainWindow, group: str) -> list[str]:
    for top in win.tree.get_children():
        if win.tree.item(top, "text") == group:
            return [win.tree.item(i, "text") for i in win.tree.get_children(top)]
    return []


def test_forms_tab_and_delete(window: MainWindow, work_xlam_with_form: Path, monkeypatch):
    window.load_path(work_xlam_with_form)
    assert window.draft is not None
    assert _group_children(window, "UserForms") == ["EntryForm"]

    # Forms tab: form -> frame -> textbox, plus the button.
    assert window.forms_tree.get_children() == ("form::EntryForm",)
    assert window.forms_tree.get_children("form::EntryForm") == (
        "form::EntryForm/fraMain",
        "form::EntryForm/cmdOK",
    )
    assert window.forms_tree.get_children("form::EntryForm/fraMain") == (
        "form::EntryForm/fraMain/txtName",
    )
    window.forms_tree.selection_set("form::EntryForm/cmdOK")
    window.forms_tree.event_generate("<<TreeviewSelect>>")
    window.root.update()
    props = {
        window.form_props.item(i, "text"): window.form_props.item(i, "values")[0]
        for i in window.form_props.get_children()
    }
    assert props["Caption"] == "OK"
    assert str(window.form_code_btn.cget("state")) == "normal"
    window.goto_form_code()
    form = window.draft.find_current("EntryForm")
    assert form is not None and window.current_module_id == form.id

    # Delete: the prompt names the layout; the form leaves the tree and is marked in the tab.
    prompts: list[str] = []
    monkeypatch.setattr(mw.messagebox, "askyesno", lambda _t, msg: prompts.append(msg) or True)
    window.delete_module()
    assert prompts and "layout" in prompts[0]
    assert form.is_deleted
    assert _group_children(window, "UserForms") == []
    assert "deleted" in window.forms_tree.item("form::EntryForm", "text")

    window.undo()
    assert not form.is_deleted
    assert _group_children(window, "UserForms") == ["EntryForm"]


def test_rename_userform_refused(window: MainWindow, work_xlam_with_form: Path, monkeypatch):
    window.load_path(work_xlam_with_form)
    form = window.draft.find_current("EntryForm")
    window.tree.selection_set(form.id)
    window.root.update()
    warnings: list[str] = []
    monkeypatch.setattr(mw.messagebox, "showwarning", lambda _t, msg: warnings.append(msg))
    window.rename_module()
    assert warnings and "renaming it is disabled" in warnings[0]
    assert form.current_name == "EntryForm"
