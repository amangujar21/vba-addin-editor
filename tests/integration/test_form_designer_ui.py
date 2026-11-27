"""GUI: the UserForms designer — edit, add, move, rename, code, new form, import, save."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from types import SimpleNamespace

import pytest
from conftest import build_form_source, build_xlam

from vba_addin_editor.services.document_service import DocumentService
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


def _commit(designer, key: str, value: str) -> None:
    row = designer.rows[key]
    row.var.set(value)
    row.commit()


def _click(designer, x: float, y: float) -> SimpleNamespace:
    return SimpleNamespace(x=x, y=y)


def test_designer_edits_are_undoable_and_saved(window: MainWindow, work_xlam_with_form: Path):
    window.load_path(work_xlam_with_form)
    window._show_view(2)
    designer = window.form_designer
    draft = window.draft
    assert draft is not None
    form = draft.find_current("EntryForm")
    assert window.forms_tree.get_children() == ("form::EntryForm",)

    # Caption of a control, then undo.
    designer.select("EntryForm", "cmdOK")
    assert designer.rows["caption"].var.get() == "OK"
    _commit(designer, "caption", "Go")
    assert designer.control().caption == "Go"
    assert form.form_storage is not None
    window.undo()
    assert designer.control().caption == "OK"
    form = draft.find_current("EntryForm")  # undo rebuilds the module objects
    assert form is not None and form.form_storage is None and not draft.is_dirty()

    # Rename: the code follows.
    designer.select("EntryForm", "cmdOK")
    _commit(designer, "name", "cmdGo")
    assert designer.control_name == "cmdGo"
    assert "Private Sub cmdGo_Click()" in form.body
    assert window.forms_tree.exists("form::EntryForm/cmdGo")

    # Geometry from the inspector.
    _commit(designer, "left", "150")
    _commit(designer, "width", "90")
    control = designer.control()
    assert (control.left, control.width) == (150, 90)

    # Drag it 30 points right and 15 down, as the mouse does.
    x0, y0, _x1, _y1 = designer.control_box("cmdGo")
    start = (x0 + 5, y0 + 5)
    designer.press(_click(designer, *start))
    designer.motion(_click(designer, start[0] + 40, start[1] + 20))
    designer.release(_click(designer, start[0] + 40, start[1] + 20))
    moved = designer.control()
    assert (moved.left, moved.top) == (180, control.top + 15)

    # Add a control into the frame (the selected container).
    designer.select("EntryForm", "fraMain")
    assert designer.add_control("CheckBox")
    assert designer.control_name == "CheckBox1"
    frame = designer.design().control("fraMain")
    assert "CheckBox1" in [c.name for c in frame.children]

    # Form caption and size.
    designer.select("EntryForm", None)
    _commit(designer, "caption", "Entry form")
    _commit(designer, "width", "320")
    assert designer.design().caption == "Entry form" and designer.design().width == 320

    # Double-click a control: its event procedure, in the code view.
    x0, y0, _x1, _y1 = designer.control_box("cmdGo")
    designer.double_click(_click(designer, x0 + 3, y0 + 3))
    assert window.current_module_id == form.id
    assert "Private Sub cmdGo_Click()" in window.editor.get_text()

    window._show_view(2)
    x0, y0, _x1, _y1 = designer.control_box("CheckBox1")
    designer.double_click(_click(designer, x0 + 3, y0 + 3))
    assert "Private Sub CheckBox1_Click()" in form.body

    # Delete a control.
    window._show_view(2)
    designer.select("EntryForm", "CheckBox1")
    designer.delete_selected()
    assert designer.design().control("CheckBox1") is None

    result = window.save_service.save_addin(draft)
    assert result.kind == "success", result
    reopened = DocumentService().open(work_xlam_with_form)
    design = reopened.baseline.form("EntryForm")
    assert design.caption == "Entry form"
    assert design.control("cmdGo").left == 180
    assert "Private Sub CheckBox1_Click()" in reopened.find_current("EntryForm").body


def test_new_and_imported_forms(window: MainWindow, tmp_path: Path, monkeypatch):
    target = build_xlam(tmp_path / "work" / "Target.xlam")
    source = build_form_source(tmp_path / "src" / "Source.xlam")
    window.load_path(target)
    draft = window.draft
    assert draft is not None

    monkeypatch.setattr(MainWindow, "_prompt_name", lambda self, title, initial, action=None: "frmNew")
    created = window.new_userform()
    assert created is not None and created.current_name == "frmNew"
    assert window.form_designer.module_id == created.id
    assert window.form_designer.add_control("CommandButton")

    monkeypatch.setattr(mw, "choose", lambda *_a, **_k: 0)
    asked: list[str] = []
    monkeypatch.setattr(mw.messagebox, "askyesno", lambda _t, msg: asked.append(msg) or True)
    monkeypatch.setattr(mw.messagebox, "showinfo", lambda *_a, **_k: None)
    assert window.import_userform(source)
    assert asked and "UserForm EntryForm" in asked[0]
    imported = draft.find_current("EntryForm")
    assert imported is not None and imported.form_storage is not None
    assert window.form_designer.module_id == imported.id
    assert draft.find_current("Helpers") is not None

    window.undo()
    assert draft.find_current("EntryForm") is None
    window.redo()
    assert draft.find_current("EntryForm") is not None

    result = window.save_service.save_addin(draft)
    assert result.kind == "success", result
    reopened = DocumentService().open(target)
    assert {f.name for f in reopened.baseline.forms} == {"frmNew", "EntryForm"}
    assert [c.name for c in reopened.baseline.form("frmNew").controls] == ["CommandButton1"]
