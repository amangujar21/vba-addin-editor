"""Import Button from Another Add-in in the GUI: dialog, apply, undo, save."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path

import pytest
from conftest import build_port_source as build_source
from conftest import build_xlam_with_ribbon
from helpers import make_service

from vba_addin_editor.services import port_service as ps
from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.ui import main_window as mw
from vba_addin_editor.ui.ribbon_dialogs import ImportControlDialog

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


def test_dialog_previews_and_returns_the_plan(window, tmp_path: Path):
    source = DocumentService().open(build_source(tmp_path / "src" / "Source.xlam"))
    draft = window.draft
    controls = ps.portable_controls(source)
    dialog = ImportControlDialog(
        window.root,
        source_name="Source.xlam",
        controls=controls,
        targets_for=lambda c: ps.port_targets(draft, c),
        plan=lambda c, w, s, g: ps.plan_port(source, draft, c, w, suffix=s, group_label=g, check_references=False),
        suffix=ps.default_suffix(draft),
        show_part=False,
    )
    window.root.update()
    index = next(i for i, c in enumerate(controls) if c.control_id == "btnReport")
    dialog.listbox.selection_clear(0, "end")
    dialog.listbox.selection_set(index)
    dialog._control_changed()
    preview = dialog.preview.get("1.0", "end")
    assert "RunReport → RunReport_Ribbon" in preview
    assert str(dialog.import_btn.cget("state")) == "normal"
    # A bad suffix blocks the import.
    dialog.suffix_var.set("no spaces allowed")
    dialog._refresh()
    assert str(dialog.import_btn.cget("state")) == "disabled"
    dialog.suffix_var.set("_X")
    dialog.confirm()
    assert dialog.result is not None and dialog.result.ok
    assert any("RunReport_X" in new for _w, _o, new in dialog.result.renames)


def test_apply_is_one_undo_step_and_saves(window, tmp_path: Path, monkeypatch):
    monkeypatch.setattr(mw.messagebox, "showinfo", lambda *a, **k: None)
    source = window.load_port_source(build_source(tmp_path / "src" / "Source.xlam"))
    assert source is not None
    draft = window.draft
    control = next(c for c in ps.portable_controls(source) if c.control_id == "btnReport")
    where = next(t for t in ps.port_targets(draft, control) if t.container.display.endswith("Empty"))
    plan = ps.plan_port(source, draft, control, where, suffix="_Ribbon", check_references=False)
    before_xml = draft.xml_part_by_path(PART).text
    assert window.apply_port_plan(plan)
    window.root.update()
    assert window.draft.find_current("RibbonCode") is not None
    assert window.tree.exists(plan.new_module_ids[0])
    assert 'onAction="RunReport_Ribbon"' in window.draft.xml_part_by_path(PART).text
    assert window.ribbon_designer.part_path == PART

    window.undo()
    assert window.draft.find_current("RibbonCode") is None
    assert window.draft.xml_part_by_path(PART).text == before_xml
    assert not window.draft.binary_parts
    window.redo()
    assert window.draft.find_current("Helpers") is not None
    assert window.draft.binary_parts

    result = make_service().save_addin(window.draft)
    assert result.kind == "success", result
    reopened = DocumentService().open(window.draft.baseline.path)
    assert reopened.find_current("Counter") is not None


def test_same_file_is_refused(window, monkeypatch):
    shown: list[str] = []
    monkeypatch.setattr(mw.messagebox, "showinfo", lambda _t, text, **k: shown.append(text))
    assert window.load_port_source(Path(window.draft.baseline.path)) is None
    assert "this is the file that is open" in shown[0]
