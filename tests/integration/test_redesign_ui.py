"""GUI: activity rail, breadcrumbs, change bars, command palette, review dialog, ribbon panel."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path

import pytest
from conftest import build_xlam_with_ribbon

from vba_addin_editor.ui import main_window as mw
from vba_addin_editor.ui import review_dialog as rd
from vba_addin_editor.ui.code_editor import changed_lines
from vba_addin_editor.ui.command_palette import (
    COMMAND,
    MODULE,
    PROCEDURE,
    PaletteItem,
    filter_items,
)

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


def test_changed_lines_marks_added_and_modified_lines():
    assert changed_lines("a\nb\nc", "a\nb\nc") == {}
    assert changed_lines("a\nb\nc", "a\nB\nc\nd") == {2: "mod", 4: "add"}
    assert changed_lines(None, "x\ny") == {1: "add", 2: "add"}


def test_palette_filters_by_section_and_prefix():
    noop = lambda: None
    items = [
        PaletteItem(COMMAND, "Save File", "File", "Ctrl+S", noop),
        PaletteItem(COMMAND, "Add Ribbon Menu", "Ribbon", "", noop),
        PaletteItem(MODULE, "ReportTools", "Modules", "", noop),
        PaletteItem(PROCEDURE, "ExportReport", "ReportTools · Sub", "", noop),
    ]
    assert [title for title, _ in filter_items(items, "")] == ["COMMANDS"]
    sections = dict(filter_items(items, "report"))
    assert [i.label for i in sections["MODULES"]] == ["ReportTools"]
    assert [i.label for i in sections["PROCEDURES"]] == ["ExportReport"]
    assert "COMMANDS" not in sections
    assert [t for t, _ in filter_items(items, "@")] == ["PROCEDURES"]
    assert [i.label for _t, group in filter_items(items, ">ribbon") for i in group] == ["Add Ribbon Menu"]


def test_rail_switches_views_and_welcome_hides_it(window):
    assert window.rail.winfo_manager() == "pack"
    window.rail_buttons[1].command()
    assert str(window.editor_notebook.select()) == str(window.xml_tab)
    assert window.rail_buttons[1]._active and not window.rail_buttons[0]._active
    window.rail_buttons[0].command()
    assert window.editor_notebook.index("current") == 0


def test_breadcrumb_follows_module_and_procedure(window):
    module = next(m for m in window.draft.modules if m.pyopenvba_kind == "standard")
    window._open_module_at(module.id, 1)
    assert window.crumb_module.cget("text") == module.current_name
    assert window.crumb_group.cget("text") == mw.GROUP_MODULES
    body = window.editor.get_text().split("\n")
    line = next(i for i, text in enumerate(body, 1) if text.lstrip().lower().startswith(("sub ", "public sub ")))
    window.editor.goto_position(line + 1, 1)
    window._update_procedure_crumb()
    name = mw._PROC_DECL.match(body[line - 1]).group(1)
    assert window.proc_button.cget("text").startswith(name)


def test_review_dialog_lists_changes_with_diff(window, monkeypatch):
    module = window.draft.find_current("Module1") or next(
        m for m in window.draft.modules if m.pyopenvba_kind == "standard"
    )
    window._open_module_at(module.id, 1)
    window.editor.text.insert("end", "\n' added line\n")
    window._flush_all_editors()
    seen: list[rd.ReviewView] = []
    monkeypatch.setattr(mw, "ask_review", lambda _p, view, **k: seen.append(view) or rd.CANCEL)
    window.save()
    assert window.draft.is_dirty()
    (view,) = seen
    (entry,) = view.entries
    assert entry.badge == "M" and entry.name == module.current_name and entry.added >= 1
    dialog = rd.ReviewDialog(window.root, view, backup_var=window.auto_backup_var)
    try:
        shown = dialog.diff.get("1.0", "end")
        assert "+ ' added line" in shown
        assert dialog.backup_check.cget("text") == "Keep a backup of the current file"
    finally:
        dialog.win.destroy()


def test_review_revert_discards_after_confirming(window, monkeypatch):
    module = next(m for m in window.draft.modules if m.pyopenvba_kind == "standard")
    window._open_module_at(module.id, 1)
    window.editor.text.insert("end", "\n' to be reverted\n")
    window._flush_all_editors()
    monkeypatch.setattr(mw, "ask_review", lambda *a, **k: rd.REVERT)
    monkeypatch.setattr(mw.messagebox, "askyesno", lambda *a, **k: True)
    window.save()
    assert not window.draft.is_dirty()
    assert window.session.draft is window.draft


def test_ribbon_panel_shows_outline_and_missing_callback(window):
    window._open_xml_at(PART, 1)
    window.root.update()
    panel = window.ribbon_panel
    assert panel.winfo_manager() == "pack"
    labels = [panel.tree.item(i, "text").split("  ", 1)[-1] for i in panel._rows]
    assert labels[:3] == ["Tools", "Main", "Say Hello"]
    assert any(panel.tree.item(i, "values") == ("large",) for i in panel._rows)
    texts = [w.cget("text") for row in panel.check.winfo_children() for w in row.winfo_children()]
    assert "NoSuchMacro: no macro" in texts and "Create stub" in texts
    window._open_xml_at("[Content_Types].xml", 1)
    assert panel.winfo_manager() == ""


def test_successful_save_shows_a_toast_not_a_dialog(window, monkeypatch):
    module = next(m for m in window.draft.modules if m.pyopenvba_kind == "standard")
    window._open_module_at(module.id, 1)
    window.editor.text.insert("end", "\n' saved\n")
    monkeypatch.setattr(mw, "ask_review", lambda *a, **k: rd.SAVE)
    infos: list[str] = []
    monkeypatch.setattr(mw.messagebox, "showinfo", lambda _t, msg, **k: infos.append(msg))
    window.save()
    assert not window.draft.is_dirty()
    assert infos == []
    assert window.toast.visible and window.toast.title.cget("text").startswith("Saved ")


def test_review_diff_numbers_lines_from_the_hunk(window):
    from vba_addin_editor.services.review_service import _diff

    diff = _diff("a\nb\nc\n", "a\nB\nc\nd\n", "M")
    added, removed = rd.diff_stats(diff)
    entry = rd.ReviewEntry("M", "M", "VBA", "", diff, added, removed)
    view = rd.ReviewView("x.xlam", "C:/x.xlam", ".xlam", (entry,), "")
    dialog = rd.ReviewDialog(window.root, view)
    try:
        lines = dialog.diff.get("1.0", "end").splitlines()
        assert lines[:4] == ["  Line 1", "     1    a", "     2  − b", "     2  + B"]
        assert (added, removed) == (2, 1)
    finally:
        dialog.win.destroy()
