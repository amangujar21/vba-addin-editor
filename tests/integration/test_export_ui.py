"""Copying between open add-ins: the copy dialog, Export to / Import from, drag onto a tab."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path

import pytest
from conftest import build_port_source as build_source
from conftest import build_xlam, build_xlam_with_ribbon
from helpers import make_service

from vba_addin_editor.services import port_service as ps
from vba_addin_editor.services import ribbon_service as rs
from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.ui import main_window as mw
from vba_addin_editor.ui.export_dialog import CHECKED, INCLUDED, ExportDialog

PART = "customUI/customUI14.xml"


@pytest.fixture()
def root(tmp_path: Path, monkeypatch):
    monkeypatch.setenv("VBAAE_SESSION_ROOT", str(tmp_path / "sessions"))
    try:
        tk_root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    tk_root.withdraw()
    try:
        yield tk_root
    finally:
        tk_root.destroy()


@pytest.fixture(autouse=True)
def _notes(monkeypatch):
    """The "please check" note after a copy is collected, not shown."""
    shown: list[str] = []
    monkeypatch.setattr(mw.messagebox, "showinfo", lambda _t, msg, **_k: shown.append(msg))
    return shown


@pytest.fixture()
def window(root, tmp_path: Path):
    """Two tabs: the target (Ribbon.xlam), then the source (Source.xlam), which is showing."""
    win = mw.MainWindow(root)
    win.save_service.process_probe = lambda _p: False
    win.load_path(build_xlam_with_ribbon(tmp_path / "work" / "Ribbon.xlam"))
    win.load_path(build_source(tmp_path / "src" / "Source.xlam"))
    root.update()
    return win


def _dialog(win, source, target, preselect=()):
    return ExportDialog(
        win.root,
        source=source,
        target_name=target.baseline.path.name,
        targets_for=lambda c: ps.port_targets(target, c),
        plan=lambda items, suffix, label: ps.plan_export(
            source, target, items, suffix=suffix, group_label=label, check_references=False
        ),
        suffix="_Ribbon",
        preselect=preselect,
    )


def _key(draft, control_id: str) -> str:
    control = next(c for c in ps.portable_controls(draft) if c.control_id == control_id)
    return f"c:{control.key}"


def test_dialog_ticks_several_things_and_plans_them_together(window):
    source_tab, target_tab = window.active_tab, window.tabs[0]
    source, target = source_tab.draft, target_tab.draft
    dialog = _dialog(window, source, target)
    # The ribbon is nested: tab › group › controls.
    tab_row, group_row = _key(source, "tabSrc"), _key(source, "grpReports")
    assert dialog.tree.parent(group_row) == tab_row
    assert dialog.tree.parent(_key(source, "btnReport")) == group_row
    assert "Tick a ribbon control" in dialog.preview.get("1.0", "end")
    assert str(dialog.import_btn.cget("state")) == "disabled"

    dialog.check(_key(source, "btnReport"))
    dialog.check(_key(source, "btnOther"))
    dialog.check("m:RibbonCode.AfterReport")
    # Loose controls share one new group by default.
    assert all(dialog.where[k] == dialog.where[_key(source, "btnReport")] for k in dialog.where)
    text = dialog.preview.get("1.0", "end")
    assert "RunReport → RunReport_Ribbon" in text and "Unrelated → Unrelated_Ribbon" in text
    dialog.confirm()
    plan = dialog.result
    assert plan is not None and plan.ok
    assert window.apply_export(target_tab, plan)
    assert window.active_tab is target_tab  # the copy is shown where it landed
    ribbon = target_tab.draft.xml_part_by_path(PART).text
    assert 'onAction="RunReport_Ribbon"' in ribbon and 'onAction="Unrelated_Ribbon"' in ribbon
    assert "Public Sub AfterReport_Ribbon()" in target_tab.draft.find_current("RibbonCode").body
    assert source.find_current("RibbonCode").body.count("RunReport_Ribbon") == 0  # source untouched
    # One undo step in the target.
    window.undo()
    assert target_tab.draft.find_current("RibbonCode") is None
    assert 'RunReport_Ribbon' not in target_tab.draft.xml_part_by_path(PART).text


def test_ticking_a_tab_includes_what_is_in_it(window):
    source, target = window.active_tab.draft, window.tabs[0].draft
    dialog = _dialog(window, source, target, preselect=[_key(source, "btnReport")])
    dialog.check(_key(source, "tabSrc"))
    assert dialog.tree.item(_key(source, "btnReport"), "text").startswith(INCLUDED)
    assert dialog.tree.item(_key(source, "tabSrc"), "text").startswith(CHECKED)
    items = dialog.items()
    assert [i.control.control_id for i in items] == ["tabSrc"]
    assert items[0].where.describe() == "New tab on the ribbon"
    assert dialog.plan is not None and dialog.plan.ok, dialog.plan.describe()
    dialog.win.destroy()


def test_no_ribbon_in_target_says_so(window, tmp_path: Path):
    plain = DocumentService().open(build_xlam(tmp_path / "plain" / "Plain.xlam"))
    source = window.active_tab.draft
    dialog = _dialog(window, source, plain, preselect=[_key(source, "btnReport")])
    assert "Create a ribbon" in dialog.preview.get("1.0", "end")
    dialog.check(_key(source, "btnReport"), False)
    dialog.check("m:RibbonCode.AfterReport")
    assert dialog.plan is not None and dialog.plan.ok
    dialog.win.destroy()


def test_export_to_a_tab_from_the_outline_and_designer(window, monkeypatch):
    source_tab, target_tab = window.active_tab, window.tabs[0]
    asked: list[tuple] = []

    def fake_ask(source, target, *, preselect=(), title=""):
        asked.append((source, target, list(preselect), title))
        items = [ps.ExportItem.of_macro("RibbonCode", "AfterReport")]
        return ps.plan_export(source, target, items, suffix="_R", check_references=False)

    monkeypatch.setattr(window, "ask_export", fake_ask)
    monkeypatch.setattr(mw, "choose", lambda *_a, **_k: 0)  # the first other tab: Ribbon.xlam
    module = source_tab.draft.find_current("RibbonCode")
    body = module.body
    window._open_module_at(module.id, body[: body.index("Public Sub AfterReport")].count("\n") + 2)
    assert window.export_current_macro()
    source, target, preselect, title = asked[-1]
    assert source is source_tab.draft and target is target_tab.draft
    assert preselect == ["m:RibbonCode.AfterReport"] and title == "Copy to Ribbon.xlam"
    assert window.active_tab is target_tab
    assert "AfterReport_R" in target_tab.draft.find_current("RibbonCode").body

    # From the Ribbon view of the source: its selection is what is offered.
    window.select_tab(source_tab.key)
    window._show_view(1)
    element = next(e for e in window.ribbon_designer.elements if e.attr("id") == "btnReport")
    window.ribbon_designer.select_start(element.start)
    window.export_ribbon_selection(target_tab.key)
    assert asked[-1][2] == [f"c:{PART}@{element.start}"]


def test_drop_on_a_tab_copies_there(window, monkeypatch):
    source_tab, target_tab = window.active_tab, window.tabs[0]
    seen = []
    monkeypatch.setattr(window, "export_to", lambda key, preselect=(): seen.append((key, preselect)) or True)
    monkeypatch.setattr(window.tab_strip, "tab_at", lambda _x, _y: target_tab.key)
    assert window.drag_over_tabs(10, 10)
    assert window.tab_strip.chips[target_tab.key].drop
    assert window.drop_on_tab(10, 10, ["m:RibbonCode.AfterReport"])
    assert seen == [(target_tab.key, ["m:RibbonCode.AfterReport"])]
    assert not window.tab_strip.chips[target_tab.key].drop
    # Over its own tab nothing happens.
    monkeypatch.setattr(window.tab_strip, "tab_at", lambda _x, _y: source_tab.key)
    assert not window.drag_over_tabs(10, 10)
    assert not window.drop_on_tab(10, 10, ["m:RibbonCode.AfterReport"])


def test_import_from_an_open_tab_then_save(window, monkeypatch):
    source_tab, target_tab = window.active_tab, window.tabs[0]
    window.select_tab(target_tab.key)
    monkeypatch.setattr(mw, "choose", lambda *_a, **_k: 0)  # Source.xlam, the other open tab

    def fake_ask(source, target, *, preselect=(), title=""):
        assert source is source_tab.draft and title == "Copy into Ribbon.xlam"
        items = [ps.ExportItem.of_module("Counter"), ps.ExportItem.of_module("Helpers")]
        return ps.plan_export(source, target, items, suffix="_R", check_references=False)

    monkeypatch.setattr(window, "ask_export", fake_ask)
    assert window.import_from()
    draft = target_tab.draft
    assert draft.find_current("Counter") is not None and draft.find_current("Helpers") is not None
    assert "Copied 2 new module(s)" in window.status.cget("text")
    assert rs.new_issues(draft) == []
    assert make_service().save_addin(draft).kind == "success"
    reopened = DocumentService().open(draft.baseline.path)
    assert reopened.find_current("Counter") is not None
