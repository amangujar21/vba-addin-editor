"""GUI: the drop-down a ribbon menu shows in the Ribbon view (open, add, insert, submenus)."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path

import pytest
from conftest import build_xlam_with_ribbon

from vba_addin_editor.services import ribbon_service as rs
from vba_addin_editor.ui import main_window as mw
from vba_addin_editor.ui import menu_dropdown as md
from vba_addin_editor.ui import ribbon_designer as rd

PART = "customUI/customUI14.xml"

RIBBON = """<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="tabTools" label="Tools">
        <group id="grpMenus" label="Menus">
          <button id="btnHello" label="Say Hello" size="large" onAction="Hello" />
          <menu id="mnuMore" label="More">
            <button id="btnA" label="First" onAction="RunA" />
            <menuSeparator id="sepA" />
            <button id="btnB" label="Second" onAction="NoSuchMacro" />
            <menu id="mnuSub" label="Deeper">
              <button id="btnC" label="Inside" onAction="RunA" />
            </menu>
          </menu>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
"""


@pytest.fixture()
def window(tmp_path: Path, monkeypatch):
    monkeypatch.setenv("VBAAE_SESSION_ROOT", str(tmp_path / "sessions"))
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    root.withdraw()
    path = build_xlam_with_ribbon(tmp_path / "work" / "Ribbon.xlam", RIBBON)
    win = mw.MainWindow(root)
    win.load_path(path)
    module = win.draft.find_current("Module1")  # type: ignore[union-attr]
    module.body += "\nPublic Sub RunA(control As IRibbonControl)\n    Hello\nEnd Sub\n"
    win._reload_editors_from_draft()
    root.update()
    try:
        yield win
    finally:
        root.destroy()


@pytest.fixture()
def designer(window) -> rd.RibbonDesigner:
    window._show_view(1)
    window.root.update()
    return window.ribbon_designer


def _el(d: rd.RibbonDesigner, ident: str) -> rs.RibbonElement:
    return next(e for e in d.elements if e.attr("id") == ident)


def _text(window) -> str:
    return window.draft.xml_part_by_path(PART).text


def _children(d: rd.RibbonDesigner, ident: str) -> list[str | None]:
    parent = _el(d, ident)
    return [e.attr("id") for e in d.elements if e.parent == parent.index and e.complete]


def _labels(panel: md.DropPanel) -> list[str]:
    return [row.label if row.kind != "separator" else "---" for row in panel.rows]


def _open(d: rd.RibbonDesigner, ident: str = "mnuMore") -> md.DropPanel:
    d._pick(_el(d, ident).index)  # a click on the menu in the drawing
    d.update()
    assert d.dropdown.is_open
    return d.dropdown.panels[0]


def test_clicking_a_menu_drops_down_its_macros(designer):
    d = designer
    panel = _open(d)
    assert d.selected_element().attr("id") == "mnuMore"
    assert _labels(panel) == ["First", "---", "Second", "Deeper"]
    first, _sep, second, deeper = panel.rows
    assert (first.macro, first.state) == ("RunA", "ok")
    assert (second.macro, second.state) == ("NoSuchMacro", "missing")
    assert deeper.kind == "submenu"
    assert panel.winfo_manager() == "place"

    # The same click again folds it up; a click on another control too.
    d._pick(_el(d, "mnuMore").index)
    assert not d.dropdown.is_open
    _open(d)
    d._pick(_el(d, "btnHello").index)
    assert not d.dropdown.is_open and not panel.winfo_manager()


def test_clicking_a_row_opens_its_macro(window, designer):
    d = designer
    panel = _open(d)
    panel.press(40, (panel.rows[0].top + panel.rows[0].bottom) // 2)
    assert not d.dropdown.is_open
    module = window.draft.find_current("Module1")
    assert window.current_module_id == module.id
    assert window.editor_notebook.index("current") == 0
    line = window.editor.text.get("insert linestart", "insert lineend")
    assert line.startswith("Public Sub RunA(")


def test_submenu_opens_beside_its_row(designer):
    d = designer
    panel = _open(d)
    deeper = panel.rows[3]
    d.dropdown.activate(0, deeper)
    assert len(d.dropdown.chain) == 2
    sub = d.dropdown.panels[1]
    assert _labels(sub) == ["Inside"]
    assert sub.x >= panel.x + panel.width - 2 or sub.x + sub.width <= panel.x + 2
    # A plain row at the first level folds the submenu away again.
    d.dropdown.close_below(0)
    assert len(d.dropdown.chain) == 1 and not sub.winfo_manager()


def test_add_a_macro_between_two_rows(window, designer):
    d = designer
    panel = _open(d)
    before = _text(window)
    # The line between "First" and the separator offers the chips.
    y = panel.slot_y(1)
    panel.hover_at(40, y)
    assert panel.hover == ("slot", 1)
    kinds = {kind for *_box, kind, slot in panel.chips if slot == 1}
    assert kinds == {"button", "menu", "separator"}
    x1, y1, _x2, _y2, _kind, _slot = next(c for c in panel.chips if c[4] == "button" and c[5] == 1)
    panel.press(x1 + 2, y1 + 2)
    picker = d.dropdown.picker
    assert picker.shown
    picker.var.set("hel")
    assert [name for name, _label in picker.choices][:1] == ["Hello"]
    picker.pick()
    assert not picker.shown
    # Hello takes no IRibbonControl, so a small callback that calls it is added too.
    assert _children(d, "mnuMore") == ["btnA", "btnHello2", "sepA", "btnB", "mnuSub"]
    assert 'id="btnHello2" label="Hello" onAction="HelloCallback"' in _text(window)
    body = window.draft.find_current("Module1").body
    assert "Public Sub HelloCallback(control As IRibbonControl)\n    Hello\nEnd Sub" in body
    assert d.dropdown.is_open and _labels(d.dropdown.panels[0])[1] == "Hello"
    window.undo()
    assert _text(window) == before
    assert "HelloCallback" not in window.draft.find_current("Module1").body


def test_new_macro_from_the_footer(window, designer):
    d = designer
    panel = _open(d)
    footer = [c for c in panel.chips if c[5] == 4]
    assert {c[4] for c in footer} == {"button", "menu", "separator"}
    d.dropdown.add(0, 4, "button")
    picker = d.dropdown.picker
    picker.var.set("Format table")
    assert picker.choices[0] == ("FormatTable", "Format table")
    picker.pick()
    assert _children(d, "mnuMore")[-1] == "btnFormatTable"
    assert 'label="Format table" onAction="FormatTable"' in _text(window)
    body = window.draft.find_current("Module1").body
    assert "Public Sub FormatTable(control As IRibbonControl)" in body
    assert d.dropdown.panels[0].rows[-1].state == "ok"


def test_add_a_submenu_and_name_it_in_place(window, designer):
    d = designer
    _open(d)
    d.dropdown.add(0, 0, "menu")
    first = _children(d, "mnuMore")[0]
    assert first == "mnuNewMenu"
    panel = d.dropdown.panels[0]
    entry = panel._rename
    assert entry is not None and entry.get() == "New Menu"
    entry.delete(0, "end")
    entry.insert(0, "Reports")
    panel.finish_rename(True)  # Enter
    d.update()
    # A fresh submenu takes its id from its first name.
    assert '<menu id="mnuReports" label="Reports">' in _text(window)
    assert d.dropdown.is_open and d.dropdown.panels[0].rows[0].label == "Reports"


def test_separator_and_row_menu_actions(window, designer, monkeypatch):
    d = designer
    panel = _open(d)
    d.dropdown.add(0, 1, "separator")
    assert _children(d, "mnuMore")[:3] == ["btnA", "sep1", "sepA"]
    assert '<menuSeparator id="sep1" />' in _text(window)

    panel = d.dropdown.panels[0]
    menu = d.dropdown.prepare_context_menu(0, panel.rows[0])
    assert menu is not None
    d.dropdown.context_action(0, _el(d, "btnA"), "down")
    assert _children(d, "mnuMore")[:2] == ["sep1", "btnA"]
    monkeypatch.setattr(rd.messagebox, "askyesno", lambda *a, **k: True)
    d.dropdown.context_action(0, _el(d, "btnB"), "delete")
    assert "btnB" not in _children(d, "mnuMore")
    d.dropdown.context_action(0, _el(d, "btnA"), "change")
    d.dropdown.picker.var.set("Hello")
    d.dropdown.picker.pick()
    assert 'id="btnA" label="First" onAction="HelloCallback"' in _text(window)


def test_add_a_macro_button_to_a_group(window, designer):
    d = designer
    group = _el(d, "grpMenus")
    d.add_to_group(group, "macro")
    d.dropdown.picker.var.set("RunA")
    d.dropdown.picker.pick()
    assert _children(d, "grpMenus")[-1] == "btnRunA"
    # It takes the size of its neighbour (the menu is normal-size).
    assert '<button id="btnRunA" label="Run A" onAction="RunA" />' in _text(window)


def test_new_macro_name():
    assert md.new_macro_name("format table") == "FormatTable"
    assert md.new_macro_name("Do_It2") == "Do_It2"
    assert md.new_macro_name("2fast") == "fast"
    assert md.new_macro_name("!!") == ""
