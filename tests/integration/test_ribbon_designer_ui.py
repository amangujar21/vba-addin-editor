"""GUI: the Ribbon view (designer) and the PROCEDURES outline of the Code view."""

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
from vba_addin_editor.ui import ribbon_designer as rd

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


@pytest.fixture()
def designer(window) -> rd.RibbonDesigner:
    window._show_view(1)
    window.root.update()
    return window.ribbon_designer


def _el(d: rd.RibbonDesigner, ident: str) -> rs.RibbonElement:
    return next(e for e in d.elements if e.attr("id") == ident)


def _select(d: rd.RibbonDesigner, ident: str) -> None:
    assert d.select_start(_el(d, ident).start)


def _type(field, value: str) -> None:
    field.var.set(value)
    field.commit()


def _text(window) -> str:
    return window.draft.xml_part_by_path(PART).text


def _parent_id(d: rd.RibbonDesigner, ident: str) -> str | None:
    parent = d.parent_of(_el(d, ident))
    return parent.attr("id") if parent is not None else None


def _children(d: rd.RibbonDesigner, ident: str) -> list[str | None]:
    parent = _el(d, ident)
    return [e.attr("id") for e in d.elements if e.parent == parent.index and e.complete]


def _rows(tree, parent: str = ""):
    for iid in tree.get_children(parent):
        yield iid
        yield from _rows(tree, iid)


def _row_labels(d: rd.RibbonDesigner) -> list[str]:
    return [d.tree.item(i, "text").split("  ", 1)[-1] for i in _rows(d.tree)]


# -- selecting and the inspector -------------------------------------------------------


def test_selection_fills_the_inspector(designer):
    d = designer
    _select(d, "btnHello")
    form = d.inspector
    assert form.heading.cget("text") == "BUTTON"
    assert form.label.var.get() == "Say Hello"
    assert form.macro.var.get() == "Hello"
    assert form.macro.link.cget("text") == "Open"
    assert form.ident.var.get() == "btnHello"
    assert form._size_shown and str(form.size_buttons["large"].cget("style")) == "SegOn.TButton"
    assert d.peek.location is not None
    assert d.peek.text.get("1.0", "end").startswith("Public Sub Hello()")

    # A control in a menu has no size; its other callbacks get rows of their own.
    _select(d, "tglX")
    assert not form._size_shown
    assert [item.attr for item in form._extra_fields] == ["getPressed"]
    assert form._extra_fields[0].var.get() == "GetToggle"

    # Containers have no macro.
    _select(d, "grpMain")
    assert form.heading.cget("text") == "GROUP" and not form.macro.shown


def test_label_commit_is_one_undo_step(window, designer):
    d = designer
    before = _text(window)
    _select(d, "btnHello")
    _type(d.inspector.label, "Greet")
    assert '<button id="btnHello" label="Greet" size="large" onAction="Hello" />' in _text(window)
    assert d.selected_element().attr("id") == "btnHello"
    window.undo()
    assert _text(window) == before
    assert d.inspector.label.var.get() == "Say Hello"


def test_typing_then_picking_another_row_commits_to_the_first(window, designer):
    d = designer
    _select(d, "btnHello")
    d.inspector.label.var.set("Greet")  # typed, not yet confirmed
    d.tree.selection_set(f"e{_el(d, 'btnMissing').index}")
    d._on_select()  # what the click does; the entry's focus-out commit comes with it
    assert 'id="btnHello" label="Greet"' in _text(window)
    assert d.selected_element().attr("id") == "btnMissing"
    assert d.inspector.label.var.get() == "Broken"


def test_macro_commit_adds_a_wrapper_for_a_plain_sub(window, designer):
    d = designer
    xml_before, body_before = _text(window), window.draft.find_current("Module1").body
    _select(d, "btnMissing")
    assert d.inspector.macro.link.cget("text") == "Create Sub"
    _type(d.inspector.macro, "Hello")
    assert 'id="btnMissing" label="Broken" onAction="HelloCallback"' in _text(window)
    body = window.draft.find_current("Module1").body
    assert "Public Sub HelloCallback(control As IRibbonControl)\n    Hello\nEnd Sub" in body
    assert d.inspector.macro.link.cget("text") == "Open"
    window.undo()  # text and wrapper together
    assert _text(window) == xml_before
    assert window.draft.find_current("Module1").body == body_before


def test_clearing_the_macro_removes_the_attribute(window, designer):
    _select(designer, "btnMissing")
    _type(designer.inspector.macro, "")
    assert '<button id="btnMissing" label="Broken" />' in _text(window)


def test_id_commit_validates(window, designer):
    d = designer
    _select(d, "btnMissing")
    _type(d.inspector.ident, "btnHello")
    assert "already has the id" in window.status.cget("text")
    _type(d.inspector.ident, "1bad")
    assert "starts with a letter" in window.status.cget("text")
    assert 'id="btnMissing"' in _text(window)
    assert d.inspector.ident.var.get() == "btnMissing"  # the box goes back to what is in the file
    _type(d.inspector.ident, "btnFixed")
    assert 'id="btnFixed"' in _text(window) and 'id="btnMissing"' not in _text(window)
    assert d.selected_element().attr("id") == "btnFixed"


def test_size_toggle_and_undo(window, designer):
    d = designer
    _select(d, "btnHello")
    d.inspector._set_size("normal")
    assert '<button id="btnHello" label="Say Hello" onAction="Hello" />' in _text(window)
    assert str(d.inspector.size_buttons["normal"].cget("style")) == "SegOn.TButton"
    window.undo()
    assert 'id="btnHello" label="Say Hello" size="large"' in _text(window)


# -- adding and deleting --------------------------------------------------------------


def test_add_button_names_its_id_after_the_label(window, designer):
    d = designer
    _select(d, "grpEmpty")
    before = _text(window)
    assert d.add_control("button")
    new = d.selected_element()
    assert new is not None and new.attr("id") == "btnNewButton"
    assert _parent_id(d, "btnNewButton") == "grpEmpty"
    assert new.attr("onAction") is None  # no macro until one is chosen
    _type(d.inspector.label, "Run Report")
    assert d.selected_element().attr("id") == "btnRunReport"
    assert 'label="Run Report"' in _text(window)
    window.undo()
    window.undo()
    assert _text(window) == before


def test_add_menu_group_and_separator(window, designer):
    d = designer
    _select(d, "grpEmpty")
    assert d.add_control("menu")
    assert d.selected_element().attr("id") == "mnuNewMenu"
    assert _parent_id(d, "mnuNewMenu") == "grpEmpty"

    _select(d, "grpMain")
    assert d.add_control("group")
    assert d.selected_element().attr("id") == "grpNewGroup"
    assert _children(d, "tabTools") == ["grpMain", "grpNewGroup", "grpMenus", "grpEmpty"]

    _select(d, "btnHello")
    assert d.add_control("separator")
    assert _children(d, "grpMain") == ["btnHello", "sep1", "btnMissing"]
    assert rs.local_name(_el(d, "sep1").name) == "separator"

    # In a menu, a separator is a menuSeparator.
    _select(d, "tglX")
    assert d.add_control("separator")
    assert rs.local_name(_el(d, "sep2").name) == "menuSeparator"
    assert rs.new_issues(window.draft) == []


def test_delete_a_control_selects_its_neighbour(window, designer):
    d = designer
    before = _text(window)
    _select(d, "btnMissing")
    assert d.delete_selected()
    assert 'id="btnMissing"' not in _text(window)
    assert d.selected_element().attr("id") == "btnHello"
    window.undo()
    assert _text(window) == before


def test_delete_a_container_asks_first(window, designer, monkeypatch):
    d = designer
    answers: list[bool] = [False]
    monkeypatch.setattr(rd.messagebox, "askyesno", lambda *a, **k: answers.pop(0))
    _select(d, "grpMenus")
    assert not d.delete_selected()
    assert 'id="tglX"' in _text(window)
    answers.append(True)
    assert d.delete_selected()
    text = _text(window)
    assert 'id="grpMenus"' not in text and 'id="mnuMore"' not in text and 'id="tglX"' not in text
    assert "Deleted 3 controls" in window.status.cget("text")


def test_multi_select_delete(window, designer):
    d = designer
    d._select_iids([f"e{_el(d, 'btnHello').index}", f"e{_el(d, 'btnMissing').index}"])
    d._show_selection()
    assert d.inspector.heading.cget("text") == "2 CONTROLS SELECTED"
    assert d.delete_selected()
    assert _children(d, "grpMain") == []


# -- moving ------------------------------------------------------------------------------


def test_move_selected_keeps_the_selection(window, designer):
    d = designer
    _select(d, "btnMissing")
    assert d.move_selected(-1)
    assert _children(d, "grpMain") == ["btnMissing", "btnHello"]
    assert d.selected_element().attr("id") == "btnMissing"
    assert not d.move_selected(-1)  # already first
    assert _children(d, "grpMain") == ["btnMissing", "btnHello"]


def test_move_to_into_before_after_first(window, designer):
    d = designer
    # into: last child of the target.
    assert d.move_to(_el(d, "btnHello"), _el(d, "grpEmpty"), "into")
    assert _children(d, "grpEmpty") == ["btnHello"]
    assert d.selected_element().attr("id") == "btnHello"
    # first: before the target's first child.
    assert d.move_to(_el(d, "btnMissing"), _el(d, "grpMenus"), "first")
    assert _children(d, "grpMenus") == ["btnMissing", "mnuMore"]
    # before a sibling, in another parent.
    assert d.move_to(_el(d, "btnHello"), _el(d, "btnMissing"), "before")
    assert _children(d, "grpMenus") == ["btnHello", "btnMissing", "mnuMore"]
    # after: into a menu, where a size makes no sense, so it goes.
    assert d.move_to(_el(d, "btnHello"), _el(d, "tglX"), "after")
    assert _children(d, "mnuMore") == ["tglX", "btnHello"]
    assert _el(d, "btnHello").attr("size") is None
    assert rs.new_issues(window.draft) == []


def test_drop_target_refuses_schema_invalid_drops(designer):
    d = designer
    row = lambda ident: f"e{_el(d, ident).index}"
    # Not onto itself, not into its own child.
    assert d.drop_target(_el(d, "grpMain"), row("grpMain"), 0.5) is None
    assert d.drop_target(_el(d, "grpMain"), row("btnHello"), 0.5) is None
    # A button cannot go into or beside a tab.
    assert d.drop_target(_el(d, "btnHello"), row("tabTools"), 0.5) is None
    # A group can go into a tab, but not into a menu or beside a button.
    target, where = d.drop_target(_el(d, "grpEmpty"), row("tabTools"), 0.5)
    assert target.attr("id") == "tabTools" and where == "into"
    assert d.drop_target(_el(d, "grpEmpty"), row("tglX"), 0.5) is None
    # A toggle cannot hold a button: beside it, by pointer height.
    assert d.drop_target(_el(d, "btnHello"), row("tglX"), 0.2)[1] == "before"
    assert d.drop_target(_el(d, "btnHello"), row("tglX"), 0.8)[1] == "after"
    # A group holds a button and sits in a tab that doesn't: into.
    assert d.drop_target(_el(d, "btnHello"), row("grpEmpty"), 0.1)[1] == "into"


# -- filter, problems and finding things -------------------------------------------------


def test_filter_and_problems_only(window, designer):
    d = designer
    d.filter_var.set("hello")
    window.root.update()
    labels = _row_labels(d)
    assert "Say Hello" in labels and "Tools" in labels and "Main" in labels
    assert "Broken" not in labels and "Toggle" not in labels
    d.filter_var.set("")
    assert "Broken" in _row_labels(d)

    assert "need" in d.health.cget("text")
    d.toggle_problems()
    labels = _row_labels(d)
    assert "Broken" in labels and "Toggle" in labels and "Say Hello" not in labels
    assert "problems only" in d.health.cget("text")
    d.toggle_problems()
    assert "Say Hello" in _row_labels(d)


def test_ribbon_events_row_follows_the_filter(window, designer):
    d = designer
    part = window.draft.xml_part_by_path(PART)

    def mutate() -> None:
        part.text = part.text.replace("<customUI ", '<customUI onLoad="RibbonLoaded" ', 1)

    window._batch_edit(mutate)
    d.refresh()
    assert _row_labels(d)[0] == "Ribbon events"
    d.filter_var.set("hello")
    assert "Ribbon events" not in _row_labels(d)
    d.filter_var.set("RibbonLoaded")
    assert _row_labels(d) == ["Ribbon events"]
    d.filter_var.set("")
    d.toggle_problems()  # RibbonLoaded is missing: a problem of its own
    assert "Ribbon events" in _row_labels(d)


def test_select_start_clears_a_filter_that_hides_it(designer):
    d = designer
    d.filter_var.set("hello")
    _select(d, "tglX")
    assert d.filter_var.get() == ""
    assert d.selected_element().attr("id") == "tglX"


def test_select_offset_picks_the_innermost_control(window, designer):
    d = designer
    text = _text(window)
    assert d.select_offset(PART, text.index('id="tglX"'))
    assert d.selected_element().attr("id") == "tglX"
    assert d.select_offset(PART, text.index('<group id="grpEmpty"') + 1)
    assert d.selected_element().attr("id") == "grpEmpty"


def test_design_view_from_the_xml_cursor(window):
    window._open_xml_at(PART, 1)
    offset = window.xml_editor.get_text().index('onAction="NoSuchMacro"')
    window.xml_editor.text.mark_set("insert", f"1.0+{offset}c")
    window.show_ribbon_design()
    window.root.update()
    assert window._designer_shown()
    assert window.ribbon_designer.selected_element().attr("id") == "btnMissing"


def test_context_menu_states(designer):
    d = designer

    def states() -> dict[str, str]:
        menu = d.context_menu
        return {
            str(menu.entrycget(i, "label")): str(menu.entrycget(i, "state"))
            for i in range((menu.index("end") or 0) + 1)
            if menu.type(i) == "command"
        }

    d.prepare_context_menu(f"e{_el(d, 'grpMain').index}")
    assert d.selected_element().attr("id") == "grpMain"
    found = states()
    assert found["Go to Macro"] == "disabled"
    assert found["Delete"] == "normal" and found["Rename"] == "normal"
    d.prepare_context_menu(f"e{_el(d, 'btnHello').index}")
    found = states()
    assert found["Go to Macro"] == "normal" and found["Paste Icon from Clipboard"] == "normal"


def test_goto_macro_opens_the_code_view(window, designer):
    _select(designer, "btnHello")
    assert designer.goto_macro()
    window.root.update()
    assert window.draft.module_by_id(window.current_module_id).current_name == "Module1"
    line = int(window.editor.text.index("insert").split(".")[0])
    assert window.editor.get_text().split("\n")[line - 1].startswith("Public Sub Hello")


def test_designer_edits_save_and_reopen(window, designer):
    d = designer
    _select(d, "grpEmpty")
    assert d.add_control("button")
    _type(d.inspector.label, "Run Report")
    _type(d.inspector.macro, "Hello")
    _select(d, "btnHello")
    _type(d.inspector.label, "Greet")
    assert rs.new_issues(window.draft) == []

    result = make_service().save_addin(window.draft)
    assert result.kind == "success", result
    reopened = DocumentService().open(window.draft.baseline.path)
    assert "Public Sub HelloCallback(control As IRibbonControl)" in reopened.find_current("Module1").body
    with zipfile.ZipFile(window.draft.baseline.path) as package:
        saved = package.read(PART).decode("utf-8")
    assert 'id="btnRunReport" label="Run Report" size="large" onAction="HelloCallback"' in saved
    assert 'id="btnHello" label="Greet"' in saved
    # Untouched lines stay byte-identical, the comment included.
    assert '<!-- <button id="btnOld" onAction="Ghost" /> -->' in saved


# -- the PROCEDURES outline -----------------------------------------------------------------


def _open_module1(window, extra: str = ""):
    m1 = window.draft.find_current("Module1")
    if extra:
        m1.body += extra
    window._open_module_at(m1.id, 1)
    window.editor.set_text(m1.body)
    window._refresh_outline()
    return m1


def _outline(window) -> list[tuple[str, str]]:
    tree = window.outline
    return [(tree.item(i, "text"), tree.item(i, "values")[0]) for i in tree.get_children()]


def test_outline_rows_and_marks(window):
    _open_module1(
        window,
        "\nPublic Function Total() As Long\nEnd Function\n"
        "Private Property Get Count() As Long\nEnd Property\n"
        "Public Property Let Count(v As Long)\nEnd Property\n"
        "Public Sub Plain()\nEnd Sub\n",
    )
    assert [name for name, _mark in _outline(window)] == ["Hello", "Total", "Count", "Count", "Plain"]
    marks = [mark for _name, mark in _outline(window)]
    assert marks[0] == mw.icon("ribbon")  # btnHello runs it
    assert marks[1:] == ["fn", "get", "let", ""]
    private = [i for i in window.outline.get_children() if "locked" in window.outline.item(i, "tags")]
    assert [window.outline.item(i, "text") for i in private] == ["Count"]


def test_outline_follows_the_cursor(window):
    m1 = _open_module1(window, "\nPublic Sub Plain()\n    x = 1\nEnd Sub\n")
    lines = m1.body.split("\n")
    window.editor.goto_position(lines.index("    x = 1") + 1, 1)
    window._follow_outline()
    selected = window.outline.selection()
    assert [window.outline.item(i, "text") for i in selected] == ["Plain"]
    window.editor.goto_position(lines.index("Public Sub Hello()") + 2, 1)
    window._follow_outline()
    assert [window.outline.item(i, "text") for i in window.outline.selection()] == ["Hello"]


def test_add_procedure_from_the_outline(window, monkeypatch):
    m1 = _open_module1(window)
    before = m1.body
    monkeypatch.setattr(window, "_prompt_name", lambda *_a: "MakeReport")
    window.add_procedure()
    window._refresh_outline()
    body = window.draft.find_current("Module1").body
    assert "Public Sub MakeReport()\n    \nEnd Sub" in body
    assert "MakeReport" in [name for name, _mark in _outline(window)]
    assert "Added Sub MakeReport" in window.status.cget("text")
    window.undo()
    assert window.draft.find_current("Module1").body == before

    errors: list[str] = []
    monkeypatch.setattr(mw.messagebox, "showerror", lambda _t, msg, **k: errors.append(msg))
    monkeypatch.setattr(window, "_prompt_name", lambda *_a: "Hello")
    window.add_procedure()
    assert errors and "already has a procedure named Hello" in errors[0]
