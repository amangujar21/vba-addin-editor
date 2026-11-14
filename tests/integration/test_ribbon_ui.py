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
    assert window.ribbon_designer.selected_element().attr("id") == "btnTidy"
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



def test_add_menu_with_macros_is_one_undo_step_and_saves(window):
    from vba_addin_editor.ui.ribbon_dialogs import MenuSpec

    m1 = window.draft.find_current("Module1")
    m1.body += (
        "\nPublic Sub ExportReport(control As IRibbonControl)\nEnd Sub\n"
        "\nPublic Sub ClearReport(control As IRibbonControl)\nEnd Sub\n"
        "\nPublic Sub NotACallback()\nEnd Sub\n"
    )
    ready = rs.ribbon_ready_procedures(window.draft)
    assert [p.name for p in ready] == ["ExportReport", "ClearReport"]
    part = window.draft.xml_part_by_path(PART)
    before = part.text
    group = next(c for c in rs.find_containers(part.text, PART) if c.control_id == "grpMain")
    spec = MenuSpec(group, "mnuReports", "Reports & More", "", "large", "", "", tuple(ready))
    window.apply_ribbon_menu(spec)
    window.root.update()

    assert (
        '          <menu id="mnuReports" label="Reports &amp; More" size="large">\n'
        '            <button id="btnExportReport" label="Export Report" onAction="ExportReport" />\n'
        '            <button id="btnClearReport" label="Clear Report" onAction="ClearReport" />\n'
        "          </menu>\n"
        "        </group>"
    ) in part.text
    assert window.ribbon_designer.selected_element().attr("id") == "mnuReports"
    assert rs.new_issues(window.draft) == []
    # The new menu is itself somewhere buttons and further menus can go.
    assert "mnuReports" in {c.control_id for c in rs.find_containers(part.text, PART)}

    window.undo()
    assert part.text == before
    window.redo()
    result = make_service().save_addin(window.draft)
    assert result.kind == "success", result
    with zipfile.ZipFile(window.draft.baseline.path) as package:
        assert 'id="mnuReports"' in package.read(PART).decode("utf-8")


def test_add_menu_dialog_defaults_to_the_container_under_the_cursor(window, monkeypatch):
    from vba_addin_editor.ui import ribbon_dialogs

    seen = {}

    def fake_wait(self):
        seen["container"] = self._selected_container().control_id
        seen["id"] = self.id_var.get()
        self.label_var.set("Tools")
        seen["id_after_label"] = self.id_var.get()
        self.id_var.set("mnuMore")
        seen["duplicate"] = self.validate()
        self.id_var.set("mnuTools")
        self.confirm()
        return self.result

    monkeypatch.setattr(ribbon_dialogs.AddMenuDialog, "wait", fake_wait)
    _cursor_on(window, 'id="tglX"')  # inside the existing "More" menu
    window.add_ribbon_menu()
    window.root.update()
    assert seen["container"] == "mnuMore"
    assert (seen["id"], seen["id_after_label"]) == ("mnuMenu", "mnuTools")
    assert "already exists" in seen["duplicate"]
    text = window.draft.xml_part_by_path(PART).text
    # Nested in a menu: no size attribute, and an empty menu is closed on its own line.
    assert (
        '            <menu id="mnuTools" label="Tools">\n'
        "            </menu>\n"
        "          </menu>"
    ) in text


def _png_icon():
    import io

    from PIL import Image

    from vba_addin_editor.services import icon_service

    out = io.BytesIO()
    Image.new("RGBA", (64, 40), (10, 160, 90, 255)).save(out, "PNG")
    return icon_service.icon_from_bytes(out.getvalue())


def test_alt_up_moves_control_and_undo_restores(window):
    part = window.draft.xml_part_by_path(PART)
    before = part.text
    _cursor_on(window, 'id="btnMissing"')
    assert window.xml_editor.text.bind("<Alt-Up>")  # the shortcut is bound on the editor
    assert window.move_ribbon_control(-1)
    window.root.update()
    assert part.text.index('id="btnMissing"') < part.text.index('id="btnHello"')
    assert window.xml_editor.get_text() == part.text
    assert window.xml_editor.text.get("sel.first", "sel.last").startswith('<button id="btnMissing"')
    # The cursor stays on the moved control, so it can go up again: it is first now.
    assert not window.move_ribbon_control(-1)
    assert "already the first" in window.status.cget("text")
    window.undo()
    assert part.text == before


def test_move_buttons_need_ribbon_xml(window):
    window._open_xml_at("[Content_Types].xml", 1)
    assert not window.move_ribbon_control(1)
    assert "ribbon XML" in window.status.cget("text")


def test_context_menu_has_structure_actions(window):
    xml_menu = window.xml_editor.text._vbaae_context_menu
    labels = [
        xml_menu.menu.entrycget(i, "label")
        for i in range(xml_menu.menu.index("end") + 1)
        if xml_menu.menu.type(i) == "command"
    ]
    for label in (
        "Move Up (Alt+Up)",
        "Move Down (Alt+Down)",
        "Group Selected into Menu…",
        "Set Icon from Clipboard",
        "Set Icon from File…",
    ):
        assert label in labels


def test_group_selection_into_menu_with_icon_saves(window, monkeypatch):
    from vba_addin_editor.ui.ribbon_dialogs import GroupMenuSpec

    part = window.draft.xml_part_by_path(PART)
    before = part.text
    window._open_xml_at(PART, 1)
    text = window.xml_editor.get_text()
    start = text.index('<button id="btnHello"')
    end = text.index("NoSuchMacro")
    window.xml_editor.text.tag_add("sel", f"1.0+{start}c", f"1.0+{end}c")
    window.xml_editor.text.mark_set("insert", f"1.0+{end}c")
    seen = {}

    class FakeDialog:
        def __init__(self, _root, **kwargs):
            seen.update(kwargs)

        def wait(self):
            return GroupMenuSpec("mnuTools", "Tools", "", "large", "", _png_icon())

    monkeypatch.setattr(mw, "GroupMenuDialog", FakeDialog)
    window.group_selection_into_menu()
    window.root.update()
    assert seen["controls"] == ['button "Say Hello"', 'button "Broken"']
    assert seen["in_group"] is True
    assert '<menu id="mnuTools" label="Tools" size="large" image="rIdImg1">' in part.text
    assert 'label="Say Hello" onAction="Hello" />' in part.text  # size dropped inside the menu
    assert window.xml_tree.exists("xml::customUI/_rels/customUI14.xml.rels")

    window.undo()
    assert part.text == before
    assert window.draft.binary_parts == [] and window.draft.new_xml_parts() == []
    window.redo()
    assert len(window.draft.binary_parts) == 1

    result = make_service().save_addin(window.draft)
    assert result.kind == "success", result
    with zipfile.ZipFile(window.draft.baseline.path) as package:
        assert "customUI/images/mnuTools.png" in package.namelist()


def test_set_icon_from_clipboard_on_cursor_control(window, monkeypatch):
    from vba_addin_editor.services import icon_service

    errors = []
    monkeypatch.setattr(mw.messagebox, "showerror", lambda _t, msg, **k: errors.append(msg))
    monkeypatch.setattr(icon_service, "icon_from_clipboard", _png_icon)
    _cursor_on(window, 'id="btnHello"')
    window.set_ribbon_icon_from_clipboard()
    window.root.update()
    assert not errors
    part = window.draft.xml_part_by_path(PART)
    assert 'onAction="Hello" image="rIdImg1"' in part.text
    assert "Set a custom icon on button" in window.status.cget("text")
    assert window.xml_editor.text.get("sel.first", "sel.last").startswith('<button id="btnHello"')

    # Nothing on the clipboard: an error, and the draft is unchanged.
    def empty():
        raise icon_service.IconError("The clipboard holds no image.")

    monkeypatch.setattr(icon_service, "icon_from_clipboard", empty)
    text = part.text
    window.set_ribbon_icon_from_clipboard()
    assert errors == ["The clipboard holds no image."]
    assert part.text == text


def test_add_button_with_custom_icon(window):
    m1 = window.draft.find_current("Module1")
    part = window.draft.xml_part_by_path(PART)
    group = next(c for c in rs.find_containers(part.text, PART) if c.control_id == "grpMain")
    spec = ButtonSpec(group, "btnPic", "Picture", "", "large", "", "", _png_icon())
    window.apply_ribbon_button(m1.id, "Hello", None, spec)
    window.root.update()
    assert '<button id="btnPic" label="Picture" size="large" onAction="Hello" image="rIdImg1" />' in part.text
    assert [b.path for b in window.draft.binary_parts] == ["customUI/images/btnPic.png"]


def test_set_icon_explains_in_a_dialog_when_it_cannot_run(window, monkeypatch):
    from vba_addin_editor.services import icon_service

    infos = []
    monkeypatch.setattr(mw.messagebox, "showinfo", lambda _t, msg, **k: infos.append(msg))
    monkeypatch.setattr(icon_service, "icon_from_clipboard", _png_icon)
    before = window.draft.xml_part_by_path(PART).text

    # Cursor on the root element, not on a control that can show an icon.
    _cursor_on(window, "<customUI")
    window.set_ribbon_icon_from_clipboard()
    assert len(infos) == 1 and "Put the cursor inside" in infos[0]

    # A VBA module is showing instead of the ribbon XML.
    window.current_xml_part_path = None
    window.set_ribbon_icon_from_clipboard()
    assert len(infos) == 2 and "Open the ribbon XML" in infos[1]
    assert window.draft.xml_part_by_path(PART).text == before


def test_opening_a_file_with_060_icon_links_offers_the_repair(window, monkeypatch):
    window.apply_ribbon_icon(PART, window.draft.xml_part_by_path(PART).text.index("<button"), _png_icon())
    rels_path = "customUI/_rels/customUI14.xml.rels"
    rels = window.draft.xml_part_by_path(rels_path)
    rels.text = rels.text.replace(rs.IMAGE_REL_TYPE, rs.LEGACY_IMAGE_REL_TYPE)
    assert window.save_service.save_addin(window.draft).kind == "success"

    asked = []
    monkeypatch.setattr(mw.messagebox, "askyesno", lambda _t, msg, **k: asked.append(msg) or True)
    window.load_path(window.draft.baseline.path)
    assert len(asked) == 1 and "shows blank" in asked[0]
    assert rs.LEGACY_IMAGE_REL_TYPE not in window.draft.xml_part_by_path(rels_path).text
    assert window.draft.is_dirty()  # repaired in the draft; written only on save


def test_remove_unused_icons_command(window, monkeypatch):
    infos = []
    monkeypatch.setattr(mw.messagebox, "showinfo", lambda _t, msg, **k: infos.append(msg))
    assert not window.remove_unused_ribbon_icons()
    assert infos == ["There are no unused icons in this file."]

    start = window.draft.xml_part_by_path(PART).text.index("<button")
    window.apply_ribbon_icon(PART, start, _png_icon())
    assert window.save_service.save_addin(window.draft).kind == "success"
    ribbon = window.draft.xml_part_by_path(PART)
    ribbon.text = ribbon.text.replace(' image="rIdImg1"', "")
    window._reload_editors_from_draft()

    assert window.remove_unused_ribbon_icons()
    assert window.draft.removed_member_paths() == ["customUI/images/btnHello.png"]
    assert "Removed 1 unused image" in window.status.cget("text")
    window.undo()
    assert window.draft.removed_parts == []
