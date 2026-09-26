from __future__ import annotations

from pathlib import Path

from conftest import RIBBON, build_xlam_with_ribbon

from vba_addin_editor.domain.document import ModuleDisplayKind, ModuleDraft, new_module_id
from vba_addin_editor.services import ribbon_service as rs
from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.services.search_service import parse_procedures


def _open(tmp_path: Path, xml: str = RIBBON):
    return DocumentService().open(build_xlam_with_ribbon(tmp_path / "Ribbon.xlam", xml))


def _add_module(draft, name: str, body: str, *, standard: bool = True) -> ModuleDraft:
    mod = ModuleDraft(
        id=new_module_id(),
        origin_name=None,
        current_name=name,
        body=body,
        kind=ModuleDisplayKind.STANDARD if standard else ModuleDisplayKind.CLASS,
        pyopenvba_kind="standard" if standard else "other",
        is_new=True,
        is_deleted=False,
        destructive_ops_safe=True,
        can_delete=True,
        can_rename=True,
        project_item_kind="standard" if standard else "class",
    )
    draft.modules.append(mod)
    return mod


def test_find_callbacks_skips_comments_and_reports_positions():
    callbacks = rs.find_callbacks(RIBBON, "customUI/customUI14.xml")
    names = [cb.name for cb in callbacks]
    assert names == ["Hello", "NoSuchMacro", "OnToggle", "GetToggle"]
    hello = callbacks[0]
    assert RIBBON[hello.offset : hello.offset + hello.length] == "Hello"
    assert hello.label == "Say Hello" and hello.control_id == "btnHello"
    assert hello.line == 6
    assert RIBBON.split("\n")[hello.line - 1][hello.column - 1 :].startswith("Hello")


def test_callback_at_value_and_tag():
    hello = RIBBON.index('onAction="Hello"') + len('onAction="')
    assert rs.callback_at(RIBBON, hello + 2).name == "Hello"
    # Cursor on the tag name, not an attribute: prefer onAction.
    tag = RIBBON.index("<toggleButton") + 3
    assert rs.callback_at(RIBBON, tag).name == "OnToggle"
    assert rs.callback_at(RIBBON, RIBBON.index("<tabs>") + 2) is None


def test_resolve_statuses(tmp_path: Path):
    draft = _open(tmp_path)
    _add_module(draft, "Ribbon", "Public Sub OnToggle(control As IRibbonControl, pressed As Boolean)\nEnd Sub\n")
    _add_module(draft, "Other", "Sub OnToggle()\nEnd Sub\n")
    index = rs.procedure_index(draft)
    assert rs.resolve(index, "Hello").status == rs.STATUS_OK
    assert rs.resolve(index, "Hello").matches[0].module_name == "Module1"
    assert rs.resolve(index, "NoSuchMacro").status == rs.STATUS_MISSING
    assert rs.resolve(index, "OnToggle").status == rs.STATUS_AMBIGUOUS
    assert rs.resolve(index, "Ribbon.OnToggle").status == rs.STATUS_OK
    assert rs.resolve(index, "'My Addin.xlam'!Module1.Hello").status == rs.STATUS_OK
    assert rs.resolve(index, "Module1.OnToggle").status == rs.STATUS_MISSING
    # Class module procedures cannot be ribbon callbacks.
    assert rs.resolve(index, "Ping").status == rs.STATUS_NOT_CALLABLE
    assert rs.resolve(index, "OrdinaryClass.Ping").status == rs.STATUS_NOT_CALLABLE


def test_check_ribbon_and_new_issues(tmp_path: Path):
    draft = _open(tmp_path)
    messages = [i.message for i in rs.check_ribbon(draft)]
    assert any("NoSuchMacro" in m for m in messages)
    assert any("OnToggle" in m for m in messages)
    assert rs.new_issues(draft) == []  # pre-existing problems do not nag on save
    part = rs.ribbon_parts(draft)[0]
    part.text = part.text.replace('id="tglX"', 'id="btnHello"')
    new = rs.new_issues(draft)
    assert len(new) == 1 and "Duplicate control id" in new[0].message


def test_find_containers_and_insert_matches_indentation():
    containers = rs.find_containers(RIBBON, "p")
    assert [c.control_id for c in containers] == ["grpMain", "grpMenus", "mnuMore", "grpEmpty"]
    assert containers[2].display == "Tools › Menus › More"
    xml = rs.build_button_xml(control_id="btnNew", label="New", on_action="DoNew", size="large")
    text, offset, length = rs.insert_into_container(RIBBON, containers[0], xml)
    assert text[offset : offset + length] == xml
    line = text.split("\n")[text.count("\n", 0, offset)]
    assert line == "          " + xml
    assert text.split("\n")[text.count("\n", 0, offset) + 1].strip() == "</group>"
    # Empty single-line group: closing tag keeps its own line and indent.
    text, offset, _ = rs.insert_into_container(RIBBON, containers[3], xml)
    assert '<group id="grpEmpty" label="Empty">\n          ' + xml + "\n        </group>" in text
    assert rs.find_containers(text, "p")  # still scans cleanly


def test_build_button_xml_escapes():
    xml = rs.build_button_xml(
        control_id="b", label='Fish & "Chips" <1>', on_action="X", supertip="a&b"
    )
    assert 'label="Fish &amp; &quot;Chips&quot; &lt;1&gt;"' in xml
    assert 'supertip="a&amp;b"' in xml and "size=" not in xml


def test_names_and_stubs():
    assert rs.unique_control_id({"btnHello"}, "Hello") == "btnHello2"
    assert rs.unique_control_id(set(), "stack_Sheets") == "btnStack_Sheets"
    assert rs.suggested_label("InsertSourceNoteCallback") == "Insert Source Note"
    assert rs.callback_stub("T", "onAction", "toggleButton").startswith(
        "Public Sub T(control As IRibbonControl, pressed As Boolean)"
    )
    assert "ByRef returnedVal" in rs.callback_stub("G", "getLabel", "button")
    assert "index As Integer" in rs.callback_stub("G", "getItemLabel", "dropDown")
    assert rs.callback_stub("L", "onLoad", "customUI").startswith("Public Sub L(ribbon As IRibbonUI)")
    wrapper = rs.callback_stub("GoCallback", "onAction", "button", calls="Go")
    assert wrapper == "Public Sub GoCallback(control As IRibbonControl)\n    Go\nEnd Sub\n"
    body, line = rs.append_procedure("Sub A()\nEnd Sub\n\n\n", wrapper)
    assert body.split("\n")[line - 1].startswith("Public Sub GoCallback")
    assert rs.append_procedure("", wrapper) == (wrapper, 1)
    assert rs.is_ribbon_ready("Public Sub X(ByVal control As Office.IRibbonControl)")
    assert not rs.is_ribbon_ready("Sub X()") and not rs.takes_arguments("Sub X()")
    assert rs.takes_arguments("Sub X(a As Long)")


def test_procedure_lines_survive_continuations():
    body = "Sub A(x As Long, _\n      y As Long)\nEnd Sub\n\nSub B()\nEnd Sub\n"
    procs = parse_procedures(body, "m")
    assert [(p.name, p.line) for p in procs] == [("A", 1), ("B", 5)]
    assert "y As Long" in procs[0].signature
    assert rs.procedure_at(body, "m", 6).name == "B"
    assert rs.procedure_at(body, "m", 2).name == "A"


def test_callbacks_for_procedure(tmp_path: Path):
    draft = _open(tmp_path)
    m1 = draft.find_current("Module1")
    found = rs.callbacks_for_procedure(draft, m1.id, "hello")
    assert [cb.control_id for cb in found] == ["btnHello"]


def test_ribbon_check_skips_module_parsing_without_ribbon(work_xlam: Path, monkeypatch):
    # Pre-save check must not parse every module when there is no customUI.
    draft = DocumentService().open(work_xlam)
    monkeypatch.setattr(
        rs, "_procedures", lambda *_a: (_ for _ in ()).throw(AssertionError("parsed"))
    )
    assert rs.ribbon_entries(draft) == []
    assert rs.new_issues(draft) == []
