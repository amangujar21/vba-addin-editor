"""UserForms: copy between add-ins byte for byte, create, edit layouts, save and verify."""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

import pytest
from conftest import (
    build_form_source,
    build_port_source,
    build_pptm,
    build_xlam,
    build_xlam_with_form,
)
from helpers import make_service
from pyopenvba import ExcelFile
from pyopenvba.cfb import CFB
from pyopenvba.vba import parse_project_stream

from vba_addin_editor.adapters import pyopenvba_adapter as pa
from vba_addin_editor.adapters.pyopenvba_adapter import PyOpenVBAAdapter, vba_entry_for
from vba_addin_editor.adapters.source_codec import fresh_designer_header
from vba_addin_editor.domain.changes import compute_changes
from vba_addin_editor.services import form_service as fs
from vba_addin_editor.services import port_service as ps
from vba_addin_editor.services import ribbon_service as rs
from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.services.recovery_service import (
    _form_storage_from_json,
    _form_storage_to_json,
)


def _cfb(path: Path) -> CFB:
    with zipfile.ZipFile(path) as package:
        return CFB.from_bytes(package.read(vba_entry_for(path)))


def _vb_base(header: str) -> str:
    match = re.search(r'VB_Base\s*=\s*"([^"]+)"', header)
    assert match is not None, header
    return match.group(1)


def _open(path: Path):
    return DocumentService().open(path)


def _save(draft) -> None:
    result = make_service().save_addin(draft)
    assert result.kind == "success", result


@pytest.fixture()
def source(tmp_path: Path):
    return _open(build_form_source(tmp_path / "src" / "FormSource.xlam"))


# -- copying between add-ins ---------------------------------------------------------------


@pytest.mark.parametrize("target_kind", ["xlam", "ppam", "pptm"])
def test_copied_form_keeps_its_layout_byte_for_byte(tmp_path: Path, source, target_kind: str):
    work = tmp_path / "work"
    work.mkdir()
    if target_kind == "xlam":
        target_path = build_xlam(work / "Target.xlam")
    else:
        built = build_pptm(work / "Target.pptm")
        target_path = built if target_kind == "pptm" else built.rename(work / "Target.ppam")
    target = _open(target_path)
    assert [f.name for f in ps.portable_forms(source)] == ["EntryForm"]

    plan = ps.plan_form_port(source, target, ["EntryForm"], suffix="_T", check_references=False)
    assert plan.ok, plan.describe()
    assert any("UserForm EntryForm" in line for line in plan.summary)
    assert any("Helpers" in line for line in plan.summary)  # the function the form calls
    ps.apply_port(target, plan)
    assert compute_changes(target).has_vba_changes
    _save(target)

    reopened = _open(target_path)
    form = reopened.find_current("EntryForm")
    assert form is not None and form.project_item_kind == "userform"
    assert form.can_delete
    assert reopened.find_current("Helpers") is not None
    assert "Greeting(txtName.Value)" in form.body
    # The layout is exactly the source's: every stream and every storage CLSID.
    src_storage = source.baseline.form("EntryForm").storage
    assert reopened.baseline.form("EntryForm").storage == src_storage
    assert reopened.baseline.form("EntryForm").control_count() == 3
    # Fresh class ids; declared a form; Microsoft Forms referenced.
    assert _vb_base(form.hidden_header) != _vb_base(source.find_current("EntryForm").hidden_header)
    cfb = _cfb(target_path)
    parsed = parse_project_stream(cfb.get_stream("PROJECT"), code_page=reopened.baseline.code_page)
    assert "EntryForm" in parsed.base_classes
    refs = PyOpenVBAAdapter().project_references(target_path)
    assert any("0d452ee1-e08f-101a-852e-02608c4d0bb4" in guid.casefold() for _n, guid in refs)


def test_second_copy_is_renamed_and_reuses_identical_code(tmp_path: Path, source):
    target_path = build_xlam(tmp_path / "work" / "Target.xlam")
    target = _open(target_path)
    ps.apply_port(target, ps.plan_form_port(source, target, ["EntryForm"], suffix="_T", check_references=False))
    _save(target)

    target = _open(target_path)
    plan = ps.plan_form_port(source, target, ["EntryForm"], suffix="_Two", check_references=False)
    assert plan.ok, plan.describe()
    assert ("module", "EntryForm", "EntryForm_Two") in plan.renames
    assert any("Reuses identical code" in line for line in plan.summary)
    ps.apply_port(target, plan)
    _save(target)

    reopened = _open(target_path)
    names = {f.name for f in reopened.baseline.forms}
    assert names == {"EntryForm", "EntryForm_Two"}
    first, second = reopened.find_current("EntryForm"), reopened.find_current("EntryForm_Two")
    assert _vb_base(first.hidden_header) != _vb_base(second.hidden_header)
    vbframe = reopened.baseline.form("EntryForm_Two").storage.stream("\x03VBFrame").decode("cp1252")
    assert re.search(r"^Begin \{[0-9A-F-]+\} EntryForm_Two", vbframe, re.MULTILINE)
    # Apart from the name in its header, the layout is the source's.
    src = source.baseline.form("EntryForm").storage
    assert pa.form_storage_named(src, "EntryForm_Two", 1252) == reopened.baseline.form("EntryForm_Two").storage


def test_ribbon_button_that_shows_a_form_brings_the_form(tmp_path: Path):
    src_path = build_port_source(tmp_path / "src" / "Source.xlam")
    with ExcelFile(src_path) as host:
        form = host.add_form("ReportForm", caption="Report")
        form.add_control("Label", "lblInfo")
        host.save()
    source = _open(src_path)
    helpers_mod = source.find_current("Helpers")
    helpers_mod.body = helpers_mod.body.replace("c.Bump", "c.Bump\n    ReportForm.Show")
    from conftest import build_xlam_with_ribbon

    target = _open(build_xlam_with_ribbon(tmp_path / "work" / "Ribbon.xlam"))
    control = next(c for c in ps.portable_controls(source) if c.control_id == "btnReport")
    where = next(t for t in ps.port_targets(target, control) if t.container.element == "group")
    plan = ps.plan_port(source, target, control, where, suffix="_T", check_references=False)
    assert plan.ok, plan.describe()
    assert any(line.startswith("UserForm ReportForm") for line in plan.summary)
    ps.apply_port(target, plan)
    _save(target)
    reopened = _open(target.baseline.path)
    assert reopened.baseline.form("ReportForm").control_count() == 1
    assert rs.new_issues(reopened) == []


def test_form_without_readable_layout_is_refused(tmp_path: Path, source):
    target = _open(build_xlam(tmp_path / "work" / "Target.xlam"))
    module = source.find_current("EntryForm")
    source.baseline.forms[0].__dict__["storage"] = None  # as if unreadable
    assert module is not None
    plan = ps.plan_form_port(source, target, ["EntryForm"], suffix="_T", check_references=False)
    assert not plan.ok and "could not be read" in plan.problems[0]


def test_same_file_and_missing_form_are_refused(tmp_path: Path, source):
    assert "this is the file" in ps.plan_form_port(source, source, ["EntryForm"], suffix="_T").problems[0]
    target = _open(build_xlam(tmp_path / "work" / "Target.xlam"))
    assert "no UserForm named" in ps.plan_form_port(source, target, ["Nope"], suffix="_T").problems[0]


# -- new forms and layout edits --------------------------------------------------------------


def test_new_blank_form_saves_as_a_form(tmp_path: Path):
    path = build_xlam(tmp_path / "work" / "Blank.xlam")
    draft = _open(path)
    name = fs.default_form_name(draft)
    assert name == "UserForm1"
    module = fs.add_blank_form(draft, name, "Hello")
    fs.edit_layout(draft, module, lambda e: e.add_control("CommandButton", "CommandButton1", left=12, top=12))
    _save(draft)

    reopened = _open(path)
    form = reopened.find_current("UserForm1")
    assert form is not None and form.project_item_kind == "userform"
    design = reopened.baseline.form("UserForm1")
    assert design.caption == "Hello"
    assert [c.name for c in design.controls] == ["CommandButton1"]


def test_layout_edits_save_exactly_and_verify(tmp_path: Path):
    path = build_xlam_with_form(tmp_path / "work" / "Form.xlam")
    draft = _open(path)
    module = draft.find_current("EntryForm")
    before_body = module.body

    def edit(e: pa.FormLayoutEditor) -> None:
        e.set_form_caption("Edited")
        e.set_form_size(300, 200)
        e.add_control("Label", "lblHint", left=6, top=150, width=120)
        e.set_caption("lblHint", "Type a name")
        e.set_geometry("cmdOK", left=200, top=160, width=72, height=24)
        e.set_tip("cmdOK", "Close the form")
        e.set_font("cmdOK", "Segoe UI", 9, True)
        e.rename_control("txtName", "txtFullName")

    edited = fs.edit_layout(draft, module, edit)
    assert draft.changed_form_layouts() == [module]
    assert compute_changes(draft).layouts_modified
    assert "UserForm layout" in "\n".join(compute_changes(draft).summary_lines())
    _save(draft)

    reopened = _open(path)
    design = reopened.baseline.form("EntryForm")
    assert design.storage == pa.form_storage_named(edited, "EntryForm", reopened.baseline.code_page)
    assert design.caption == "Edited" and design.width == 300 and design.height == 200
    cmd = design.control("cmdOK")
    assert (cmd.left, cmd.top, cmd.width, cmd.height) == (200, 160, 72, 24)
    assert cmd.tip == "Close the form" and cmd.font == ("Segoe UI", 9.0, True)
    assert design.control("lblHint").caption == "Type a name"
    assert design.control("txtFullName") is not None and design.control("txtName") is None
    assert reopened.find_current("EntryForm").body == before_body


def test_editing_back_to_the_original_layout_is_no_change(tmp_path: Path):
    draft = _open(build_xlam_with_form(tmp_path / "work" / "Form.xlam"))
    module = draft.find_current("EntryForm")
    fs.edit_layout(draft, module, lambda e: e.set_caption("cmdOK", "Go"))
    assert draft.is_dirty()
    fs.edit_layout(draft, module, lambda e: e.set_caption("cmdOK", "OK"))
    assert module.form_storage is None and not draft.is_dirty()


def test_bad_edits_are_refused_without_touching_the_draft(tmp_path: Path):
    draft = _open(build_xlam_with_form(tmp_path / "work" / "Form.xlam"))
    module = draft.find_current("EntryForm")
    with pytest.raises(fs.FormError):
        fs.edit_layout(draft, module, lambda e: e.add_control("CommandButton", "cmdOK"))
    with pytest.raises(fs.FormError):
        fs.edit_layout(draft, module, lambda e: e.remove_control("nope"))
    with pytest.raises(fs.FormError):
        fs.edit_layout(draft, module, lambda e: e.rename_control("cmdOK", "fraMain"))
    assert module.form_storage is None and not draft.is_dirty()


def test_verification_rejects_a_layout_other_than_the_drafts(tmp_path: Path, source):
    target = _open(build_xlam(tmp_path / "work" / "Target.xlam"))
    ps.apply_port(target, ps.plan_form_port(source, target, ["EntryForm"], suffix="_T", check_references=False))
    adapter = PyOpenVBAAdapter()
    data = target.baseline.package_bytes
    built = adapter.build_candidate_bytes(data, target, allow_signature_removal=False)
    assert adapter.verify_candidate(data, built, target).ok
    module = target.find_current("EntryForm")
    module.form_storage = pa.edit_form_storage(
        module.form_storage, "EntryForm", 1252, lambda e: e.set_caption("cmdOK", "Changed")
    )
    result = adapter.verify_candidate(data, built, target)
    assert not result.ok
    assert any("differs from the draft" in p for p in result.problems)


# -- code behind the form, persistence --------------------------------------------------------


def test_rename_in_code_follows_references_and_handlers():
    body = (
        "Private Sub cmdOK_Click()\n"
        '    cmdOK.Caption = "cmdOK"  \' cmdOK in a comment\n'
        "    Me.cmdOK.Enabled = False\n"
        "    cmdOKay = 1\n"
        "End Sub\n"
    )
    out = fs.rename_in_code(body, "cmdOK", "cmdGo")
    assert "Private Sub cmdGo_Click()" in out
    assert 'cmdGo.Caption = "cmdOK"' in out
    assert "Me.cmdGo.Enabled" in out
    assert "cmdOKay = 1" in out


def test_event_handlers():
    body, line = fs.ensure_handler("Option Explicit\n", "cmdOK", "Click")
    assert "Private Sub cmdOK_Click()" in body
    assert body.splitlines()[line - 2] == "Private Sub cmdOK_Click()"
    again, line2 = fs.ensure_handler(body, "cmdOK", "Click")
    assert again == body and line2 == line
    assert fs.find_handler(body, "cmdOK", "Click") == line - 1
    assert fs.event_for("TextBox") == "Change"


def test_designer_header_gets_new_ids_and_name():
    header = (
        'Attribute VB_Name = "frmA"\n'
        'Attribute VB_Base = "0{11111111-1111-1111-1111-111111111111}{22222222-2222-2222-2222-222222222222}"\n'
        "Attribute VB_PredeclaredId = True"
    )
    fresh = fresh_designer_header(header, "frmB")
    assert 'Attribute VB_Name = "frmB"' in fresh
    assert "1111" not in fresh and "2222" not in fresh
    assert re.search(r'VB_Base = "0\{[0-9A-F-]{36}\}\{[0-9A-F-]{36}\}"', fresh)
    assert "VB_PredeclaredId = True" in fresh


def test_recovery_keeps_a_form_storage(source):
    storage = source.baseline.form("EntryForm").storage
    assert _form_storage_from_json(_form_storage_to_json(storage)) == storage
    assert _form_storage_from_json(None) is None
    with pytest.raises(ValueError):
        _form_storage_from_json({"streams": [["../f", ""]], "storages": []})
    with pytest.raises(ValueError):
        _form_storage_from_json({"streams": [["o", ""]], "storages": []})  # no f stream


def test_storage_round_trips_through_a_project(source):
    storage = source.baseline.form("EntryForm").storage
    cfb = pa._scratch_cfb()
    pa._write_storage(cfb, [], "Copy", storage)
    again = CFB.from_bytes(cfb.to_bytes())
    assert pa._read_storage(again, ["Copy"]) == storage
    assert pa.read_form_design(storage, "Copy", 1252).control_count() == 3


def test_ppam_saves_with_pyopenvba_6_6(tmp_path: Path):
    """pyOpenVBA 6.3.5+ opens .ppam read only; the adapter's host must still save one."""
    built = build_pptm(tmp_path / "Deck.pptm")
    ppam = built.rename(tmp_path / "Deck.ppam")
    draft = _open(ppam)
    module = next(m for m in draft.modules if m.pyopenvba_kind == "standard")
    module.body += "\nPublic Sub Added()\nEnd Sub\n"
    _save(draft)
    assert "Public Sub Added()" in next(
        m.body for m in _open(ppam).modules if m.current_name == module.current_name
    )


def test_layout_edit_drops_type_info_version(tmp_path: Path):
    """Excel refuses an edited form whose header keeps TypeInfoVer (measured live)."""
    from vba_addin_editor.domain.document import FormStorage

    draft = _open(build_xlam_with_form(tmp_path / "work" / "Form.xlam"))
    module = draft.find_current("EntryForm")
    base = draft.form_storage_of(module)
    raw = base.stream("\x03VBFrame").decode("cp1252")
    with_version = raw.replace("End\r\n", "   TypeInfoVer     =   13\r\nEnd\r\n", 1)
    versioned = FormStorage(
        streams=tuple(
            (p, with_version.encode("cp1252") if p == "\x03VBFrame" else d) for p, d in base.streams
        ),
        storages=base.storages,
    )
    module.form_storage = versioned
    edited = fs.edit_layout(draft, module, lambda e: e.set_caption("cmdOK", "Go"))
    assert "TypeInfoVer" not in edited.stream("\x03VBFrame").decode("cp1252")
    assert pa.same_layout(versioned, pa.edit_form_storage(versioned, "EntryForm", 1252, lambda e: None))
