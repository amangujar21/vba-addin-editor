"""Porting a ribbon control and its macros from one add-in into another."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from conftest import PORT_ICON_PNG, build_xlam_with_ribbon
from conftest import build_port_source as build_source
from helpers import make_service

from vba_addin_editor.domain.document import ModuleDisplayKind, ModuleDraft, new_module_id
from vba_addin_editor.services import port_service as ps
from vba_addin_editor.services import ribbon_service as rs
from vba_addin_editor.services.document_service import DocumentService

PART = "customUI/customUI14.xml"
RELS = "customUI/_rels/customUI14.xml.rels"

def _add_module(draft, name: str, body: str) -> ModuleDraft:
    mod = ModuleDraft(
        id=new_module_id(),
        origin_name=None,
        current_name=name,
        body=body,
        kind=ModuleDisplayKind.STANDARD,
        pyopenvba_kind="standard",
        is_new=True,
        is_deleted=False,
        destructive_ops_safe=True,
        can_delete=True,
        can_rename=True,
        project_item_kind="standard",
    )
    draft.modules.append(mod)
    return mod


@pytest.fixture()
def pair(tmp_path: Path):
    source = DocumentService().open(build_source(tmp_path / "src" / "Source.xlam"))
    target = DocumentService().open(build_xlam_with_ribbon(tmp_path / "work" / "Ribbon.xlam"))
    return source, target


def _control(source, control_id: str) -> ps.PortableControl:
    return next(c for c in ps.portable_controls(source) if c.control_id == control_id)


def _target(target, control, display: str, *, new_group: bool = False) -> ps.PortTarget:
    return next(
        t
        for t in ps.port_targets(target, control)
        if t.container.display.endswith(display) and t.new_group == new_group
    )


def _body(draft, name: str) -> str:
    module = draft.find_current(name)
    assert module is not None, [m.current_name for m in draft.modules]
    return module.body


def test_lists_controls_that_run_macros(pair):
    source, _target_draft = pair
    controls = {c.control_id: c for c in ps.portable_controls(source)}
    assert set(controls) == {"grpReports", "btnReport", "btnOther", "mnuMore", "btnMenuItem"}
    assert controls["btnReport"].display == "Source › Reports › Run Report"
    assert controls["btnReport"].callbacks == ("RunReport",)


def test_port_copies_the_closure_and_renames(pair):
    source, target = pair
    # The target already has a different BuildReport, and a module named Helpers.
    _add_module(target, "Helpers", "Public Sub BuildReport()\nEnd Sub\n")
    control = _control(source, "btnReport")
    where = _target(target, control, "Empty")
    plan = ps.plan_port(source, target, control, where, suffix="_Ribbon", check_references=False)
    assert plan.ok, plan.describe()
    draft = plan.draft

    code = _body(draft, "RibbonCode")
    assert "Public Sub RunReport_Ribbon(control As IRibbonControl)" in code
    assert "' Runs the report." in code  # its comment came along
    assert "Public Sub AfterReport_Ribbon()" in code  # run by name: renamed
    assert 'Application.OnTime Now, "AfterReport_Ribbon"' in code
    assert "Application.Run \"'Ribbon.xlam'!RibbonCode.AfterReport_Ribbon\"" in code
    assert "BuildReport_Ribbon APP_TITLE" in code
    assert 'Helpers_Ribbon.BuildReport_Ribbon "x"' in code
    assert "Private gRibbon As IRibbonUI" in code and "Public Const APP_TITLE" in code
    assert code.startswith("' Imported from Source.xlam")
    assert "Option Explicit" in code
    assert "Unrelated" not in code and "OnLoadRibbon" not in code
    # The ribbon variable is set by a hook that becomes the target's onLoad.
    assert "Public Sub RibbonLoaded_Ribbon(loadedRibbon As IRibbonUI)\n    Set gRibbon = loadedRibbon\nEnd Sub" in code

    helpers = _body(draft, "Helpers_Ribbon")
    assert "#If VBA7 Then\nPrivate Declare PtrSafe Function GetTickCount" in helpers
    assert "#End If" in helpers
    assert "Public Type ReportRow\n    Name As String" in helpers  # member names untouched
    assert "Public Enum ReportKind" in helpers
    assert "Public Sub BuildReport_Ribbon(title As String)" in helpers
    assert "r.Name = title" in helpers
    assert "Utility()" in helpers and "Public Function Utility() As Long" in helpers
    assert "NotNeeded" not in helpers
    # Declarations come before procedures.
    assert helpers.index("End Enum") < helpers.index("Public Sub BuildReport_Ribbon")

    counter = draft.find_current("Counter")
    assert counter is not None and counter.pyopenvba_kind == "other" and "Public Sub Bump()" in counter.body
    assert "VB_Base" in counter.hidden_header

    ribbon = draft.xml_part_by_path(PART).text
    assert '<button id="btnReport" label="Run Report" size="large" onAction="RunReport_Ribbon" image="rIdImg1" />' in ribbon
    assert 'onLoad="RibbonLoaded_Ribbon"' in ribbon
    assert draft.binary_part_by_path("customUI/images/btnReport.png") is not None
    assert any("btnOther" in w for w in plan.warnings)  # InvalidateControl of a control not copied
    renames = {(old, new) for _what, old, new in plan.renames}
    assert {("RunReport", "RunReport_Ribbon"), ("Helpers", "Helpers_Ribbon"), ("BuildReport", "BuildReport_Ribbon")} <= renames
    # The target draft itself is untouched until apply.
    assert target.find_current("RibbonCode") is None
    assert rs.new_issues(draft) == []


def test_ported_add_in_saves_and_reopens(pair):
    source, target = pair
    control = _control(source, "btnReport")
    plan = ps.plan_port(
        source, target, control, _target(target, control, "Main"), suffix="_Ribbon", check_references=False
    )
    assert plan.ok, plan.describe()
    ps.apply_port(target, plan)
    result = make_service().save_addin(target)
    assert result.kind == "success", result
    reopened = DocumentService().open(target.baseline.path)
    assert "RunReport_Ribbon" in _body(reopened, "RibbonCode")
    counter = reopened.find_current("Counter")
    assert counter is not None and counter.kind == ModuleDisplayKind.CLASS
    ribbon = reopened.xml_part_by_path(PART).text
    assert 'onAction="RunReport_Ribbon"' in ribbon
    with zipfile.ZipFile(target.baseline.path) as package:
        assert package.read("customUI/images/btnReport.png") == PORT_ICON_PNG
        assert 'Extension="png"' in package.read("[Content_Types].xml").decode("utf-8")


def test_existing_onload_calls_the_hook(tmp_path: Path):
    source = DocumentService().open(build_source(tmp_path / "src" / "Source.xlam"))
    xml = rs_ribbon_with_onload()
    target = DocumentService().open(build_xlam_with_ribbon(tmp_path / "work" / "Ribbon.xlam", xml))
    _add_module(target, "Startup", "Public Sub StartRibbon(ByVal rib As IRibbonUI)\n    Debug.Print 1\nEnd Sub\n")
    control = _control(source, "btnReport")
    plan = ps.plan_port(
        source, target, control, _target(target, control, "Empty"), suffix="_T", check_references=False
    )
    assert plan.ok, plan.describe()
    startup = _body(plan.draft, "Startup")
    assert startup == "Public Sub StartRibbon(ByVal rib As IRibbonUI)\n    Debug.Print 1\n    RibbonLoaded_T rib\nEnd Sub\n"
    assert 'onLoad="StartRibbon"' in plan.draft.xml_part_by_path(PART).text


def rs_ribbon_with_onload() -> str:
    from conftest import RIBBON

    return RIBBON.replace("<customUI ", '<customUI onLoad="StartRibbon" ', 1)


def test_identical_code_is_reused_and_ids_renamed(pair):
    source, target = pair
    # The target already has Utility, identically, and a control called btnReport.
    _add_module(target, "Shared", "Public Function Utility() As Long\n    Utility = 42\nEnd Function\n")
    part = target.xml_part_by_path(PART)
    part.text = part.text.replace('id="btnHello"', 'id="btnReport"')
    control = _control(source, "btnReport")
    plan = ps.plan_port(
        source, target, control, _target(target, control, "Empty"), suffix="_T", check_references=False
    )
    assert plan.ok, plan.describe()
    helpers = _body(plan.draft, "Helpers")
    assert "Function Utility" not in helpers and "Utility()" in helpers
    assert any("Utility (Shared)" in s for s in plan.summary)
    ribbon = plan.draft.xml_part_by_path(PART).text
    assert '<button id="btnReport_T" label="Run Report"' in ribbon
    # The control id string in the code follows the rename.
    assert 'gRibbon.InvalidateControl "btnReport_T"' in _body(plan.draft, "RibbonCode")


def test_menu_item_into_group_and_new_group(pair):
    source, target = pair
    item = _control(source, "btnMenuItem")
    plan = ps.plan_port(
        source, target, item, _target(target, item, "Tools", new_group=True),
        suffix="_T", group_label="Imported", check_references=False,
    )
    assert plan.ok, plan.describe()
    ribbon = plan.draft.xml_part_by_path(PART).text
    assert 'description=' not in ribbon.split('id="grpBtnMenuItem"')[1]
    assert '<group id="grpBtnMenuItem" label="Imported">' in ribbon
    assert any("description" in w for w in plan.warnings)
    # The button with a size cannot go into a menu with it.
    button = _control(source, "btnReport")
    plan = ps.plan_port(
        source, target, button, _target(target, button, "More"), suffix="_T", check_references=False
    )
    assert plan.ok, plan.describe()
    menu_xml = plan.draft.xml_part_by_path(PART).text
    assert '<button id="btnReport" label="Run Report" onAction="RunReport_T"' in menu_xml


def test_whole_group_goes_onto_a_tab(pair):
    source, target = pair
    group = _control(source, "grpReports")
    targets = ps.port_targets(target, group)
    assert [t.container.element for t in targets] == ["tab"]
    plan = ps.plan_port(source, target, group, targets[0], suffix="_T", check_references=False)
    assert plan.ok, plan.describe()
    code = _body(plan.draft, "RibbonCode")
    assert "Public Sub Unrelated_T(control As IRibbonControl)" in code
    ribbon = plan.draft.xml_part_by_path(PART).text
    assert '<group id="grpReports" label="Reports">' in ribbon
    assert ribbon.count('onAction="RunReport_T"') == 2
    assert rs.new_issues(plan.draft) == []


def test_refusals(pair, tmp_path: Path):
    source, target = pair
    control = _control(source, "btnReport")
    where = _target(target, control, "Empty")
    assert "suffix" in ps.plan_port(source, target, control, where, suffix="bad-suffix").problems[0]
    assert "this is the file" in ps.plan_port(target, target, control, where, suffix="_T").problems[0]
    # A UserForm the macro uses cannot be copied.
    module = source.find_current("Helpers")
    module.body = module.body.replace("c.Bump", "c.Bump\n    MyForm.Show")
    form = _add_module(source, "MyForm", "")
    form.pyopenvba_kind = "other"
    form.project_item_kind = "userform"
    plan = ps.plan_port(source, target, control, where, suffix="_T", check_references=False)
    assert not plan.ok and "UserForm MyForm" in plan.problems[0]


def test_cross_host_port_warns(tmp_path: Path):
    source = DocumentService().open(build_source(tmp_path / "src" / "Source.pptm", powerpoint=True))
    target = DocumentService().open(build_xlam_with_ribbon(tmp_path / "work" / "Ribbon.xlam"))
    control = _control(source, "btnReport")
    plan = ps.plan_port(
        source, target, control, _target(target, control, "Empty"), suffix="_T", check_references=False
    )
    assert plan.ok, plan.describe()
    assert any("Powerpoint file" in w or "PowerPoint" in w for w in plan.warnings)


def test_lexer_and_items():
    body = 'Dim a As Long, b(1 To 2, 3) As String _\n    \nRem note\nPublic Sub X(): a = &HFF: Debug.Print "a""b" \' c\nEnd Sub\n'
    toks = ps._lex(body)
    assert [t.text for t in toks if t.kind == "str"] == ['a"b']
    assert not any(t.kind == "id" and t.text in ("note", "c", "HFF") for t in toks)
    items = ps._parse_items(body, toks)
    assert [sorted(i.names) for i in items] == [["a", "b"], ["x"]]
    assert items[1].kind == "proc"
