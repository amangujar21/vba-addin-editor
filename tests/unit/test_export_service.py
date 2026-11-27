"""Exporting several things at once: tabs, controls, plain macros and whole modules."""

from __future__ import annotations

from pathlib import Path

import pytest
from conftest import PORT_ICON_PNG, build_xlam, build_xlam_with_ribbon
from conftest import build_port_source as build_source
from helpers import make_service

from vba_addin_editor.domain.document import BinaryPartDraft
from vba_addin_editor.services import port_service as ps
from vba_addin_editor.services import ribbon_service as rs
from vba_addin_editor.services.document_service import DocumentService

PART = "customUI/customUI14.xml"


@pytest.fixture()
def pair(tmp_path: Path):
    source = DocumentService().open(build_source(tmp_path / "src" / "Source.xlam"))
    target = DocumentService().open(build_xlam_with_ribbon(tmp_path / "work" / "Ribbon.xlam"))
    return source, target


def _control(source, control_id: str) -> ps.PortableControl:
    return next(c for c in ps.portable_controls(source) if c.control_id == control_id)


def _where(target, control, display: str = "", *, new_group: bool = False) -> ps.PortTarget:
    return next(
        t
        for t in ps.port_targets(target, control)
        if t.container.display.endswith(display) and t.new_group == new_group
    )


def _body(draft, name: str) -> str:
    module = draft.find_current(name)
    assert module is not None, [m.current_name for m in draft.modules]
    return module.body


def test_whole_tab_becomes_a_new_tab(pair):
    source, target = pair
    tab = _control(source, "tabSrc")
    assert tab.element == "tab"
    targets = ps.port_targets(target, tab)
    assert [t.container.element for t in targets] == ["tabs"]
    assert targets[0].describe() == "New tab on the ribbon"
    plan = ps.plan_export(
        source, target, [ps.ExportItem.of_control(tab, targets[0])], suffix="_T", check_references=False
    )
    assert plan.ok, plan.describe()
    ribbon = plan.draft.xml_part_by_path(PART).text
    elements = rs.element_tree(ribbon)
    tabs = [rs.unescape(e.attr("id") or "") for e in elements if rs.local_name(e.name) == "tab"]
    assert tabs == ["tabTools", "tabSrc"]
    # Every control, its callbacks renamed, its image and imageMso along with it.
    assert ribbon.count('onAction="RunReport_T"') == 2
    assert 'onAction="Unrelated_T"' in ribbon and 'imageMso="Help"' in ribbon
    assert plan.draft.binary_part_by_path("customUI/images/btnReport.png") is not None
    assert 'mnuMore_T' in ribbon  # the target already has an mnuMore
    assert rs.new_issues(plan.draft) == []
    # The first control's place is reported, for showing it after the import.
    assert plan.part_path == PART and ribbon[plan.element_offset :].startswith('<tab id="tabSrc"')


def test_several_controls_share_a_new_group_and_code_once(pair):
    source, target = pair
    report, other = _control(source, "btnReport"), _control(source, "btnOther")
    where = _where(target, report, "Tools", new_group=True)
    items = [ps.ExportItem.of_control(report, where), ps.ExportItem.of_control(other, where)]
    plan = ps.plan_export(source, target, items, suffix="_T", group_label="From Source", check_references=False)
    assert plan.ok, plan.describe()
    ribbon = plan.draft.xml_part_by_path(PART).text
    assert ribbon.count("<group ") == 4  # Main, Menus, Empty + one new group
    group = ribbon.split('<group id="grpBtnReport" label="From Source">')[1].split("</group>")[0]
    assert 'id="btnReport"' in group and 'id="btnOther"' in group
    code = _body(plan.draft, "RibbonCode")
    assert code.count("Sub RunReport_T(") == 1 and "Sub Unrelated_T(" in code
    assert sum(1 for m in plan.draft.modules if m.current_name.startswith("RibbonCode")) == 1
    assert rs.new_issues(plan.draft) == []


def test_controls_into_different_places(pair):
    source, target = pair
    report, item = _control(source, "btnReport"), _control(source, "btnMenuItem")
    items = [
        ps.ExportItem.of_control(report, _where(target, report, "Empty")),
        ps.ExportItem.of_control(item, _where(target, item, "More")),
    ]
    plan = ps.plan_export(source, target, items, suffix="_T", check_references=False)
    assert plan.ok, plan.describe()
    ribbon = plan.draft.xml_part_by_path(PART).text
    empty = ribbon.split('<group id="grpEmpty" label="Empty">')[1].split("</group>")[0]
    menu = ribbon.split('<menu id="mnuMore" label="More">')[1].split("</menu>")[0]
    assert 'id="btnReport"' in empty and 'id="btnMenuItem"' in menu
    assert rs.new_issues(plan.draft) == []


def test_control_inside_a_chosen_group_is_copied_once(pair):
    source, target = pair
    group, button = _control(source, "grpReports"), _control(source, "btnReport")
    items = [
        ps.ExportItem.of_control(button, _where(target, button, "Empty")),
        ps.ExportItem.of_control(group, _where(target, group)),
    ]
    plan = ps.plan_export(source, target, items, suffix="_T", check_references=False)
    assert plan.ok, plan.describe()
    assert plan.draft.xml_part_by_path(PART).text.count('id="btnReport"') == 1


def test_plain_macro_gets_a_new_name_with_what_it_needs(pair):
    source, target = pair
    macros = {m.display: m for m in ps.portable_macros(source)}
    assert macros["RibbonCode.RunReport"].ribbon and not macros["RibbonCode.AfterReport"].ribbon
    assert "Helpers.BuildReport" in macros
    plan = ps.plan_export(
        source, target, [ps.ExportItem.of_macro("RibbonCode", "AfterReport")], suffix="_T",
        check_references=False,
    )
    assert plan.ok, plan.describe()
    code = _body(plan.draft, "RibbonCode")
    assert "Public Sub AfterReport_T()" in code
    assert "RunReport" not in code
    assert plan.draft.xml_part_by_path(PART).text == target.xml_part_by_path(PART).text
    # A procedure with arguments it calls keeps its name unless it clashes.
    plan = ps.plan_export(
        source, target, [ps.ExportItem.of_macro("Helpers", "BuildReport")], suffix="_T",
        check_references=False,
    )
    assert plan.ok, plan.describe()
    helpers = _body(plan.draft, "Helpers")
    assert "Public Sub BuildReport_T(title As String)" in helpers  # asked for by name: renamed
    assert "Public Function Utility() As Long" in helpers  # needed, no clash: kept
    assert plan.draft.find_current("Counter") is not None  # the class it uses


def test_whole_module_renames_its_macros(pair):
    source, target = pair
    names = {m.name: m.kind for m in ps.portable_modules(source)}
    assert {"RibbonCode": "standard", "Helpers": "standard", "Counter": "class"}.items() <= names.items()
    assert "ThisWorkbook" not in names
    plan = ps.plan_export(
        source, target, [ps.ExportItem.of_module("RibbonCode")], suffix="_T", check_references=False
    )
    assert plan.ok, plan.describe()
    code = _body(plan.draft, "RibbonCode")
    # Every procedure, macros (no required arguments) renamed; the others kept.
    assert "Public Sub AfterReport_T()" in code
    assert "Public Sub Unrelated(control As IRibbonControl)" in code
    assert "Public Sub OnLoadRibbon(ribbon As IRibbonUI)" in code
    assert 'Application.OnTime Now, "AfterReport_T"' in code


def test_class_module_and_several_kinds_at_once(pair):
    source, target = pair
    items = [
        ps.ExportItem.of_module("Counter"),
        ps.ExportItem.of_macro("Helpers", "BuildReport"),
        ps.ExportItem.of_control(_control(source, "btnOther"), _where(target, _control(source, "btnOther"), "Empty")),
    ]
    plan = ps.plan_export(source, target, items, suffix="_T", check_references=False)
    assert plan.ok, plan.describe()
    assert sum(1 for m in plan.draft.modules if m.current_name.startswith("Counter")) == 1
    assert "Sub Unrelated_T(" in _body(plan.draft, "RibbonCode")
    assert "Sub BuildReport_T(" in _body(plan.draft, "Helpers")


def test_unsaved_source_icon_is_exported(pair):
    source, target = pair
    # The source tab changed the icon without saving: the export takes the new one.
    new_icon = PORT_ICON_PNG + b"\x00"
    source.binary_parts.append(BinaryPartDraft(path="customUI/images/new.png", data=new_icon, content_type="image/png"))
    rels = source.xml_part_by_path("customUI/_rels/customUI14.xml.rels")
    rels.text = rels.text.replace("images/report.png", "images/new.png")
    control = _control(source, "btnReport")
    plan = ps.plan_port(source, target, control, _where(target, control, "Empty"), suffix="_T", check_references=False)
    assert plan.ok, plan.describe()
    assert plan.draft.binary_part_by_path("customUI/images/btnReport.png").data == new_icon


def test_macro_into_add_in_without_ribbon_saves(tmp_path: Path):
    source = DocumentService().open(build_source(tmp_path / "src" / "Source.xlam"))
    target = DocumentService().open(build_xlam(tmp_path / "work" / "Plain.xlam"))
    plan = ps.plan_export(
        source, target, [ps.ExportItem.of_macro("RibbonCode", "RunReport")], suffix="_P",
        check_references=False,
    )
    assert plan.ok, plan.describe()
    # Its ribbon variable cannot be set without a ribbon: said, not silently dropped.
    assert any("no ribbon" in w for w in plan.warnings)
    ps.apply_port(target, plan)
    assert make_service().save_addin(target).kind == "success"
    reopened = DocumentService().open(target.baseline.path)
    assert "Public Sub RunReport_P(control As IRibbonControl)" in _body(reopened, "RibbonCode")


def test_export_refusals(pair):
    source, target = pair
    assert "Choose at least one" in ps.plan_export(source, target, [], suffix="_T").problems[0]
    plan = ps.plan_export(source, target, [ps.ExportItem.of_macro("RibbonCode", "Nope")], suffix="_T")
    assert "no procedure named Nope" in plan.problems[0]
    plan = ps.plan_export(source, target, [ps.ExportItem.of_macro("Counter", "Bump")], suffix="_T")
    assert "not in a standard module" in plan.problems[0]
    plan = ps.plan_export(source, target, [ps.ExportItem.of_module("ThisWorkbook")], suffix="_T")
    assert not plan.ok


def test_whole_tab_into_an_add_in_without_a_ribbon(tmp_path: Path):
    source = DocumentService().open(build_source(tmp_path / "src" / "Source.xlam"))
    target = DocumentService().open(build_xlam(tmp_path / "work" / "Plain.xlam"))
    tab = _control(source, "tabSrc")
    assert [c.element for c in ps.portable_controls(source)].count("tab") == 1
    targets = ps.port_targets(target, tab)
    assert len(targets) == 1 and targets[0].new_ribbon
    assert "new ribbon" in targets[0].describe()
    # A button still needs a ribbon to go into.
    assert ps.port_targets(target, _control(source, "btnReport")) == []
    plan = ps.plan_export(source, target, [ps.ExportItem.of_control(tab, targets[0])], suffix="_P", check_references=False)
    assert plan.ok, plan.describe()
    ribbon = plan.draft.xml_part_by_path(PART)
    assert ribbon is not None and ribbon.is_new
    assert '<tab id="tabSrc" label="Source">' in ribbon.text and 'onAction="RunReport_P"' in ribbon.text
    assert plan.part_path == PART and ribbon.text[plan.element_offset :].startswith("<tab ")
    assert any("A ribbon for this add-in" in s for s in plan.summary)
    assert 'onLoad="RibbonLoaded_P"' in ribbon.text
    assert target.xml_part_by_path(PART) is None  # the target itself is untouched
    ps.apply_port(target, plan)
    assert make_service().save_addin(target).kind == "success"
    reopened = DocumentService().open(target.baseline.path)
    assert rs.new_issues(reopened) == [] and reopened.binary_part_by_path("customUI/images/btnReport.png") is None
    with __import__("zipfile").ZipFile(target.baseline.path) as package:
        assert "customUI/images/btnReport.png" in package.namelist()
        assert "customUI14.xml" in package.read("_rels/.rels").decode("utf-8")
