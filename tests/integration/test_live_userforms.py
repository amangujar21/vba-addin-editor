"""Release gate for UserForms, on Office-authored fixtures.

A UserForm Excel made (``tests/fixtures/xlam/FormSource.xlam``: a frame,
a MultiPage, pictures, code that calls another module) is copied into
Office-authored add-ins (.xlam, .ppam, .pptm) and edited, then saved through
the real pipeline. Each test checks the result at the parser level. With
``VBAAE_LIVE_OFFICE=1`` it also loads the saved file in real Excel or
PowerPoint (``scripts/office_check.py``) and checks what Office reports for
every control: the form compiles and runs, pictures included.

Fixtures come from ``python scripts/make_office_fixtures.py`` and are never
committed; without them every test here skips.
"""

from __future__ import annotations

import importlib.util
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

from vba_addin_editor.adapters.pyopenvba_adapter import PyOpenVBAAdapter
from vba_addin_editor.domain.document import ModuleDisplayKind, ModuleDraft, new_module_id
from vba_addin_editor.services import form_service as fs
from vba_addin_editor.services import port_service as ps
from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.services.save_service import SaveService

ROOT = Path(__file__).resolve().parents[2]
FIXTURES = ROOT / "tests" / "fixtures"
FORM_SOURCE = FIXTURES / "xlam" / "FormSource.xlam"
TARGETS = [
    pytest.param(FIXTURES / "xlam" / "RealAddin.xlam", id="xlam"),
    pytest.param(FIXTURES / "ppam" / "RealAddin.ppam", id="ppam"),
    pytest.param(FIXTURES / "pptm" / "RealPresentation.pptm", id="pptm"),
]
OFFICE = os.environ.get("VBAAE_LIVE_OFFICE") == "1"


def _office_check_module():
    spec = importlib.util.spec_from_file_location("office_check", ROOT / "scripts" / "office_check.py")
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _require(*paths: Path) -> None:
    missing = [str(p.relative_to(ROOT)) for p in paths if not p.exists()]
    if missing:
        pytest.skip(f"Live fixture(s) missing: {missing}. Run scripts/make_office_fixtures.py.")


def _add_report_module(draft) -> None:
    check = _office_check_module()
    draft.modules.append(
        ModuleDraft(
            id=new_module_id(),
            origin_name=None,
            current_name=check.REPORT_MODULE,
            body=check.REPORT_CODE,
            kind=ModuleDisplayKind.STANDARD,
            pyopenvba_kind="standard",
            is_new=True,
            is_deleted=False,
            destructive_ops_safe=True,
            can_delete=True,
            can_rename=True,
            project_item_kind="standard",
        )
    )


def _office_report(path: Path, *forms: str) -> str:
    """What real Office says about ``forms`` in ``path``."""
    run = subprocess.run(
        [sys.executable, str(ROOT / "scripts" / "office_check.py"), str(path), *forms],
        capture_output=True,
        text=True,
        timeout=180,
        check=False,
    )
    output = run.stdout + run.stderr
    assert run.returncode == 0 and "OFFICE CHECK OK" in output, output
    return output


def _line(report: str, control: str) -> str:
    match = re.search(rf"^\s+{control}:.*$", report, re.MULTILINE)
    assert match is not None, report
    return match.group(0)


@pytest.mark.live
@pytest.mark.parametrize("target", TARGETS)
def test_office_form_copied_into_real_addin(target: Path, tmp_path: Path):
    _require(FORM_SOURCE, target)
    work = tmp_path / target.name
    shutil.copy2(target, work)
    service = DocumentService()
    source = service.open(FORM_SOURCE)
    draft = service.open(work)

    plan = ps.plan_form_port(source, draft, ["frmDemo"], suffix="_Copy")
    assert plan.ok, plan.describe()
    ps.apply_port(draft, plan)
    _add_report_module(draft)
    result = SaveService(adapter=PyOpenVBAAdapter()).save_addin(draft)
    assert result.kind == "success", result

    reopened = service.open(work)
    form = reopened.find_current("frmDemo")
    assert form is not None and form.project_item_kind == "userform"
    assert reopened.find_current("FormHelpers") is not None
    assert reopened.baseline.form("frmDemo").storage == source.baseline.form("frmDemo").storage
    assert reopened.find_current("ModuleKeep").body == service.open(target).find_current("ModuleKeep").body

    if OFFICE:
        report = _office_report(work, "frmDemo")
        assert "FORM frmDemo caption=Demo Form controls=13" in report
        assert "pic=True" in _line(report, "imgLogo")
        assert "pic=True" in _line(report, "cmdOK")
        assert "value=World" in _line(report, "txtName")  # Initialize ran, calling FormHelpers
        assert "value=Alpha" in _line(report, "cboKind")
        assert "cap=[Options]" in _line(report, "fraOptions")
        assert "cap=[On page 1]" in _line(report, "lblInner")


@pytest.mark.live
def test_office_form_layout_edited_in_place(tmp_path: Path):
    _require(FORM_SOURCE)
    work = tmp_path / "FormSource.xlam"
    shutil.copy2(FORM_SOURCE, work)
    service = DocumentService()
    draft = service.open(work)
    form = draft.find_current("frmDemo")
    assert form is not None

    def edit(e) -> None:
        e.set_form_caption("Edited Form")
        e.set_form_size(360, 280)
        e.add_control("CommandButton", "cmdCancel", left=130, top=236, width=80, height=26)
        e.set_caption("cmdCancel", "Cancel")
        e.set_tip("cmdCancel", "Close without saving")
        e.set_geometry("cmdOK", left=220, top=236, width=90)
        e.set_font("lblTitle", "Segoe UI", 10, True)
        e.remove_control("lstItems")
        e.rename_control("lblStatus", "lblResult")
        e.add_control("TextBox", "txtInFrame", container="fraOptions", left=6, top=50, width=100)
        e.add_control("Label", "lblOnPage2", container="Page2", left=4, top=4)

    fs.edit_layout(draft, form, edit)
    form.body = fs.rename_in_code(form.body, "lblStatus", "lblResult")
    _add_report_module(draft)
    result = SaveService(adapter=PyOpenVBAAdapter()).save_addin(draft)
    assert result.kind == "success", result

    design = service.open(work).baseline.form("frmDemo")
    assert design.caption == "Edited Form"
    assert design.control("lstItems") is None and design.control("lblResult") is not None

    if OFFICE:
        report = _office_report(work, "frmDemo")
        assert "caption=Edited Form controls=15" in report
        assert "@130,236 80x26" in _line(report, "cmdCancel")
        assert "tip=[Close without saving]" in _line(report, "cmdCancel")
        assert "@220,236 90x" in _line(report, "cmdOK")
        assert "font=Segoe UI/10/b" in _line(report, "lblTitle")
        assert "lstItems" not in report
        assert _line(report, "txtInFrame") and _line(report, "lblOnPage2")


@pytest.mark.live
@pytest.mark.parametrize("target", [TARGETS[0], TARGETS[1]])
def test_new_form_with_every_control_kind(target: Path, tmp_path: Path):
    """A form made in the designer from nothing, one control of each kind it offers."""
    _require(target)
    work = tmp_path / target.name
    shutil.copy2(target, work)
    service = DocumentService()
    draft = service.open(work)
    module = fs.add_blank_form(draft, "frmBuilt", "Built here")

    def edit(e) -> None:
        e.set_form_size(420, 330)
        top = 6.0
        for index, kind in enumerate(fs.CONTROL_KINDS):
            column, row = divmod(index, 7)
            e.add_control(kind, f"{kind}1", left=6 + column * 200, top=top + row * 44)
        e.set_caption("CommandButton1", "Run")
        e.add_control("OptionButton", "OptionInFrame", container="Frame1", left=6, top=6)

    fs.edit_layout(draft, module, edit)
    module.body, _handler_line = fs.ensure_handler(module.body, "CommandButton1", "Click")
    _add_report_module(draft)
    result = SaveService(adapter=PyOpenVBAAdapter()).save_addin(draft)
    assert result.kind == "success", result
    reopened = service.open(work).baseline.form("frmBuilt")
    assert reopened is not None and {c.name for c in reopened.walk()} >= {f"{k}1" for k in fs.CONTROL_KINDS}

    if OFFICE:
        report = _office_report(work, "frmBuilt")
        assert "FORM frmBuilt caption=Built here" in report
        for kind in fs.CONTROL_KINDS:
            assert f":{kind}" in _line(report, f"{kind}1"), kind
        assert "cap=[Run]" in _line(report, "CommandButton1")
        assert _line(report, "OptionInFrame")
