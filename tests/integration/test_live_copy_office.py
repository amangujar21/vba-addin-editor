"""Live gate for copying between add-ins open in two tabs.

Two copies of the Office-authored fixture are opened as two tabs. The first
(the source) gets, through the editor, a module of macros, a class module, and
a ribbon tab whose buttons run the macros (one with an imageMso, one with a
custom icon), and is saved. Its whole tab, two plain macros and a whole
module are then copied into the second tab (the target) with the copy dialog's
plan, and the target is saved.

Real Excel / PowerPoint then load *both* add-ins at once and are asked:

- the copied macros run under their new names, and each add-in keeps its own
  state (the target's counter is not the source's);
- a macro name the source still uses resolves to the source alone (no clash);
- both tabs are on the ribbon Office built, and clicking the copied buttons
  runs the target's macros, the source's buttons the source's.

Needs ``VBAAE_LIVE_OFFICE=1`` and the fixtures from
``scripts/make_office_fixtures.py``. PowerPoint runs one instance, so its test
skips when PowerPoint is already open.
"""

from __future__ import annotations

import io
import shutil
import time
import tkinter as tk
import zipfile
from pathlib import Path

import pytest
from test_live_reload_office import (  # shared live helpers (same folder)
    CLICKER,
    FIXTURES,
    OFFICE,
    _pid,
    _procs,
    _run_until,
    _switch_to,
)

from vba_addin_editor.domain.document import ModuleDisplayKind, ModuleDraft, new_module_id
from vba_addin_editor.services import icon_service
from vba_addin_editor.services import port_service as ps
from vba_addin_editor.services import ribbon_service as rs
from vba_addin_editor.services import starter_service as ss
from vba_addin_editor.ui import main_window as mw
from vba_addin_editor.ui.ribbon_dialogs import ButtonSpec

pytestmark = [
    pytest.mark.live,
    pytest.mark.skipif(not OFFICE, reason="set VBAAE_LIVE_OFFICE=1 to drive real Office"),
]

PART = "customUI/customUI14.xml"
SOURCE_TAB, COPIED_TAB = "VBAAE Source", "VBAAE Copied"
MODULE, CLASS = "CopyMacros", "CopyCounter"
CODE = """Option Explicit

Public gHits As Long

Public Sub CopyHello(control As IRibbonControl)
    gHits = gHits + 1
End Sub

Public Sub CopyIcon(control As IRibbonControl)
    gHits = gHits + 1000
End Sub

Public Sub CopyPlain()
    Dim c As New CopyCounter
    gHits = gHits + c.Step
End Sub

Public Function CopyHits() As Long
    CopyHits = gHits
End Function

Public Function CopyWho() As String
    CopyWho = ThisName()
End Function

Private Function ThisName() As String
    ThisName = "copied code"
End Function
"""
CLASS_CODE = """Option Explicit

Public Function Step() As Long
    Step = 10
End Function
"""


@pytest.fixture
def root():
    try:
        tk_root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    tk_root.withdraw()
    yield tk_root
    tk_root.destroy()


@pytest.fixture
def ui(monkeypatch):
    shown: list[str] = []

    def record(*args, **_kwargs):
        shown.append(str(args[1] if len(args) > 1 else ""))
        return True

    for name in ("showinfo", "showwarning", "showerror", "askyesno", "askokcancel"):
        monkeypatch.setattr(f"vba_addin_editor.ui.main_window.messagebox.{name}", record)
    monkeypatch.setattr("vba_addin_editor.ui.main_window.ask_review", lambda *_a, **_k: "save")
    return shown


def _png() -> bytes:
    from PIL import Image

    out = io.BytesIO()
    Image.new("RGBA", (32, 32), (220, 40, 40, 255)).save(out, "PNG")
    return out.getvalue()


def _module(name: str, body: str, *, standard: bool) -> ModuleDraft:
    return ModuleDraft(
        id=new_module_id(), origin_name=None, current_name=name, body=body,
        kind=ModuleDisplayKind.STANDARD if standard else ModuleDisplayKind.CLASS,
        pyopenvba_kind="standard" if standard else "other", is_new=True, is_deleted=False,
        destructive_ops_safe=True, can_delete=True, can_rename=True,
        project_item_kind="standard" if standard else "class",
    )


def _save(window: mw.MainWindow, root, shown: list[str]) -> None:
    window.save()
    root.update()
    result = window._last_result
    assert result is not None and result.kind == "success", (result, shown)
    assert not window.draft.is_dirty()


def _make_source(window: mw.MainWindow, root, shown) -> None:
    """Macros, a class and a ribbon tab with two buttons, added through the editor."""
    draft = window.draft
    code, counter = _module(MODULE, CODE, standard=True), _module(CLASS, CLASS_CODE, standard=False)
    window._batch_edit(lambda: draft.modules.extend([code, counter]))
    assert window.apply_starter_ribbon(ss.StarterRibbonSpec(SOURCE_TAB, "Copy Tools"))
    for control_id, label, macro in (("btnCopyHello", "Copy Hello", "CopyHello"), ("btnCopyIcon", "Copy Icon", "CopyIcon")):
        part = draft.xml_part_by_path(PART)
        group = next(c for c in rs.find_containers(part.text, PART) if c.element == "group")
        window.apply_ribbon_button(code.id, macro, None, ButtonSpec(group, control_id, label, "HappyFace", "large", "", ""))
    part = draft.xml_part_by_path(PART)
    button = next(e for e in rs.element_tree(part.text) if e.attr("id") == "btnCopyIcon")
    assert window.apply_ribbon_icon(PART, button.start, icon_service.icon_from_bytes(_png()))
    assert rs.check_ribbon(draft) == []
    _save(window, root, shown)


def _copy(window: mw.MainWindow, source_tab, target_tab, monkeypatch) -> ps.PortPlan:
    """Export the whole tab, two plain macros and the class from the source tab into the target."""
    source = source_tab.draft
    tab = next(c for c in ps.portable_controls(source) if c.element == "tab")
    seen: list[ps.PortPlan] = []

    def ask(src, tgt, *, preselect=(), title=""):
        assert src is source and tgt is target_tab.draft
        where = ps.port_targets(tgt, tab)[0]
        items = [
            ps.ExportItem.of_control(tab, where),
            ps.ExportItem.of_macro(MODULE, "CopyPlain"),
            ps.ExportItem.of_macro(MODULE, "CopyHits"),
            ps.ExportItem.of_macro(MODULE, "CopyWho"),
        ]
        plan = ps.plan_export(src, tgt, items, suffix=ps.default_suffix(tgt))
        seen.append(plan)
        return plan

    monkeypatch.setattr(window, "ask_export", ask)
    assert window.select_tab(source_tab.key)
    assert window.export_to(target_tab.key)
    assert window.active_tab is target_tab
    plan = seen[0]
    assert plan.ok, plan.describe()
    # A second tab with the same label would be ambiguous to find; a user may rename it too.
    part = target_tab.draft.xml_part_by_path(PART)
    window._batch_edit(lambda: setattr(part, "text", part.text.replace(f'label="{SOURCE_TAB}"', f'label="{COPIED_TAB}"')))
    return plan


def _click(pid: int, tab: str, label: str) -> str:
    import subprocess

    args = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(CLICKER),
            "-ProcessId", str(pid), "-Tab", tab, "-Button", label, "-Click"]
    run = subprocess.run(args, capture_output=True, text=True, timeout=90, check=False)
    return (run.stdout + run.stderr).strip()


def _check_package(path: Path) -> None:
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        rels = package.read("customUI/_rels/customUI14.xml.rels").decode("utf-8")
        ribbon = package.read(PART).decode("utf-8")
    assert any(n.startswith("customUI/images/") for n in names), names
    assert rs.IMAGE_REL_TYPE in rels and rs.LEGACY_IMAGE_REL_TYPE not in rels
    assert 'imageMso="HappyFace"' in ribbon


def _workflow(root, ui, monkeypatch, tmp_path: Path, fixture: Path, extension: str):
    source_path = tmp_path / f"CopySource{extension}"
    target_path = tmp_path / f"CopyTarget{extension}"
    shutil.copy2(fixture, source_path)
    shutil.copy2(fixture, target_path)
    window = mw.MainWindow(root)
    assert window.load_path(target_path) and window.load_path(source_path)
    source_tab, target_tab = window.tab_for_path(source_path), window.tab_for_path(target_path)
    _make_source(window, root, ui)
    plan = _copy(window, source_tab, target_tab, monkeypatch)
    suffix = ps.default_suffix(target_tab.draft)
    renames = {old: new for _what, old, new in plan.renames}
    assert renames["CopyHello"] == f"CopyHello{suffix}" and renames["CopyPlain"] == f"CopyPlain{suffix}"
    _save(window, root, ui)
    _check_package(target_path)
    body = target_tab.draft.find_current(MODULE).body
    assert f"Public Sub CopyPlain{suffix}()" in body and "Private Function ThisName" in body
    assert target_tab.draft.find_current(CLASS) is not None
    return source_path, target_path, suffix


def _assert_in_office(run, src: str, tgt: str, suffix: str, pid: int, hwnd: int) -> None:
    # Copied macros run under their new names; each add-in keeps its own state.
    assert _run_until(lambda: run(f"{tgt}!CopyWho{suffix}"), "copied code") == "copied code"
    run(f"{tgt}!CopyPlain{suffix}")
    assert run(f"{tgt}!CopyHits{suffix}") == 10
    assert run(f"{src}!{MODULE}.CopyHits") == 0
    # The name the source keeps runs the source's macro.
    run(f"{src}!CopyPlain")
    assert run(f"{src}!{MODULE}.CopyHits") == 10 and run(f"{tgt}!CopyHits{suffix}") == 10
    # Both tabs are on the ribbon; each add-in's buttons run its own macros.
    _switch_to(hwnd)
    assert "CLICKED" in (out := _click(pid, COPIED_TAB, "Copy Hello")), out
    assert _run_until(lambda: run(f"{tgt}!CopyHits{suffix}"), 11) == 11
    assert "CLICKED" in (out := _click(pid, COPIED_TAB, "Copy Icon")), out
    assert _run_until(lambda: run(f"{tgt}!CopyHits{suffix}"), 1011) == 1011
    assert "CLICKED" in (out := _click(pid, SOURCE_TAB, "Copy Hello")), out
    assert _run_until(lambda: run(f"{src}!{MODULE}.CopyHits"), 11) == 11
    assert run(f"{tgt}!CopyHits{suffix}") == 1011


def test_excel_copy_between_two_open_add_ins(root, ui, monkeypatch, tmp_path: Path):
    fixture = FIXTURES / "xlam" / "RealAddin.xlam"
    if not fixture.exists():
        pytest.skip("run scripts/make_office_fixtures.py")
    source, target, suffix = _workflow(root, ui, monkeypatch, tmp_path, fixture, ".xlam")
    import pythoncom
    import win32com.client as com

    live_pids, stop_new = _procs()
    before = live_pids("EXCEL.EXE")
    pythoncom.CoInitialize()
    app = com.DispatchEx("Excel.Application")
    try:
        app.Visible = True
        app.Workbooks.Add()
        for path in (source, target):
            assert app.Workbooks.Open(str(path)).IsAddin
        excel = app
        _assert_in_office(
            lambda macro: excel.Run(macro), f"'{source.name}'", f"'{target.name}'", suffix,
            _pid(int(app.Hwnd)), int(app.Hwnd),
        )
    finally:
        app.DisplayAlerts = False
        for book in list(app.Workbooks):
            book.Close(False)
        app.Quit()
        app = None
        pythoncom.CoUninitialize()
        stop_new("EXCEL.EXE", before)


def test_powerpoint_copy_between_two_open_add_ins(root, ui, monkeypatch, tmp_path: Path):
    fixture = FIXTURES / "ppam" / "RealAddin.ppam"
    if not fixture.exists():
        pytest.skip("run scripts/make_office_fixtures.py")
    live_pids, stop_new = _procs()
    before = live_pids("POWERPNT.EXE")
    if before:
        pytest.skip("PowerPoint runs one instance: close it before this test")
    source, target, suffix = _workflow(root, ui, monkeypatch, tmp_path, fixture, ".ppam")
    import pythoncom
    import win32com.client as com
    import win32gui

    pythoncom.CoInitialize()
    app = com.DispatchEx("PowerPoint.Application")
    try:
        app.Visible = -1
        app.AutomationSecurity = 1
        app.Presentations.Add(WithWindow=-1)
        for path in (source, target):
            app.AddIns.Add(str(path)).Loaded = -1
        time.sleep(1)
        hwnd = win32gui.FindWindow("PPTFrameClass", None)
        assert hwnd, "PowerPoint's window was not found"
        powerpoint = app
        _assert_in_office(
            lambda macro: powerpoint.Run(macro), source.name, target.name, suffix, _pid(hwnd), hwnd
        )
    finally:
        for addin in list(app.AddIns):
            if Path(str(addin.FullName)) in (source, target):
                addin.Loaded = 0
                app.AddIns.Remove(addin.Name)
        for presentation in list(app.Presentations):
            presentation.Saved = -1
            presentation.Close()
        app.Quit()
        app = None
        pythoncom.CoUninitialize()
        stop_new("POWERPNT.EXE", before)
