"""Live gate for saving while Office has the add-in loaded.

Real Excel / PowerPoint load a copy of the Office-authored fixture. The editor
window then adds a macro and a ribbon button and saves *while the add-in stays
loaded*: the save closes it in Office, commits, and loads it again. Office is
then asked directly:

- the new macro runs (``Application.Run``), so the reloaded VBA is the saved one;
- the new button is on the ribbon Office built, found by its label through UI
  Automation (``scripts/ribbon_click.ps1``) and clicked, which runs its macro.

A second round adds another macro and button to the same group the same way,
the "keep adding things while Office is open" workflow.

Needs ``VBAAE_LIVE_OFFICE=1`` and the fixtures from
``scripts/make_office_fixtures.py``. Office is started visibly (a ribbon is
only built for a window) as its own instance and quit at the end. PowerPoint
runs one instance, so its test skips when PowerPoint is already open.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import time
import tkinter as tk
from pathlib import Path

import pytest

from vba_addin_editor.domain.document import ModuleDisplayKind, ModuleDraft, new_module_id
from vba_addin_editor.platform import office_live
from vba_addin_editor.services import ribbon_service as rs
from vba_addin_editor.services import starter_service as ss
from vba_addin_editor.ui import main_window as mw
from vba_addin_editor.ui.ribbon_dialogs import ButtonSpec

REAL_LOCATE = office_live.locate  # before the autouse fixture swaps it out
ROOT = Path(__file__).resolve().parents[2]
FIXTURES = ROOT / "tests" / "fixtures"
CLICKER = ROOT / "scripts" / "ribbon_click.ps1"
PART = "customUI/customUI14.xml"
OFFICE = os.environ.get("VBAAE_LIVE_OFFICE") == "1"
TAB = "VBAAE Live"

pytestmark = [
    pytest.mark.live,
    pytest.mark.skipif(not OFFICE, reason="set VBAAE_LIVE_OFFICE=1 to drive real Office"),
]

MODULE = "LiveMacros"
CODE_V1 = """Option Explicit

Public gClicks As Long

Public Sub LiveHello(control As IRibbonControl)
    gClicks = gClicks + 1
End Sub

Public Function LiveClicks() As Long
    LiveClicks = gClicks
End Function

Public Function LiveProbe() As String
    LiveProbe = "v1"
End Function
"""
CODE_V2_EXTRA = """
Public Sub LiveSecond(control As IRibbonControl)
    gClicks = gClicks + 100
End Sub

Public Function LiveProbe2() As String
    LiveProbe2 = "v2"
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
    """Real Office lookup; message boxes recorded; the review answers Save."""
    monkeypatch.setattr(office_live, "locate", REAL_LOCATE)
    shown: list[str] = []

    def record(*args, **_kwargs):
        shown.append(str(args[1] if len(args) > 1 else ""))
        return True

    for name in ("showinfo", "showwarning", "showerror", "askyesno", "askokcancel"):
        monkeypatch.setattr(f"vba_addin_editor.ui.main_window.messagebox.{name}", record)
    monkeypatch.setattr("vba_addin_editor.ui.main_window.ask_review", lambda *_a, **_k: "save")
    return shown


def _pid(hwnd: int) -> int:
    import win32process

    return win32process.GetWindowThreadProcessId(hwnd)[1]


def _switch_to(hwnd: int) -> None:
    """Bring Office to the front, as a user does to use the new button.

    Office rebuilds a reloaded add-in's ribbon only once its window is active
    again; while the editor has the foreground the new tab is not there yet.
    Windows lets a process that is not in the foreground take it only after
    input, hence the Alt tap when the plain call is refused.
    """
    import win32api
    import win32con
    import win32gui

    for _attempt in range(3):
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            win32gui.SetForegroundWindow(hwnd)
        except Exception:  # noqa: BLE001, S110 - refused by the foreground lock
            pass
        if win32gui.GetForegroundWindow() == hwnd:
            return
        win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
        win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)
        time.sleep(0.2)
    assert win32gui.GetForegroundWindow() == hwnd, "could not bring Office to the front"


def _ribbon_button(pid: int, label: str, *, click: bool) -> str:
    args = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(CLICKER),
            "-ProcessId", str(pid), "-Tab", TAB, "-Button", label]
    if click:
        args.append("-Click")
    run = subprocess.run(args, capture_output=True, text=True, timeout=90, check=False)
    return (run.stdout + run.stderr).strip()


def _add_macros(window: mw.MainWindow, code: str) -> ModuleDraft:
    draft = window.draft
    assert draft is not None
    existing = draft.find_current(MODULE)
    if existing is not None:
        window._batch_edit(lambda: setattr(existing, "body", existing.body + code))
        return existing
    module = ModuleDraft(
        id=new_module_id(), origin_name=None, current_name=MODULE, body=code,
        kind=ModuleDisplayKind.STANDARD, pyopenvba_kind="standard", is_new=True,
        is_deleted=False, destructive_ops_safe=True, can_delete=True, can_rename=True,
        project_item_kind="standard",
    )
    window._batch_edit(lambda: draft.modules.append(module))
    return module


def _add_button(window: mw.MainWindow, module: ModuleDraft, control_id: str, label: str, macro: str) -> None:
    draft = window.draft
    assert draft is not None
    if draft.xml_part_by_path(PART) is None:
        assert window.apply_starter_ribbon(ss.StarterRibbonSpec(TAB, "Live Tools"))
    part = draft.xml_part_by_path(PART)
    assert part is not None and part.text is not None
    group = next(c for c in rs.find_containers(part.text, PART) if c.element == "group")
    spec = ButtonSpec(group, control_id, label, "HappyFace", "large", "", "")
    window.apply_ribbon_button(module.id, macro, None, spec)
    assert rs.check_ribbon(draft) == []


def _save_live(window: mw.MainWindow, root, shown: list[str], host: str) -> float:
    started = time.perf_counter()
    window.save()
    root.update()
    elapsed = time.perf_counter() - started
    result = window._last_result
    assert result is not None and result.kind == "success", (result, shown)
    assert result.details.get("live_host") == host, result.details
    assert result.details.get("live_reloaded") is True, result.details
    assert f"reloaded in {host}" in window.status.cget("text")
    assert window.draft is not None and not window.draft.is_dirty()
    return elapsed


def _run_until(call, expect, *, timeout: float = 20.0):
    deadline = time.monotonic() + timeout
    value = None
    while time.monotonic() < deadline:
        try:
            value = call()
            if value == expect:
                return value
        except Exception as exc:  # noqa: BLE001 - Office still loading the add-in
            value = exc
        time.sleep(0.25)
    return value


def _rounds(window, root, ui, host, *, run, hwnd, work) -> list[float]:
    """Both rounds of the live workflow; returns the two save times."""
    times = []
    assert window.load_path(work)
    module = _add_macros(window, CODE_V1)
    _add_button(window, module, "btnLiveHello", "Live Hello", "LiveHello")
    times.append(_save_live(window, root, ui, host))

    pid = _pid(hwnd)
    assert _run_until(lambda: run("LiveProbe"), "v1") == "v1"
    _switch_to(hwnd)
    out = _ribbon_button(pid, "Live Hello", click=True)
    assert "CLICKED" in out, out
    assert _run_until(lambda: run("LiveClicks"), 1) == 1  # the click ran the new macro

    # Round two: more macros and a second button while the add-in stays loaded.
    module = _add_macros(window, CODE_V2_EXTRA)
    _add_button(window, module, "btnLiveSecond", "Live Second", "LiveSecond")
    times.append(_save_live(window, root, ui, host))

    assert _run_until(lambda: run("LiveProbe2"), "v2") == "v2"
    _switch_to(hwnd)
    assert "FOUND" in _ribbon_button(pid, "Live Hello", click=False)
    out = _ribbon_button(pid, "Live Second", click=True)
    assert "CLICKED" in out, out
    assert _run_until(lambda: run("LiveClicks"), 100) == 100  # fresh project: only the new click
    return times


def _procs():
    sys.path.insert(0, str(ROOT / "scripts"))
    from _office_procs import live_pids, stop_new  # pyright: ignore[reportMissingImports]

    return live_pids, stop_new


def test_excel_add_macros_and_buttons_while_loaded(root, ui, tmp_path: Path):
    source = FIXTURES / "xlam" / "RealAddin.xlam"
    if not source.exists():
        pytest.skip("run scripts/make_office_fixtures.py")
    import pythoncom
    import win32com.client as com

    live_pids, stop_new = _procs()
    work = tmp_path / "LiveAddin.xlam"
    shutil.copy2(source, work)
    before = live_pids("EXCEL.EXE")
    pythoncom.CoInitialize()
    app = com.DispatchEx("Excel.Application")
    try:
        app.Visible = True
        app.Workbooks.Add()
        book = app.Workbooks.Open(str(work))
        assert book.IsAddin
        book = None
        hwnd = int(app.Hwnd)

        excel = app

        def run(macro):
            return excel.Run(f"'{work.name}'!{MODULE}.{macro}")

        window = mw.MainWindow(root)
        window.settings.live_reload = True  # GUI tests share a settings file
        times = _rounds(window, root, ui, "Excel", run=run, hwnd=hwnd, work=work)
        # The add-in is still the one loaded from the same path.
        assert any(Path(str(b.FullName)) == work for b in app.Workbooks) or app.Workbooks(work.name)
        print(f"\nExcel live saves: {times[0]:.2f}s, {times[1]:.2f}s")
    finally:
        app.DisplayAlerts = False
        for b in list(app.Workbooks):
            b.Close(False)
        app.Quit()
        app = None
        pythoncom.CoUninitialize()
        stop_new("EXCEL.EXE", before)


def test_powerpoint_add_macros_and_buttons_while_loaded(root, ui, tmp_path: Path):
    source = FIXTURES / "ppam" / "RealAddin.ppam"
    if not source.exists():
        pytest.skip("run scripts/make_office_fixtures.py")
    live_pids, stop_new = _procs()
    before = live_pids("POWERPNT.EXE")
    if before:
        pytest.skip("PowerPoint runs one instance: close it before this test")
    import pythoncom
    import win32com.client as com

    work = tmp_path / "LiveAddin.ppam"
    shutil.copy2(source, work)
    pythoncom.CoInitialize()
    app = com.DispatchEx("PowerPoint.Application")
    try:
        app.Visible = -1
        app.AutomationSecurity = 1
        app.Presentations.Add(WithWindow=-1)
        addin = app.AddIns.Add(str(work))
        addin.Loaded = -1
        addin = None
        import win32gui

        # PowerPoint runs one instance (this test's), so its frame is the only one.
        hwnd = win32gui.FindWindow("PPTFrameClass", None)
        assert hwnd, "PowerPoint's window was not found"

        powerpoint = app

        def run(macro):
            return powerpoint.Run(f"{work.name}!{MODULE}.{macro}")

        window = mw.MainWindow(root)
        window.settings.live_reload = True  # GUI tests share a settings file
        times = _rounds(window, root, ui, "PowerPoint", run=run, hwnd=hwnd, work=work)
        loaded = [a for a in app.AddIns if Path(str(a.FullName)) == work]
        assert loaded and int(loaded[0].Loaded)
        print(f"\nPowerPoint live saves: {times[0]:.2f}s, {times[1]:.2f}s")
    finally:
        for a in list(app.AddIns):
            if Path(str(a.FullName)) == work:
                a.Loaded = 0
                app.AddIns.Remove(a.Name)
        for p in list(app.Presentations):
            p.Saved = -1
            p.Close()
        app.Quit()
        app = None
        pythoncom.CoUninitialize()
        stop_new("POWERPNT.EXE", before)
