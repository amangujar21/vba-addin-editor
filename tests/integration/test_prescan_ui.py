"""After a save the status bar says when antivirus has finished with the file."""

from __future__ import annotations

import gc
import threading
import tkinter as tk
from pathlib import Path

import pytest

from vba_addin_editor.platform import prescan
from vba_addin_editor.platform import windows_file_ops as wfo
from vba_addin_editor.ui import main_window as mw


@pytest.fixture()
def window(work_xlam: Path, tmp_path: Path, monkeypatch):
    monkeypatch.setenv("VBAAE_SESSION_ROOT", str(tmp_path / "sessions"))
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    root.withdraw()
    win = mw.MainWindow(root)
    win.backup_service.catalog_path = tmp_path / "catalog.json"
    win.save_service.process_probe = lambda p: False
    monkeypatch.setattr(mw.messagebox, "askokcancel", lambda *a, **k: True)
    monkeypatch.setattr(mw.messagebox, "showinfo", lambda *a, **k: None)
    win.load_path(work_xlam)
    root.update()
    try:
        yield win
    finally:
        root.destroy()


def _edit_and_save(win, marker: str) -> None:
    module = win.draft.find_current("Module1")
    win._open_module_at(module.id, 1)
    win.editor.text.insert("end", f"' {marker}\n")
    win.save()
    assert not win.draft.is_dirty()


def _hold_scans(win) -> tuple[threading.Event, list[Path]]:
    release = threading.Event()
    opened: list[Path] = []

    def opener(path: Path) -> bool:
        opened.append(path)
        release.wait(10)
        return prescan.open_for_read(path)

    win._prescan = prescan.Prescan(opener)
    return release, opened


def _pump_until_idle(win) -> None:
    win._prescan.settle()
    win._poll_prescan()


def test_note_changes_from_scanning_to_ready_for_office(window, work_xlam):
    release, opened = _hold_scans(window)
    assert window.scan_status.cget("text") == ""
    was_enabled = gc.isenabled()
    _edit_and_save(window, "one")
    assert opened == [work_xlam]
    assert window.scan_status.cget("text") == f"Antivirus scanning {work_xlam.name}…"
    assert not gc.isenabled()  # the pre-read thread must never collect Tk variables
    release.set()
    _pump_until_idle(window)
    assert window.scan_status.cget("text") == f"{work_xlam.name} ready for Office"
    assert gc.isenabled() == was_enabled


def test_a_second_save_waits_for_the_pre_read_then_succeeds(window, work_xlam):
    release, opened = _hold_scans(window)
    _edit_and_save(window, "one")
    threading.Timer(0.3, release.set).start()
    _edit_and_save(window, "two")  # would be "file in use" with the handle still open
    assert opened == [work_xlam, work_xlam]
    _pump_until_idle(window)
    assert window.scan_status.cget("text") == f"{work_xlam.name} ready for Office"
    assert wfo.probe_exclusive_access(work_xlam).ok  # nothing left holding the file


def test_a_blocked_save_starts_no_pre_read(window, work_xlam, monkeypatch):
    _, opened = _hold_scans(window)
    monkeypatch.setattr(mw.messagebox, "showwarning", lambda *a, **k: None)
    window.save_service.process_probe = lambda p: True
    module = window.draft.find_current("Module1")
    window._open_module_at(module.id, 1)
    window.editor.text.insert("end", "' blocked\n")
    window.save()
    assert opened == [] and window.scan_status.cget("text") == ""


def test_closing_the_window_mid_scan_gives_gc_back(window, work_xlam):
    release, _ = _hold_scans(window)
    was_enabled = gc.isenabled()
    _edit_and_save(window, "one")
    window._cancel_afters()
    assert gc.isenabled() == was_enabled
    release.set()
    window._prescan.settle()
