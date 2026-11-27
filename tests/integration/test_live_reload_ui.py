"""The window's side of saving while Office has the add-in loaded (no real Office).

Office is replaced by a fake ``office_live.locate`` / ``HostSwap``; the real
ones are exercised in ``test_live_reload_office.py``.
"""

from __future__ import annotations

import os
import tempfile
import tkinter as tk
from pathlib import Path
from typing import ClassVar

import pytest

from vba_addin_editor.platform import office_live
from vba_addin_editor.platform import windows_processes as wp
from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.ui import main_window as mw


@pytest.fixture
def root():
    try:
        tk_root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    tk_root.withdraw()
    yield tk_root
    tk_root.destroy()


class FakeSwap:
    made: ClassVar[list[FakeSwap]] = []
    loaded = True

    def __init__(self, path: Path, *, discard_unsaved: bool = False) -> None:
        self.path = path
        self.label = "Excel"
        self.discard_unsaved = discard_unsaved
        self.calls: list[str] = []
        FakeSwap.made.append(self)

    def __enter__(self):
        return self

    def __exit__(self, *_exc) -> None:
        self.calls.append("exit")

    def unload(self) -> bool:
        self.calls.append("unload")
        return FakeSwap.loaded

    def reload(self) -> None:
        self.calls.append("reload")


@pytest.fixture
def office(monkeypatch):
    """Excel is running; ``state["live"]`` is what asking it returns."""
    state = {"live": office_live.LiveState(office_live.EXCEL, loaded=True)}
    FakeSwap.made = []
    FakeSwap.loaded = True
    monkeypatch.setattr(office_live, "available", lambda: True)
    monkeypatch.setattr(office_live, "locate", lambda _path: state["live"])
    monkeypatch.setattr(office_live, "HostSwap", FakeSwap)
    running = wp.HostProcessProbe(corresponding_host_running=True, enumeration_failed=False, label="Excel")
    monkeypatch.setattr(mw.wp, "probe_host_process", lambda _path: running)
    return state


@pytest.fixture
def dialogs(monkeypatch):
    calls: dict[str, list] = {"askyesno": [], "showinfo": [], "showwarning": [], "showerror": []}
    answers = {"askyesno": True}

    def record(name):
        def show(*args, **_kwargs):
            calls[name].append(args[1] if len(args) > 1 else "")
            return answers.get(name)

        return show

    for name in calls:
        monkeypatch.setattr(f"vba_addin_editor.ui.main_window.messagebox.{name}", record(name))
    monkeypatch.setattr("vba_addin_editor.ui.main_window.ask_review", lambda *_a, **_k: "save")
    calls["answers"] = answers  # type: ignore[assignment]
    return calls


@pytest.fixture(autouse=True)
def _live_reload_on():
    """GUI tests share one settings file: leave live reload on for whoever runs next."""
    yield
    # Where MainWindow keeps settings under test (see its session root).
    base = os.environ.get("VBAAE_SESSION_ROOT") or str(Path(tempfile.gettempdir()) / "vbaae-test-sessions")
    path = Path(base) / "settings.json"
    saved = mw.load_settings(path)
    if not saved.live_reload:
        saved.live_reload = True
        mw.save_settings(saved, path)


def _window(root, path: Path) -> mw.MainWindow:
    window = mw.MainWindow(root)
    window.live_reload_var.set(True)
    window.settings.live_reload = True
    assert window.load_path(path)
    window.save_service.process_probe = lambda _p: True  # Excel is running
    module = window.draft.find_current("Module1")
    window._open_module_at(module.id, 1)
    window.editor.text.insert("end", "\n' live edit\n")
    window._flush_all_editors()
    return window


def _body(path: Path) -> str:
    return DocumentService().open(path).find_current("Module1").body


def test_save_with_addin_loaded_reloads_it(work_xlam: Path, root, office, dialogs):
    window = _window(root, work_xlam)
    window.save()
    root.update()
    assert window._last_result.kind == "success"
    assert [s.calls for s in FakeSwap.made] == [["unload", "reload", "exit"]]
    assert not FakeSwap.made[0].discard_unsaved
    assert "live edit" in _body(work_xlam)
    assert "reloaded in Excel" in window.status.cget("text")
    assert window._queued_save is None and dialogs["askyesno"] == []
    assert not window._prescan.busy  # Office already read the new file


def test_unsaved_changes_in_office_are_only_discarded_when_agreed(work_xlam: Path, root, office, dialogs):
    office["live"] = office_live.LiveState(office_live.EXCEL, loaded=True, unsaved=True)
    window = _window(root, work_xlam)
    dialogs["answers"]["askyesno"] = False
    window.save()
    root.update()
    assert len(dialogs["askyesno"]) == 1 and "unsaved changes" in dialogs["askyesno"][0]
    assert FakeSwap.made == [] and window.draft.is_dirty()
    assert "live edit" not in _body(work_xlam)

    dialogs["answers"]["askyesno"] = True
    window.save()
    root.update()
    assert window._last_result.kind == "success"
    assert FakeSwap.made[0].discard_unsaved is True


def test_office_running_without_the_addin_saves_directly(work_xlam: Path, root, office, dialogs):
    office["live"] = office_live.LiveState(office_live.EXCEL, loaded=False)
    FakeSwap.loaded = False
    window = _window(root, work_xlam)
    window.save()
    root.update()
    assert window._last_result.kind == "success"
    assert window._queued_save is None and dialogs["askyesno"] == []
    assert "reloaded" not in window.status.cget("text")
    assert "live edit" in _body(work_xlam)


@pytest.mark.parametrize("why", ["setting_off", "office_unreachable"])
def test_without_live_reload_the_save_waits_for_office(work_xlam: Path, root, office, dialogs, why):
    if why == "office_unreachable":
        office["live"] = office_live.LiveState(office_live.EXCEL, loaded=False, error="RPC failed")
    window = _window(root, work_xlam)
    if why == "setting_off":
        window.live_reload_var.set(False)
        window._on_live_reload_toggled()
    dialogs["answers"]["askyesno"] = False  # decline the offer to save once Office closes
    window.save()
    root.update()
    assert FakeSwap.made == []
    assert len(dialogs["askyesno"]) == 1 and "automatically" in dialogs["askyesno"][0]
    assert window.draft.is_dirty() and "live edit" not in _body(work_xlam)


def test_live_reload_setting_persists(work_xlam: Path, root, office, dialogs):
    window = _window(root, work_xlam)
    window.live_reload_var.set(False)
    window._on_live_reload_toggled()
    again = mw.MainWindow(root)
    assert again.settings.live_reload is False
