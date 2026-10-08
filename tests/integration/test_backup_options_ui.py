"""Backups on request: File -> Back Up Now, and the automatic-backup switch."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path

import pytest

from vba_addin_editor.platform.settings import load_settings
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
    win.load_path(work_xlam)
    root.update()
    try:
        yield win
    finally:
        root.destroy()


def _backups(path: Path) -> list[Path]:
    return sorted(path.parent.glob(f"{path.stem} - backup *{path.suffix}"))


def _edit_and_save(win, monkeypatch, marker: str):
    shown: list[str] = []
    monkeypatch.setattr(mw.messagebox, "askokcancel", lambda _t, msg, **k: shown.append(msg) or True)
    monkeypatch.setattr(mw.messagebox, "showinfo", lambda _t, msg, **k: shown.append(msg))
    module = win.draft.find_current("Module1")
    win._open_module_at(module.id, 1)
    win.editor.text.insert("end", f"' {marker}\n")
    win.save()
    assert not win.draft.is_dirty(), shown
    return shown


def test_auto_backup_is_on_by_default_and_saves_keep_a_backup(window, work_xlam, monkeypatch):
    assert window.auto_backup_var.get() is True
    before = work_xlam.read_bytes()
    shown = _edit_and_save(window, monkeypatch, "one")
    assert "A same-folder backup is created" in shown[0]
    (backup,) = _backups(work_xlam)
    assert backup.read_bytes() == before
    assert window.backup_service.list_records(work_xlam)[0].operation_type == "save"


def test_switching_auto_backup_off_saves_without_one_and_is_remembered(
    window, work_xlam, tmp_path, monkeypatch
):
    window.auto_backup_var.set(False)
    window._on_auto_backup_toggled()
    assert load_settings(tmp_path / "sessions" / "settings.json").auto_backup is False
    shown = _edit_and_save(window, monkeypatch, "one")
    assert "Automatic backup is off" in shown[0]
    assert "Backup:" not in shown[-1]
    assert _backups(work_xlam) == []
    assert "one" in window.draft.find_current("Module1").body
    # A second window (the next run) starts with the saved choice.
    again = mw.MainWindow(tk.Toplevel(window.root))
    assert again.auto_backup_var.get() is False and again.save_service.auto_backup is False


def test_back_up_now_copies_the_file_on_disk_not_the_draft(window, work_xlam, monkeypatch):
    on_disk = work_xlam.read_bytes()
    module = window.draft.find_current("Module1")
    window._open_module_at(module.id, 1)
    window.editor.text.insert("end", "' unsaved\n")
    window.backup_now()
    (backup,) = _backups(work_xlam)
    assert backup.read_bytes() == on_disk
    assert "Unsaved edits are not in it" in window.status.cget("text")
    assert window.draft.is_dirty()
    assert window.backup_service.list_records(work_xlam)[0].operation_type == "manual_backup"


def test_back_up_now_after_a_save_holds_the_saved_version(window, work_xlam, monkeypatch):
    window.auto_backup_var.set(False)
    window._on_auto_backup_toggled()
    _edit_and_save(window, monkeypatch, "saved-edit")
    window.backup_now()
    (backup,) = _backups(work_xlam)
    assert backup.read_bytes() == work_xlam.read_bytes()
