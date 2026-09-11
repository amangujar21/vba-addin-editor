from __future__ import annotations

import tkinter as tk
from pathlib import Path

import pytest

from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.ui.main_window import MainWindow


def _window(root) -> MainWindow:
    root.withdraw()
    return MainWindow(root)


def test_save_button_enabled_when_host_was_running_at_open(work_xlam: Path, monkeypatch):
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    monkeypatch.setattr(
        "vba_addin_editor.services.document_service.wp.probe_host_process",
        lambda path: _probe(True),
    )
    try:
        window = _window(root)
        window.load_path(work_xlam)
        assert str(window.save_btn.cget("state")) == "normal"
        module_id = next(m.id for m in window.draft.modules if m.current_name == "Module1")
        window.tree.selection_set(module_id)
        root.update()
        window.editor.text.insert("end", "\n' edited\n")
        window._flush_all_editors()
        window.save_service.process_probe = lambda _p: False
        monkeypatch.setattr(
            "vba_addin_editor.platform.windows_processes.probe_host_process",
            lambda path: _probe(False),
        )
        window._refresh_host_status()
        window._refresh_state()
        assert str(window.save_btn.cget("state")) == "normal"
        shown = []
        monkeypatch.setattr(
            "vba_addin_editor.ui.main_window.messagebox.askokcancel",
            lambda *_a, **_k: True,
        )
        monkeypatch.setattr(
            "vba_addin_editor.ui.main_window.messagebox.showinfo",
            lambda *_a, **_k: shown.append("ok"),
        )
        window.save_btn.invoke()
        root.update()
        assert window.draft is not None
        assert not window.draft.is_dirty()
        reopened = DocumentService().open(work_xlam)
        body = next(m for m in reopened.modules if m.current_name == "Module1").body
        assert "edited" in body
    finally:
        root.destroy()


def test_blocked_save_leaves_button_retryable(work_xlam: Path, monkeypatch):
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    try:
        window = _window(root)
        window.load_path(work_xlam)
        module_id = next(m.id for m in window.draft.modules if m.current_name == "Module1")
        window.tree.selection_set(module_id)
        root.update()
        window.editor.text.insert("end", "\n' retry-me\n")
        window._flush_all_editors()
        window.save_service.process_probe = lambda _p: True
        monkeypatch.setattr(
            "vba_addin_editor.ui.main_window.messagebox.askokcancel",
            lambda *_a, **_k: True,
        )
        monkeypatch.setattr(
            "vba_addin_editor.ui.main_window.messagebox.showwarning",
            lambda *_a, **_k: None,
        )
        monkeypatch.setattr(
            "vba_addin_editor.ui.main_window.messagebox.showinfo",
            lambda *_a, **_k: None,
        )
        window.save_btn.invoke()
        root.update()
        assert window.draft.is_dirty()
        assert str(window.save_btn.cget("state")) == "normal"
        window.save_service.process_probe = lambda _p: False
        window.save_btn.invoke()
        root.update()
        assert not window.draft.is_dirty()
    finally:
        root.destroy()


def test_close_cancel_keeps_dirty_draft(work_xlam: Path, monkeypatch):
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    destroyed = []
    try:
        window = _window(root)
        window.load_path(work_xlam)
        module_id = next(m.id for m in window.draft.modules if m.current_name == "Module1")
        window.tree.selection_set(module_id)
        root.update()
        window.editor.text.insert("end", "\n' keep\n")
        window._flush_all_editors()
        monkeypatch.setattr(
            "vba_addin_editor.ui.main_window.messagebox.askyesnocancel",
            lambda *_a, **_k: None,
        )
        original_destroy = window.root.destroy
        monkeypatch.setattr(window.root, "destroy", lambda: destroyed.append(True) or original_destroy())
        window.on_close()
        assert destroyed == []
        assert window.draft.is_dirty()
    finally:
        if not destroyed:
            root.destroy()


def test_gui_save_publishes_session_baseline(work_xlam: Path, monkeypatch):
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    try:
        window = _window(root)
        window.load_path(work_xlam)
        assert window.session is not None
        old_hash = window.session.baseline_sha256
        module_id = next(m.id for m in window.draft.modules if m.current_name == "Module1")
        window.tree.selection_set(module_id)
        root.update()
        window.editor.text.insert("end", "\n' session-pub\n")
        window._flush_all_editors()
        monkeypatch.setattr(
            "vba_addin_editor.ui.main_window.messagebox.askokcancel",
            lambda *_a, **_k: True,
        )
        monkeypatch.setattr(
            "vba_addin_editor.ui.main_window.messagebox.showinfo",
            lambda *_a, **_k: None,
        )
        window.save_btn.invoke()
        root.update()
        assert window.session.baseline_sha256 != old_hash
        assert window.session.captured_path.read_bytes() == work_xlam.read_bytes()
        assert (window.session.session_dir / "complete.marker").exists()
    finally:
        root.destroy()


def _probe(running: bool):
    from vba_addin_editor.platform.windows_processes import HostProcessProbe

    return HostProcessProbe(
        corresponding_host_running=running,
        enumeration_failed=False,
        label="Excel",
        exe_name="EXCEL.EXE",
    )
