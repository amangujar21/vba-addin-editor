"""Several add-ins open at once, one per tab."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from types import SimpleNamespace

import pytest
from conftest import build_xlam, build_xlam_with_ribbon

from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.ui import main_window as mw


@pytest.fixture()
def root(tmp_path: Path, monkeypatch):
    monkeypatch.setenv("VBAAE_SESSION_ROOT", str(tmp_path / "sessions"))
    try:
        tk_root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    tk_root.withdraw()
    try:
        yield tk_root
    finally:
        try:
            tk_root.destroy()
        except tk.TclError:
            pass  # the window closed itself


@pytest.fixture()
def files(tmp_path: Path):
    a = build_xlam(tmp_path / "a" / "Alpha.xlam")
    b = build_xlam_with_ribbon(tmp_path / "b" / "Beta.xlam")
    return a, b


def _window(root) -> mw.MainWindow:
    window = mw.MainWindow(root)
    window.save_service.process_probe = lambda _p: False
    return window


def _module1(window):
    return next(m for m in window.draft.modules if m.current_name == "Module1")


def _type(window, text: str) -> None:
    """Show Module1 and type at its end, as a user would."""
    window._open_module_at(_module1(window).id, 1)
    window.editor.text.insert("end", text)


def test_two_files_two_tabs_each_keeps_its_edits(root, files):
    a, b = files
    window = _window(root)
    assert window.load_path(a) and window.load_path(b)
    assert [t.path.name for t in window.tabs] == ["Alpha.xlam", "Beta.xlam"]
    assert window.active_tab.path.name == "Beta.xlam" and window.draft.baseline.path.name == "Beta.xlam"
    assert window.tab_strip.winfo_manager() == "pack"
    assert list(window.tab_strip.chips) == [t.key for t in window.tabs]

    _type(window, "\n' typed in beta\n")
    alpha, beta = window.tabs
    assert window.select_tab(alpha.key)
    # The editor's unflushed text went into Beta's draft, and Alpha shows Alpha.
    assert "typed in beta" in _module1_of(beta).body
    assert window.draft is alpha.draft
    assert "typed in beta" not in window.editor.get_text()
    assert window.tab_strip.chips[beta.key].mark() == "●"  # unsaved changes

    _type(window, "\n' typed in alpha\n")
    window.select_tab(beta.key)
    assert "typed in alpha" in _module1_of(alpha).body
    assert "typed in alpha" not in _module1_of(beta).body
    # Each tab has its own undo history.
    window.undo()
    assert "typed in beta" not in _module1_of(beta).body
    assert "typed in alpha" in _module1_of(alpha).body


def _module1_of(tab):
    return next(m for m in tab.draft.modules if m.current_name == "Module1")


def test_opening_an_open_file_shows_its_tab(root, files):
    a, b = files
    window = _window(root)
    window.load_path(a)
    window.load_path(b)
    _type(window, "\n' unsaved\n")
    window._flush_all_editors()
    beta = window.active_tab
    window.select_tab(window.tabs[0].key)
    assert window.load_path(b)
    assert len(window.tabs) == 2 and window.active_tab is beta
    assert "unsaved" in _module1(window).body  # not re-read from disk


def test_ctrl_tab_cycles_and_drag_reorders(root, files, tmp_path: Path):
    a, b = files
    c = build_xlam(tmp_path / "c" / "Gamma.xlam")
    window = _window(root)
    for path in (a, b, c):
        window.load_path(path)
    names = lambda: [t.path.name for t in window.tabs]
    assert window.active_tab.path.name == "Gamma.xlam"
    window.next_tab(1)
    assert window.active_tab.path.name == "Alpha.xlam"
    window.next_tab(-1)
    assert window.active_tab.path.name == "Gamma.xlam"
    window.move_tab(window.active_tab.key, 0)
    assert names() == ["Gamma.xlam", "Alpha.xlam", "Beta.xlam"]
    assert [Path(p).name for p in window.settings.open_tabs] == names()

    # Dragging a tab along the strip moves it.
    root.update()
    strip = window.tab_strip
    first, last = strip.chips[strip.order[0]], strip.chips[strip.order[-1]]
    start = SimpleNamespace(x_root=first.winfo_rootx() + 5, y_root=first.winfo_rooty() + 5)
    end = SimpleNamespace(x_root=last.winfo_rootx() + last.winfo_width() - 2, y_root=start.y_root)
    strip.press(strip.order[0], start)
    strip.motion(end)
    strip.release(end)
    assert names() == ["Alpha.xlam", "Beta.xlam", "Gamma.xlam"]


def test_closing_tabs(root, files, monkeypatch):
    a, b = files
    window = _window(root)
    window.load_path(a)
    window.load_path(b)
    _type(window, "\n' will be discarded\n")
    asked = []
    monkeypatch.setattr(mw.messagebox, "askyesnocancel", lambda _t, msg, **_k: asked.append(msg) or None)
    window.close_active_tab()
    assert len(window.tabs) == 2 and "Beta.xlam has unsaved changes" in asked[0]  # cancelled
    monkeypatch.setattr(mw.messagebox, "askyesnocancel", lambda *_a, **_k: False)
    window.close_active_tab()
    assert [t.path.name for t in window.tabs] == ["Alpha.xlam"]
    assert window.draft.baseline.path.name == "Alpha.xlam"
    assert window.ribbon_designer.part_path is None  # Alpha has no ribbon
    window.close_active_tab()  # nothing to save: no question
    assert window.tabs == [] and window.draft is None
    assert window.tab_strip.winfo_manager() == ""
    assert window.welcome.winfo_manager() == "place"


def test_save_all_saves_every_changed_tab(root, files, monkeypatch):
    a, b = files
    window = _window(root)
    window.load_path(a)
    _type(window, "\n' alpha saved\n")
    window.load_path(b)
    _type(window, "\n' beta saved\n")
    reviews = []
    monkeypatch.setattr(mw, "ask_review", lambda *_a, **_k: reviews.append(1) or "save")
    monkeypatch.setattr(mw.messagebox, "showinfo", lambda *_a, **_k: None)
    window.save_all()
    assert len(reviews) == 2
    assert not any(t.is_dirty() for t in window.tabs)
    assert window.active_tab.path.name == "Beta.xlam"  # back where it started
    for path, text in ((a, "alpha saved"), (b, "beta saved")):
        reopened = DocumentService().open(path)
        assert text in next(m for m in reopened.modules if m.current_name == "Module1").body


def test_window_close_asks_once_for_all_tabs(root, files, monkeypatch):
    a, b = files
    window = _window(root)
    window.load_path(a)
    _type(window, "\n' x\n")
    window.load_path(b)
    _type(window, "\n' y\n")
    asked = []
    monkeypatch.setattr(mw.messagebox, "askyesnocancel", lambda _t, msg, **_k: asked.append(msg) or None)
    window.on_close()
    assert len(asked) == 1 and "2 add-ins have unsaved changes" in asked[0]
    assert "Alpha.xlam" in asked[0] and "Beta.xlam" in asked[0]
    assert len(window.tabs) == 2  # cancelled: still open


def test_tabs_reopen_in_a_new_window(root, files):
    a, b = files
    window = _window(root)
    window.load_path(a)
    window.load_path(b)
    window.select_tab(window.tabs[0].key)
    window.on_close()  # nothing unsaved: closes

    second = tk.Tk()
    second.withdraw()
    try:
        again = _window(second)
        assert again.reopen_tabs() == 2
        assert [t.path.name for t in again.tabs] == ["Alpha.xlam", "Beta.xlam"]
        assert again.active_tab.path.name == "Alpha.xlam"
        # Off: nothing is reopened.
        again.reopen_tabs_var.set(False)
        again._on_reopen_tabs_toggled()
        for tab in again.tabs:
            again.close_tab(tab.key)
        assert again.reopen_tabs() == 0
        again.reopen_tabs_var.set(True)
        again._on_reopen_tabs_toggled()
    finally:
        second.destroy()


def test_closing_a_background_tab_keeps_the_active_one(root, files):
    a, b = files
    window = _window(root)
    window.load_path(a)
    window.load_path(b)
    beta = window.active_tab
    window.close_tab(window.tabs[0].key)  # Alpha: no changes, closed where it is
    assert window.tabs == [beta] and window.active_tab is beta
