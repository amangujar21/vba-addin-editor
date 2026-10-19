"""GUI: themed window, header menus, start screen, recent files, zoom and tree markers."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path

import pytest

from vba_addin_editor.platform.settings import load_settings
from vba_addin_editor.ui import main_window as mw
from vba_addin_editor.ui import theme


@pytest.fixture()
def window(tmp_path: Path, monkeypatch):
    monkeypatch.setenv("VBAAE_SESSION_ROOT", str(tmp_path / "sessions"))
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    root.withdraw()
    try:
        yield mw.MainWindow(root)
    finally:
        root.destroy()


def _settings(win):
    return load_settings(win._settings_path)


def test_window_is_dark_without_a_native_menu_bar(window):
    assert theme.palette() is theme.DARK
    assert str(window.root.cget("menu")) == ""
    assert [b.cget("text") for b in window._header_buttons] == ["File", "Edit", "Module", "Ribbon", "View", "Help"]
    assert window.editor.text.cget("background") == theme.DARK.bg2
    assert ttk_style_lookup(window.root, "TButton", "background") == theme.DARK.bg3


def ttk_style_lookup(root, style: str, option: str) -> str:
    return str(tk.ttk.Style(root).lookup(style, option))


def test_status_bar_and_banner_are_packed_ahead_of_the_tabs(window):
    window._set_banner("Close Excel before saving this add-in.")
    order = window.root.pack_slaves()
    tabs = order.index(window.editor_notebook)
    assert order.index(window.header) < order.index(window.banner) < tabs
    assert order.index(window.statusbar) < tabs


def test_header_menus_follow_the_busy_state(window):
    window._set_menubar_state("disabled")
    assert all(str(b.cget("state")) == "disabled" for b in window._header_buttons)
    window._set_menubar_state("normal")
    assert all(str(b.cget("state")) == "normal" for b in window._header_buttons)


def test_start_screen_until_a_file_is_open(window, work_xlam: Path):
    assert window.welcome.winfo_manager() == "place"
    assert not window.save_btn.winfo_manager()
    window.load_path(work_xlam)
    assert window.welcome.winfo_manager() == ""
    assert window.save_btn.winfo_manager() == "pack"
    assert window.file_label.cget("text") == work_xlam.name
    assert _settings(window).recent_files == [str(work_xlam)]


def test_recent_file_that_moved_is_forgotten(window, work_xlam: Path, monkeypatch):
    shown: list[str] = []
    monkeypatch.setattr(mw.messagebox, "showinfo", lambda _t, msg, **k: shown.append(msg))
    missing = work_xlam.with_name("Gone.xlam")
    window.settings.recent_files = [str(missing), str(work_xlam)]
    window.open_recent(missing)
    assert shown and "Gone.xlam" in shown[0]
    assert _settings(window).recent_files == [str(work_xlam)]
    window._fill_recent_menu()
    assert window.recent_menu.entrycget(0, "label") == str(work_xlam)


def test_zoom_changes_every_editor_and_is_remembered(window):
    font = theme.code_font(window.root)
    window.zoom(1)
    window.zoom(1)
    assert int(window.root.tk.call("font", "configure", font, "-size")) == theme.CODE_SIZE + 2
    assert str(window.xml_editor.text.cget("font")) == font
    assert _settings(window).editor_font_size == theme.CODE_SIZE + 2
    window.zoom(0)
    assert _settings(window).editor_font_size == theme.CODE_SIZE


def test_light_theme_restyles_the_open_window(window):
    window.theme_var.set("light")
    window._on_theme_chosen()
    try:
        assert theme.palette() is theme.LIGHT
        assert window.editor.text.cget("background") == theme.LIGHT.bg2
        assert ttk_style_lookup(window.root, "TButton", "background") == theme.LIGHT.bg3
        assert _settings(window).theme == "light"
    finally:
        theme.apply_theme(window.root, "dark")


def test_tree_marks_edited_and_locked_modules(window, work_xlam: Path):
    window.load_path(work_xlam)
    module = window.draft.find_current("Module1")
    window._open_module_at(module.id, 1)
    window.editor.text.insert("end", "' edited\n")
    window._flush_all_editors()
    window._refresh_tree()
    window._refresh_state()
    assert window.tree.item(module.id, "text") == "Module1  ●"
    assert window.tree.item(module.id, "tags") == ("dirty",)
    locked = next(m for m in window.draft.modules if not (m.can_delete or m.can_rename))
    assert window.tree.item(locked.id, "text").startswith("🔒 ")
    assert window.tree.item(locked.id, "tags") == ("locked",)
    assert window.dirty_label.cget("text") == "●  1 unsaved"
