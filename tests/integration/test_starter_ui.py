"""File → New Add-in and Ribbon → Create Ribbon in the GUI."""

from __future__ import annotations

import tkinter as tk
import zipfile
from pathlib import Path

import pytest
from conftest import build_xlam
from helpers import make_service

from vba_addin_editor.services import ribbon_service as rs
from vba_addin_editor.services import starter_service as ss
from vba_addin_editor.ui import main_window as mw
from vba_addin_editor.ui.ribbon_dialogs import ButtonSpec

PART = ss.RIBBON_PART_PATH


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
        tk_root.destroy()


@pytest.mark.parametrize("extension", [".xlam", ".ppam"])
def test_new_file_with_ribbon_then_button_saves(root, tmp_path: Path, monkeypatch, extension):
    win = mw.MainWindow(root)
    read_back: list[Path] = []
    real_read = mw.DocumentService.open
    monkeypatch.setattr(
        mw.DocumentService, "open", lambda self, p: read_back.append(p) or real_read(self, p)
    )
    path = tmp_path / "out" / f"Fresh{extension}"
    path.parent.mkdir()
    assert win.create_new_file(path, ss.StarterRibbonSpec("Fresh", "Tools"))
    root.update()
    assert read_back == []  # opened from the bytes it wrote
    assert win.draft is not None and win.draft.baseline.path == path.resolve()
    assert win.draft.is_dirty() and win.ribbon_designer.part_path == PART
    assert "Save File" in win.status.cget("text")

    # Add a macro and a button for it into the new group.
    m1 = win.draft.find_current("Module1")
    m1.body += "Public Sub SayHi(control As IRibbonControl)\n    MsgBox \"hi\"\nEnd Sub\n"
    win._open_module_at(m1.id, 1)
    win.editor.set_text(m1.body)
    part = win.draft.xml_part_by_path(PART)
    group = rs.find_containers(part.text, PART)[0]
    spec = ButtonSpec(group, "btnHi", "Say Hi", "HappyFace", "large", "", "")
    win.apply_ribbon_button(m1.id, "SayHi", None, spec)
    assert rs.check_ribbon(win.draft) == []

    result = make_service().save_addin(win.draft)
    assert result.kind == "success", result
    with zipfile.ZipFile(path) as package:
        assert 'onAction="SayHi"' in package.read(PART).decode("utf-8")
        assert "customUI/customUI14.xml" in package.read("_rels/.rels").decode("utf-8")


def test_new_file_without_ribbon(root, tmp_path: Path):
    win = mw.MainWindow(root)
    path = tmp_path / "Plain.pptm"
    assert win.create_new_file(path, None)
    assert not win.draft.is_dirty()
    assert rs.ribbon_parts(win.draft) == []


def test_new_file_refuses_an_existing_name(root, tmp_path: Path, monkeypatch):
    errors: list[str] = []
    monkeypatch.setattr(mw.messagebox, "showerror", lambda _t, msg, **k: errors.append(msg))
    win = mw.MainWindow(root)
    path = tmp_path / "Taken.xlam"
    path.write_bytes(b"x")
    assert not win.create_new_file(path, None)
    assert "already exists" in errors[0] and path.read_bytes() == b"x"
    assert win.draft is None


def test_new_file_dialog_adds_the_chosen_type_extension(root, tmp_path: Path, monkeypatch):
    win = mw.MainWindow(root)

    def fake_save_dialog(**kwargs):
        kwargs["typevariable"].set("PowerPoint Add-in")
        return str(tmp_path / "Typed")

    made: list = []
    monkeypatch.setattr(mw.filedialog, "asksaveasfilename", fake_save_dialog)
    monkeypatch.setattr(mw.StarterRibbonDialog, "wait", lambda self: None)
    monkeypatch.setattr(mw.StarterRibbonDialog, "__init__", _quiet_dialog(skipped=True))
    monkeypatch.setattr(win, "create_new_file", lambda p, spec: made.append((p, spec)))
    win.new_file()
    assert made == [(tmp_path / "Typed.ppam", None)]


def _quiet_dialog(*, skipped: bool):
    def init(self, *_a, **_k):
        self.skipped = skipped
        self.result = None

    return init


def test_create_ribbon_command_is_one_undo_step(root, tmp_path: Path, monkeypatch):
    win = mw.MainWindow(root)
    win.load_path(build_xlam(tmp_path / "work" / "Plain.xlam"))
    monkeypatch.setattr(win, "ask_starter_ribbon", lambda _intro: ss.StarterRibbonSpec("T", "G"))
    rels_before = win.draft.xml_part_by_path("_rels/.rels").text
    assert win.create_ribbon()
    root.update()
    assert win.draft.xml_part_by_path(PART) is not None
    assert win.xml_tree.exists("xml::" + PART)
    win.undo()
    assert win.draft.xml_part_by_path(PART) is None
    assert win.draft.xml_part_by_path("_rels/.rels").text == rels_before
    assert not win.draft.is_dirty()
    win.redo()
    assert win.draft.xml_part_by_path(PART) is not None

    infos: list[str] = []
    monkeypatch.setattr(mw.messagebox, "showinfo", lambda _t, msg, **k: infos.append(msg))
    assert not win.create_ribbon()
    assert "already has ribbon XML" in infos[0]


def test_add_button_offers_to_create_the_ribbon(root, tmp_path: Path, monkeypatch):
    win = mw.MainWindow(root)
    win.load_path(build_xlam(tmp_path / "work" / "Plain.xlam"))
    m1 = win.draft.find_current("Module1")
    win._open_module_at(m1.id, 2)  # inside Public Sub Hello()
    asked: list[str] = []

    def ask(intro):
        asked.append(intro)
        return ss.StarterRibbonSpec("Mine", "Things")

    shown: list = []

    class FakeDialog:
        def __init__(self, _root, **kwargs):
            shown.append(kwargs["containers"])

        def wait(self):
            return None

    monkeypatch.setattr(win, "ask_starter_ribbon", ask)
    monkeypatch.setattr(mw, "AddButtonDialog", FakeDialog)
    win.add_ribbon_button()
    assert "no ribbon yet" in asked[0]
    assert [c.control_id for c in shown[0]] == ["grpThings"]

    # Declining leaves the file without a ribbon and shows no button form.
    other = mw.MainWindow(root)
    other.load_path(build_xlam(tmp_path / "work2" / "Plain.xlam"))
    m1 = other.draft.find_current("Module1")
    other._open_module_at(m1.id, 2)
    shown.clear()
    monkeypatch.setattr(other, "ask_starter_ribbon", lambda _intro: None)
    other.add_ribbon_button()
    assert shown == [] and rs.ribbon_parts(other.draft) == []
