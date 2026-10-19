"""Saving while Office holds or rewrites the add-in: wait, then merge onto its version."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path

import pytest
from helpers import make_service

from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.ui.main_window import MainWindow


@pytest.fixture
def root():
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    root.withdraw()
    yield root
    root.destroy()


@pytest.fixture
def dialogs(monkeypatch):
    calls: dict[str, list] = {"askokcancel": [], "askyesno": [], "showinfo": [], "showwarning": [], "showerror": []}
    answers = {"askokcancel": True, "askyesno": True}

    def record(name):
        def show(*args, **_kwargs):
            calls[name].append(args[1] if len(args) > 1 else "")
            return answers.get(name)

        return show

    for name in calls:
        monkeypatch.setattr(f"vba_addin_editor.ui.main_window.messagebox.{name}", record(name))
    calls["answers"] = answers  # type: ignore[assignment]
    return calls


def _edit_module1(window: MainWindow, root, text: str) -> None:
    module_id = next(m.id for m in window.draft.modules if m.current_name == "Module1")
    window.tree.selection_set(module_id)
    root.update()
    window.editor.text.insert("end", text)
    window._flush_all_editors()


def _module_body(path: Path, name: str) -> str | None:
    draft = DocumentService().open(path)
    return next((m.body for m in draft.modules if m.current_name == name), None)


def _office_saves(path: Path, edits: dict[str, tuple[str, str]]) -> None:
    """Stand-in for Excel rewriting the add-in after the editor opened it."""
    draft = DocumentService().open(path)
    for name, (old, new) in edits.items():
        mod = next(m for m in draft.modules if m.current_name == name)
        mod.body = mod.body.replace(old, new) if old else mod.body + new
    assert make_service().save_addin(draft).kind == "success"


def test_save_waits_for_office_then_saves(work_xlam: Path, root, dialogs):
    window = MainWindow(root)
    window.load_path(work_xlam)
    _edit_module1(window, root, "\n' queued\n")
    window.save_service.process_probe = lambda _p: True

    window.save_btn.invoke()
    root.update()

    assert len(dialogs["askyesno"]) == 1
    assert window._queued_save is not None
    assert window.wait_bar.winfo_manager() == "pack"
    assert window.draft.is_dirty()
    window._poll_queued_save()  # Office still running: keeps waiting
    assert window._queued_save is not None

    window.save_service.process_probe = lambda _p: False
    window._poll_queued_save()
    root.update()

    assert window._queued_save is None
    assert window.wait_bar.winfo_manager() == ""
    assert not window.draft.is_dirty()
    assert "queued" in (_module_body(work_xlam, "Module1") or "")
    assert len(dialogs["askokcancel"]) == 1  # reviewed once, before waiting


def test_declining_to_wait_keeps_draft(work_xlam: Path, root, dialogs):
    window = MainWindow(root)
    window.load_path(work_xlam)
    _edit_module1(window, root, "\n' not-queued\n")
    window.save_service.process_probe = lambda _p: True
    dialogs["answers"]["askyesno"] = False

    window.save_btn.invoke()
    root.update()

    assert window._queued_save is None
    assert window.draft.is_dirty()
    assert str(window.save_btn.cget("state")) == "normal"


def test_queued_save_reviews_again_after_further_edits(work_xlam: Path, root, dialogs):
    window = MainWindow(root)
    window.load_path(work_xlam)
    _edit_module1(window, root, "\n' first\n")
    window.save_service.process_probe = lambda _p: True
    window.save_btn.invoke()
    root.update()
    window.editor.text.insert("end", "' second\n")

    window.save_service.process_probe = lambda _p: False
    window._poll_queued_save()
    root.update()

    assert len(dialogs["askokcancel"]) == 2
    assert not window.draft.is_dirty()
    assert "second" in (_module_body(work_xlam, "Module1") or "")


def test_cancel_queued_save_writes_nothing(work_xlam: Path, root, dialogs):
    before = work_xlam.read_bytes()
    window = MainWindow(root)
    window.load_path(work_xlam)
    _edit_module1(window, root, "\n' cancelled\n")
    window.save_service.process_probe = lambda _p: True
    window.save_btn.invoke()
    root.update()

    window.wait_cancel_btn.invoke()
    window.save_service.process_probe = lambda _p: False
    window._poll_queued_save()
    root.update()

    assert window._queued_save is None
    assert window.draft.is_dirty()
    assert work_xlam.read_bytes() == before


def test_save_merges_onto_file_office_rewrote(work_xlam: Path, root, dialogs):
    window = MainWindow(root)
    window.load_path(work_xlam)
    _edit_module1(window, root, "\n' mine\n")
    _office_saves(work_xlam, {"ThisWorkbook": ("", "' from office\r\n")})

    window.save_btn.invoke()
    root.update()

    assert not dialogs["showwarning"] and not dialogs["showerror"]
    assert not window.draft.is_dirty()
    assert "mine" in (_module_body(work_xlam, "Module1") or "")
    assert "from office" in (_module_body(work_xlam, "ThisWorkbook") or "")
    assert any("applied on top" in text for text in dialogs["showinfo"])


def test_overlapping_office_change_goes_to_conflicts(work_xlam: Path, root, dialogs, monkeypatch):
    window = MainWindow(root)
    window.load_path(work_xlam)
    module_id = next(m.id for m in window.draft.modules if m.current_name == "Module1")
    window.tree.selection_set(module_id)
    root.update()
    text = window.editor.text.get("1.0", "end-1c").replace("ORIGINAL", "MINE")
    window.editor.set_text(text)
    window.editor.text.event_generate("<<Modified>>")
    window.draft.module_by_id(module_id).body = text
    _office_saves(work_xlam, {"Module1": ("ORIGINAL", "OFFICE")})
    after_office = work_xlam.read_bytes()
    opened = []
    monkeypatch.setattr(window, "_run_conflict_resolution", lambda external: opened.append(external))

    window.save_btn.invoke()
    root.update()

    assert opened, "overlapping edits must go to conflict resolution"
    assert any("overlap" in text for text in dialogs["showinfo"])
    assert window.draft.is_dirty()
    assert work_xlam.read_bytes() == after_office
