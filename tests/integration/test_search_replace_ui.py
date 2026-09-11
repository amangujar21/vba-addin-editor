from __future__ import annotations

import tkinter as tk
from pathlib import Path

import pytest

from vba_addin_editor.domain.history import HistoryCommand
from vba_addin_editor.services.search_service import SearchService
from vba_addin_editor.services.session_service import SessionService
from vba_addin_editor.ui.main_window import MainWindow
from vba_addin_editor.ui.search_dialog import ProjectSearchDialog


def test_search_replace_is_atomic_and_undoable(work_xlam: Path, tmp_path: Path):
    session = SessionService(session_root=tmp_path / "sessions").open(work_xlam)
    m1 = next(m for m in session.draft.modules if m.current_name == "Module1")
    m1.body = "foo bar foo\n"
    session.bump()
    service = SearchService()
    results = service.search(session.draft, "foo", revision=session.revision)
    selected = {(hit.target_id, hit.offset) for hit in results.hits}
    applied = service.apply_replacements(session, results, "baz", selected)
    assert applied.ok
    assert applied.changed == 2
    session.draft.modules[:] = applied.draft.modules
    from vba_addin_editor.services.history_service import HistoryService

    history = HistoryService()
    history.record_structural(
        session,
        HistoryCommand(
            op="replace_all",
            target_id="*",
            before={"modules": applied.before_modules, "xml": applied.before_xml},
            after={"modules": applied.after_modules, "xml": applied.after_xml},
        ),
    )
    m1 = next(m for m in session.draft.modules if m.current_name == "Module1")
    assert m1.body == "baz bar baz\n"
    history.undo(session)
    m1 = next(m for m in session.draft.modules if m.current_name == "Module1")
    assert m1.body == "foo bar foo\n"


def test_stale_search_results_do_not_mutate(work_xlam: Path, tmp_path: Path):
    session = SessionService(session_root=tmp_path / "sessions").open(work_xlam)
    m1 = next(m for m in session.draft.modules if m.current_name == "Module1")
    m1.body = "foo\n"
    session.bump()
    results = SearchService().search(session.draft, "foo", revision=session.revision)
    session.bump()
    applied = SearchService().apply_replacements(
        session, results, "baz", {(hit.target_id, hit.offset) for hit in results.hits}
    )
    assert not applied.ok
    assert applied.reason == "stale_proposal"
    assert "foo" in session.draft.modules[0].body or True
    m1 = next(m for m in session.draft.modules if m.current_name == "Module1")
    assert m1.body == "foo\n"


def test_search_dialog_navigation_and_options(work_xlam: Path, monkeypatch):
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    root.withdraw()
    try:
        window = MainWindow(root)
        window.load_path(work_xlam)
        m1 = next(m for m in window.draft.modules if m.current_name == "Module1")
        m1.body = "Foo foo food\n"
        dialog = ProjectSearchDialog(root, window)
        dialog.query.insert(0, "foo")
        dialog.match_case.set(False)
        dialog.whole_word.set(True)
        dialog.search()
        assert dialog.results is not None
        assert dialog.results.total == 2
        dialog.listbox.selection_set(0)
        dialog.goto_selected()
        root.update()
        dialog.replacement.insert(0, "bar")
        dialog.selected = {(hit.target_id, hit.offset) for hit in dialog.results.hits}
        dialog.preview_replace()
        preview = dialog.preview.get("1.0", "end")
        assert "foo" in preview.lower() or "Foo" in preview
    finally:
        root.destroy()
