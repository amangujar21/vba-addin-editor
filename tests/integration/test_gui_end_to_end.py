"""Scripted GUI end-to-end: open → edit in the Text widget → Save → verify on disk.

Real ReplaceFileW commit runs because temp paths are local same-volume files.
"""

from __future__ import annotations

import tkinter as tk
from pathlib import Path

import pytest

from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.ui.main_window import MainWindow


def test_gui_edit_and_save_roundtrip(work_xlam: Path):
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    root.withdraw()
    try:
        window = MainWindow(root)
        window.load_path(work_xlam)
        assert window.draft is not None

        # Select Module1 and type into the real Text widget.
        module_id = next(
            m.id for m in window.draft.modules if m.current_name == "Module1"
        )
        window.tree.selection_set(module_id)
        root.update()
        window.editor.text.insert("end", "\r\nPublic Sub AddedByGui()\r\nEnd Sub\r\n")
        window._flush_all_editors()

        # Drive the save pipeline directly (messageboxes bypassed).
        result = window.save_service.save_addin(window.draft)
        assert result.kind == "success", result
        assert result.backup_path.exists()

        reopened = DocumentService().open(work_xlam)
        m1 = next(m for m in reopened.modules if m.current_name == "Module1")
        assert "AddedByGui" in m1.body
        # GUI state refreshed: draft clean.
        assert not window.draft.is_dirty()
        root.update()
    finally:
        root.destroy()


def test_gui_xml_edit_roundtrip(work_pptm: Path):
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    root.withdraw()
    try:
        window = MainWindow(root)
        window.load_path(work_pptm)
        assert window.draft is not None
        assert window.draft.xml_parts, "XML tab should be populated for pptm"

        window.editor_notebook.select(window.xml_tab)
        window.xml_tree.selection_set("xml::docProps/core.xml")
        root.update()
        window.xml_editor.text.insert("end", "\n<!--VBAAE_XML_EDITED-->")
        window._flush_all_editors()
        assert window.draft.xml_part_by_path("docProps/core.xml").is_dirty()

        result = window.save_service.save_addin(window.draft)
        assert result.kind == "success", result
        reopened = DocumentService().open(work_pptm)
        assert "<!--VBAAE_XML_EDITED-->" in reopened.xml_part_by_path("docProps/core.xml").text
        assert not reopened.is_dirty()
    finally:
        root.destroy()


def test_gui_combined_vba_and_xml_roundtrip(work_pptm: Path):
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    root.withdraw()
    try:
        window = MainWindow(root)
        window.load_path(work_pptm)
        module_id = next(m.id for m in window.draft.modules if m.current_name == "Module1")
        window.tree.selection_set(module_id)
        root.update()
        window.editor.text.insert("end", "\r\nPublic Sub AddedByGui()\r\nEnd Sub\r\n")
        window.editor_notebook.select(window.xml_tab)
        window.xml_tree.selection_set("xml::docProps/core.xml")
        root.update()
        window.xml_editor.text.insert("end", "\n<!--VBAAE_XML_EDITED-->")
        window._flush_all_editors()

        result = window.save_service.save_addin(window.draft)
        assert result.kind == "success", result
        reopened = DocumentService().open(work_pptm)
        m1 = next(m for m in reopened.modules if m.current_name == "Module1")
        assert "AddedByGui" in m1.body
        assert "<!--VBAAE_XML_EDITED-->" in reopened.xml_part_by_path("docProps/core.xml").text
        assert not reopened.is_dirty()
    finally:
        root.destroy()
