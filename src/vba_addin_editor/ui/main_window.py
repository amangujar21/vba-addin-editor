"""Main application window (plan sections 9, 18, 43-46, 57)."""

from __future__ import annotations

import os
import re
import tempfile
import threading
import time
import tkinter as tk
from collections.abc import Callable
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import TypeVar

from vba_addin_editor.adapters.ooxml_package_adapter import XML_EDITABLE_EXTENSIONS
from vba_addin_editor.adapters.pyopenvba_adapter import AdapterError
from vba_addin_editor.domain.capabilities import RESTRICTION_MESSAGES
from vba_addin_editor.domain.changes import compute_changes, dirty_count
from vba_addin_editor.domain.document import (
    DocumentDraft,
    FormControlSnapshot,
    FormDesignSnapshot,
    ModuleDisplayKind,
    ModuleDraft,
    new_module_id,
    revert_all,
)
from vba_addin_editor.domain.history import HistoryCommand
from vba_addin_editor.platform import windows_processes as wp
from vba_addin_editor.services import ribbon_service as rs
from vba_addin_editor.services.backup_service import BackupService
from vba_addin_editor.services.conflict_service import ConflictService
from vba_addin_editor.services.diagnostics_service import dialog_text, report_for
from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.services.folder_sync_service import FolderSyncService
from vba_addin_editor.services.history_service import (
    HistoryService,
    snapshot_modules,
    snapshot_xml,
)
from vba_addin_editor.services.import_export_service import ImportExportService
from vba_addin_editor.services.recovery_service import (
    IDLE_MS,
    MAX_INTERVAL_MS,
    RecoveryOpenService,
    RecoveryService,
)
from vba_addin_editor.services.review_service import ReviewService
from vba_addin_editor.services.save_service import SaveService
from vba_addin_editor.services.search_service import SearchService
from vba_addin_editor.services.session_service import SessionService
from vba_addin_editor.services.validation_service import validate_module_name
from vba_addin_editor.ui.code_editor import CodeEditor
from vba_addin_editor.ui.conflict_dialog import ConflictDialog
from vba_addin_editor.ui.ribbon_dialogs import AddButtonDialog, RibbonButtonsDialog, choose
from vba_addin_editor.ui.search_dialog import ProjectSearchDialog
from vba_addin_editor.ui.xml_editor import XmlEditor
from vba_addin_editor.version import APP_NAME, VERSION, build_identity

_FILETYPES = [
    ("Supported Office VBA files", "*.xlam;*.ppam;*.pptm"),
    ("Excel Add-ins", "*.xlam"),
    ("PowerPoint Add-ins", "*.ppam"),
    ("PowerPoint Macro-Enabled Presentations", "*.pptm"),
]


_T = TypeVar("_T")


class MainWindow:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.draft: DocumentDraft | None = None
        self.session = None
        self.current_module_id: str | None = None
        self._saving = False
        self._live_host_running = False
        self._live_host_probe_failed = False
        self._host_label = "Office"
        self._recovery_warning = ""
        self._idle_after = None
        self._max_after = None
        self._poll_after = None
        self._last_vba_text = ""
        self._last_xml_text = ""
        self.doc_service = DocumentService()
        session_root = os.environ.get("VBAAE_SESSION_ROOT")
        if session_root:
            root_path = Path(session_root)
        elif os.environ.get("PYTEST_CURRENT_TEST"):
            root_path = Path(tempfile.gettempdir()) / "vbaae-test-sessions"
        else:
            root_path = None
        self.session_service = SessionService(
            document_service=self.doc_service, session_root=root_path
        )
        self.save_service = SaveService(adapter=self.doc_service.adapter, progress=self._on_progress)
        self.ie_service = ImportExportService(adapter=self.doc_service.adapter)
        self.backup_service = BackupService(adapter=self.doc_service.adapter)
        self.recovery_service = RecoveryService(session_root=root_path)
        self.review_service = ReviewService()
        self.conflict_service = ConflictService(self.doc_service)
        self.search_service = SearchService()
        self.folder_sync = FolderSyncService()
        self.history_service = HistoryService()
        self.recovery_open = RecoveryOpenService(self.recovery_service, self.session_service)
        self._search_dialog = None
        self._ribbon_dialog: RibbonButtonsDialog | None = None

        root.title(APP_NAME)
        root.geometry("1000x680")
        root.minsize(760, 480)
        self._build_menu()
        self._build_toolbar()
        self._build_panes()
        self._build_statusbar()
        self._bind_shortcuts()
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        root.bind("<Destroy>", self._on_root_destroy, add=True)
        root.bind("<FocusIn>", self._on_focus_in, add=True)
        self.editor.on_change = self._on_vba_edited
        self.xml_editor.on_change = self._on_xml_edited
        self.editor.on_undo = self.undo
        self.editor.on_redo = self.redo
        self.xml_editor.on_undo = self.undo
        self.xml_editor.on_redo = self.redo
        self._recovery_after = None
        self._schedule_host_poll()
        if not os.environ.get("PYTEST_CURRENT_TEST"):
            self._recovery_after = self.root.after_idle(self._offer_recovery)

    # -- construction ------------------------------------------------------

    def _build_menu(self) -> None:
        bar = tk.Menu(self.root)
        file_menu = tk.Menu(bar, tearoff=0)
        file_menu.add_command(label="Open Office VBA File…", accelerator="Ctrl+O", command=self.open_file)
        file_menu.add_command(label="Save File", accelerator="Ctrl+S", command=self.save)
        file_menu.add_command(label="Save a Copy…", accelerator="Ctrl+Shift+S", command=self.save_copy)
        file_menu.add_separator()
        file_menu.add_command(label="Restore Backup…", command=self.restore_backup)
        file_menu.add_command(label="Backup Browser…", command=self.backup_browser)
        file_menu.add_command(label="Compare External Changes…", command=self.compare_external_changes)
        file_menu.add_command(label="Open File Location", command=self.open_location)
        file_menu.add_separator()
        file_menu.add_command(label="Export Source Folder…", command=self.export_source_folder)
        file_menu.add_command(label="Preview Folder Changes…", command=self.preview_folder_changes)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_close)
        bar.add_cascade(label="File", menu=file_menu)

        edit_menu = tk.Menu(bar, tearoff=0)
        edit_menu.add_command(label="Undo", accelerator="Ctrl+Z", command=self.undo)
        edit_menu.add_command(label="Redo", accelerator="Ctrl+Y", command=self.redo)
        edit_menu.add_separator()
        edit_menu.add_command(label="Find in Project…", accelerator="Ctrl+Shift+F", command=self.project_search)
        bar.add_cascade(label="Edit", menu=edit_menu)

        module_menu = tk.Menu(bar, tearoff=0)
        module_menu.add_command(label="Add Standard Module…", command=lambda: self.add_module(standard=True))
        module_menu.add_command(label="Add Class Module…", command=lambda: self.add_module(standard=False))
        module_menu.add_command(label="Rename…", command=self.rename_module)
        module_menu.add_command(label="Delete", command=self.delete_module)
        module_menu.add_separator()
        module_menu.add_command(label="Import .bas/.cls…", command=self.import_file)
        module_menu.add_command(label="Export Current Module…", command=self.export_module)
        module_menu.add_command(label="Export All…", command=self.export_all)
        bar.add_cascade(label="Module", menu=module_menu)

        ribbon_menu = tk.Menu(bar, tearoff=0)
        ribbon_menu.add_command(label="Ribbon Buttons…", command=self.show_ribbon_buttons)
        ribbon_menu.add_command(
            label="Go to Macro / Button", accelerator="F12", command=self.goto_callback_under_cursor
        )
        ribbon_menu.add_command(
            label="Add Ribbon Button for This Macro…", command=self.add_ribbon_button
        )
        ribbon_menu.add_separator()
        ribbon_menu.add_command(label="Check Ribbon Callbacks", command=self.check_ribbon_callbacks)
        bar.add_cascade(label="Ribbon", menu=ribbon_menu)

        help_menu = tk.Menu(bar, tearoff=0)
        help_menu.add_command(label="About", command=self.show_about)
        help_menu.add_command(label="Copy Diagnostic Report", command=self.copy_last_diagnostic)
        bar.add_cascade(label="Help", menu=help_menu)
        self._last_result = None
        self.menubar = bar
        self.root.config(menu=bar)

    def _build_toolbar(self) -> None:
        bar = ttk.Frame(self.root, padding=4)
        bar.pack(fill="x")
        ttk.Button(bar, text="Open", command=self.open_file).pack(side="left", padx=2)
        self.save_btn = ttk.Button(bar, text="Save File", command=self.save, state="disabled")
        self.save_btn.pack(side="left", padx=2)
        self.review_btn = ttk.Button(bar, text="Review Changes", command=self.review_changes, state="disabled")
        self.review_btn.pack(side="left", padx=2)
        self.revert_btn = ttk.Button(bar, text="Revert All", command=self.revert_all, state="disabled")
        self.revert_btn.pack(side="left", padx=2)
        self.file_label = ttk.Label(bar, text="No add-in open")
        self.file_label.pack(side="right", padx=8)

    def _build_panes(self) -> None:
        self.editor_notebook = ttk.Notebook(self.root)
        vba_tab = ttk.Frame(self.editor_notebook)
        self._build_vba_tab(vba_tab)
        self.xml_tab = ttk.Frame(self.editor_notebook)
        self._build_xml_tab(self.xml_tab)
        self.forms_tab = ttk.Frame(self.editor_notebook)
        self._build_forms_tab(self.forms_tab)
        self.editor_notebook.add(vba_tab, text="VBA")
        self.editor_notebook.add(self.xml_tab, text="XML")
        self.editor_notebook.add(self.forms_tab, text="Forms")
        self.editor_notebook.pack(fill="both", expand=True)
        self.banner = ttk.Label(self.root, text="", background="#fff3cd", padding=4)
        self.banner.pack(fill="x", before=self.editor_notebook)
        self.banner.pack_forget()
        self._install_ribbon_actions()

    def _install_ribbon_actions(self) -> None:
        xml_menu = self.xml_editor.context_menu
        xml_menu.add_extra(
            "Go to Macro",
            self.goto_callback_under_cursor,
            lambda: self._xml_cursor_callback() is not None,
        )
        vba_menu = self.editor.context_menu
        vba_menu.add_extra(
            "Add Ribbon Button for This Macro…",
            self.add_ribbon_button,
            lambda: self._cursor_procedure()[1] is not None,
        )
        vba_menu.add_extra(
            "Show Ribbon Buttons Using This Macro",
            self.show_buttons_for_macro,
            lambda: self._cursor_procedure()[1] is not None,
        )
        self.xml_editor.text.bind("<Control-Button-1>", self._on_xml_ctrl_click)

    def _build_vba_tab(self, parent) -> None:
        panes = ttk.Panedwindow(parent, orient="horizontal")
        panes.pack(fill="both", expand=True)

        left = ttk.Frame(panes, padding=4)
        self.tree = ttk.Treeview(left, show="tree", selectmode="browse")
        tree_sb = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        tree_sb.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.on_module_selected)
        panes.add(left, weight=1)

        right = ttk.Frame(panes)
        self.editor = CodeEditor(right)
        self.editor.pack(fill="both", expand=True)
        panes.add(right, weight=4)

    def _build_xml_tab(self, parent) -> None:
        panes = ttk.Panedwindow(parent, orient="horizontal")
        panes.pack(fill="both", expand=True)

        left = ttk.Frame(panes, padding=4)
        bar = ttk.Frame(left)
        bar.pack(fill="x")
        self.validate_xml_btn = ttk.Button(
            bar, text="Validate XML", command=self._validate_selected_xml, state="disabled"
        )
        self.validate_xml_btn.pack(side="left", padx=2)
        self.xml_tree = ttk.Treeview(left, show="tree", selectmode="browse")
        tree_sb = ttk.Scrollbar(left, orient="vertical", command=self.xml_tree.yview)
        self.xml_tree.configure(yscrollcommand=tree_sb.set)
        self.xml_tree.pack(side="left", fill="both", expand=True)
        tree_sb.pack(side="right", fill="y")
        self.xml_tree.bind("<<TreeviewSelect>>", self.on_xml_part_selected)
        panes.add(left, weight=1)

        right = ttk.Frame(panes)
        self.xml_editor = XmlEditor(right)
        self.xml_editor.pack(fill="both", expand=True)
        self.xml_editor.set_validate_action(self._validate_selected_xml)
        panes.add(right, weight=4)

        self.current_xml_part_path: str | None = None

    def _build_forms_tab(self, parent) -> None:
        """Read-only UserForm layouts, as stored in the file when it was opened."""
        panes = ttk.Panedwindow(parent, orient="horizontal")
        panes.pack(fill="both", expand=True)

        left = ttk.Frame(panes, padding=4)
        bar = ttk.Frame(left)
        bar.pack(fill="x")
        self.form_code_btn = ttk.Button(
            bar, text="Go to Code", command=self.goto_form_code, state="disabled"
        )
        self.form_code_btn.pack(side="left", padx=2)
        self.forms_tree = ttk.Treeview(left, show="tree", selectmode="browse")
        tree_sb = ttk.Scrollbar(left, orient="vertical", command=self.forms_tree.yview)
        self.forms_tree.configure(yscrollcommand=tree_sb.set)
        self.forms_tree.pack(side="left", fill="both", expand=True)
        tree_sb.pack(side="right", fill="y")
        self.forms_tree.bind("<<TreeviewSelect>>", self.on_form_item_selected)
        panes.add(left, weight=1)

        right = ttk.Frame(panes, padding=4)
        self.form_heading = ttk.Label(right, text="", anchor="w", wraplength=560, justify="left")
        self.form_heading.pack(fill="x")
        props = ttk.Frame(right)
        props.pack(fill="both", expand=True, pady=(4, 0))
        self.form_props = ttk.Treeview(props, columns=("value",), show="tree headings", selectmode="browse")
        self.form_props.heading("#0", text="Property")
        self.form_props.heading("value", text="Value")
        self.form_props.column("#0", width=200, stretch=False)
        props_sb = ttk.Scrollbar(props, orient="vertical", command=self.form_props.yview)
        self.form_props.configure(yscrollcommand=props_sb.set)
        self.form_props.pack(side="left", fill="both", expand=True)
        props_sb.pack(side="right", fill="y")
        ttk.Label(
            right,
            text=(
                "Read-only view. Only properties that differ from the control's "
                "default are stored in the file, so only those are listed."
            ),
            foreground="#555555",
            wraplength=560,
            justify="left",
        ).pack(fill="x", pady=(4, 0))
        panes.add(right, weight=3)

        self._form_items: dict[str, tuple[str, FormDesignSnapshot | FormControlSnapshot]] = {}

    def _refresh_forms(self) -> None:
        self.forms_tree.delete(*self.forms_tree.get_children())
        self.form_props.delete(*self.form_props.get_children())
        self._form_items.clear()
        self.form_code_btn.config(state="disabled")
        draft = self.draft
        if draft is None:
            self.form_heading.config(text="")
            return
        if not draft.baseline.forms:
            self.form_heading.config(text="This project has no UserForms.")
            return
        self.form_heading.config(text="Select a form or control to see its stored properties.")
        deleted = {
            (m.origin_name or m.current_name).casefold() for m in draft.deleted_original_modules()
        }
        for form in draft.baseline.forms:
            label = form.name
            if form.problem:
                label += " [unreadable]"
            elif form.name.casefold() in deleted:
                label += " [deleted — pending save]"
            iid = f"form::{form.name}"
            self.forms_tree.insert("", "end", iid=iid, text=label, open=True)
            self._form_items[iid] = (form.name, form)
            self._insert_form_controls(iid, form.name, form.controls)

    def _insert_form_controls(self, parent: str, form_name: str, controls) -> None:
        for control in controls:
            iid = f"{parent}/{control.name}"
            kind = control.kind.rsplit(".", 1)[-1]
            self.forms_tree.insert(parent, "end", iid=iid, text=f"{control.name}  ({kind})", open=True)
            self._form_items[iid] = (form_name, control)
            self._insert_form_controls(iid, form_name, control.children)

    def on_form_item_selected(self, _event=None) -> None:
        selection = self.forms_tree.selection()
        if not selection or selection[0] not in self._form_items:
            return
        form_name, item = self._form_items[selection[0]]
        self.form_props.delete(*self.form_props.get_children())
        if isinstance(item, FormDesignSnapshot):
            heading = f"UserForm {item.name} — {item.control_count()} control(s)"
            if item.problem:
                heading += f"\nThe layout could not be read: {item.problem}"
        else:
            heading = f"{item.name} ({item.kind}) on {form_name}"
        self.form_heading.config(text=heading)
        for name, value in item.properties:
            self.form_props.insert("", "end", text=name, values=(value,))
        self.form_code_btn.config(
            state="normal" if self._form_code_module(form_name) is not None else "disabled"
        )

    def _form_code_module(self, form_name: str) -> ModuleDraft | None:
        if self.draft is None:
            return None
        return self.draft.find_current(form_name)

    def goto_form_code(self) -> None:
        selection = self.forms_tree.selection()
        if not selection or selection[0] not in self._form_items:
            return
        mod = self._form_code_module(self._form_items[selection[0]][0])
        if mod is not None:
            self._open_module_at(mod.id, 1)

    def _build_statusbar(self) -> None:
        self.status = ttk.Label(self.root, text="Ready", padding=3, anchor="w")
        self.status.pack(fill="x", side="bottom")

    def _bind_shortcuts(self) -> None:
        self.root.bind("<Control-o>", lambda _e: self.open_file())
        self.root.bind("<Control-s>", lambda _e: self.save())
        self.root.bind("<Control-S>", lambda _e: self.save_copy())
        self.root.bind("<Control-f>", lambda _e: self._active_text_editor().find_dialog())
        self.root.bind("<Control-h>", lambda _e: self._active_text_editor().replace_dialog())
        self.root.bind("<Control-Shift-F>", lambda _e: self.project_search())
        self.root.bind("<Control-z>", lambda _e: self.undo())
        self.root.bind("<Control-y>", lambda _e: self.redo())
        self.root.bind("<F12>", lambda _e: self.goto_callback_under_cursor())

    # -- banners / status ----------------------------------------------------

    def _set_banner(self, text: str) -> None:
        if text:
            self.banner.config(text=text)
            self.banner.pack(fill="x", before=self.status)
        else:
            self.banner.pack_forget()

    def _on_progress(self, stage: str) -> None:
        # Called from the save worker thread; Tk is only touched by the poller.
        self._progress_stage = stage

    def _run_in_background(self, work: Callable[[], _T]) -> _T:
        """Run ``work()`` on a worker thread; keep the UI painted but inert until done.

        Blocks the caller in a nested event loop so callers keep synchronous
        semantics. Saves are dominated by antivirus scans of freshly written
        packages; running them off the Tk thread keeps the window responsive.
        """
        outcome: dict = {}

        def target() -> None:
            try:
                outcome["result"] = work()
            except BaseException as exc:  # noqa: BLE001 - re-raised on the UI thread
                outcome["error"] = exc

        # Non-daemon: interpreter exit must never cut a commit in half.
        thread = threading.Thread(target=target, name="vbaae-save")
        # A bare Tcl variable name, not tk.BooleanVar: a Variable finalized by
        # the GC on the worker thread would call into Tk off the main thread.
        done = f"vbaae_bg_done_{id(thread)}"
        self.root.setvar(done, "0")
        self._progress_stage = None
        held = self._hold_ui()

        def poll() -> None:
            stage = self._progress_stage
            if stage:
                self.status.config(text=f"Saving… {stage}")
            if thread.is_alive():
                self.root.after(50, poll)
            else:
                self.root.setvar(done, "1")

        try:
            thread.start()
            self.root.after(50, poll)
            self.root.wait_variable(done)
        finally:
            self._release_ui(held)
            try:
                self.root.globalunsetvar(done)
            except tk.TclError:
                pass
        if "error" in outcome:
            raise outcome["error"]
        return outcome["result"]

    def _hold_ui(self) -> list[str]:
        """Block mouse, keyboard shortcuts and menus while a save runs."""
        held: list[str] = []
        windows = [self.root] + [w for w in self.root.winfo_children() if isinstance(w, tk.Toplevel)]
        for window in windows:
            name = str(window)
            try:
                self.root.tk.call("tk", "busy", "hold", name)
                busy = "._Busy" if name == "." else name + "_Busy"
                # Keep only the busy window's own tag so root shortcuts cannot fire.
                self.root.tk.call("bindtags", busy, (busy,))
                held.append(name)
            except tk.TclError:
                pass
        if held:
            try:
                self.root.tk.call("focus", "._Busy")
            except tk.TclError:
                pass
        self._set_menubar_state("disabled")
        self.root.update_idletasks()
        return held

    def _release_ui(self, held: list[str]) -> None:
        for name in held:
            try:
                self.root.tk.call("tk", "busy", "forget", name)
            except tk.TclError:
                pass
        self._set_menubar_state("normal")

    def _set_menubar_state(self, state: str) -> None:
        try:
            last = self.menubar.index("end")
            for i in range(0 if last is None else last + 1):
                self.menubar.entryconfig(i, state=state)
        except tk.TclError:
            pass

    def _refresh_state(self) -> None:
        draft = self.draft
        if draft is None:
            self.file_label.config(text="No add-in open")
            self.save_btn.config(state="disabled")
            self.review_btn.config(state="disabled")
            self.revert_btn.config(state="disabled")
            self.status.config(text="Open an add-in to begin.")
            return
        name = draft.baseline.path.name
        n = dirty_count(draft)
        self.file_label.config(text=f"{name} — unsaved changes: {n}")
        inplace_ok = (
            not self._saving
            and (self.session is None or (self.session.original_exists and not self.session.conflicts_pending))
        )
        self.save_btn.config(state="normal" if inplace_ok else "disabled")
        self.review_btn.config(state="normal" if n else "disabled")
        self.revert_btn.config(state="normal" if n else "disabled")
        if self.session is not None and not self.session.original_exists:
            self._set_banner(
                "The original add-in is missing. In-place Save is disabled. "
                "Use Save a Copy to write a recovered draft to a new path."
            )
        elif self.session is not None and self.session.conflicts_pending:
            self._set_banner(
                "The original file changed after this draft was captured. "
                "Resolve external changes before saving in place."
            )
        elif self._recovery_warning:
            self._set_banner(self._recovery_warning)
        elif draft.baseline.safety.password_protected:
            self._set_banner(
                "This VBA project is password-protected. VBA editing is disabled; "
                "XML package editing remains available for .xlam/.ppam/.pptm files."
            )
        elif self._live_host_probe_failed:
            self._set_banner(
                "Windows could not list running programs. Close Excel and PowerPoint "
                "before saving; the editor will check again on the next save."
            )
        elif self._live_host_running:
            self._set_banner(
                f"Close {self._host_label} before saving this add-in."
            )
        else:
            self._set_banner("")

    def _refresh_tree(self) -> None:
        draft = self.draft
        self.tree.delete(*self.tree.get_children())
        if draft is None:
            return
        self._refresh_forms()
        groups = {
            "Standard Modules": [],
            "Class Modules": [],
            "UserForms": [],
            "Object / Class / Form Code": [],
        }
        for m in draft.modules:
            if m.is_deleted:
                continue
            marker = ""
            if m.is_new:
                marker = " +"
            elif m.origin_name and m.current_name.casefold() != m.origin_name.casefold():
                marker = " →"
            elif m.original_body is not None and m.body != m.original_body:
                marker = " *"
            if m.kind == ModuleDisplayKind.CLASS:
                group = "Class Modules"
            elif m.pyopenvba_kind == "standard":
                group = "Standard Modules"
            elif m.project_item_kind == "userform":
                group = "UserForms"
            else:
                group = "Object / Class / Form Code"
            locked = not (m.can_delete or m.can_rename or m.is_new)
            label = f"{m.current_name}{marker}" + ("" if not locked else " [lock]")
            groups[group].append((m.id, label))
        for group, items in groups.items():
            if group == "UserForms" and not items:
                continue
            parent = self.tree.insert("", "end", text=group, open=True)
            for mid, label in items:
                self.tree.insert(parent, "end", iid=mid, text=label)

    # -- editor plumbing ------------------------------------------------------

    def _flush_active_vba_editor(self) -> None:
        draft = self.draft
        if draft is None or self.current_module_id is None:
            return
        mod = draft.module_by_id(self.current_module_id)
        if mod is not None:
            new_body = self.editor.get_text()
            if self.session is not None and new_body != mod.body:
                self.history_service.record_text(
                    self.session,
                    target_id=mod.id,
                    before=mod.body,
                    after=new_body,
                    now_ms=int(time.time() * 1000),
                    coalesce=True,
                )
            mod.body = new_body

    def _flush_active_xml_editor(self) -> None:
        draft = self.draft
        if draft is None or self.current_xml_part_path is None:
            return
        part = draft.xml_part_by_path(self.current_xml_part_path)
        if part is not None and part.editable:
            new_text = self.xml_editor.get_text()
            if self.session is not None and new_text != (part.text or ""):
                self.history_service.record_text(
                    self.session,
                    target_id=part.path,
                    before=part.text or "",
                    after=new_text,
                    op="xml_text",
                    now_ms=int(time.time() * 1000),
                    coalesce=True,
                )
            part.text = new_text

    def _flush_all_editors(self) -> None:
        self._flush_active_vba_editor()
        self._flush_active_xml_editor()

    def on_module_selected(self, _event=None) -> None:
        draft = self.draft
        if draft is None:
            return
        selection = self.tree.selection()
        if not selection:
            return
        mid = selection[0]
        if mid == self.current_module_id:
            return  # already shown (programmatic re-selection)
        self._flush_all_editors()
        self.current_module_id = mid
        mod = draft.module_by_id(mid)
        if mod is None:
            return
        self.editor.set_text(mod.body)
        self.editor.text.config(state="normal")  # body editing allowed; deletion is gated
        self._refresh_state()

    def on_xml_part_selected(self, _event=None) -> None:
        draft = self.draft
        if draft is None:
            return
        selection = self.xml_tree.selection()
        if not selection:
            return
        path = selection[0][len("xml::"):]
        if path == self.current_xml_part_path:
            return  # already shown (programmatic re-selection)
        self._flush_all_editors()
        part = draft.xml_part_by_path(path)
        if part is None:
            return
        self.current_xml_part_path = path
        # Re-enable before replacing widget contents in case the previously
        # selected baseline part was read-only.
        self.xml_editor.text.config(state="normal")
        self.xml_editor.set_text(part.text or "")
        if not part.editable:
            self.xml_editor.text.config(state="disabled")
        self.xml_editor.set_part_info(part)
        self.validate_xml_btn.config(state="normal")
        self._refresh_state()

    def _refresh_xml_tree(self) -> None:
        draft = self.draft
        self.xml_tree.delete(*self.xml_tree.get_children())
        if draft is None:
            return
        for part in sorted(draft.xml_parts, key=lambda p: p.path.lower()):
            marker = " *" if part.is_dirty() else ""
            iid = "xml::" + part.path
            warning = "⚠ " if not part.editable else ""
            self.xml_tree.insert("", "end", iid=iid, text=warning + part.path + marker)

    def _active_text_editor(self):
        if str(self.editor_notebook.select()) == str(self.xml_tab):
            return self.xml_editor
        return self.editor

    def _validate_selected_xml(self) -> None:
        draft = self.draft
        if draft is None or self.current_xml_part_path is None:
            return
        self._flush_active_xml_editor()
        part = draft.xml_part_by_path(self.current_xml_part_path)
        if part is None:
            return
        if not part.editable:
            messagebox.showerror(APP_NAME, part.open_problem or "This XML part is read-only.")
            return
        problems = self.doc_service.package_adapter.validate_draft_part(part)
        if problems:
            messagebox.showerror(APP_NAME, "\n".join(problems))
            return
        issues = (
            [i for i in rs.check_ribbon(draft) if i.part_path == part.path]
            if rs.is_ribbon_part(part.path)
            else []
        )
        if issues:
            messagebox.showwarning(
                APP_NAME,
                "XML is well-formed, but some ribbon callbacks have problems:\n\n"
                + "\n".join(f"Line {i.line}: {i.message}" for i in issues[:15])
                + ("\n…" if len(issues) > 15 else "")
                + "\n\nUse Ribbon → Ribbon Buttons… to fix them.",
            )
        else:
            messagebox.showinfo(APP_NAME, "XML is well-formed.")

    # -- file ops --------------------------------------------------------------

    def open_file(self) -> None:
        path = filedialog.askopenfilename(title="Open Office VBA File", filetypes=_FILETYPES)
        if not path:
            return
        self.load_path(Path(path))

    def load_path(self, path: Path) -> None:
        try:
            self._flush_all_editors()
            if self.session is not None:
                self.session_service.close(self.session)
            try:
                self.session = self.session_service.open(path)
                draft = self.session.draft
            except (AdapterError, OSError, ValueError):
                self.session = None
                draft = self.doc_service.open(path)
        except AdapterError as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return
        self.draft = draft
        self._refresh_host_status()
        self.current_module_id = None
        self.current_xml_part_path = None
        self._refresh_tree()
        self._refresh_xml_tree()
        xml_supported = path.suffix.lower() in XML_EDITABLE_EXTENSIONS
        self.editor_notebook.tab(self.xml_tab, state="normal" if xml_supported else "disabled")
        if xml_supported:
            self.xml_editor.set_part_info(None)
            self.xml_editor.text.config(state="normal")
            self.xml_editor.set_text("")
            self.validate_xml_btn.config(state="disabled")
        first = next((m.id for m in draft.modules if not m.is_deleted), None)
        if first:
            self.tree.selection_set(first)
        self._refresh_state()
        self.status.config(text=f"Opened {path.name}")

    def save(self) -> None:
        if self.draft is None or self._saving:
            return
        if self.session is not None and not self.session.original_exists:
            messagebox.showwarning(
                APP_NAME,
                "In-place Save is disabled because the original add-in is missing. "
                "Use Save a Copy to write the recovered draft to a different path.",
            )
            return
        if self.session is not None and self.session.conflicts_pending:
            messagebox.showwarning(
                APP_NAME,
                "The original file has changed. Compare and resolve external changes before saving in place.",
            )
            self.compare_external_changes()
            return
        self._flush_all_editors()
        self._checkpoint_now()
        changes = compute_changes(self.draft)
        if changes.is_empty:
            messagebox.showinfo(APP_NAME, "No changes to save.")
            return
        if not self._confirm_ribbon_issues():
            return
        if not self._review_and_confirm(operation_type="save"):
            return
        self._saving = True
        self._refresh_state()
        result = None
        published = None
        try:
            result, published = self._run_in_background(self._save_work)
        except Exception as exc:  # noqa: BLE001 - UI boundary
            from vba_addin_editor.domain.results import SaveResult

            result = SaveResult.error(
                f"Save failed unexpectedly ({type(exc).__name__}). The original file was not overwritten.",
                reason="unknown",
                details={"exception_type": type(exc).__name__},
            )
        finally:
            self._saving = False
            self._refresh_host_status()
            self._refresh_state()
        self._handle_save_result(result, published=published)

    def _save_work(self):
        """Worker-thread half of save: commit, then publish the session baseline."""
        draft = self.draft
        assert draft is not None
        result = self.save_service.save_addin(draft)
        published = None
        if result.kind == "success" and self.session is not None:
            saved = Path(str(result.details.get("saved") or draft.baseline.path))
            published = self.session_service.publish_verified_baseline(self.session, saved)
        return result, published

    def _handle_save_result(self, result, *, published=None) -> None:
        self._last_result = result
        kind = result.kind
        if kind == "success":
            warning = self.backup_service.record_success(
                original=self.draft.baseline.path if self.draft else Path("."),
                backup=result.backup_path,
                operation_type=result.operation_type,
            )
            if self.session is not None:
                saved = Path(str(result.details.get("saved") or self.draft.baseline.path))
                if published is not None:
                    self._last_result = published
                    messagebox.showerror(
                        APP_NAME,
                        (published.message or "Session baseline publication failed.")
                        + f"\n\nCommitted file: {saved}"
                        + f"\nCaptured baseline: {self.session.captured_path}"
                        + f"\nSession folder: {self.session.session_dir}",
                    )
                    self._refresh_state()
                    return
                self.recovery_service.mark_complete(self.session)
            self._refresh_tree()
            self._refresh_xml_tree()
            if self.current_module_id and self.draft:
                mod = self.draft.module_by_id(self.current_module_id)
                if mod is not None and not mod.is_deleted:
                    self.editor.set_text(mod.body)
            if self.current_xml_part_path and self.draft:
                part = self.draft.xml_part_by_path(self.current_xml_part_path)
                if part is not None:
                    self.xml_editor.set_text(part.text)
                    self.xml_editor.set_part_info(part)
            self._refresh_state()
            msg = "Saved successfully."
            if result.backup_path:
                msg += f"\n\nBackup: {result.backup_path}"
            if warning:
                msg += f"\n\n{warning}"
            messagebox.showinfo(APP_NAME, msg)
        elif kind == "no_changes":
            messagebox.showinfo(APP_NAME, result.message or "No changes to save.")
        elif kind == "blocked":
            messagebox.showwarning(APP_NAME, dialog_text(result))
        elif kind == "needs_signature_confirmation":
            self._signature_dialog()
        elif kind == "candidate_failed":
            details = "\n".join(result.problems)
            messagebox.showerror(APP_NAME, result.message + "\n\n" + details)
        elif kind == "recovery_required":
            messagebox.showerror(
                APP_NAME,
                (result.message or "Final verification failed after replacing the add-in.")
                + f"\n\nBackup: {result.backup_path}"
                + f"\nCandidate: {result.candidate_path}"
                + "\n\nUse File → Restore Backup to recover. The editor will not restore automatically.",
            )
        elif kind == "error":
            messagebox.showerror(APP_NAME, result.message or "Save failed.")
        self._refresh_state()

    def _review_and_confirm(self, *, operation_type: str, destination: str | None = None) -> bool:
        if self.draft is None:
            return False
        if self.session is None:
            return self._preflight_dialog(compute_changes(self.draft))
        model = self.review_service.build(
            self.session, operation_type=operation_type, destination=destination
        )
        lines = [
            f"Review changes to {model.destination}",
            f"Type: {model.file_type}    VBA: {model.vba_count}    XML: {model.xml_count}",
            model.backup_policy,
            "",
        ]
        for item in model.items:
            ops = "+".join(item.operations)
            label = item.new_name or item.old_name or item.target_id
            lines.append(f"{item.category} [{ops}] {label}")
            if item.unified_diff:
                snippet = "\n".join(item.unified_diff.splitlines()[:40])
                lines.append(snippet)
                lines.append("")
        if model.signature_will_be_removed:
            lines.append("This save will remove the VBA digital signature.")
            lines.append("Confirm signature removal by choosing Save.")
        if not messagebox.askokcancel(APP_NAME, "\n".join(lines)[:4000]):
            return False
        if not self.session.matches_proposal(model.binding):
            messagebox.showwarning(APP_NAME, "This review is out of date because the draft changed.")
            return False
        if model.signature_will_be_removed:
            self.draft.signed_save_confirmed = True
        return True

    def _preflight_dialog(self, changes) -> bool:
        draft = self.draft
        has_vba = changes.has_vba_changes
        lines = [f"Review changes to {draft.baseline.path.name}", ""]
        lines.extend(changes.summary_lines())
        lines += ["", "Safety"]
        lines.append("  ✓ File has not changed on disk")
        if draft.baseline.safety.password_protected:
            lines.append("  ! VBA project is password-protected")
        elif has_vba:
            lines.append("  ✓ VBA project is not password-protected")
        if draft.baseline.package_safety.opc_signature_present:
            lines.append("  ! OPC package signature detected")
        else:
            lines.append("  ✓ No OPC package signature detected")
        if has_vba and draft.baseline.safety.signature_present:
            lines.append("  ! VBA digital signature will be removed")
        lines += ["", "A backup will be created before the original is replaced."]
        confirm_label = (
            "Save and Remove Signature"
            if has_vba and draft.baseline.safety.signature_present
            else "Save File"
        )
        if not messagebox.askokcancel(APP_NAME, "\n".join(lines)):
            return False
        if has_vba and draft.baseline.safety.signature_present:
            draft.signed_save_confirmed = True
        self.status.config(text=f"Saving with: {confirm_label}")
        return True

    def _signature_dialog(self) -> None:
        text = (
            "This add-in's VBA project is digitally signed.\n\n"
            "Any code change makes the existing signature invalid; saving will "
            "remove it. Office or your organization's security policy may then "
            "block the add-in.\n\nA backup of the currently signed file will be "
            "created automatically.\n\nRemove the signature and save?"
        )
        if messagebox.askyesno(APP_NAME, text, icon="warning"):
            self.draft.signed_save_confirmed = True
            self.save()

    def save_copy(self) -> None:
        if self.draft is None:
            return
        self._flush_all_editors()
        self._checkpoint_now()
        dest_name = filedialog.asksaveasfilename(
            title="Save a Copy",
            defaultextension=self.draft.baseline.extension,
            filetypes=_FILETYPES,
        )
        if not dest_name:
            return
        if not self._confirm_ribbon_issues():
            return
        dest = Path(dest_name)
        if dest.exists() and not messagebox.askyesno(
            APP_NAME,
            f"{dest.name} already exists. Overwrite it? A backup of the "
            "destination will be created.",
        ):
            return
        if not self._review_and_confirm(operation_type="save_copy", destination=str(dest)):
            return
        dest_hash = None
        if dest.exists():
            from vba_addin_editor.platform import paths as pathmod

            dest_hash = pathmod.fingerprint(dest).sha256
        recovered = bool(self.session is not None and (
            not self.session.original_exists or self.session.conflicts_pending
        ))
        if recovered and not messagebox.askokcancel(
            APP_NAME,
            "Save recovered draft as a separate copy?\n\n"
            "Newer or missing original bytes will not be included. "
            "The destination must differ from the original path.",
        ):
            return
        draft = self.draft
        session = self.session
        allow_overwrite = dest.exists()
        self._saving = True
        self._refresh_state()
        try:
            result = self._run_in_background(
                lambda: self.save_service.save_copy(
                    draft,
                    dest,
                    source_path=session.captured_path if session is not None else None,
                    allow_overwrite=allow_overwrite,
                    recovered_copy=recovered,
                    reviewed_dest_hash=dest_hash,
                    session_dir=session.session_dir if session is not None else None,
                )
            )
        finally:
            self._saving = False
            self._refresh_state()
        result.operation_type = "save_copy"
        if result.kind == "success":
            self.backup_service.record_success(
                original=dest, backup=result.backup_path, operation_type="save_copy"
            )
            messagebox.showinfo(
                APP_NAME,
                "Copy saved. Excel/PowerPoint will continue using the original "
                "installed add-in unless you change its add-in configuration.\n\n"
                f"Destination: {dest}",
            )
        else:
            self._handle_save_result(result)

    def revert_all(self) -> None:
        if self.draft is None:
            return
        from vba_addin_editor.domain.document import revert_all

        self._flush_all_editors()
        self.draft = revert_all(self.draft)
        self.current_module_id = None
        self.current_xml_part_path = None
        self._refresh_tree()
        self._refresh_xml_tree()
        self.editor.set_text("")
        self.xml_editor.set_text("")
        self.xml_editor.set_part_info(None)
        self.validate_xml_btn.config(state="disabled")
        first = next((m.id for m in self.draft.modules if not m.is_deleted), None)
        if first:
            self.tree.selection_set(first)
        self._refresh_state()

    # -- module ops --------------------------------------------------------------

    def _selected_module(self):
        if self.draft is None:
            return None
        sel = self.tree.selection()
        return self.draft.module_by_id(sel[0]) if sel else None

    def add_module(self, *, standard: bool) -> None:
        draft = self.draft
        if draft is None:
            return
        if draft.baseline.safety.password_protected:
            messagebox.showwarning(APP_NAME, "This project is password-protected and read-only.")
            return
        kind_label = "Standard" if standard else "Class"
        name = self._prompt_name(f"New {kind_label} Module", "")
        if name is None:
            return
        problem = validate_module_name(name, draft)
        if problem:
            messagebox.showerror(APP_NAME, problem)
            return
        mod = ModuleDraft(
            id=new_module_id(),
            origin_name=None,
            current_name=name,
            body="",
            kind=ModuleDisplayKind.CLASS if not standard else ModuleDisplayKind.STANDARD,
            pyopenvba_kind="other" if not standard else "standard",
            is_new=True,
            is_deleted=False,
            destructive_ops_safe=True,
            can_delete=True,
            can_rename=True,
            project_item_kind="standard" if standard else "class",
        )
        draft.modules.append(mod)
        if self.session is not None:
            self.history_service.record_structural(
                self.session,
                HistoryCommand(
                    op="add",
                    target_id=mod.id,
                    before=None,
                    after=snapshot_modules(draft)[-1],
                ),
            )
            self._checkpoint_now()
        self._refresh_tree()
        self.tree.selection_set(mod.id)

    def rename_module(self) -> None:
        mod = self._selected_module()
        draft = self.draft
        if mod is None or draft is None:
            return
        if mod.is_new:
            name = self._prompt_name("Rename Module", mod.current_name)
            if name:
                old = mod.current_name
                mod.current_name = name
                self._record_rename(mod.id, old, name)
        else:
            if not mod.can_rename:
                messagebox.showwarning(
                    APP_NAME,
                    RESTRICTION_MESSAGES.get(
                        mod.restriction_reason or "",
                        "Renaming this module is disabled because its type cannot be verified safely.",
                    ),
                )
                return
            name = self._prompt_name("Rename Module", mod.current_name)
            if name:
                old = mod.current_name
                mod.current_name = name
                self._record_rename(mod.id, old, name)
        self._flush_all_editors()
        self._refresh_tree()
        self._refresh_state()

    def delete_module(self) -> None:
        mod = self._selected_module()
        if mod is None or self.draft is None:
            return
        if not (mod.can_delete or mod.is_new):
            messagebox.showwarning(
                APP_NAME,
                RESTRICTION_MESSAGES.get(
                    mod.restriction_reason or "",
                    "Deleting this module is disabled because its type cannot be verified safely.",
                ),
            )
            return
        if mod.project_item_kind == "userform" and not mod.is_new:
            prompt = (
                f'Delete UserForm "{mod.current_name}"?\n\nThis removes the form\'s code '
                "AND its layout (all of its controls) from the VBA project when you save. "
                "The layout cannot be exported or re-imported by this editor; the "
                "pre-save backup is the only way to get it back.\n\n"
                "It is not written to the file until Save File."
            )
        else:
            prompt = (
                f'Delete module "{mod.current_name}"?\n\nThis removes the module from the '
                "VBA project when you save. It is not written to the file until Save File."
            )
        if not messagebox.askyesno(APP_NAME, prompt):
            return
        if mod.is_new:
            self.draft.modules.remove(mod)
        else:
            mod.is_deleted = True
        if self.session is not None:
            self.history_service.record_structural(
                self.session,
                HistoryCommand(op="delete", target_id=mod.id, before=False, after=True),
            )
            self._checkpoint_now()
        self.editor.set_text("")
        self.current_module_id = None
        self._refresh_tree()
        self._refresh_state()

    def _prompt_name(self, title: str, initial: str) -> str | None:
        win = tk.Toplevel(self.root)
        win.title(title)
        win.transient(self.root)
        win.grab_set()
        var = tk.StringVar(value=initial)
        ttk.Label(win, text="Name:").grid(row=0, column=0, padx=8, pady=8)
        entry = ttk.Entry(win, textvariable=var, width=32)
        entry.grid(row=0, column=1, padx=8, pady=8)
        result: list[str | None] = [None]

        def confirm() -> None:
            result[0] = var.get().strip()
            win.destroy()

        ttk.Button(win, text="Create" if initial == "" else "Rename", command=confirm).grid(row=1, column=1, sticky="e", padx=8, pady=8)
        win.bind("<Return>", lambda _e: confirm())
        entry.focus_set()
        self.root.wait_window(win)
        return result[0]

    # -- import / export ------------------------------------------------------------

    def import_file(self) -> None:
        if self.draft is None:
            return
        path = filedialog.askopenfilename(
            title="Import Module Source",
            filetypes=[("VBA source", "*.bas;*.cls")],
        )
        if not path:
            return
        try:
            suggested, text, is_class = self.ie_service.read_import(Path(path))
            name = self._prompt_name("Import As", suggested)
            if name is None:
                return
            existing = self.draft.find_current(name)
            if existing is not None:
                choice = messagebox.askyesnocancel(
                    APP_NAME,
                    f"A module named '{name}' exists. Replace its body with the "
                    "imported file?\n\nYes = replace body, No = import under a new name, "
                    "Cancel = abort.",
                )
                if choice is None:
                    return
                if choice:
                    existing.body = text
                    self._refresh_tree()
                    return
                name = self._prompt_name("Import As", suggested + "1")
                if name is None:
                    return
            mod = ModuleDraft(
                id=new_module_id(),
                origin_name=None,
                current_name=name,
                body=text,
                kind=ModuleDisplayKind.CLASS if is_class else ModuleDisplayKind.STANDARD,
                pyopenvba_kind="other" if is_class else "standard",
                is_new=True,
                is_deleted=False,
                destructive_ops_safe=True,
                can_delete=True,
                can_rename=True,
                project_item_kind="class" if is_class else "standard",
            )
            self.draft.modules.append(mod)
            self._refresh_tree()
            self.tree.selection_set(mod.id)
        except (AdapterError, UnicodeDecodeError) as exc:
            messagebox.showerror(APP_NAME, str(exc))

    def export_module(self) -> None:
        mod = self._selected_module()
        if mod is None or self.draft is None:
            return
        ext = ".bas" if mod.pyopenvba_kind == "standard" else ".cls"
        path = filedialog.asksaveasfilename(
            title="Export Module", defaultextension=ext, initialfile=mod.current_name + ext
        )
        if not path:
            return
        try:
            self.ie_service.export_module(self.draft, mod.id, Path(path))
            self.status.config(text=f"Exported to {path}")
        except (KeyError, AdapterError) as exc:
            messagebox.showerror(APP_NAME, str(exc))

    def export_all(self) -> None:
        if self.draft is None:
            return
        folder = filedialog.askdirectory(title="Export all modules to folder")
        if not folder:
            return
        try:
            written = self.ie_service.export_all(self.draft, Path(folder))
            self.status.config(text=f"Exported {len(written)} modules to {folder}")
        except (KeyError, AdapterError) as exc:
            messagebox.showerror(APP_NAME, str(exc))

    def restore_backup(self) -> None:
        if self.draft is None:
            messagebox.showinfo(APP_NAME, "Open the add-in first, then restore a backup.")
            return
        if self.draft.is_dirty():
            choice = messagebox.askyesnocancel(
                APP_NAME, "You have unsaved changes. Save before restoring a backup?"
            )
            if choice is None:
                return
            if choice:
                self.save()
                if self.draft is not None and self.draft.is_dirty():
                    return
            else:
                self._checkpoint_now()
                self.draft = revert_all(self.draft)
                if self.session is not None:
                    self.session.draft = self.draft
                self._reload_editors_from_draft()
        path = filedialog.askopenfilename(
            title="Select Backup", filetypes=_FILETYPES
        )
        if not path:
            return
        backup = Path(path)
        inspect = self.backup_service.inspect(backup, current_path=self.draft.baseline.path)
        if inspect.problems:
            messagebox.showerror(APP_NAME, "\n".join(inspect.problems))
            return
        summary = "\n".join(inspect.summary_lines) or backup.name
        if not messagebox.askokcancel(APP_NAME, "Restore this backup?\n\n" + summary[:4000]):
            return
        result = self.backup_service.restore_transaction(
            self.draft.baseline.path,
            backup,
            expected_backup_hash=inspect.sha256,
            expected_target_hash=self.draft.baseline.file_fingerprint.sha256,
        )
        if result.kind != "success":
            messagebox.showerror(
                APP_NAME,
                (result.message or "Restore failed.")
                + "\n\n"
                + "\n".join(f"{k}: {v}" for k, v in result.details.items() if k.endswith("backup") or k in {"target", "candidate", "safety_backup", "selected_backup"})
            )
            return
        warning = result.details.get("catalog_warning")
        try:
            refreshed = self.doc_service.open(self.draft.baseline.path)
        except AdapterError as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return
        self.draft = refreshed
        if self.session is not None:
            self.session.draft = refreshed
            published = self.session_service.publish_verified_baseline(
                self.session, self.draft.baseline.path
            )
            if published is not None:
                messagebox.showerror(APP_NAME, published.message or "Session baseline publication failed.")
            else:
                self.recovery_service.checkpoint(self.session)
        self._reload_editors_from_draft()
        msg = f"Backup restored.\n\nPre-restore safety copy: {result.backup_path}"
        if warning:
            msg += f"\n\n{warning}"
        messagebox.showinfo(APP_NAME, msg)

    def open_location(self) -> None:
        if self.draft is None:
            return
        import subprocess

        subprocess.Popen(["explorer", "/select,", str(self.draft.baseline.path)])

    # -- misc ------------------------------------------------------------------------

    def review_changes(self) -> None:
        if self.draft is None:
            return
        changes = compute_changes(self.draft)
        lines = [f"Review changes to {self.draft.baseline.path.name}", ""]
        lines.extend(changes.summary_lines() or ["No changes."])
        messagebox.showinfo(APP_NAME, "\n".join(lines))

    def show_about(self) -> None:
        identity = build_identity()
        messagebox.showinfo(
            APP_NAME,
            f"{APP_NAME} {identity.get('version', VERSION)}\n"
            f"Commit: {identity.get('source_commit', 'unbuilt')}\n"
            f"Mode: {identity.get('packaged_mode', 'development')}\n"
            f"pyOpenVBA: {identity.get('pyopenvba_installed') or identity.get('pyopenvba_pin')}\n\n"
            "Edits VBA source inside installed .xlam / .ppam add-ins and .pptm\n"
            "presentations, in place, with automatic backup and verification.\n\n"
            "Ordinary class modules can be deleted and renamed when PROJECT/dir\n"
            "metadata agrees. Close the corresponding Office host before saving.\n\n"
            "Uses pyOpenVBA (MIT) — see Help → Third-party notices.",
        )

    def _refresh_host_status(self) -> None:
        if self.draft is None:
            self._live_host_running = False
            self._live_host_probe_failed = False
            return
        probe = wp.probe_host_process(self.draft.baseline.path)
        self._live_host_running = probe.corresponding_host_running
        self._live_host_probe_failed = probe.enumeration_failed
        self._host_label = probe.label or wp.corresponding_host_label(self.draft.baseline.path)

    def _on_focus_in(self, _event=None) -> None:
        if self.draft is None or self._saving:
            return
        self._refresh_host_status()
        self._refresh_state()

    def _schedule_host_poll(self) -> None:
        if os.environ.get("PYTEST_CURRENT_TEST"):
            return
        if self._poll_after is not None:
            try:
                self.root.after_cancel(self._poll_after)
            except tk.TclError:
                pass
        self._poll_after = self.root.after(5000, self._poll_host)

    def _poll_host(self) -> None:
        if self.draft is not None and not self._saving:
            self._refresh_host_status()
            self._refresh_state()
        self._schedule_host_poll()

    def _on_vba_edited(self) -> None:
        self._note_text_edit(kind="vba")

    def _on_xml_edited(self) -> None:
        self._note_text_edit(kind="xml")

    def _note_text_edit(self, *, kind: str) -> None:
        if self.session is None or self.draft is None:
            return
        if self._search_dialog is not None:
            self._search_dialog.mark_stale()
        self._schedule_autosave()

    def _schedule_autosave(self) -> None:
        if self._idle_after is not None:
            try:
                self.root.after_cancel(self._idle_after)
            except tk.TclError:
                pass
        self._idle_after = self.root.after(IDLE_MS, self._checkpoint_now)
        if self._max_after is None:
            self._max_after = self.root.after(MAX_INTERVAL_MS, self._checkpoint_now)

    def _checkpoint_now(self, _event=None) -> None:
        self._idle_after = None
        if self._max_after is not None:
            try:
                self.root.after_cancel(self._max_after)
            except tk.TclError:
                pass
            self._max_after = None
        if self.session is None:
            return
        if self._saving:
            # The save worker owns the draft; retry once it has finished.
            self._schedule_autosave()
            return
        self._flush_all_editors()
        failed = self.recovery_service.checkpoint(self.session)
        if failed is not None:
            self._recovery_warning = (
                "Draft recovery could not be written. Editing stays enabled. "
                "Use File → Save or Save Recovery Now."
            )
            self._refresh_state()
        else:
            if self._recovery_warning.startswith("Draft recovery"):
                self._recovery_warning = ""
            self.status.config(text=f"Draft recovery saved {time.strftime('%H:%M:%S')}")

    def _record_rename(self, module_id: str, old: str, new: str) -> None:
        if self.session is None or old == new:
            return
        self.history_service.record_structural(
            self.session,
            HistoryCommand(op="rename", target_id=module_id, before=old, after=new),
        )
        self._checkpoint_now()

    def undo(self) -> None:
        if self.session is None:
            return
        self._flush_all_editors()
        if self.history_service.undo(self.session) is None:
            return
        self._reload_editors_from_draft()

    def redo(self) -> None:
        if self.session is None:
            return
        self._flush_all_editors()
        if self.history_service.redo(self.session) is None:
            return
        self._reload_editors_from_draft()

    def _reload_editors_from_draft(self) -> None:
        if self.draft is None:
            return
        self._refresh_tree()
        self._refresh_xml_tree()
        if self.current_module_id:
            mod = self.draft.module_by_id(self.current_module_id)
            if mod is not None and not mod.is_deleted:
                self.editor.set_text(mod.body)
        if self.current_xml_part_path:
            part = self.draft.xml_part_by_path(self.current_xml_part_path)
            if part is not None:
                self.xml_editor.set_text(part.text or "")
        self._refresh_state()

    def project_search(self) -> None:
        if self.draft is None:
            return
        self._flush_all_editors()
        if self._search_dialog is not None:
            try:
                if self._search_dialog.win.winfo_exists():
                    self._search_dialog.win.lift()
                    return
            except tk.TclError:
                self._search_dialog = None
        dialog = ProjectSearchDialog(self.root, self)
        self._search_dialog = dialog
        self.root.update_idletasks()

    def goto_search_hit(self, hit) -> None:
        if hit.kind == "vba":
            self._open_module_at(hit.target_id, hit.line, hit.column, hit.length)
        else:
            self._open_xml_at(hit.target_id, hit.line, hit.column, hit.length)

    def _open_module_at(self, module_id: str, line: int, column: int = 1, length: int = 0) -> None:
        if self.draft is None:
            return
        mod = self.draft.module_by_id(module_id)
        if mod is None or mod.is_deleted:
            return
        self._flush_all_editors()
        self.editor_notebook.select(0)
        if self.current_module_id != module_id:
            self.current_module_id = module_id
            self.editor.text.config(state="normal")
            self.editor.set_text(mod.body)
        try:
            self.tree.selection_set(module_id)
            self.tree.see(module_id)
        except tk.TclError:
            pass
        self.editor.goto_position(line, column, length)
        self.editor.text.focus_set()
        self._refresh_state()

    def _open_xml_at(self, part_path: str, line: int, column: int = 1, length: int = 0) -> None:
        if self.draft is None:
            return
        part = self.draft.xml_part_by_path(part_path)
        if part is None or part.text is None:
            return
        self._flush_all_editors()
        self.editor_notebook.select(self.xml_tab)
        if self.current_xml_part_path != part_path:
            self.current_xml_part_path = part_path
            self.xml_editor.text.config(state="normal")
            self.xml_editor.set_text(part.text)
            if not part.editable:
                self.xml_editor.text.config(state="disabled")
            self.xml_editor.set_part_info(part)
            self.validate_xml_btn.config(state="normal")
        try:
            self.xml_tree.selection_set("xml::" + part_path)
            self.xml_tree.see("xml::" + part_path)
        except tk.TclError:
            pass
        self.xml_editor.goto_position(line, column, length)
        self.xml_editor.text.focus_set()
        self._refresh_state()

    # -- ribbon callbacks ----------------------------------------------------------

    def flush_editors(self) -> None:
        self._flush_all_editors()

    def _xml_cursor_callback(self):
        path = self.current_xml_part_path
        if self.draft is None or path is None or not rs.is_ribbon_part(path):
            return None
        offset = len(self.xml_editor.text.get("1.0", "insert"))
        return rs.callback_at(self.xml_editor.get_text(), offset, path)

    def _cursor_procedure(self):
        if self.draft is None or self.current_module_id is None:
            return None, None
        mod = self.draft.module_by_id(self.current_module_id)
        if mod is None or mod.is_deleted:
            return None, None
        line = int(self.editor.text.index("insert").split(".")[0])
        return mod, rs.procedure_at(self.editor.get_text(), mod.id, line)

    def _on_xml_ctrl_click(self, event):
        self.xml_editor.text.mark_set("insert", f"@{event.x},{event.y}")
        self.goto_callback_under_cursor()
        return "break"

    def goto_callback_under_cursor(self) -> None:
        if self.draft is None:
            return
        if self._active_text_editor() is self.editor:
            self.show_buttons_for_macro()
            return
        callback = self._xml_cursor_callback()
        if callback is None:
            self.status.config(
                text="Put the cursor on an onAction (or other callback) attribute in the ribbon XML."
            )
            return
        self.goto_callback(callback)

    def goto_callback(self, callback) -> None:
        draft = self.draft
        if draft is None:
            return
        self._flush_all_editors()
        resolution = rs.resolve(rs.procedure_index(draft), callback.name)
        if not resolution.matches:
            if messagebox.askyesno(
                APP_NAME, f"{resolution.message}\n\nCreate a callback Sub for it now?"
            ):
                self.create_callback(callback)
            return
        match = resolution.matches[0]
        if len(resolution.matches) > 1:
            index = choose(
                self.root,
                "Go to Macro",
                resolution.message or "Choose a procedure:",
                [f"{m.module_name}.{m.name}   (line {m.line})" for m in resolution.matches],
            )
            if index is None:
                return
            match = resolution.matches[index]
        module = draft.module_by_id(match.module_id)
        if module is None:
            return
        line_text = module.body.split("\n")[match.line - 1]
        found = re.search(rf"\b{re.escape(match.name)}\b", line_text)
        column = found.start() + 1 if found else 1
        self._open_module_at(match.module_id, match.line, column, len(match.name) if found else 0)
        if resolution.status != rs.STATUS_OK and resolution.message:
            self.status.config(text=resolution.message)
        else:
            self.status.config(text=f"{callback.name} → {match.module_name}, line {match.line}")

    def goto_callback_xml(self, callback) -> None:
        self._open_xml_at(callback.part_path, callback.line, callback.column, callback.length)

    def _batch_edit(self, mutate) -> None:
        """Apply a programmatic multi-target edit as one undoable step."""
        draft = self.draft
        if draft is None:
            return
        self._flush_all_editors()
        before = {"modules": snapshot_modules(draft), "xml": snapshot_xml(draft)}
        mutate()
        # Reload before checkpointing: a checkpoint flushes the (stale) editors.
        self._reload_editors_from_draft()
        if self.session is not None:
            after = {"modules": snapshot_modules(draft), "xml": snapshot_xml(draft)}
            self.history_service.record_structural(
                self.session,
                HistoryCommand(op="replace_all", target_id="*", before=before, after=after),
            )
            self._checkpoint_now()
        if self._search_dialog is not None:
            self._search_dialog.mark_stale()

    def create_callback(self, callback) -> bool:
        draft = self.draft
        if draft is None:
            return False
        self._flush_all_editors()
        if draft.baseline.safety.password_protected:
            messagebox.showwarning(APP_NAME, "This project is password-protected and read-only.")
            return False
        module_name, proc_name = rs.split_callback_name(callback.name)
        if not rs.is_valid_procedure_name(proc_name):
            messagebox.showerror(APP_NAME, f"'{callback.name}' is not a valid VBA procedure name.")
            return False
        if module_name is not None:
            target = draft.find_current(module_name)
            if target is None:
                messagebox.showerror(
                    APP_NAME,
                    f"'{callback.name}' refers to module {module_name}, which does not exist.",
                )
                return False
        else:
            standard = [
                m for m in draft.modules if not m.is_deleted and m.pyopenvba_kind == "standard"
            ]
            if not standard:
                messagebox.showerror(
                    APP_NAME, "Ribbon callbacks must live in a standard module. Add one first."
                )
                return False
            usage: dict[str, int] = {}
            for entry in rs.ribbon_entries(draft):
                for match in entry.resolution.matches:
                    usage[match.module_id] = usage.get(match.module_id, 0) + 1
            initial = max(range(len(standard)), key=lambda i: usage.get(standard[i].id, 0))
            index = choose(
                self.root,
                "Create Callback",
                f"Add  Sub {proc_name}  to which module?",
                [
                    m.current_name
                    + (f"   ({usage[m.id]} ribbon callbacks)" if usage.get(m.id) else "")
                    for m in standard
                ],
                initial,
            )
            if index is None:
                return False
            target = standard[index]
        stub = rs.callback_stub(proc_name, callback.attribute, callback.element)
        new_body, line = rs.append_procedure(target.body, stub)

        def mutate() -> None:
            target.body = new_body

        self._batch_edit(mutate)
        self._open_module_at(target.id, line + 1, 5)
        self.status.config(text=f"Created {proc_name} in {target.current_name}.")
        return True

    def add_ribbon_button(self) -> None:
        draft = self.draft
        if draft is None:
            return
        self._flush_all_editors()
        mod, proc = self._cursor_procedure()
        if mod is None or proc is None:
            messagebox.showinfo(APP_NAME, "Put the cursor inside the Sub you want a button for.")
            return
        if draft.baseline.safety.password_protected:
            messagebox.showwarning(APP_NAME, "This project is password-protected and read-only.")
            return
        parts = [p for p in rs.ribbon_parts(draft) if p.editable]
        if not parts:
            messagebox.showinfo(
                APP_NAME,
                "This file has no editable ribbon XML (customUI/customUI.xml or "
                "customUI14.xml).\n\nAdding a new ribbon part is not supported yet.",
            )
            return
        containers = [c for p in parts for c in rs.find_containers(p.text or "", p.path)]
        if not containers:
            messagebox.showinfo(
                APP_NAME, "The ribbon XML has no custom group or menu to add a button to."
            )
            return
        plan = self._plan_button_callback(mod, proc)
        if plan is None:
            return
        callback_name, wrapper, note = plan
        existing_ids = {
            value for p in rs.ribbon_parts(draft) for value, _o in rs.control_ids(p.text or "")
        }
        images = sorted(
            {v for p in parts for v in re.findall(r'imageMso="([^"]+)"', p.text or "")},
            key=str.casefold,
        )
        dialog = AddButtonDialog(
            self.root,
            callback_name=callback_name,
            note=note,
            containers=containers,
            existing_ids=existing_ids,
            image_suggestions=images,
            default_label=rs.suggested_label(proc.name),
            default_id=rs.unique_control_id(existing_ids, proc.name),
            show_part=len(parts) > 1,
        )
        spec = dialog.wait()
        if spec is not None:
            self.apply_ribbon_button(mod.id, callback_name, wrapper, spec)

    def _plan_button_callback(self, mod, proc):
        """(callback name, wrapper stub or None, note) for a button that runs proc."""
        if self.draft is None:
            return None
        if proc.kind not in ("Sub", "Function"):
            messagebox.showerror(APP_NAME, f"{proc.name} is a {proc.kind}; a button needs a Sub.")
            return None
        if mod.pyopenvba_kind != "standard":
            messagebox.showerror(
                APP_NAME,
                f"{proc.name} is in {mod.current_name}, which is not a standard module. "
                "Ribbon buttons can only call Subs in standard modules.",
            )
            return None
        if rs.is_ribbon_ready(proc.signature):
            return proc.name, None, ""
        if rs.takes_arguments(proc.signature):
            messagebox.showerror(
                APP_NAME,
                f"{proc.name} takes arguments, so a ribbon button cannot call it directly. "
                "Write a callback Sub that supplies them, then add the button from there.",
            )
            return None
        index = rs.procedure_index(self.draft)
        base = f"{proc.name}Callback"
        existing = index.get(base.casefold(), [])
        if (
            len(existing) == 1
            and existing[0].is_standard
            and rs.is_ribbon_ready(existing[0].signature)
        ):
            return base, None, f"Uses the existing {base} in {existing[0].module_name}."
        name, n = base, 2
        while name.casefold() in index:
            name, n = f"{base}{n}", n + 1
        wrapper = rs.callback_stub(name, "onAction", "button", calls=proc.name)
        note = (
            f"{proc.name} has no (control As IRibbonControl) argument, so a small {name} "
            f"Sub that calls it will be added to {mod.current_name}."
        )
        return name, wrapper, note

    def apply_ribbon_button(self, module_id: str, callback_name: str, wrapper, spec) -> None:
        draft = self.draft
        if draft is None:
            return
        part = draft.xml_part_by_path(spec.container.part_path)
        mod = draft.module_by_id(module_id)
        if part is None or part.text is None or mod is None:
            return
        element = rs.build_button_xml(
            control_id=spec.control_id,
            label=spec.label,
            on_action=callback_name,
            image_mso=spec.image_mso,
            size=spec.size,
            screentip=spec.screentip,
            supertip=spec.supertip,
        )
        new_text, offset, length = rs.insert_into_container(part.text, spec.container, element)
        new_body = rs.append_procedure(mod.body, wrapper)[0] if wrapper else mod.body

        def mutate() -> None:
            part.text = new_text
            mod.body = new_body

        self._batch_edit(mutate)
        line = new_text.count("\n", 0, offset) + 1
        column = offset - (new_text.rfind("\n", 0, offset) + 1) + 1
        self._open_xml_at(part.path, line, column, length)
        extra = f" and {callback_name} in {mod.current_name}" if wrapper else ""
        self.status.config(
            text=f'Added button "{spec.label}"{extra}. Nothing is written until Save File.'
        )

    def show_buttons_for_macro(self) -> None:
        if self.draft is None:
            return
        self._flush_all_editors()
        mod, proc = self._cursor_procedure()
        if mod is None or proc is None:
            self.status.config(
                text="Put the cursor inside a Sub to find the ribbon buttons that call it."
            )
            return
        callbacks = rs.callbacks_for_procedure(self.draft, mod.id, proc.name)
        callbacks += rs.callbacks_for_procedure(self.draft, mod.id, f"{proc.name}Callback")
        if not callbacks:
            messagebox.showinfo(
                APP_NAME,
                f"No ribbon button calls {proc.name}.\n\n"
                "Use Ribbon → Add Ribbon Button for This Macro… to add one.",
            )
        elif len(callbacks) == 1:
            self.goto_callback_xml(callbacks[0])
        else:
            self.show_ribbon_buttons(initial_filter=proc.name)

    def show_ribbon_buttons(self, initial_filter: str = "", problems_only: bool = False) -> None:
        if self.draft is None:
            return
        if not rs.ribbon_parts(self.draft):
            messagebox.showinfo(APP_NAME, "This file has no ribbon XML (customUI).")
            return
        dialog = self._ribbon_dialog
        try:
            alive = dialog is not None and bool(dialog.win.winfo_exists())
        except tk.TclError:
            alive = False
        if dialog is None or not alive:
            dialog = self._ribbon_dialog = RibbonButtonsDialog(self.root, self, initial_filter)
        else:
            dialog.filter_var.set(initial_filter)
            dialog.win.lift()
        dialog.problems_only.set(problems_only)
        dialog.refresh()

    def check_ribbon_callbacks(self) -> None:
        if self.draft is None:
            return
        self._flush_all_editors()
        if not rs.ribbon_parts(self.draft):
            messagebox.showinfo(APP_NAME, "This file has no ribbon XML (customUI).")
            return
        issues = rs.check_ribbon(self.draft)
        if not issues:
            total = len(rs.ribbon_entries(self.draft))
            messagebox.showinfo(
                APP_NAME,
                f"All {total} ribbon callbacks point to existing macros, "
                "and control ids are unique.",
            )
            return
        text = "\n".join(i.describe() for i in issues[:15]) + ("\n…" if len(issues) > 15 else "")
        if messagebox.askyesno(
            APP_NAME,
            f"{len(issues)} ribbon problem(s):\n\n{text}\n\nOpen the Ribbon Buttons list?",
            icon="warning",
        ):
            self.show_ribbon_buttons(problems_only=True)

    def _confirm_ribbon_issues(self) -> bool:
        """Warn (never block) about ribbon problems introduced by this draft."""
        if self.draft is None:
            return True
        try:
            issues = rs.new_issues(self.draft)
        except Exception:  # noqa: BLE001 - a checker bug must never block saving
            return True
        if not issues:
            return True
        text = "\n".join(i.describe() for i in issues[:12]) + ("\n…" if len(issues) > 12 else "")
        return messagebox.askyesno(
            APP_NAME,
            "Your edits introduce ribbon problems that will make buttons fail in Office:\n\n"
            f"{text}\n\nSave anyway?",
            icon="warning",
        )

    def backup_browser(self) -> None:
        if self.draft is None:
            messagebox.showinfo(APP_NAME, "Open an add-in first.")
            return
        records = self.backup_service.list_records(self.draft.baseline.path)
        legacy = self.backup_service.discover_legacy(self.draft.baseline.path)
        lines = ["Backup catalog", ""]
        for rec in records:
            lines.append(f"{rec.utc_time}  {Path(rec.backup_path).name}  {rec.sha256[:12]}")
        if legacy:
            lines.append("")
            lines.append("Same-folder backups:")
            lines.extend(f"  {path.name}" for path in legacy)
        messagebox.showinfo(APP_NAME, "\n".join(lines)[:4000])

    def export_source_folder(self) -> None:
        if self.session is None:
            messagebox.showinfo(APP_NAME, "Open an add-in first.")
            return
        folder = filedialog.askdirectory(title="Export Source Folder")
        if not folder:
            return
        try:
            written = self.folder_sync.export_folder(self.session, Path(folder))
            self.status.config(text=f"Exported source folder manifest {written}")
        except AdapterError as exc:
            messagebox.showerror(APP_NAME, str(exc))

    def preview_folder_changes(self) -> None:
        if self.session is None:
            messagebox.showinfo(APP_NAME, "Open an add-in first.")
            return
        folder = filedialog.askdirectory(title="Folder with vbaae-project.json")
        if not folder:
            return
        preview = self.folder_sync.preview(self.session, Path(folder))
        if preview.problems:
            messagebox.showerror(APP_NAME, "\n".join(preview.problems))
            return
        lines = [f"{change.operation}: {change.logical_name}" for change in preview.changes if change.operation != "noop"]
        if not lines:
            messagebox.showinfo(APP_NAME, "No folder changes to apply.")
            return
        if not messagebox.askokcancel(APP_NAME, "Apply currently selected folder edits?\n\n" + "\n".join(lines)):
            return
        try:
            self.folder_sync.apply(self.session, preview, history_service=self.history_service)
        except AdapterError as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return
        self.recovery_service.checkpoint(self.session)
        self._reload_editors_from_draft()
        self.review_changes()

    def copy_last_diagnostic(self) -> None:
        if self._last_result is None:
            messagebox.showinfo(APP_NAME, "No operation has been recorded yet.")
            return
        include = messagebox.askyesno(APP_NAME, "Include full paths in the diagnostic report?")
        report = report_for(
            self._last_result,
            include_full_paths=bool(include),
            extension=self.draft.baseline.extension if self.draft else None,
        )
        self.root.clipboard_clear()
        self.root.clipboard_append(report)
        preview = tk.Toplevel(self.root)
        preview.title("Diagnostic Report")
        text = tk.Text(preview, width=90, height=24)
        text.pack(fill="both", expand=True)
        text.insert("1.0", report)
        text.config(state="disabled")

    def _adopt_session(self, session) -> None:
        if self.session is not None and self.session is not session:
            self.session_service.close(self.session)
        self.session = session
        self.draft = session.draft
        self.current_module_id = None
        self.current_xml_part_path = None
        path = session.original_path
        xml_supported = path.suffix.lower() in XML_EDITABLE_EXTENSIONS
        self.editor_notebook.tab(self.xml_tab, state="normal" if xml_supported else "disabled")
        self._reload_editors_from_draft()
        first = next((m.id for m in session.draft.modules if not m.is_deleted), None)
        if first:
            try:
                self.tree.selection_set(first)
            except tk.TclError:
                pass
        self._refresh_host_status()
        self._refresh_state()
        self.status.config(text=f"Recovered {path.name}")

    def _offer_recovery(self) -> None:
        listings = self.recovery_service.list_recoverable()
        if not listings:
            return
        first = listings[0]
        if first.invalid:
            messagebox.showwarning(
                APP_NAME,
                f"A recovery folder is present but invalid ({first.reason}). "
                "It was left in place for inspection/export.",
            )
            return
        choice = messagebox.askyesnocancel(
            APP_NAME,
            f"A recovered draft was found for {first.filename} "
            f"(revision {first.revision}, {first.checkpoint_utc}).\n\n"
            "Yes = Recover now, No = Keep for later, Cancel = Delete recovery.",
        )
        if choice is None:
            self.recovery_service.delete_recovery(first.directory)
            return
        if choice is False:
            return
        opened = self.recovery_open.open_recoverable(first.directory)
        if opened.status == "invalid" or opened.session is None:
            messagebox.showerror(
                APP_NAME,
                "Recovery data is invalid and was left in place for inspection.\n\n"
                + (opened.reason or "invalid"),
            )
            return
        self._adopt_session(opened.session)
        if opened.status == "missing_source":
            messagebox.showwarning(
                APP_NAME,
                "The original add-in is missing. In-place Save is disabled. "
                "Use Save a Copy to write the recovered draft to a different path. "
                "Newer or missing source bytes are not included.",
            )
            return
        if opened.status == "needs_conflict":
            if opened.external is None:
                messagebox.showwarning(
                    APP_NAME,
                    "The original file changed and could not be parsed. "
                    "In-place Save is blocked. Use Save a Copy or Cancel.",
                )
                return
            self._run_conflict_resolution(opened.external)

    def compare_external_changes(self) -> None:
        if self.session is None or self.draft is None:
            return
        snapshot, _digest, error = self.conflict_service.snapshot_external(self.draft.baseline.path)
        if snapshot is None:
            messagebox.showerror(
                APP_NAME,
                "The current file could not be compared: " + (error or "unknown"),
            )
            return
        self._run_conflict_resolution(snapshot)

    def _run_conflict_resolution(self, external) -> None:
        if self.session is None:
            return
        proposal = self.conflict_service.compare(self.session, external)
        if proposal.parse_failed:
            messagebox.showwarning(
                APP_NAME,
                "The current file could not be parsed. Cancel, save a recovered copy, or discard.",
            )
            return
        if not any(item.requires_choice for item in proposal.items):
            if not messagebox.askokcancel(
                APP_NAME,
                "No unresolved choices. Apply auto-resolved external baseline adoption?",
            ):
                return
            result = self.conflict_service.commit(
                self.session,
                proposal,
                external,
                live_external_hash=external.file_fingerprint.sha256,
                external_path=self.session.original_path,
                session_service=self.session_service,
                recovery_service=self.recovery_service,
            )
            if result.kind != "success":
                messagebox.showerror(APP_NAME, result.message or result.reason or "Apply failed.")
                return
            self.draft = self.session.draft
            self._reload_editors_from_draft()
            return
        dialog = ConflictDialog(self.root, proposal)
        if os.environ.get("PYTEST_CURRENT_TEST"):
            return
        choice = dialog.wait()
        if choice != "apply":
            return
        result = self.conflict_service.commit(
            self.session,
            proposal,
            external,
            live_external_hash=external.file_fingerprint.sha256,
            external_path=self.session.original_path,
            session_service=self.session_service,
            recovery_service=self.recovery_service,
        )
        if result.kind != "success":
            messagebox.showerror(APP_NAME, result.message or result.reason or "Apply failed.")
            return
        self.draft = self.session.draft
        self._reload_editors_from_draft()

    def on_close(self) -> None:
        if self._saving:
            return  # never tear down mid-commit
        self._flush_all_editors()
        if self.draft is not None and self.draft.is_dirty():
            self._refresh_host_status()
            extra = ""
            if self._live_host_running:
                extra = (
                    f"\n\nClose {self._host_label} before saving this add-in. "
                    "Saving now will be blocked until that application is closed."
                )
            choice = messagebox.askyesnocancel(
                APP_NAME, "You have unsaved changes. Save before exiting?" + extra
            )
            if choice is None:
                return
            if choice:
                self.save()
                if self.draft is not None and self.draft.is_dirty():
                    return  # save failed or was blocked; keep window open
            else:
                if self.session is not None:
                    self.recovery_service.mark_discard(self.session)
        if self.session is not None:
            self.session_service.close(self.session)
        self._cancel_afters()
        self.root.destroy()

    def _on_root_destroy(self, event) -> None:
        if event.widget is self.root:
            self._cancel_afters()

    def _cancel_afters(self) -> None:
        for attr in ("_idle_after", "_max_after", "_poll_after", "_recovery_after"):
            handle = getattr(self, attr, None)
            if handle is None:
                continue
            try:
                self.root.after_cancel(handle)
            except tk.TclError:
                pass
            setattr(self, attr, None)


def run_gui() -> None:
    root = tk.Tk()
    MainWindow(root)
    root.mainloop()
