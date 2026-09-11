"""Main application window (plan sections 9, 18, 43-46, 57)."""

from __future__ import annotations

import os
import tempfile
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from vba_addin_editor.adapters.ooxml_package_adapter import XML_EDITABLE_EXTENSIONS
from vba_addin_editor.adapters.pyopenvba_adapter import AdapterError
from vba_addin_editor.domain.capabilities import RESTRICTION_MESSAGES
from vba_addin_editor.domain.changes import compute_changes, dirty_count
from vba_addin_editor.domain.document import (
    DocumentDraft,
    ModuleDisplayKind,
    ModuleDraft,
    new_module_id,
)
from vba_addin_editor.domain.history import HistoryCommand
from vba_addin_editor.platform import windows_processes as wp
from vba_addin_editor.services.backup_service import BackupService
from vba_addin_editor.services.conflict_service import ConflictService
from vba_addin_editor.services.diagnostics_service import dialog_text, report_for
from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.services.folder_sync_service import FolderSyncService
from vba_addin_editor.services.history_service import HistoryService, snapshot_modules
from vba_addin_editor.services.import_export_service import ImportExportService
from vba_addin_editor.services.recovery_service import IDLE_MS, MAX_INTERVAL_MS, RecoveryService
from vba_addin_editor.services.review_service import ReviewService
from vba_addin_editor.services.save_service import SaveService
from vba_addin_editor.services.search_service import SearchService
from vba_addin_editor.services.session_service import SessionService
from vba_addin_editor.services.validation_service import validate_module_name
from vba_addin_editor.ui.code_editor import CodeEditor
from vba_addin_editor.ui.xml_editor import XmlEditor
from vba_addin_editor.version import APP_NAME, VERSION, build_identity

_FILETYPES = [
    ("Supported Office VBA files", "*.xlam;*.ppam;*.pptm"),
    ("Excel Add-ins", "*.xlam"),
    ("PowerPoint Add-ins", "*.ppam"),
    ("PowerPoint Macro-Enabled Presentations", "*.pptm"),
]


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

        help_menu = tk.Menu(bar, tearoff=0)
        help_menu.add_command(label="About", command=self.show_about)
        help_menu.add_command(label="Copy Diagnostic Report", command=self.copy_last_diagnostic)
        bar.add_cascade(label="Help", menu=help_menu)
        self._last_result = None
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
        self.editor_notebook.add(vba_tab, text="VBA")
        self.editor_notebook.add(self.xml_tab, text="XML")
        self.editor_notebook.pack(fill="both", expand=True)
        self.banner = ttk.Label(self.root, text="", background="#fff3cd", padding=4)
        self.banner.pack(fill="x", before=self.editor_notebook)
        self.banner.pack_forget()

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

    # -- banners / status ----------------------------------------------------

    def _set_banner(self, text: str) -> None:
        if text:
            self.banner.config(text=text)
            self.banner.pack(fill="x", before=self.status)
        else:
            self.banner.pack_forget()

    def _on_progress(self, stage: str) -> None:
        self.status.config(text=f"Saving… {stage}")
        self.root.update_idletasks()

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
        can_save = not self._saving
        self.save_btn.config(state="normal" if can_save else "disabled")
        self.review_btn.config(state="normal" if n else "disabled")
        self.revert_btn.config(state="normal" if n else "disabled")
        if self._recovery_warning:
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
        groups = {
            "Standard Modules": [],
            "Class Modules": [],
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
            else:
                group = "Object / Class / Form Code"
            locked = not (m.can_delete or m.can_rename or m.is_new)
            label = f"{m.current_name}{marker}" + ("" if not locked else " [lock]")
            groups[group].append((m.id, label))
        for group, items in groups.items():
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
        self._flush_all_editors()
        mid = selection[0]
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
        self._flush_all_editors()
        path = selection[0][len("xml::"):]
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
        self._flush_all_editors()
        self._checkpoint_now()
        changes = compute_changes(self.draft)
        if changes.is_empty:
            messagebox.showinfo(APP_NAME, "No changes to save.")
            return
        if not self._review_and_confirm(operation_type="save"):
            return
        self._saving = True
        self._refresh_state()
        result = None
        try:
            result = self.save_service.save_addin(self.draft)
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
        self._handle_save_result(result)

    def _handle_save_result(self, result) -> None:
        self._last_result = result
        kind = result.kind
        if kind == "success":
            warning = self.backup_service.record_success(
                original=self.draft.baseline.path if self.draft else Path("."),
                backup=result.backup_path,
                operation_type=result.operation_type,
            )
            if self.session is not None:
                self.session.history.clear()
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
            if messagebox.askyesno(
                APP_NAME,
                "Final verification failed after replacing the add-in.\n\nRestore the backup now?",
            ):
                try:
                    self.backup_service.restore(self.draft.baseline.path, result.backup_path)
                    messagebox.showinfo(APP_NAME, "Backup restored.")
                except AdapterError as exc:
                    messagebox.showerror(APP_NAME, str(exc))
            else:
                messagebox.showwarning(APP_NAME, f"Keep this backup safe: {result.backup_path}")
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
        result = self.save_service.save_copy(
            self.draft,
            dest,
            source_path=self.session.captured_path if self.session is not None else None,
            allow_overwrite=dest.exists(),
            reviewed_dest_hash=dest_hash,
            session_dir=self.session.session_dir if self.session is not None else None,
        )
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
        if not messagebox.askyesno(
            APP_NAME,
            f'Delete module "{mod.current_name}"?\n\nThis removes the module from the '
            "VBA project when you save. It is not written to the file until Save File.",
        ):
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
        path = filedialog.askopenfilename(
            title="Select Backup", filetypes=_FILETYPES
        )
        if not path:
            return
        try:
            safety = self.backup_service.restore(self.draft.baseline.path, Path(path))
            messagebox.showinfo(
                APP_NAME, f"Backup restored.\n\nPre-restore safety copy: {safety}"
            )
            self.load_path(self.draft.baseline.path)
        except AdapterError as exc:
            messagebox.showerror(APP_NAME, str(exc))

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
        if self.draft is None:
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
        if self.draft is not None:
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
        query = self._prompt_name("Find in Project", "")
        if not query:
            return
        revision = self.session.revision if self.session is not None else 0
        results = self.search_service.search(
            self.draft, query, revision=revision, include_xml=True
        )
        lines = [f"{results.total} matches for {query!r}", ""]
        for hit in self.search_service.page(results, 0):
            lines.append(f"{hit.path}:{hit.line}:{hit.column}  {hit.snippet}")
        messagebox.showinfo(APP_NAME, "\n".join(lines)[:4000] or "No matches.")
        if results.hits:
            first = results.hits[0]
            if first.kind == "vba":
                self.tree.selection_set(first.target_id)
                self.editor.goto_line(first.line)
            else:
                self.editor_notebook.select(self.xml_tab)
                self.xml_tree.selection_set("xml::" + first.target_id)

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
        if not messagebox.askokcancel(APP_NAME, "Apply selected folder edits?\n\n" + "\n".join(lines)):
            return
        for change in preview.changes:
            if change.operation in {"edit", "add"}:
                change.selected = True
        try:
            self.folder_sync.apply(self.session, preview)
        except AdapterError as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return
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

    def _offer_recovery(self) -> None:
        listings = self.recovery_service.list_recoverable()
        if not listings:
            return
        first = listings[0]
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
        source = Path(first.source_path) if first.source_path else None
        if source is None or not source.exists():
            messagebox.showwarning(
                APP_NAME,
                "The original add-in is missing. Recover the draft, then use Save a Copy.",
            )
            return
        self.load_path(source)
        if self.session is None:
            return
        try:
            data = self.recovery_service.load_checkpoint(first.directory)
            self.recovery_service.apply_checkpoint(self.session, data)
            self.recovery_service.revalidate_capabilities(self.session.draft)
            self.draft = self.session.draft
            self._reload_editors_from_draft()
        except (OSError, ValueError, KeyError) as exc:
            messagebox.showerror(APP_NAME, f"Recovery could not be applied: {type(exc).__name__}")

    def compare_external_changes(self) -> None:
        if self.session is None or self.draft is None:
            return
        try:
            external = self.doc_service.open(self.draft.baseline.path)
        except AdapterError as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return
        proposal = self.conflict_service.compare(self.session, external.baseline)
        unresolved = [item for item in proposal.items if item.requires_choice]
        if not unresolved:
            messagebox.showinfo(APP_NAME, "No unresolved external differences.")
            return
        lines = [f"{item.category} {item.property_name} {item.target_id}" for item in unresolved]
        messagebox.showwarning(APP_NAME, "External changes need review:\n\n" + "\n".join(lines)[:4000])

    def on_close(self) -> None:
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
