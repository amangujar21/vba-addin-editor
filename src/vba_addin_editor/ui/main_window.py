"""Main application window (plan sections 9, 18, 43-46, 57)."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from vba_addin_editor.adapters.ooxml_package_adapter import XML_EDITABLE_EXTENSIONS
from vba_addin_editor.adapters.pyopenvba_adapter import AdapterError
from vba_addin_editor.domain.changes import compute_changes, dirty_count
from vba_addin_editor.domain.document import (
    DocumentDraft,
    ModuleDisplayKind,
    ModuleDraft,
    new_module_id,
)
from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.services.import_export_service import BackupService, ImportExportService
from vba_addin_editor.services.save_service import SaveService
from vba_addin_editor.services.validation_service import validate_module_name
from vba_addin_editor.ui.code_editor import CodeEditor
from vba_addin_editor.ui.xml_editor import XmlEditor
from vba_addin_editor.version import APP_NAME, VERSION

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
        self.current_module_id: str | None = None
        self._saving = False
        self.doc_service = DocumentService()
        self.save_service = SaveService(adapter=self.doc_service.adapter, progress=self._on_progress)
        self.ie_service = ImportExportService(adapter=self.doc_service.adapter)
        self.backup_service = BackupService(adapter=self.doc_service.adapter)

        root.title(APP_NAME)
        root.geometry("1000x680")
        root.minsize(760, 480)
        self._build_menu()
        self._build_toolbar()
        self._build_panes()
        self._build_statusbar()
        self._bind_shortcuts()
        root.protocol("WM_DELETE_WINDOW", self.on_close)

    # -- construction ------------------------------------------------------

    def _build_menu(self) -> None:
        bar = tk.Menu(self.root)
        file_menu = tk.Menu(bar, tearoff=0)
        file_menu.add_command(label="Open Office VBA File…", accelerator="Ctrl+O", command=self.open_file)
        file_menu.add_command(label="Save File", accelerator="Ctrl+S", command=self.save)
        file_menu.add_command(label="Save a Copy…", accelerator="Ctrl+Shift+S", command=self.save_copy)
        file_menu.add_separator()
        file_menu.add_command(label="Restore Backup…", command=self.restore_backup)
        file_menu.add_command(label="Open File Location", command=self.open_location)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_close)
        bar.add_cascade(label="File", menu=file_menu)

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
        bar.add_cascade(label="Help", menu=help_menu)
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
        can_save = not draft.baseline.safety.host_process_running and not self._saving
        self.save_btn.config(state="normal" if can_save else "disabled")
        self.review_btn.config(state="normal" if n else "disabled")
        self.revert_btn.config(state="normal" if n else "disabled")
        if draft.baseline.safety.password_protected:
            self._set_banner(
                "This VBA project is password-protected. VBA editing is disabled; "
                "XML package editing remains available for .ppam/.pptm files."
            )
        elif draft.baseline.safety.host_process_running:
            host = "Excel" if draft.baseline.host_kind == "excel" else "PowerPoint"
            self._set_banner(
                f"{host} is currently running. Close {host} before saving."
            )
        else:
            self._set_banner("")

    def _refresh_tree(self) -> None:
        draft = self.draft
        self.tree.delete(*self.tree.get_children())
        if draft is None:
            return
        groups = {"Standard Modules": [], "Object / Class / Form Code": []}
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
            group = "Standard Modules" if m.pyopenvba_kind == "standard" else "Object / Class / Form Code"
            label = f"{m.current_name}{marker}" + ("" if m.destructive_ops_safe else " [lock]")
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
            mod.body = self.editor.get_text()

    def _flush_active_xml_editor(self) -> None:
        draft = self.draft
        if draft is None or self.current_xml_part_path is None:
            return
        part = draft.xml_part_by_path(self.current_xml_part_path)
        if part is not None and part.editable:
            part.text = self.xml_editor.get_text()

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
            draft = self.doc_service.open(path)
        except AdapterError as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return
        self.draft = draft
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
        changes = compute_changes(self.draft)
        if changes.is_empty:
            messagebox.showinfo(APP_NAME, "No changes to save.")
            return
        if not self._preflight_dialog(changes):
            return
        self._saving = True
        self._refresh_state()
        try:
            result = self.save_service.save_addin(self.draft)
        finally:
            self._saving = False
        self._handle_save_result(result)

    def _handle_save_result(self, result) -> None:
        kind = result.kind
        if kind == "success":
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
            messagebox.showinfo(APP_NAME, msg)
        elif kind == "no_changes":
            messagebox.showinfo(APP_NAME, result.message or "No changes to save.")
        elif kind == "blocked":
            messagebox.showwarning(APP_NAME, result.message or "Save blocked.")
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
        path = filedialog.asksaveasfilename(
            title="Save a Copy",
            defaultextension=self.draft.baseline.extension,
            filetypes=_FILETYPES,
        )
        if not path:
            return
        result = self.save_service.save_copy(self.draft, Path(path))
        if result.kind == "success":
            messagebox.showinfo(
                APP_NAME,
                "Copy saved. Excel/PowerPoint will continue using the original "
                "installed add-in unless you change its add-in configuration.",
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
        )
        draft.modules.append(mod)
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
                mod.current_name = name
        else:
            if not mod.destructive_ops_safe:
                messagebox.showwarning(
                    APP_NAME,
                    "This module is an object/form or host-bound component; renaming it "
                    "is disabled in this version.",
                )
                return
            name = self._prompt_name("Rename Module", mod.current_name)
            if name:
                mod.current_name = name
        self._flush_all_editors()
        self._refresh_tree()
        self._refresh_state()

    def delete_module(self) -> None:
        mod = self._selected_module()
        if mod is None or self.draft is None:
            return
        if not mod.destructive_ops_safe:
            messagebox.showwarning(
                APP_NAME,
                "Deleting this module (object/form or host-bound code) is disabled in "
                "this version because its subtype cannot be verified safely.",
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
        messagebox.showinfo(
            APP_NAME,
            f"{APP_NAME} {VERSION}\n\n"
            "Edits VBA source inside installed .xlam / .ppam add-ins and .pptm\n"
            "presentations, in place, with automatic backup and verification.\n\n"
            "Limitations: no VBA compile validation; password-protected projects "
            "are read-only; UserForm layout cannot be created or edited.\n\n"
            "Uses pyOpenVBA (MIT) — see Help → Third-party notices.",
        )

    def on_close(self) -> None:
        self._flush_all_editors()
        if self.draft is not None and self.draft.is_dirty():
            host = "Excel" if self.draft.baseline.host_kind == "excel" else "PowerPoint"
            if self.draft.baseline.safety.host_process_running:
                messagebox.showwarning(
                    APP_NAME,
                    f"You have unsaved changes, but {host} is running, so the add-in "
                    "cannot be saved now. Exit anyway?",
                )
                self.root.destroy()
                return
            choice = messagebox.askyesnocancel(
                APP_NAME, "You have unsaved changes. Save before exiting?"
            )
            if choice is None:
                return
            if choice:
                self.save()
                if self.draft is not None and self.draft.is_dirty():
                    return  # save failed or was blocked; keep window open
        self.root.destroy()


def run_gui() -> None:
    root = tk.Tk()
    MainWindow(root)
    root.mainloop()
