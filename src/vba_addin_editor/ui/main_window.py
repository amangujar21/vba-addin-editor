"""Main application window (plan sections 9, 18, 43-46, 57)."""

from __future__ import annotations

import gc
import hashlib
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
from vba_addin_editor.adapters.pyopenvba_adapter import AdapterError, enable_source_disk_cache
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
from vba_addin_editor.platform.paths import appdata_root
from vba_addin_editor.platform.prescan import READY, Prescan
from vba_addin_editor.platform.settings import load_settings, save_settings
from vba_addin_editor.services import icon_service
from vba_addin_editor.services import port_service as ps
from vba_addin_editor.services import ribbon_service as rs
from vba_addin_editor.services import starter_service as ss
from vba_addin_editor.services.backup_service import BackupService
from vba_addin_editor.services.conflict_service import ConflictService
from vba_addin_editor.services.diagnostics_service import dialog_text, report_for
from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.services.folder_sync_service import FolderSyncService
from vba_addin_editor.services.history_service import (
    HistoryService,
    snapshot_modules,
    snapshot_package,
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
from vba_addin_editor.services.search_service import SearchService, parse_procedures
from vba_addin_editor.services.session_service import SessionService
from vba_addin_editor.services.validation_service import validate_module_name
from vba_addin_editor.ui import theme
from vba_addin_editor.ui.code_editor import CodeEditor
from vba_addin_editor.ui.command_palette import (
    COMMAND,
    MODULE,
    PROCEDURE,
    CommandPalette,
    PaletteItem,
)
from vba_addin_editor.ui.conflict_dialog import ConflictDialog
from vba_addin_editor.ui.review_dialog import REVERT, SAVE, ReviewView, ask_review, entry_from_item
from vba_addin_editor.ui.ribbon_designer import RibbonDesigner
from vba_addin_editor.ui.ribbon_dialogs import (
    AddButtonDialog,
    AddMenuDialog,
    GroupMenuDialog,
    ImportControlDialog,
    RibbonButtonsDialog,
    StarterRibbonDialog,
    ask_image_file,
    choose,
)
from vba_addin_editor.ui.search_dialog import ProjectSearchDialog
from vba_addin_editor.ui.widgets import (
    RailButton,
    Sidebar,
    Spinner,
    Toast,
    Tooltip,
    hover_row,
    icon,
)
from vba_addin_editor.ui.xml_editor import XmlEditor
from vba_addin_editor.version import APP_NAME, PYOPENVBA_PIN, VERSION, build_identity

_FILETYPES = [
    ("Supported Office VBA files", "*.xlam;*.ppam;*.pptm"),
    ("Excel Add-ins", "*.xlam"),
    ("PowerPoint Add-ins", "*.ppam"),
    ("PowerPoint Macro-Enabled Presentations", "*.pptm"),
]
_NEW_FILETYPES = [
    ("Excel Add-in", "*.xlam"),
    ("PowerPoint Add-in", "*.ppam"),
    ("PowerPoint Macro-Enabled Presentation", "*.pptm"),
]


_T = TypeVar("_T")

# Save blocks that clear once Office closes or lets go of the file.
_HOST_WAIT_REASONS = frozenset({"office_running", "file_in_use"})
QUEUED_SAVE_POLL_MS = 500
MAX_RECENT_FILES = 8
MIN_CODE_SIZE, MAX_CODE_SIZE = 7, 28
_FILE_KINDS = {".xlam": "Excel add-in", ".ppam": "PowerPoint add-in", ".pptm": "PowerPoint presentation"}
# Sidebar groups, in display order.
GROUP_MODULES, GROUP_CLASSES, GROUP_FORMS = "Modules", "Class modules", "Forms"
# Start of a procedure, or its end, for the breadcrumb (Tcl regular expression).
_PROC_EDGE_TCL = (
    r"^[ \t]*((public|private|friend)[ \t]+)?(static[ \t]+)?"
    r"(sub|function|property[ \t]+(get|let|set))[ \t]+[a-z_][a-z0-9_]*"
    r"|^[ \t]*end[ \t]+(sub|function|property)"
)
_PROC_DECL = re.compile(
    r"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)",
    re.IGNORECASE,
)
CURSOR_FOLLOW_MS = 150
HEADER_PATH_MIN_WIDTH = 1440  # narrower: the folder path moves to the file name tooltip
HEADER_COMMAND_MIN_WIDTH = 1100  # narrower: the command box becomes a search icon
OUTLINE_REFRESH_MS = 400
# Badge in the PROCEDURES list; short so the narrow column never truncates it.
OUTLINE_MARKS = {"Sub": "", "Function": "fn", "Property Get": "get", "Property Let": "let", "Property Set": "set"}


def _shorten(text: str, limit: int = 56) -> str:
    """Keep both ends of a long path, joined by an ellipsis."""
    if len(text) <= limit:
        return text
    keep = (limit - 1) // 2
    return text[:keep] + "…" + text[-keep:]


def _when(path: Path) -> str:
    """Today / Yesterday / Nov 2 for a recent file's last change."""
    try:
        stamp = path.stat().st_mtime
    except OSError:
        return "missing"
    day = time.localtime(stamp)
    now = time.localtime()
    days = (time.mktime(now[:3] + (0, 0, 0, 0, 0, -1)) - time.mktime(day[:3] + (0, 0, 0, 0, 0, -1))) // 86400
    if days <= 0:
        return "Today"
    if days == 1:
        return "Yesterday"
    return time.strftime("%b %d", day).replace(" 0", " ")


def _draft_signature(draft: DocumentDraft) -> str:
    """Digest of everything a save would write, to tell whether a reviewed draft changed."""
    digest = hashlib.sha256()
    for mod in draft.modules:
        digest.update(repr((mod.id, mod.current_name, mod.body, mod.is_deleted)).encode("utf-8", "surrogatepass"))
    for part in draft.xml_parts:
        digest.update(repr((part.path, part.text, part.is_new)).encode("utf-8", "surrogatepass"))
    for binary in draft.binary_parts:
        digest.update(repr((binary.path, hashlib.sha256(binary.data).hexdigest())).encode("utf-8"))
    digest.update(repr(("removed", draft.removed_member_paths())).encode("utf-8"))
    return digest.hexdigest()


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
        # Draft signature a confirmed save is waiting to write once Office lets go.
        self._queued_save: str | None = None
        self._queue_after = None
        self._prescan = Prescan()
        self._prescan_after = None
        self._prescan_holds_gc = False
        self._gc_holds = 0
        self._gc_was_enabled = False
        self._last_vba_text = ""
        self._last_xml_text = ""
        self._follow_after = None
        self._shown_view = ""
        # Where a ribbon edit is shown afterwards: the view used last, "design" or "xml".
        self._ribbon_surface = "design"
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
        # Tests keep their settings beside their temporary sessions.
        self._settings_path = root_path / "settings.json" if root_path is not None else None
        self.settings = load_settings(self._settings_path)
        self.save_service.auto_backup = self.settings.auto_backup
        self.auto_backup_var = tk.BooleanVar(master=root, value=self.settings.auto_backup)
        if root_path is None:
            enable_source_disk_cache(appdata_root() / "source-cache" / f"pyopenvba-{PYOPENVBA_PIN}")

        root.title(APP_NAME)
        root.geometry("1280x800")
        root.minsize(820, 520)
        # Before any widget exists: classic Tk widgets read the option database once.
        theme.apply_theme(root, self.settings.theme, self.settings.editor_font_size)
        theme.set_title_bar(root)
        self._build_menu()
        self._build_toolbar()
        self._build_panes()
        self._build_statusbar()
        self.command_palette = CommandPalette(root, self._palette_items)
        self.toast = Toast(root)
        self._bind_shortcuts()
        for tree in (self.tree, self.outline, self.xml_tree, self.forms_tree):
            self._configure_tree_tags(tree)
        self._refresh_state()
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        root.bind("<Destroy>", self._on_root_destroy, add=True)
        root.bind("<FocusIn>", self._on_focus_in, add=True)
        self.editor.on_change = self._on_vba_edited
        self.xml_editor.on_change = self._on_xml_edited
        self.editor.on_undo = self.undo
        self.editor.on_redo = self.redo
        self.xml_editor.on_undo = self.undo
        self.xml_editor.on_redo = self.redo
        self.editor.baseline_provider = self._module_baseline
        self.xml_editor.baseline_provider = self._xml_baseline
        self._recovery_after = None
        self._schedule_host_poll()
        if not os.environ.get("PYTEST_CURRENT_TEST"):
            self._recovery_after = self.root.after_idle(self._offer_recovery)

    # -- construction ------------------------------------------------------

    def _header_menu(self, label: str) -> tk.Menu:
        """A header button with its drop-down menu (the dark stand-in for a Win32 menu bar)."""
        button = ttk.Menubutton(self.header, text=label, style="Header.TMenubutton", takefocus=False)
        button.pack(side="left")
        menu = tk.Menu(button, tearoff=0)
        button.configure(menu=menu)
        self._header_buttons.append(button)
        self._menus.append(menu)
        # Alt+<first letter> opens it, as on a native menu bar; F10 opens the first.
        self.root.bind(f"<Alt-KeyPress-{label[0].lower()}>", lambda _e: self._post_header_menu(button))
        return menu

    def _post_header_menu(self, button: ttk.Menubutton) -> str:
        if str(button.cget("state")) != "disabled":
            menu = self.root.nametowidget(button.cget("menu"))
            menu.tk_popup(button.winfo_rootx(), button.winfo_rooty() + button.winfo_height())
        return "break"

    def _build_menu(self) -> None:
        self.header = ttk.Frame(self.root, style="Chrome.TFrame", padding=(10, 6, 8, 6))
        self.header.pack(fill="x", side="top")
        self.header_line = ttk.Frame(self.root, style="Line.TFrame", height=1)
        self.header_line.pack(fill="x", side="top")
        ttk.Label(self.header, text="V", style="Logo.TLabel").pack(side="left", padx=(2, 10))
        self._header_buttons: list[ttk.Menubutton] = []
        self._menus: list[tk.Menu] = []
        self._sidebars: list[Sidebar] = []
        self.root.bind("<F10>", lambda _e: self._post_header_menu(self._header_buttons[0]))

        file_menu = self._header_menu("File")
        file_menu.add_command(label="New Add-in…", accelerator="Ctrl+N", command=self.new_file)
        file_menu.add_command(label="Open Office VBA File…", accelerator="Ctrl+O", command=self.open_file)
        self.recent_menu = tk.Menu(file_menu, tearoff=0, postcommand=self._fill_recent_menu)
        self._menus.append(self.recent_menu)
        file_menu.add_cascade(label="Open Recent", menu=self.recent_menu)
        file_menu.add_command(label="Save File", accelerator="Ctrl+S", command=self.save)
        file_menu.add_command(label="Save a Copy…", accelerator="Ctrl+Shift+S", command=self.save_copy)
        file_menu.add_separator()
        file_menu.add_command(label="Back Up Now", command=self.backup_now)
        file_menu.add_checkbutton(
            label="Back Up Automatically on Save",
            variable=self.auto_backup_var,
            command=self._on_auto_backup_toggled,
        )
        file_menu.add_command(label="Restore Backup…", command=self.restore_backup)
        file_menu.add_command(label="Backup Browser…", command=self.backup_browser)
        file_menu.add_command(label="Compare External Changes…", command=self.compare_external_changes)
        file_menu.add_command(label="Open File Location", command=self.open_location)
        file_menu.add_separator()
        file_menu.add_command(label="Export Source Folder…", command=self.export_source_folder)
        file_menu.add_command(label="Preview Folder Changes…", command=self.preview_folder_changes)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_close)

        edit_menu = self._header_menu("Edit")
        edit_menu.add_command(label="Undo", accelerator="Ctrl+Z", command=self.undo)
        edit_menu.add_command(label="Redo", accelerator="Ctrl+Y", command=self.redo)
        edit_menu.add_separator()
        edit_menu.add_command(label="Find…", accelerator="Ctrl+F", command=self.find_in_editor)
        edit_menu.add_command(
            label="Replace…", accelerator="Ctrl+H", command=lambda: self.find_in_editor(replace=True)
        )
        edit_menu.add_command(
            label="Find Next", accelerator="F3", command=lambda: self._active_text_editor().find_next()
        )
        edit_menu.add_command(label="Find in Project…", accelerator="Ctrl+Shift+F", command=self.project_search)
        edit_menu.add_separator()
        edit_menu.add_command(label="Revert All Changes…", command=self.confirm_revert_all)

        module_menu = self._header_menu("Module")
        module_menu.add_command(label="Add Standard Module…", command=lambda: self.add_module(standard=True))
        module_menu.add_command(label="Add Class Module…", command=lambda: self.add_module(standard=False))
        module_menu.add_command(label="Rename…", command=self.rename_module)
        module_menu.add_command(label="Delete", command=self.delete_module)
        module_menu.add_separator()
        module_menu.add_command(label="Import .bas/.cls…", command=self.import_file)
        module_menu.add_command(label="Export Current Module…", command=self.export_module)
        module_menu.add_command(label="Export All…", command=self.export_all)

        ribbon_menu = self._header_menu("Ribbon")
        ribbon_menu.add_command(label="Design Ribbon", accelerator="Ctrl+2", command=self.show_ribbon_design)
        ribbon_menu.add_separator()
        ribbon_menu.add_command(label="Add Button", command=lambda: self.add_ribbon_control("button"))
        ribbon_menu.add_command(
            label="Add Ribbon Button for This Macro…", command=self.add_ribbon_button
        )
        ribbon_menu.add_command(label="Add Menu", command=self.add_ribbon_menu)
        ribbon_menu.add_command(label="Add Group", command=lambda: self.add_ribbon_control("group"))
        ribbon_menu.add_command(label="Add Separator", command=lambda: self.add_ribbon_control("separator"))
        ribbon_menu.add_command(
            label="Import Button from Another Add-in…", command=self.import_ribbon_control
        )
        ribbon_menu.add_separator()
        ribbon_menu.add_command(
            label="Move Control Up", accelerator="Alt+Up", command=lambda: self.move_ribbon_control(-1)
        )
        ribbon_menu.add_command(
            label="Move Control Down",
            accelerator="Alt+Down",
            command=lambda: self.move_ribbon_control(1),
        )
        ribbon_menu.add_command(
            label="Group Selected into Menu…", command=self.group_selection_into_menu
        )
        ribbon_menu.add_command(
            label="Set Icon from Clipboard", command=self.set_ribbon_icon_from_clipboard
        )
        ribbon_menu.add_command(label="Set Icon from File…", command=self.set_ribbon_icon_from_file)
        ribbon_menu.add_separator()
        ribbon_menu.add_command(
            label="Go to Macro / Button", accelerator="F12", command=self.goto_callback_under_cursor
        )
        ribbon_menu.add_command(label="Ribbon Buttons…", command=self.show_ribbon_buttons)
        ribbon_menu.add_command(label="Check Ribbon Callbacks", command=self.check_ribbon_callbacks)
        ribbon_menu.add_separator()
        ribbon_menu.add_command(label="Create Ribbon…", command=self.create_ribbon)
        ribbon_menu.add_command(label="Repair Custom Icons", command=self.repair_ribbon_icons)
        ribbon_menu.add_command(label="Remove Unused Icons", command=self.remove_unused_ribbon_icons)

        view_menu = self._header_menu("View")
        view_menu.add_command(
            label="Command Palette…", accelerator="Ctrl+K", command=lambda: self.open_command_palette()
        )
        view_menu.add_separator()
        for index, label in enumerate(("Code", "Ribbon", "UserForms", "Package XML")):
            view_menu.add_command(
                label=label, accelerator=f"Ctrl+{index + 1}", command=lambda i=index: self._show_view(i)
            )
        view_menu.add_separator()
        self.theme_var = tk.StringVar(master=self.root, value=theme.palette().name)
        view_menu.add_radiobutton(
            label="Dark Theme", value="dark", variable=self.theme_var, command=self._on_theme_chosen
        )
        view_menu.add_radiobutton(
            label="Light Theme", value="light", variable=self.theme_var, command=self._on_theme_chosen
        )
        view_menu.add_separator()
        view_menu.add_command(label="Zoom In", accelerator="Ctrl+=", command=lambda: self.zoom(1))
        view_menu.add_command(label="Zoom Out", accelerator="Ctrl+-", command=lambda: self.zoom(-1))
        view_menu.add_command(label="Reset Zoom", accelerator="Ctrl+0", command=lambda: self.zoom(0))

        help_menu = self._header_menu("Help")
        help_menu.add_command(label="About", command=self.show_about)
        help_menu.add_command(label="Copy Diagnostic Report", command=self.copy_last_diagnostic)
        self._last_result = None

    def _build_toolbar(self) -> None:
        """File identity, the command box and the save actions, in the header."""
        bar = self.header
        colors = theme.palette()
        self.save_btn = ttk.Button(
            bar, text="Save   Ctrl+S", style="Accent.TButton", command=self.save, state="disabled", takefocus=False
        )
        self.review_btn = ttk.Button(
            bar, text="", style="Changes.TButton", command=self.review_changes, state="disabled", takefocus=False
        )
        Tooltip(self.review_btn, "Review changes")
        self.file_label = ttk.Label(bar, text="", style="Title.TLabel")
        self.file_label.pack(side="left", padx=(14, 0))
        self.path_label = ttk.Label(bar, text="", style="Path.TLabel")
        # Whether the Office host that would load the file is running.
        self.host_pill = tk.Frame(
            bar,
            background=colors.bg0,
            highlightthickness=1,
            highlightbackground=colors.border_strong,
            highlightcolor=colors.border_strong,
        )
        self.host_dot = tk.Label(
            self.host_pill, text="●", font=theme.SMALL_FONT, background=colors.bg0, foreground=colors.ok
        )
        self.host_dot.pack(side="left", padx=(7, 3))
        self.host_text = tk.Label(
            self.host_pill, text="", font=theme.SMALL_FONT, background=colors.bg0, foreground=colors.muted
        )
        self.host_text.pack(side="left", padx=(0, 8), pady=1)
        self.command_btn = ttk.Button(
            bar,
            text="Search modules, code or commands" + " " * 8 + "Ctrl+K",
            style="Command.TButton",
            width=42,
            command=lambda: self.open_command_palette(),
            takefocus=False,
        )
        # Last: when the header is narrow, the command box gives way first.
        self.command_btn.pack(side="left", expand=True, padx=12)
        self._command_text = str(self.command_btn.cget("text"))
        self.file_tip = Tooltip(self.file_label, "")
        bar.bind("<Configure>", lambda _e: self._fit_header(), add=True)

    def _fit_header(self) -> None:
        """Like the mockup on small screens: drop the folder path, then shrink the command box."""
        width = self.header.winfo_width()
        if self.draft is not None and width >= HEADER_PATH_MIN_WIDTH:
            if not self.path_label.winfo_manager() and self.save_btn.winfo_manager():
                self.path_label.pack(side="left", padx=(8, 0), after=self.file_label)
        elif self.path_label.winfo_manager():
            self.path_label.pack_forget()
        compact = width < HEADER_COMMAND_MIN_WIDTH
        self.command_btn.configure(
            text=icon("search") if compact else self._command_text,
            width=0 if compact else 42,
            style="CommandIcon.TButton" if compact else "Command.TButton",
        )

    def _show_header_actions(self, show: bool) -> None:
        """Path, host state, changes and Save appear once a file is open."""
        right = (self.save_btn, self.review_btn)
        if not show:
            for widget in (*right, self.path_label, self.host_pill):
                widget.pack_forget()
        elif not self.save_btn.winfo_manager():
            # Packed ahead of the labels: when the header is narrow the file name gives way, not Save.
            self.save_btn.pack(side="right", padx=(6, 0), before=self.file_label)
            self.review_btn.pack(side="right", padx=(6, 0), before=self.file_label)
            self.host_pill.pack(side="left", padx=(10, 0), after=self.file_label)
            self._fit_header()

    def _build_panes(self) -> None:
        # The activity rail picks the view; the notebook keeps the pages but shows no tabs.
        self.rail = ttk.Frame(self.root, style="Chrome.TFrame", padding=(0, 6, 0, 6))
        self.rail.pack(side="left", fill="y")
        self.rail_line = ttk.Frame(self.root, style="Line.TFrame", width=1)
        self.rail_line.pack(side="left", fill="y")
        self.editor_notebook = ttk.Notebook(self.root, style="Hidden.TNotebook")
        self.vba_tab = ttk.Frame(self.editor_notebook)
        self._build_vba_tab(self.vba_tab)
        self.ribbon_tab = ttk.Frame(self.editor_notebook)
        self.ribbon_designer = RibbonDesigner(self.ribbon_tab, self)
        self.ribbon_designer.pack(fill="both", expand=True)
        self.forms_tab = ttk.Frame(self.editor_notebook)
        self._build_forms_tab(self.forms_tab)
        self.xml_tab = ttk.Frame(self.editor_notebook)
        self._build_xml_tab(self.xml_tab)
        self.editor_notebook.add(self.vba_tab, text="Code")
        self.editor_notebook.add(self.ribbon_tab, text="Ribbon")
        self.editor_notebook.add(self.forms_tab, text="UserForms")
        self.editor_notebook.add(self.xml_tab, text="Package XML")
        self.editor_notebook.pack(fill="both", expand=True)
        self.editor_notebook.bind("<<NotebookTabChanged>>", lambda _e: self._on_view_changed(), add=True)
        self._build_rail()
        self._build_welcome()
        self.banner = ttk.Label(self.root, text="", style="Banner.TLabel", padding=(14, 7))
        # Shown while a save waits for the Office host to release the file.
        colors = theme.palette()
        self.wait_bar = ttk.Frame(self.root, padding=(14, 6, 10, 6), style="Banner.TFrame")
        self.wait_spinner = Spinner(self.wait_bar, background=colors.banner_bg, color=colors.warn)
        self.wait_spinner.pack(side="left", padx=(0, 10))
        self.wait_title = ttk.Label(self.wait_bar, text="", style="BannerBold.TLabel")
        self.wait_title.pack(side="left")
        self.wait_label = ttk.Label(self.wait_bar, text="", style="Banner.TLabel")
        self.wait_label.pack(side="left", padx=(6, 0))
        self.wait_cancel_btn = ttk.Button(
            self.wait_bar, text="Cancel auto-save", style="Banner.TButton", command=self.cancel_queued_save
        )
        self.wait_cancel_btn.pack(side="right")
        self.wait_sub = ttk.Label(
            self.wait_bar,
            text="Keep this window open · edits made meanwhile are reviewed again",
            style="BannerSub.TLabel",
        )
        self.wait_sub.pack(side="left", padx=(14, 0))
        self._install_ribbon_actions()

    def _build_rail(self) -> None:
        # In notebook page order: the rail button at index i shows page i.
        views = (
            ("code", "Code  ·  modules and macros"),
            ("ribbon", "Ribbon  ·  tabs, groups and buttons"),
            ("form", "UserForms"),
            ("xml", "Package XML  ·  ribbon source and other parts"),
        )
        self.rail_buttons: list[RailButton] = []
        for index, (name, tip) in enumerate(views):
            button = RailButton(self.rail, icon(name), tip, lambda i=index: self._show_view(i))
            button.pack(side="top", pady=1)
            self.rail_buttons.append(button)
        self._rail_extra = [
            RailButton(self.rail, icon("search"), "Find in project  Ctrl+Shift+F", self.project_search),
            RailButton(self.rail, icon("history"), "Backups", self.backup_browser),
        ]
        for button in self._rail_extra:
            button.pack(side="top", pady=1)
        settings = RailButton(self.rail, icon("gear"), "Settings", lambda: self._post_settings_menu(settings))
        settings.pack(side="bottom", pady=1)
        self._rail_extra.append(settings)
        menu = self.settings_menu = tk.Menu(self.rail, tearoff=0)
        self._menus.append(menu)
        menu.add_radiobutton(label="Dark Theme", value="dark", variable=self.theme_var, command=self._on_theme_chosen)
        menu.add_radiobutton(
            label="Light Theme", value="light", variable=self.theme_var, command=self._on_theme_chosen
        )
        menu.add_separator()
        menu.add_checkbutton(
            label="Back Up Automatically on Save", variable=self.auto_backup_var, command=self._on_auto_backup_toggled
        )
        menu.add_separator()
        menu.add_command(label="Zoom In", accelerator="Ctrl+=", command=lambda: self.zoom(1))
        menu.add_command(label="Zoom Out", accelerator="Ctrl+-", command=lambda: self.zoom(-1))
        menu.add_command(label="Reset Zoom", accelerator="Ctrl+0", command=lambda: self.zoom(0))
        self._sync_rail()

    def _post_settings_menu(self, button) -> None:
        self.settings_menu.tk_popup(button.winfo_rootx() + button.winfo_width(), button.winfo_rooty())

    def _show_view(self, index: int) -> None:
        if self.draft is None:
            return
        self.editor_notebook.select(index)
        self._on_view_changed()

    def _current_view(self) -> str:
        try:
            return str(self.editor_notebook.select())
        except tk.TclError:
            return ""

    def _designer_shown(self) -> bool:
        return self._current_view() == str(self.ribbon_tab)

    def _on_view_changed(self) -> None:
        """A view was picked: bring it up to date with the draft."""
        view = self._current_view()
        if view == self._shown_view:
            self._sync_rail()
            return
        self._shown_view = view
        if view == str(self.ribbon_tab):
            self._flush_all_editors()
            self._ribbon_surface = "design"
            self.ribbon_designer.refresh()
        elif view == str(self.xml_tab) and rs.is_ribbon_part(self.current_xml_part_path or ""):
            self._ribbon_surface = "xml"
        self._sync_rail()

    def _sync_rail(self) -> None:
        try:
            current = self.editor_notebook.index("current")
        except tk.TclError:
            current = 0
        for index, button in enumerate(getattr(self, "rail_buttons", [])):
            button.set_active(index == current)
        if hasattr(self, "cursor_label"):  # the status bar is built after the rail
            self._update_cursor_label()

    def _build_sidebar(
        self,
        parent,
        title: str,
        *,
        add=None,
        add_tip: str = "",
        filter_var: tk.StringVar | None = None,
    ):
        """Sidebar column: heading, optional filter box and a tree."""
        side = Sidebar(parent, title, add=add, add_tip=add_tip, filter_var=filter_var)
        self._sidebars.append(side)
        return side, side.tree

    def _crumb_bar(self, parent):
        bar = ttk.Frame(parent, style="Editor.TFrame", padding=(14, 4, 8, 4))
        bar.pack(fill="x", side="top")
        ttk.Frame(parent, style="Line.TFrame", height=1).pack(fill="x", side="top")
        return bar

    def _build_welcome(self) -> None:
        """What a first-time user sees until a file is open: one clear thing to do."""
        self.welcome = ttk.Frame(self.root, style="Editor.TFrame")
        inner = ttk.Frame(self.welcome, style="Editor.TFrame")
        inner.place(relx=0.5, rely=0.44, anchor="center")
        ttk.Label(inner, text="V", style="LogoBig.TLabel").pack(anchor="w")
        ttk.Label(inner, text="Open an add-in", style="EditorHero.TLabel").pack(anchor="w", pady=(18, 0))
        # Text set by _show_welcome: it depends on the backup setting.
        self.welcome_note = ttk.Label(inner, style="EditorMuted.TLabel", justify="left", wraplength=540)
        self.welcome_note.pack(anchor="w", pady=(4, 20))
        colors = theme.palette()
        self.welcome_card = tk.Frame(
            inner, background=colors.bg2, highlightthickness=1, highlightbackground=colors.border_strong
        )
        self.welcome_card.pack(fill="x")
        card = ttk.Frame(self.welcome_card, style="Editor.TFrame", padding=(26, 22))
        card.pack(fill="x")
        ttk.Label(
            card, text="Open an .xlam, .ppam or .pptm file, or start a new one", style="EditorFaint.TLabel"
        ).pack()
        buttons = ttk.Frame(card, style="Editor.TFrame")
        buttons.pack(pady=(12, 0))
        ttk.Button(
            buttons, text="Open file   Ctrl+O", style="Accent.TButton", padding=(16, 6), command=self.open_file
        ).pack(side="left")
        ttk.Button(buttons, text="New add-in   Ctrl+N", style="EditorGhost.TButton", command=self.new_file).pack(
            side="left", padx=(8, 0)
        )
        self.welcome_recent = ttk.Frame(inner, style="Editor.TFrame", width=540)
        self.welcome_recent.pack(fill="x", pady=(24, 0))
        self.welcome.place(in_=self.editor_notebook, x=0, y=0, relwidth=1, relheight=1)

    def _show_welcome(self, show: bool) -> None:
        if not show:
            self.welcome.place_forget()
            if not self.rail.winfo_manager():
                self.rail.pack(side="left", fill="y", before=self.editor_notebook)
                self.rail_line.pack(side="left", fill="y", before=self.editor_notebook)
            return
        self.rail.pack_forget()
        self.rail_line.pack_forget()
        backup_note = " The original is kept as a backup when you save." if self.settings.auto_backup else ""
        self.welcome_note.config(
            text="Edit VBA and ribbon XML inside .xlam, .ppam and .pptm files — in place, "
            "without the VBA editor. Office does not need to be installed or running." + backup_note
        )
        for child in self.welcome_recent.winfo_children():
            child.destroy()
        recent = self._recent_files()
        colors = theme.palette()
        if recent:
            ttk.Label(self.welcome_recent, text="RECENT", style="EditorCaption.TLabel").pack(
                anchor="w", pady=(0, 6), padx=(10, 0)
            )
        for path in recent[:5]:
            row = tk.Frame(self.welcome_recent, background=colors.bg2)
            row.pack(fill="x")
            tag_bg, tag_fg = colors.file_tag(path.suffix)
            tag = tk.Label(
                row,
                text=path.suffix.lstrip(".").upper() or "FILE",
                font=theme.TAG_FONT,
                width=6,
                background=tag_bg,
                foreground=tag_fg,
                pady=2,
            )
            tag.pack(side="left", padx=(10, 12), pady=8)
            name = tk.Label(row, text=path.name, font=theme.BOLD_FONT, background=colors.bg2, foreground=colors.text)
            name.pack(side="left")
            when = tk.Label(
                row, text=_when(path), font=theme.SMALL_FONT, background=colors.bg2, foreground=colors.faint
            )
            when.pack(side="right", padx=(12, 10))
            where = tk.Label(
                row,
                text=_shorten(str(path.parent), 44),
                font=theme.SMALL_FONT,
                background=colors.bg2,
                foreground=colors.faint,
                anchor="w",
            )
            where.pack(side="left", padx=(12, 0), fill="x", expand=True)
            hover_row([row, name, where, when], colors.bg2, colors.bg3, lambda p=path: self.open_recent(p))
        self.welcome.place(in_=self.editor_notebook, x=0, y=0, relwidth=1, relheight=1)
        self.welcome.lift()

    # -- recent files ----------------------------------------------------------

    def _recent_files(self) -> list[Path]:
        return [Path(p) for p in self.settings.recent_files if isinstance(p, str) and p]

    def _remember_recent(self, path: Path) -> None:
        entry = str(path)
        kept = [p for p in self.settings.recent_files if isinstance(p, str) and p.casefold() != entry.casefold()]
        self.settings.recent_files = [entry, *kept][:MAX_RECENT_FILES]
        save_settings(self.settings, self._settings_path)

    def _forget_recent(self, path: Path) -> None:
        entry = str(path).casefold()
        self.settings.recent_files = [
            p for p in self.settings.recent_files if isinstance(p, str) and p.casefold() != entry
        ]
        save_settings(self.settings, self._settings_path)

    def _fill_recent_menu(self) -> None:
        menu = self.recent_menu
        menu.delete(0, "end")
        recent = self._recent_files()
        for path in recent:
            menu.add_command(label=str(path), command=lambda p=path: self.open_recent(p))
        if not recent:
            menu.add_command(label="No recent files", state="disabled")

    def open_recent(self, path: Path) -> None:
        if not path.exists():
            self._forget_recent(path)
            messagebox.showinfo(APP_NAME, f"{path.name} is no longer at\n{path.parent}")
            if self.draft is None:
                self._show_welcome(True)
            return
        self.load_path(path)

    # -- view ------------------------------------------------------------------

    def zoom(self, step: int) -> None:
        """Grow (1), shrink (-1) or reset (0) the code font in every editor."""
        size = theme.CODE_SIZE if step == 0 else self.settings.editor_font_size + step
        size = min(max(size, MIN_CODE_SIZE), MAX_CODE_SIZE)
        if size == self.settings.editor_font_size:
            return
        self.settings.editor_font_size = size
        self.root.tk.call("font", "configure", theme.code_font(self.root), "-size", size)
        for editor in (self.editor, self.xml_editor):
            editor._redraw_line_numbers()
        save_settings(self.settings, self._settings_path)
        self._update_status_cells()
        self.status.config(text=f"Code font size {size}")

    def _on_theme_chosen(self) -> None:
        name = self.theme_var.get()
        if name == self.settings.theme:
            return
        self.settings.theme = name
        save_settings(self.settings, self._settings_path)
        self._apply_theme()

    def _apply_theme(self) -> None:
        """Restyle the open window: ttk widgets follow their styles; the rest is set here."""
        colors = theme.apply_theme(self.root, self.settings.theme, self.settings.editor_font_size)
        menu_colors = theme.menu_options(colors)
        for menu in self._menus + [self.editor.context_menu.menu, self.xml_editor.context_menu.menu]:
            menu.configure(**menu_colors)
        for editor in (self.editor, self.xml_editor):
            editor.apply_palette()
        for tree in (self.tree, self.outline, self.xml_tree, self.forms_tree):
            self._configure_tree_tags(tree)
        for button in self.rail_buttons + self._rail_extra:
            button.apply_palette()
        for side in self._sidebars:
            side.apply_palette()
        self.host_pill.configure(
            background=colors.bg0, highlightbackground=colors.border_strong, highlightcolor=colors.border_strong
        )
        self.host_text.configure(background=colors.bg0)
        self.host_dot.configure(background=colors.bg0)
        self.welcome_card.configure(background=colors.bg2, highlightbackground=colors.border_strong)
        self.ribbon_designer.apply_palette()
        palette_window = self.command_palette.window
        if palette_window is not None:
            palette_window.destroy()  # rebuilt with the new colours when next opened
            self.command_palette.window = None
        self.toast.frame.destroy()
        self.toast = Toast(self.root)
        self.wait_spinner.configure(background=colors.banner_bg)
        theme.set_title_bar(self.root)
        self._refresh_state()
        if self.draft is None:
            self._show_welcome(True)

    def _configure_tree_tags(self, tree: ttk.Treeview) -> None:
        colors = theme.palette()
        tree.tag_configure("group", foreground=colors.faint, font=theme.SMALL_FONT)
        tree.tag_configure("dirty", foreground=colors.warn)
        tree.tag_configure("new", foreground=colors.ok)
        tree.tag_configure("renamed", foreground=colors.info)
        tree.tag_configure("locked", foreground=colors.muted)
        tree.tag_configure("readonly", foreground=colors.warn)

    def _install_ribbon_actions(self) -> None:
        xml_menu = self.xml_editor.context_menu
        xml_menu.add_extra(
            "Go to Macro",
            self.goto_callback_under_cursor,
            lambda: self._xml_cursor_callback() is not None,
        )
        xml_menu.add_extra(
            "Add Ribbon Menu Here…",
            self.add_ribbon_menu,
            lambda: rs.is_ribbon_part(self.current_xml_part_path or ""),
        )
        xml_menu.add_extra(
            "Move Up (Alt+Up)", lambda: self.move_ribbon_control(-1), self._ribbon_edit_enabled
        )
        xml_menu.add_extra(
            "Move Down (Alt+Down)", lambda: self.move_ribbon_control(1), self._ribbon_edit_enabled
        )
        xml_menu.add_extra(
            "Group Selected into Menu…", self.group_selection_into_menu, self._ribbon_edit_enabled
        )
        xml_menu.add_extra(
            "Set Icon from Clipboard", self.set_ribbon_icon_from_clipboard, self._ribbon_edit_enabled
        )
        xml_menu.add_extra(
            "Set Icon from File…", self.set_ribbon_icon_from_file, self._ribbon_edit_enabled
        )
        for key, direction in (("<Alt-Up>", -1), ("<Alt-Down>", 1)):
            self.xml_editor.text.bind(
                key, lambda _e, d=direction: (self.move_ribbon_control(d), "break")[1]
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
        self.tree_filter = tk.StringVar(master=self.root)
        self.add_module_menu = tk.Menu(self.root, tearoff=0)
        self._menus.append(self.add_module_menu)
        self.add_module_menu.add_command(label="Standard Module…", command=lambda: self.add_module(standard=True))
        self.add_module_menu.add_command(label="Class Module…", command=lambda: self.add_module(standard=False))
        self.add_module_menu.add_separator()
        self.add_module_menu.add_command(label="Import .bas/.cls…", command=self.import_file)
        # Two lists share the sidebar: the modules, and the procedures of the open one.
        side = ttk.Frame(panes, width=264)
        side.pack_propagate(False)
        lists = ttk.Panedwindow(side, orient="vertical")
        lists.pack(fill="both", expand=True)
        modules, self.tree = self._build_sidebar(
            lists,
            "MODULES",
            add=self._post_add_module_menu,
            add_tip="Add module",
            filter_var=self.tree_filter,
        )
        self.tree.bind("<<TreeviewSelect>>", self.on_module_selected)
        self.tree_filter.trace_add("write", lambda *_a: self._refresh_tree())
        lists.add(modules, weight=3)
        outline, self.outline = self._build_sidebar(
            lists, "PROCEDURES", add=self.add_procedure, add_tip="New Sub in this module"
        )
        self.outline.column("badge", width=44, anchor="e")
        self.outline.bind("<<TreeviewSelect>>", self._on_outline_selected)
        self.outline.bind("<Button-3>", self._on_outline_right_click)
        self.outline_menu = tk.Menu(self.root, tearoff=0)
        self._menus.append(self.outline_menu)
        self.outline_menu.add_command(label="Add Ribbon Button for This Macro…", command=self.add_ribbon_button)
        self.outline_menu.add_command(label="Show Its Ribbon Button", command=self.show_buttons_for_macro)
        lists.add(outline, weight=2)
        self._outline_rows: dict[str, int] = {}
        self._outline_signature: tuple | None = None
        self._outline_after = None
        self._outline_idle = None
        self._outline_selecting = False
        panes.add(side, weight=0)

        main = ttk.Frame(panes, style="Editor.TFrame")
        bar = self._crumb_bar(main)
        self.crumb_group = ttk.Label(bar, text="", style="Crumb.TLabel")
        self.crumb_group.pack(side="left")
        ttk.Label(bar, text=icon("right"), style="CrumbSep.TLabel").pack(side="left", padx=6)
        self.crumb_module = ttk.Label(bar, text="", style="CrumbStrong.TLabel")
        self.crumb_module.pack(side="left")
        self.crumb_proc_sep = ttk.Label(bar, text=icon("right"), style="CrumbSep.TLabel")
        self.crumb_proc_sep.pack(side="left", padx=(6, 0))
        self.proc_button = ttk.Menubutton(bar, text="", style="Crumb.TMenubutton", takefocus=False)
        self.proc_button.pack(side="left")
        self.proc_menu = tk.Menu(self.proc_button, tearoff=0, postcommand=self._fill_procedure_menu)
        self.proc_button.configure(menu=self.proc_menu)
        self._menus.append(self.proc_menu)
        self.crumb_note = ttk.Label(bar, text="", style="CrumbMuted.TLabel")
        self.crumb_note.pack(side="right", padx=(8, 4))
        self.editor = CodeEditor(main)
        self.editor.pack(fill="both", expand=True)
        panes.add(main, weight=1)

    def _post_add_module_menu(self) -> None:
        if self.draft is None:
            return
        x, y = self.root.winfo_pointerxy()
        self.add_module_menu.tk_popup(x, y)

    def _build_xml_tab(self, parent) -> None:
        panes = ttk.Panedwindow(parent, orient="horizontal")
        panes.pack(fill="both", expand=True)
        self.xml_filter = tk.StringVar(master=self.root)
        side, self.xml_tree = self._build_sidebar(panes, "PACKAGE PARTS", filter_var=self.xml_filter)
        self.xml_tree.bind("<<TreeviewSelect>>", self.on_xml_part_selected)
        self.xml_filter.trace_add("write", lambda *_a: self._refresh_xml_tree())
        panes.add(side, weight=0)

        main = ttk.Frame(panes, style="Editor.TFrame")
        bar = self._crumb_bar(main)
        self.xml_crumb_folder = ttk.Label(bar, text="", style="Crumb.TLabel")
        self.xml_crumb_folder.pack(side="left")
        self.xml_crumb_sep = ttk.Label(bar, text=icon("right"), style="CrumbSep.TLabel")
        self.xml_crumb_sep.pack(side="left", padx=6)
        self.xml_crumb_file = ttk.Label(bar, text="", style="CrumbStrong.TLabel")
        self.xml_crumb_file.pack(side="left")
        self.validate_xml_btn = ttk.Button(
            bar, text="Validate XML", style="EditorText.TButton", command=self._validate_selected_xml,
            state="disabled", takefocus=False,
        )
        self.validate_xml_btn.pack(side="right")
        self.xml_crumb_note = ttk.Label(bar, text="", style="CrumbMuted.TLabel")
        self.xml_crumb_note.pack(side="right", padx=(8, 8))
        # Shown for ribbon XML only: the same control in the Ribbon view.
        self.design_btn = ttk.Button(
            bar, text="Design view", style="EditorText.TButton", command=self.show_ribbon_design, takefocus=False
        )
        self.xml_editor = XmlEditor(main)
        self.xml_editor.pack(fill="both", expand=True)
        self.xml_editor.set_validate_action(self._validate_selected_xml)
        panes.add(main, weight=1)

        self.current_xml_part_path: str | None = None

    def _build_forms_tab(self, parent) -> None:
        """Read-only UserForm layouts, as stored in the file when it was opened."""
        panes = ttk.Panedwindow(parent, orient="horizontal")
        panes.pack(fill="both", expand=True)
        side, self.forms_tree = self._build_sidebar(panes, "USERFORMS")
        self.forms_tree.column("badge", width=0, minwidth=0)
        self.forms_tree.bind("<<TreeviewSelect>>", self.on_form_item_selected)
        panes.add(side, weight=0)

        main = ttk.Frame(panes, style="Editor.TFrame")
        bar = self._crumb_bar(main)
        ttk.Label(bar, text=icon("lock"), style="CrumbSep.TLabel").pack(side="left", padx=(0, 6))
        self.form_heading = ttk.Label(bar, text="", style="CrumbStrong.TLabel", anchor="w")
        self.form_heading.pack(side="left", fill="x", expand=True)
        self.form_code_btn = ttk.Button(
            bar, text="Go to code", style="EditorText.TButton", command=self.goto_form_code,
            state="disabled", takefocus=False,
        )
        self.form_code_btn.pack(side="right")
        ttk.Label(
            main,
            text=(
                "Layout is read-only. Only properties that differ from the control's "
                "default are stored in the file, so only those are listed."
            ),
            style="EditorInfo.TLabel",
            padding=(14, 8),
        ).pack(fill="x", side="bottom")
        props = ttk.Frame(main, style="Editor.TFrame", padding=(6, 6, 0, 0))
        props.pack(fill="both", expand=True)
        self.form_props = ttk.Treeview(
            props, columns=("value",), show="tree headings", selectmode="browse", style="Editor.Treeview"
        )
        self.form_props.heading("#0", text="Property", anchor="w")
        self.form_props.heading("value", text="Value", anchor="w")
        self.form_props.column("#0", width=220, stretch=False)
        props_sb = ttk.Scrollbar(props, orient="vertical", command=self.form_props.yview)
        self.form_props.configure(yscrollcommand=theme.autohide(props_sb, side="right", fill="y"))
        self.form_props.pack(side="left", fill="both", expand=True)
        panes.add(main, weight=1)

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
            self.form_heading.config(text="This project has no UserForms")
            return
        self.form_heading.config(text="Select a form or control to see its stored properties")
        deleted = {
            (m.origin_name or m.current_name).casefold() for m in draft.deleted_original_modules()
        }
        orphans = draft.baseline.orphan_form_names()
        for form in draft.baseline.forms:
            label = form.name
            if form.problem:
                label += " [unreadable]"
            elif form.name.casefold() in orphans:
                label += " [layout only — no code module]"
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
            heading = f"{item.name}  ·  {item.control_count()} control(s)"
            if item.problem:
                heading += f"  ·  the layout could not be read: {item.problem}"
        else:
            heading = f"{form_name}  ›  {item.name}  ({item.kind.rsplit('.', 1)[-1]})"
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
        self.statusbar = ttk.Frame(self.root, style="Chrome.TFrame", padding=(6, 0))
        # Packed ahead of the rail and the pages, so a small window shrinks the editor, not
        # the status bar, and the bar runs the whole width under the rail.
        self.statusbar.pack(fill="x", side="bottom", before=self.rail)
        ttk.Frame(self.root, style="Line.TFrame", height=1).pack(fill="x", side="bottom", before=self.rail)

        def cell(text: str = "", command=None, tip: str = "") -> ttk.Label:
            label = ttk.Label(self.statusbar, text=text, padding=(8, 4), style="Status.TLabel")
            label.pack(side="right")
            if command is not None:
                label.configure(cursor="hand2")
                label.bind("<Button-1>", lambda _e: command())
            if tip:
                Tooltip(label, tip)
            return label

        self.zoom_label = cell(command=lambda: self.zoom(0), tip="Code size · click to reset")
        self.backup_label = cell(command=self._toggle_auto_backup, tip="Click to switch automatic backup on or off")
        self.lang_label = cell()
        self.cursor_label = cell()
        # Whether antivirus has finished with the last saved file.
        self.scan_status = ttk.Label(self.statusbar, text="", padding=(0, 4, 8, 4), anchor="e", style="Status.TLabel")
        self.scan_status.pack(side="right")
        self.scan_icon = ttk.Label(self.statusbar, text="", padding=(8, 4, 4, 4), style="StatusIcon.TLabel")
        self.scan_icon.pack(side="right")
        self.status = ttk.Label(self.statusbar, text="Ready", padding=(8, 4), anchor="w", style="Status.TLabel")
        self.status.pack(fill="x", side="left", expand=True)
        for editor in (self.editor, self.xml_editor):
            editor.bind("<<CursorMove>>", lambda _e: self._update_cursor_label(), add=True)
        self._update_status_cells()

    def _toggle_auto_backup(self) -> None:
        self.auto_backup_var.set(not self.auto_backup_var.get())
        self._on_auto_backup_toggled()

    def _update_status_cells(self) -> None:
        open_file = self.draft is not None
        self.backup_label.config(
            text=("Backup on save" if self.settings.auto_backup else "No backup on save") if open_file else ""
        )
        percent = round(100 * self.settings.editor_font_size / theme.CODE_SIZE)
        self.zoom_label.config(text=f"{percent}%" if open_file else "")

    def _update_cursor_label(self) -> None:
        try:
            tab = str(self.editor_notebook.select())
        except tk.TclError:
            return
        if self.draft is None:
            self.cursor_label.config(text="")
            self.lang_label.config(text="")
            return
        if tab in (str(self.forms_tab), str(self.ribbon_tab)):
            self.cursor_label.config(text="")
            self.lang_label.config(text="UserForm" if tab == str(self.forms_tab) else "Ribbon")
            return
        editor = self._active_text_editor()
        if (editor is self.editor and self.current_module_id is None) or (
            editor is self.xml_editor and self.current_xml_part_path is None
        ):
            self.cursor_label.config(text="")
            self.lang_label.config(text="")
            return
        line, column = editor.cursor_position().split(".")
        self.cursor_label.config(text=f"Ln {line}, Col {int(column) + 1}")
        self.lang_label.config(text="VBA" if editor is self.editor else "XML")
        if self._follow_after is None:
            self._follow_after = self.root.after(CURSOR_FOLLOW_MS, self._follow_cursor)

    def _follow_cursor(self) -> None:
        """Breadcrumb and outline follow the VBA cursor, a moment after it moves."""
        self._follow_after = None
        if self.draft is None:
            return
        if self._active_text_editor() is self.editor:
            self._update_procedure_crumb()
            self._follow_outline()

    # -- breadcrumbs -----------------------------------------------------------

    def _module_group(self, mod: ModuleDraft) -> str:
        if mod.kind == ModuleDisplayKind.CLASS:
            return GROUP_CLASSES
        if mod.pyopenvba_kind == "standard":
            return GROUP_MODULES
        if mod.project_item_kind == "userform":
            return GROUP_FORMS
        return self._document_group()

    def _document_group(self) -> str:
        if self.draft is not None and self.draft.baseline.path.suffix.lower() == ".xlam":
            return "Excel objects"
        return "Document modules"

    def _update_crumbs(self) -> None:
        draft = self.draft
        self._refresh_outline()
        mod = draft.module_by_id(self.current_module_id) if draft and self.current_module_id else None
        if mod is None or mod.is_deleted:
            self.crumb_group.config(text=draft.baseline.path.name if draft else "")
            self.crumb_module.config(text="No module open" if draft else "")
            self.crumb_proc_sep.pack_forget()
            self.proc_button.pack_forget()
            self.crumb_note.config(text="", style="CrumbMuted.TLabel")
        else:
            self.crumb_group.config(text=self._module_group(mod))
            self.crumb_module.config(text=mod.current_name)
            if not self.proc_button.winfo_manager():
                self.crumb_proc_sep.pack(side="left", padx=(6, 0), after=self.crumb_module)
                self.proc_button.pack(side="left", after=self.crumb_proc_sep)
            self._update_procedure_crumb()
            note, style = "", "CrumbMuted.TLabel"
            if draft is not None and draft.baseline.safety.password_protected:
                note, style = "Password-protected project: VBA is read-only", "CrumbWarn.TLabel"
            elif mod.is_new:
                note, style = "New module · not saved yet", "CrumbOk.TLabel"
            elif mod.origin_name and mod.current_name.casefold() != mod.origin_name.casefold():
                note, style = f"Renamed from {mod.origin_name}", "CrumbInfo.TLabel"
            elif not (mod.can_delete or mod.can_rename):
                note = "🔒 Code is editable; this module can't be renamed or deleted"
            self.crumb_note.config(text=note, style=style)
        part = draft.xml_part_by_path(self.current_xml_part_path) if draft and self.current_xml_part_path else None
        if part is None:
            self.xml_crumb_folder.config(text=draft.baseline.path.name if draft else "")
            self.xml_crumb_file.config(text="No part open" if draft else "")
            self.xml_crumb_note.config(text="", style="CrumbMuted.TLabel")
        else:
            folder, _sep, name = part.path.rpartition("/")
            self.xml_crumb_folder.config(text=folder or "package root")
            self.xml_crumb_file.config(text=name)
            if not part.editable:
                self.xml_crumb_note.config(text="Read-only: malformed when opened", style="CrumbWarn.TLabel")
            elif part.is_new:
                self.xml_crumb_note.config(text="New part · not saved yet", style="CrumbOk.TLabel")
            elif part.is_dirty():
                self.xml_crumb_note.config(text="Edited", style="CrumbWarn.TLabel")
            else:
                self.xml_crumb_note.config(text="", style="CrumbMuted.TLabel")

    def _current_procedure(self) -> str | None:
        """Name of the procedure the VBA cursor is in, by a backward scan of the text."""
        text = self.editor.text
        start = "insert lineend"
        for _attempt in range(2):
            try:
                found = text.search(_PROC_EDGE_TCL, start, "1.0", backwards=True, regexp=True, nocase=True)
            except tk.TclError:
                return None
            if not found:
                return None
            line = text.get(f"{found} linestart", f"{found} lineend")
            match = _PROC_DECL.match(line)
            if match:
                return match.group(1)
            # An End line: inside that procedure only when the cursor sits on the End line itself.
            if text.index(f"{found} linestart") != text.index("insert linestart"):
                return None
            start = f"{found} linestart"
        return None

    def _update_procedure_crumb(self) -> None:
        if self.current_module_id is None:
            return
        name = self._current_procedure() or "(General)"
        self.proc_button.config(text=f"{name}  ▾")

    def _fill_procedure_menu(self) -> None:
        menu = self.proc_menu
        menu.delete(0, "end")
        draft = self.draft
        mod = draft.module_by_id(self.current_module_id) if draft and self.current_module_id else None
        if mod is None:
            return
        self._flush_active_vba_editor()
        procedures = parse_procedures(mod.body, mod.id)
        for proc in procedures[:400]:
            menu.add_command(
                label=f"{proc.name}", accelerator=proc.kind,
                command=lambda p=proc: self._open_module_at(mod.id, p.line),
            )
        if not procedures:
            menu.add_command(label="No procedures", state="disabled")

    # -- procedures of the open module ---------------------------------------------

    def _schedule_outline(self) -> None:
        if self._outline_after is not None:
            self.root.after_cancel(self._outline_after)
        self._outline_after = self.root.after(OUTLINE_REFRESH_MS, self._refresh_outline)

    def _linked_macro_names(self) -> set[str]:
        """Casefolded names of the procedures some ribbon control runs."""
        if self.draft is None:
            return set()
        return {entry.callback.procedure_name.casefold() for entry in rs.ribbon_entries(self.draft)}

    def _refresh_outline(self) -> None:
        """List the open module's procedures; a mark shows the ones a ribbon control runs."""
        if self._outline_after is not None:
            try:
                self.root.after_cancel(self._outline_after)
            except tk.TclError:
                pass
            self._outline_after = None
        draft = self.draft
        mod = draft.module_by_id(self.current_module_id) if draft and self.current_module_id else None
        tree = self.outline
        if draft is None or mod is None or mod.is_deleted:
            if self._outline_signature is not None:
                tree.delete(*tree.get_children())
                self._outline_rows.clear()
                self._outline_signature = None
            return
        linked = self._linked_macro_names()
        rows = []
        for proc in parse_procedures(self.editor.get_text(), mod.id):
            name = proc.name.casefold()
            on_ribbon = name in linked or name + "callback" in linked
            mark = icon("ribbon") if on_ribbon else OUTLINE_MARKS.get(proc.kind, "")
            rows.append((proc.name, proc.line, mark, proc.visibility.casefold() == "private"))
        signature = (mod.id, tuple(rows))
        if signature != self._outline_signature:
            self._outline_signature = signature
            top = tree.yview()[0]
            self._outline_selecting = True
            try:
                tree.delete(*tree.get_children())
                self._outline_rows.clear()
                for name, line, mark, private in rows:
                    iid = tree.insert("", "end", text=name, values=(mark,), tags=("locked",) if private else ())
                    self._outline_rows[iid] = line
                tree.yview_moveto(top)
            finally:
                self._outline_selecting = False
        self._follow_outline()

    def _follow_outline(self) -> None:
        """Highlight the procedure the VBA cursor is in."""
        if not self._outline_rows:
            return
        line = int(self.editor.text.index("insert").split(".")[0])
        inside = self._current_procedure() is not None
        current = None
        for iid, start in self._outline_rows.items():
            if start <= line:
                current = iid
        self._outline_selecting = True
        try:
            if current is None or not inside:
                self.outline.selection_set(())
            elif self.outline.selection() != (current,):
                self.outline.selection_set(current)
                self.outline.see(current)
        finally:
            if self._outline_idle is None:
                self._outline_idle = self.root.after_idle(self._end_outline_selecting)

    def _end_outline_selecting(self) -> None:
        self._outline_idle = None
        self._outline_selecting = False

    def _on_outline_selected(self, _event=None) -> None:
        selection = self.outline.selection()
        if self._outline_selecting or not selection or self.current_module_id is None:
            return
        line = self._outline_rows.get(selection[0])
        if line is not None:
            self._open_module_at(self.current_module_id, line)
            self.editor.text.yview(f"{max(line - 2, 1)}.0")

    def _on_outline_right_click(self, event) -> str:
        iid = self.outline.identify_row(event.y)
        if iid:
            self.outline.selection_set(iid)
            self._on_outline_selected()
            self.outline_menu.tk_popup(event.x_root, event.y_root)
        return "break"

    def add_procedure(self) -> None:
        """A new Sub at the end of the open module, with the cursor inside it."""
        draft = self.draft
        mod = draft.module_by_id(self.current_module_id) if draft and self.current_module_id else None
        if draft is None or mod is None or mod.is_deleted:
            return
        if draft.baseline.safety.password_protected:
            messagebox.showwarning(APP_NAME, "This project is password-protected and read-only.")
            return
        name = self._prompt_name("New Sub", "")
        if not name:
            return
        self._flush_all_editors()
        if not rs.is_valid_procedure_name(name):
            messagebox.showerror(APP_NAME, f"'{name}' is not a valid VBA procedure name.")
            return
        if any(p.name.casefold() == name.casefold() for p in parse_procedures(mod.body, mod.id)):
            messagebox.showerror(APP_NAME, f"{mod.current_name} already has a procedure named {name}.")
            return
        new_body, line = rs.append_procedure(mod.body, f"Public Sub {name}()\n    \nEnd Sub\n")

        def mutate() -> None:
            mod.body = new_body

        self._batch_edit(mutate)
        self._open_module_at(mod.id, line + 1, 5)
        self.status.config(text=f"Added Sub {name} to {mod.current_name}.")

    def _bind_shortcuts(self) -> None:
        self.root.bind("<Control-o>", lambda _e: self.open_file())
        self.root.bind("<Control-n>", lambda _e: self.new_file())
        self.root.bind("<Control-s>", lambda _e: self.save())
        self.root.bind("<Control-S>", lambda _e: self.save_copy())
        # Reached only when focus is outside an editor's text; the text widget
        # handles these keys itself and stops them there.
        self.root.bind("<Control-f>", lambda _e: self.find_in_editor())
        self.root.bind("<Control-h>", lambda _e: self.find_in_editor(replace=True))
        self.root.bind("<Control-Shift-F>", lambda _e: self.project_search())
        self.root.bind("<Control-z>", lambda _e: self.undo())
        self.root.bind("<Control-y>", lambda _e: self.redo())
        self.root.bind("<F12>", lambda _e: self.goto_callback_under_cursor())
        self.root.bind("<Control-k>", lambda _e: self.open_command_palette())
        for editor in (self.editor, self.xml_editor):
            # Tk's Text binds Ctrl+K to "delete to end of line"; here it opens the palette.
            editor.text.bind("<Control-k>", lambda _e: self.open_command_palette() or "break")
        for sequence, step in (("<Control-equal>", 1), ("<Control-plus>", 1), ("<Control-minus>", -1), ("<Control-0>", 0)):
            self.root.bind(sequence, lambda _e, step=step: self.zoom(step))
        for index in range(4):
            self.root.bind(f"<Control-Key-{index + 1}>", lambda _e, i=index: self._show_view(i))

    # -- command palette ----------------------------------------------------------

    def open_command_palette(self, prefix: str = "") -> str:
        if not self._saving:
            self._flush_all_editors()
            self.command_palette.open(prefix)
        return "break"

    def _palette_items(self) -> list[PaletteItem]:
        """Every enabled menu command, plus the open file's modules, parts and procedures."""
        items: list[PaletteItem] = []
        for button in self._header_buttons:
            menu = self.root.nametowidget(button.cget("menu"))
            group = str(button.cget("text"))
            last = menu.index("end")
            for index in range(0 if last is None else last + 1):
                kind = menu.type(index)
                if kind not in ("command", "checkbutton", "radiobutton"):
                    continue
                if str(menu.entrycget(index, "state")) == "disabled":
                    continue
                label = str(menu.entrycget(index, "label")).rstrip("…").rstrip(".")
                if label == "Command Palette":
                    continue
                items.append(
                    PaletteItem(
                        COMMAND,
                        label,
                        group,
                        str(menu.entrycget(index, "accelerator")),
                        lambda m=menu, i=index: m.invoke(i),
                    )
                )
        for path in self._recent_files():
            items.append(PaletteItem(COMMAND, f"Open {path.name}", "Recent", "", lambda p=path: self.open_recent(p)))
        draft = self.draft
        if draft is None:
            return items
        for mod in draft.modules:
            if not mod.is_deleted:
                items.append(
                    PaletteItem(
                        MODULE, mod.current_name, self._module_group(mod), "",
                        lambda mid=mod.id: self._open_module_at(mid, 1),
                    )
                )
        for part in draft.xml_parts:
            if part.text is not None:
                items.append(
                    PaletteItem(MODULE, part.path, "XML part", "", lambda p=part.path: self._open_xml_at(p, 1))
                )
        for locations in rs.procedure_index(draft).values():
            for loc in locations:
                items.append(
                    PaletteItem(
                        PROCEDURE, loc.name, f"{loc.module_name} · {loc.kind}", "",
                        lambda loc=loc: self._goto_procedure(loc),
                    )
                )
        return items

    def _goto_procedure(self, loc) -> None:
        module = self.draft.module_by_id(loc.module_id) if self.draft else None
        if module is None:
            return
        lines = module.body.split("\n")
        line_text = lines[loc.line - 1] if 0 < loc.line <= len(lines) else ""
        found = re.search(rf"\b{re.escape(loc.name)}\b", line_text)
        column = found.start() + 1 if found else 1
        self._open_module_at(loc.module_id, loc.line, column, len(loc.name) if found else 0)

    # -- banners / status ----------------------------------------------------

    def _set_banner(self, text: str) -> None:
        if text:
            self.banner.config(text=text, wraplength=max(self.root.winfo_width() - 40, 400))
            self.banner.pack(fill="x", side="top", before=self.statusbar)
        else:
            self.banner.pack_forget()

    def _on_progress(self, stage: str) -> None:
        # Called from the save worker thread; Tk is only touched by the poller.
        self._progress_stage = stage

    def _run_in_background(self, work: Callable[[], _T], label: str = "Saving…") -> _T:
        """Run ``work()`` on a worker thread; keep the UI painted but inert until done.

        Blocks the caller in a nested event loop so callers keep synchronous
        semantics. Parsing or rebuilding a large VBA project takes a moment;
        running it off the Tk thread keeps the window responsive.
        """
        outcome: dict = {}

        def target() -> None:
            try:
                if self._prescan.busy:
                    # Its handle would fail the exclusive-access probe.
                    self._progress_stage = "waiting for the antivirus scan of the last save"
                    self._prescan.settle()
                    self._progress_stage = None
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
        self.status.config(text=label)
        held = self._hold_ui()
        # The same goes for garbage: a closed dialog's Tk variables must not be
        # collected by the worker, so cyclic GC waits until it has finished.
        self._hold_gc()

        def poll() -> None:
            stage = self._progress_stage
            if stage:
                self.status.config(text=f"{label} {stage}")
            if thread.is_alive():
                self.root.after(50, poll)
            else:
                self.root.setvar(done, "1")

        try:
            thread.start()
            self.root.after(50, poll)
            self.root.wait_variable(done)
        finally:
            if thread.is_alive():
                thread.join()
            self._release_gc()
            self._release_ui(held)
            try:
                self.root.globalunsetvar(done)
            except tk.TclError:
                pass
        if "error" in outcome:
            raise outcome["error"]
        return outcome["result"]

    def _hold_gc(self) -> None:
        """Pause cyclic GC while a worker thread runs; holds nest."""
        if self._gc_holds == 0:
            self._gc_was_enabled = gc.isenabled()
            gc.disable()
        self._gc_holds += 1

    def _release_gc(self) -> None:
        self._gc_holds -= 1
        if self._gc_holds == 0 and self._gc_was_enabled:
            gc.enable()

    def _start_prescan(self, path: Path) -> None:
        """Let antivirus scan the file just written now, not when Office loads it."""
        self._prescan.start(path)
        if not self._prescan_holds_gc:
            self._hold_gc()
            self._prescan_holds_gc = True
        if self._prescan_after is not None:
            self.root.after_cancel(self._prescan_after)
        self._poll_prescan()

    def _poll_prescan(self) -> None:
        self._prescan_after = None
        name = self._prescan.path.name if self._prescan.path is not None else ""
        if self._prescan.busy:
            self.scan_status.config(text=f"Antivirus scanning {name}…")
            self.scan_icon.config(text=icon("shield"), style="StatusIcon.TLabel")
            self._prescan_after = self.root.after(200, self._poll_prescan)
            return
        self._release_prescan_gc()
        ready = self._prescan.state == READY
        self.scan_status.config(text=f"{name} ready for Office" if ready else "")
        self.scan_icon.config(text=icon("shield") if ready else "", style="StatusOk.TLabel")

    def _release_prescan_gc(self) -> None:
        if self._prescan_holds_gc:
            self._prescan_holds_gc = False
            self._release_gc()

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
        for button in self._header_buttons + [self.command_btn]:
            try:
                button.configure(state=state)
            except tk.TclError:
                pass

    def _refresh_state(self) -> None:
        draft = self.draft
        self._show_wait_bar()
        self._update_status_cells()
        if draft is None:
            self.file_label.config(text="")
            self.save_btn.config(state="disabled")
            self.review_btn.config(state="disabled")
            self.status.config(text="Open an add-in to begin.")
            self._show_header_actions(False)
            self._show_welcome(True)
            self._update_cursor_label()
            self._update_crumbs()
            return
        self._show_welcome(False)
        self._show_header_actions(True)
        name = draft.baseline.path.name
        n = dirty_count(draft)
        self.file_label.config(text=name)
        self.path_label.config(text=_shorten(str(draft.baseline.path.parent), 40))
        self.file_tip.text = str(draft.baseline.path)
        self.review_btn.config(text=f"●  {n} change{'s' if n != 1 else ''}" if n else "No changes")
        self.save_btn.config(text="Waiting…" if self._queued_save is not None else "Save   Ctrl+S")
        self._update_host_pill()
        self._update_cursor_label()
        self._update_crumbs()
        inplace_ok = (
            not self._saving
            and (self.session is None or (self.session.original_exists and not self.session.conflicts_pending))
        )
        self.save_btn.config(state="normal" if inplace_ok else "disabled")
        self.review_btn.config(state="normal" if n and inplace_ok else "disabled")
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
        elif self._live_host_running and self._queued_save is None:
            self._set_banner(
                f"Close {self._host_label} before saving this add-in."
            )
        else:
            self._set_banner("")

    def _show_wait_bar(self) -> None:
        if self._queued_save is None or self.draft is None:
            self.wait_bar.pack_forget()
            return
        label = wp.corresponding_host_label(self.draft.baseline.path)
        self.wait_title.config(text=f"{label} is open.")
        self.wait_label.config(
            text=f"Close it and {self.draft.baseline.path.name} will be saved automatically."
        )
        self.wait_bar.pack(fill="x", side="top", after=self.header_line)

    def _update_host_pill(self) -> None:
        colors = theme.palette()
        label = self._host_label or "Office"
        if self._live_host_probe_failed:
            text, dot = f"Can't check {label}", colors.faint
        elif self._live_host_running:
            text, dot = f"{label} running", colors.warn
        else:
            text, dot = f"{label} closed", colors.ok
        self.host_text.config(text=text)
        self.host_dot.config(foreground=dot)

    def _refresh_tree(self) -> None:
        draft = self.draft
        self.tree.delete(*self.tree.get_children())
        if draft is None:
            return
        self._refresh_forms()
        query = self.tree_filter.get().strip().casefold()
        groups: dict[str, list] = {self._document_group(): [], GROUP_MODULES: [], GROUP_CLASSES: [], GROUP_FORMS: []}
        for m in draft.modules:
            if m.is_deleted:
                continue
            if query and query not in m.current_name.casefold():
                continue
            badge, tags = "", ()
            label = m.current_name
            if m.is_new:
                badge, tags = "A", ("new",)
            elif m.origin_name and m.current_name.casefold() != m.origin_name.casefold():
                badge, tags = "R", ("renamed",)
                label += f"   (was {m.origin_name})"
            elif m.original_body is not None and m.body != m.original_body:
                badge, tags = "M", ("dirty",)
            locked = not (m.can_delete or m.can_rename or m.is_new)
            if locked:
                tags = tags or ("locked",)
                badge = badge or "🔒"
            groups.setdefault(self._module_group(m), []).append((m.id, label, badge, tags))
        for group, items in groups.items():
            if not items:
                continue
            parent = self.tree.insert("", "end", text=group, open=True, tags=("group",))
            for mid, label, badge, tags in items:
                self.tree.insert(parent, "end", iid=mid, text=label, values=(badge,), tags=tags)
        current = self.current_module_id
        if current and self.tree.exists(current):
            # on_module_selected ignores re-selection of the module already shown.
            self.tree.selection_set(current)
            self.tree.see(current)

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
        mod = draft.module_by_id(mid)
        if mod is None:
            return  # a group heading
        self._flush_all_editors()
        self.current_module_id = mid
        self.editor.set_text(mod.body)
        self.editor.text.config(state="normal")  # body editing allowed; deletion is gated
        self._refresh_state()
        if mod.restriction_reason and not mod.is_new:
            # Newcomers see a lock in the tree; say what it means.
            reason = RESTRICTION_MESSAGES.get(mod.restriction_reason, mod.restriction_reason)
            self.status.config(text=f"{mod.current_name}: {reason} Its code can still be edited.")

    def on_xml_part_selected(self, _event=None) -> None:
        draft = self.draft
        if draft is None:
            return
        selection = self.xml_tree.selection()
        if not selection:
            return
        if not selection[0].startswith("xml::"):
            return  # a group heading
        path = selection[0][len("xml::"):]
        if path == self.current_xml_part_path:
            return  # already shown (programmatic re-selection)
        self._flush_all_editors()
        part = draft.xml_part_by_path(path)
        if part is None:
            return
        self.current_xml_part_path = path
        if rs.is_ribbon_part(path):
            self._ribbon_surface = "xml"
        # Re-enable before replacing widget contents in case the previously
        # selected baseline part was read-only.
        self.xml_editor.text.config(state="normal")
        self.xml_editor.set_text(part.text or "")
        if not part.editable:
            self.xml_editor.text.config(state="disabled")
        self.xml_editor.set_part_info(part)
        self.validate_xml_btn.config(state="normal")
        self._refresh_ribbon_views()
        self._refresh_state()

    def _refresh_xml_tree(self) -> None:
        draft = self.draft
        self.xml_tree.delete(*self.xml_tree.get_children())
        if draft is None:
            return
        query = self.xml_filter.get().strip().casefold()
        groups: dict[str, list] = {"Ribbon": [], "Package": []}
        for part in sorted(draft.xml_parts, key=lambda p: p.path.lower()):
            if query and query not in part.path.casefold():
                continue
            dirty = part.is_dirty()
            warning = "⚠ " if not part.editable else ""
            tags = ("readonly",) if not part.editable else ("new",) if part.is_new else ("dirty",) if dirty else ()
            badge = "A" if part.is_new else "M" if dirty else ""
            group = "Ribbon" if part.path.casefold().startswith("customui") else "Package"
            groups[group].append(("xml::" + part.path, warning + part.path, badge, tags))
        for group, items in groups.items():
            if not items:
                continue
            parent = self.xml_tree.insert("", "end", iid=f"group::{group}", text=group, open=True, tags=("group",))
            for iid, label, badge, tags in items:
                self.xml_tree.insert(parent, "end", iid=iid, text=label, values=(badge,), tags=tags)
        current = "xml::" + (self.current_xml_part_path or "")
        if self.current_xml_part_path and self.xml_tree.exists(current):
            self.xml_tree.selection_set(current)
            self.xml_tree.see(current)

    def _active_text_editor(self):
        if str(self.editor_notebook.select()) == str(self.xml_tab):
            return self.xml_editor
        return self.editor

    def find_in_editor(self, replace: bool = False) -> None:
        if self._designer_shown():
            self.ribbon_designer.focus_filter()  # in the Ribbon view, Ctrl+F filters the controls
            return
        self._active_text_editor().show_find(replace=replace)

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

    def load_path(self, path: Path, *, written=None) -> bool:
        """Open ``path``; ``written`` = (bytes, fingerprint) of a file this editor just made."""
        if self._saving:
            return False
        self._flush_all_editors()

        def work():
            try:
                session = self.session_service.open(path, written=written)
            except (AdapterError, OSError, ValueError):
                return None, self.doc_service.open(path)
            return session, session.draft

        self._saving = True  # keeps autosave and close away while the worker runs
        try:
            session, draft = self._run_in_background(work, f"Opening {path.name}…")
        except (AdapterError, OSError) as exc:
            self._saving = False
            messagebox.showerror(
                APP_NAME,
                str(exc)
                if isinstance(exc, AdapterError)
                else f"{path.name} could not be opened.\n\n{exc}",
            )
            self._refresh_state()
            return False
        finally:
            self._saving = False
        if self.session is not None:
            self.session_service.close(self.session)
        self.session = session
        self.draft = draft
        self._refresh_host_status()
        self.current_module_id = None
        self.current_xml_part_path = None
        self._refresh_tree()
        self._refresh_xml_tree()
        xml_supported = path.suffix.lower() in XML_EDITABLE_EXTENSIONS
        for tab in (self.xml_tab, self.ribbon_tab):
            self.editor_notebook.tab(tab, state="normal" if xml_supported else "disabled")
        if xml_supported:
            self.xml_editor.set_part_info(None)
            self.xml_editor.text.config(state="normal")
            self.xml_editor.set_text("")
            self.validate_xml_btn.config(state="disabled")
        first = next((m.id for m in draft.modules if not m.is_deleted), None)
        if first:
            self.tree.selection_set(first)
        self._ribbon_surface = "design"
        self._refresh_ribbon_views()
        self._refresh_state()
        self._remember_recent(path)
        self.status.config(text=f"Opened {path.name}")
        self._offer_icon_repair()
        return True

    # -- new file / first ribbon -------------------------------------------------------

    def new_file(self) -> None:
        """Make a new, empty add-in file (optionally with a ribbon tab) and open it."""
        if self._saving:
            return
        kind = tk.StringVar(master=self.root, value=_NEW_FILETYPES[0][0])
        chosen = filedialog.asksaveasfilename(
            title="New Add-in",
            filetypes=_NEW_FILETYPES,
            typevariable=kind,
            confirmoverwrite=False,
        )
        if not chosen:
            return
        path = Path(chosen)
        if path.suffix.lower() not in (".xlam", ".ppam", ".pptm"):
            # A name typed without an extension takes the chosen file type's.
            pattern = dict(_NEW_FILETYPES).get(kind.get(), "*.xlam")
            path = path.with_name(path.name + pattern.lstrip("*"))
        if path.exists():
            messagebox.showerror(
                APP_NAME, f"{path.name} already exists. Choose a new file name, or open it instead."
            )
            return
        dialog = StarterRibbonDialog(
            self.root,
            title="New Add-in",
            intro=f"Give {path.name} its own ribbon tab with a group for your macros' buttons?",
            optional=True,
        )
        spec = dialog.wait()
        if spec is None and not dialog.skipped:
            return
        self.create_new_file(path, spec)

    def create_new_file(self, path: Path, ribbon: ss.StarterRibbonSpec | None) -> bool:
        """Write an empty ``path`` (.xlam/.ppam/.pptm), open it, and add ``ribbon`` as an edit."""
        try:
            written = ss.create_blank_file(path)
        except ss.StarterError as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return False
        except OSError as exc:
            messagebox.showerror(APP_NAME, f"{path.name} could not be created.\n\n{exc}")
            return False
        if not self.load_path(path, written=written):
            return False
        if ribbon is not None and not self.apply_starter_ribbon(ribbon):
            return True
        tail = " with a ribbon tab. Save File writes the ribbon" if ribbon else ""
        self.status.config(text=f"Created {path.name}{tail}.")
        return True

    def create_ribbon(self) -> bool:
        """Ribbon → Create Ribbon…: give a file without ribbon XML a first tab."""
        draft = self.draft
        if draft is None or self._saving:
            return False
        self._flush_all_editors()
        problem = ss.ribbon_blocker(draft)
        if problem:
            messagebox.showinfo(APP_NAME, problem)
            return False
        spec = self.ask_starter_ribbon(
            f"{draft.baseline.path.name} has no ribbon yet. Add a tab with one group?"
        )
        return spec is not None and self.apply_starter_ribbon(spec)

    def ask_starter_ribbon(self, intro: str) -> ss.StarterRibbonSpec | None:
        return StarterRibbonDialog(self.root, title="Create Ribbon", intro=intro).wait()

    def apply_starter_ribbon(self, spec: ss.StarterRibbonSpec) -> bool:
        """Add a first ribbon (customUI14.xml and its links) to the draft as one undoable edit."""
        draft = self.draft
        if draft is None:
            return False
        self._flush_all_editors()
        problem = ss.ribbon_blocker(draft) or ss.validate_spec(spec)
        if problem:
            messagebox.showinfo(APP_NAME, problem)
            return False
        added: list = []
        failures: list[str] = []

        def mutate() -> None:
            try:
                added.append(ss.add_starter_ribbon(draft, spec))
            except ss.StarterError as exc:  # pragma: no cover - checked just above
                failures.append(str(exc))

        self._batch_edit(mutate)
        if failures or not added:
            messagebox.showerror(APP_NAME, failures[0] if failures else "No ribbon was added.")
            return False
        part = added[0]
        text = part.text or ""
        _tab_id, group_id = spec.ids()
        at = max(text.find(f'id="{rs.xml_attr(group_id)}"'), 0)
        self._reveal_ribbon(part.path, at, 0)
        self.status.config(
            text=f'Added ribbon tab "{spec.tab_label}" with group "{spec.group_label}". '
            "Nothing is written until Save File."
        )
        return True

    def _ensure_ribbon(self, purpose: str) -> bool:
        """When the file has no editable ribbon XML, offer to create one. True = there is one now."""
        draft = self.draft
        if draft is None:
            return False
        if any(p.editable for p in rs.ribbon_parts(draft)):
            return True
        problem = ss.ribbon_blocker(draft)
        if problem:
            messagebox.showinfo(
                APP_NAME, f"This file has no editable ribbon XML to put {purpose} in.\n\n{problem}"
            )
            return False
        spec = self.ask_starter_ribbon(
            f"This file has no ribbon yet. Add a tab with one group to put {purpose} in?"
        )
        return spec is not None and self.apply_starter_ribbon(spec)

    def _offer_icon_repair(self) -> None:
        """Version 0.6.0 linked custom icons so that Office shows them blank: offer the fix."""
        if self.draft is None or not rs.broken_image_link_parts(self.draft):
            return
        if messagebox.askyesno(
            APP_NAME,
            f"{self.draft.baseline.path.name} has custom ribbon icons that Office shows "
            "blank (an earlier version of this editor linked them wrongly).\n\n"
            "Repair them now? Nothing is written until you save.",
        ):
            self.repair_ribbon_icons()

    def repair_ribbon_icons(self) -> bool:
        """Relink custom icons an earlier version saved with a type Office ignores."""
        draft = self.draft
        if draft is None:
            return False
        self._flush_all_editors()
        if not rs.broken_image_link_parts(draft):
            messagebox.showinfo(APP_NAME, "No custom icons need repairing.")
            return False
        changed: list[str] = []

        def mutate() -> None:
            changed.extend(rs.repair_image_links(draft))

        self._batch_edit(mutate)
        self.status.config(
            text=f"Repaired the icon links in {', '.join(changed)}. "
            "Nothing is written until Save File."
        )
        return True

    def remove_unused_ribbon_icons(self) -> bool:
        """Drop images (and their links) that no ribbon control shows any more."""
        draft = self.draft
        if draft is None:
            return False
        self._flush_all_editors()
        if draft.baseline.package_safety.opc_signature_present:
            messagebox.showinfo(
                APP_NAME, "This file has a package digital signature; removing images would break it."
            )
            return False
        if not any(not link.used for link in icon_service.image_links(draft)):
            messagebox.showinfo(APP_NAME, "There are no unused icons in this file.")
            return False
        removed: list[str] = []

        def mutate() -> None:
            removed.extend(icon_service.remove_unused_images(draft))

        self._batch_edit(mutate)
        what = (
            f"{len(removed)} unused image(s): {', '.join(removed)}"
            if removed
            else "unused icon links (their images are still shown elsewhere)"
        )
        self.status.config(text=f"Removed {what}. Nothing is written until Save File.")
        return True

    def save(self) -> None:
        if self.draft is None or self._saving:
            return
        self.cancel_queued_save()
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
        path = self.draft.baseline.path
        blocked = self.save_service.host_block(path)
        if blocked is not None and blocked.reason == "file_in_use":
            blocked = self._wait_for_free_file(path)
        if blocked is not None:
            self._offer_queued_save(blocked)
            return
        self._save_confirmed()

    def _wait_for_free_file(self, path: Path):
        """Give a briefly held file (antivirus, sync, indexing) a few seconds, off the Tk thread."""
        wait = self.save_service.lock_wait_seconds
        self._saving = True
        self._refresh_state()
        try:
            return self._run_in_background(
                lambda: self.save_service.host_block(path, wait=wait),
                f"Waiting for {path.name} to be free…",
            )
        finally:
            self._saving = False
            self._refresh_state()

    def _save_confirmed(self, *, merged_external: bool = False) -> None:
        """Run the reviewed save; recover from Office holding or rewriting the file."""
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
        if result.kind == "blocked" and result.reason in _HOST_WAIT_REASONS:
            # Office started, or something grabbed the file, after the check.
            self._offer_queued_save(result)
            return
        if (
            result.kind == "blocked"
            and result.reason == "external_change"
            and not merged_external
            and self.session is not None
        ):
            if self._adopt_external_changes():
                self._save_confirmed(merged_external=True)
            return
        note = ""
        if merged_external and result.kind == "success" and self.draft is not None:
            note = (
                f"{self.draft.baseline.path.name} had been saved by another program "
                f"(such as {wp.corresponding_host_label(self.draft.baseline.path)}) "
                "after you opened it. Your edits were applied on top of that version."
            )
        self._handle_save_result(result, published=published, note=note)

    # -- waiting for Office / adopting its changes ---------------------------

    def _offer_queued_save(self, blocked) -> None:
        """Office (or another program) holds the file: offer to save once it is free."""
        assert self.draft is not None
        label = wp.corresponding_host_label(self.draft.baseline.path)
        if blocked.reason not in _HOST_WAIT_REASONS:
            messagebox.showwarning(APP_NAME, dialog_text(blocked))
            return
        text = (
            f"{blocked.message}\n\n"
            f"Your changes stay in the editor. The editor can save "
            f"{self.draft.baseline.path.name} automatically as soon as {label} "
            "has closed and the file is free.\n\n"
            "Save automatically when it is free?"
        )
        if messagebox.askyesno(APP_NAME, text):
            self._queue_save()

    def _queue_save(self) -> None:
        assert self.draft is not None
        self._flush_all_editors()
        self._queued_save = _draft_signature(self.draft)
        self._refresh_state()
        self._schedule_queued_save_poll()

    def cancel_queued_save(self) -> None:
        if self._queue_after is not None:
            try:
                self.root.after_cancel(self._queue_after)
            except tk.TclError:
                pass
            self._queue_after = None
        if self._queued_save is not None:
            self._queued_save = None
            self._refresh_state()

    def _schedule_queued_save_poll(self) -> None:
        self._queue_after = self.root.after(QUEUED_SAVE_POLL_MS, self._poll_queued_save)

    def _poll_queued_save(self) -> None:
        self._queue_after = None
        if self._queued_save is None or self.draft is None:
            self.cancel_queued_save()
            return
        if self._saving or self.save_service.host_block(self.draft.baseline.path) is not None:
            self._schedule_queued_save_poll()
            return
        self._flush_all_editors()
        reviewed = self._queued_save
        self._queued_save = None
        self._refresh_state()
        if not self.draft.is_dirty():
            return
        if _draft_signature(self.draft) == reviewed:
            self._save_confirmed()
        else:
            self.save()  # edited since the review: review again

    def _adopt_external_changes(self) -> bool:
        """Rebase the draft onto a file Office rewrote after it was opened.

        Non-overlapping changes merge without asking (the three-way rules of
        ConflictService); overlapping ones open the conflict dialog and the
        user saves again afterwards. True when the draft now sits on the
        current file and can be saved.
        """
        assert self.draft is not None and self.session is not None
        path = self.draft.baseline.path
        session = self.session
        label = wp.corresponding_host_label(path)
        self._flush_all_editors()

        def work():
            snapshot, _digest, error = self.conflict_service.snapshot_external(path)
            if snapshot is None:
                return None, None, error
            return snapshot, self.conflict_service.compare(session, snapshot), None

        self._saving = True
        try:
            snapshot, proposal, error = self._run_in_background(
                work, f"Reading the changed {path.name}…"
            )
        except (AdapterError, OSError) as exc:
            snapshot, proposal, error = None, None, str(exc)
        finally:
            self._saving = False
        if snapshot is None or proposal is None:
            messagebox.showwarning(
                APP_NAME,
                f"{path.name} changed on disk after you opened it, and the new "
                f"version could not be read ({error or 'unknown'}).\n\n"
                "Nothing was overwritten and your draft is kept. Close any program "
                "using the file and Save again.",
            )
            self._refresh_state()
            return False
        if proposal.parse_failed or any(item.requires_choice for item in proposal.items):
            messagebox.showinfo(
                APP_NAME,
                f"{path.name} was saved by another program (such as {label}) after "
                "you opened it, and some of those changes overlap yours.\n\n"
                "Choose which version to keep for each, then Save again.",
            )
            self._run_conflict_resolution(snapshot)
            self._refresh_state()
            return False
        current = (
            self.draft.module_by_id(self.current_module_id) if self.current_module_id else None
        )
        current_name = current.current_name if current is not None else None
        result = self.conflict_service.commit(
            session,
            proposal,
            snapshot,
            live_external_hash=snapshot.file_fingerprint.sha256,
            external_path=session.original_path,
            session_service=self.session_service,
            recovery_service=self.recovery_service,
        )
        if result.kind != "success":
            messagebox.showerror(
                APP_NAME,
                f"{path.name} changed on disk after you opened it, and your edits "
                "could not be applied to the new version: "
                + (result.message or result.reason or "unknown")
                + "\n\nNothing was overwritten and your draft is kept.",
            )
            self._refresh_state()
            return False
        self.draft = session.draft
        self._reload_editors_from_draft()
        if current_name is not None and (
            self.current_module_id is None or self.draft.module_by_id(self.current_module_id) is None
        ):
            moved = self.draft.find_current(current_name)
            if moved is not None:
                self._open_module_at(moved.id, 1)
        return True

    def _save_work(self):
        """Worker-thread half of save: commit, then publish the session baseline."""
        draft = self.draft
        assert draft is not None
        result = self.save_service.save_addin(draft)
        published = None
        if result.kind == "success" and self.session is not None:
            saved = Path(str(result.details.get("saved") or draft.baseline.path))
            # The refreshed baseline holds the committed bytes; no need to read them back.
            published = self.session_service.publish_verified_baseline(
                self.session, saved, draft.baseline.package_bytes
            )
        return result, published

    def _handle_save_result(self, result, *, published=None, note: str = "") -> None:
        self._last_result = result
        kind = result.kind
        if kind == "success":
            if self.draft is not None:
                self._start_prescan(
                    Path(str(result.details.get("saved") or self.draft.baseline.path))
                )
            warning = self.backup_service.record_success(
                original=self.draft.baseline.path if self.draft else Path("."),
                backup=result.backup_path,
                operation_type=result.operation_type,
                sha256=result.details.get("hash_backup"),
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
            self._refresh_ribbon_views()
            self._refresh_state()
            name = self.draft.baseline.path.name if self.draft else ""
            if note or warning:
                msg = "Saved successfully."
                if note:
                    msg += f"\n\n{note}"
                if result.backup_path:
                    msg += f"\n\nBackup: {result.backup_path}"
                if warning:
                    msg += f"\n\n{warning}"
                messagebox.showinfo(APP_NAME, msg)
            else:
                kept = (
                    f"Backup kept next to the file: {Path(str(result.backup_path)).name}."
                    if result.backup_path
                    else "No backup was made: automatic backup is off."
                )
                self.toast.show(f"Saved {name}", kept + " The antivirus check runs in the status bar.")
                self.status.config(
                    text=f"Saved {name} · " + ("backup kept" if result.backup_path else "no backup")
                )
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
            # Without a backup file, the copy captured when the add-in was opened
            # (or last saved) is the pre-save version.
            previous = result.backup_path or (
                self.session.captured_path if self.session is not None else None
            )
            messagebox.showerror(
                APP_NAME,
                "The add-in was replaced but final verification failed."
                + "\n\n"
                + "\n".join(result.problems)
                + f"\n\nPrevious version: {previous}"
                + "\n\nUse File → Restore Backup with that file to recover. "
                "The editor will not restore automatically.",
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
        in_place = operation_type == "save"
        warnings = list(model.encoding_warnings)
        if model.signature_will_be_removed:
            warnings.insert(
                0,
                "This save removes the VBA digital signature. Office or your organization's "
                "security policy may then block the add-in.",
            )
        extra = [f"VBA changes: {model.vba_count}   ·   XML changes: {model.xml_count}"]
        if not in_place:
            extra.append(model.backup_policy)
        view = ReviewView(
            file_name=self.draft.baseline.path.name,
            destination=str(model.destination),
            file_type=str(model.file_type),
            entries=tuple(entry_from_item(item) for item in model.items),
            backup_text=self._backup_policy_text() if in_place else model.backup_policy,
            confirm_label=(
                ("Save and Remove Signature" if model.signature_will_be_removed else "Save")
                if in_place
                else "Save Copy"
            ),
            warnings=tuple(warnings),
            in_place=in_place,
            extra=tuple(extra),
        )
        choice = ask_review(
            self.root,
            view,
            backup_var=self.auto_backup_var if in_place else None,
            on_backup_toggled=self._on_auto_backup_toggled,
        )
        if choice == REVERT:
            self.confirm_revert_all()
            return False
        if choice != SAVE:
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
        lines += ["", self._backup_policy_text()]
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
            "block the add-in.\n\n"
            + (
                "A backup of the currently signed file will be created automatically."
                if self.save_service.auto_backup
                else "Automatic backup is off: the signed file will NOT be kept. "
                "Use File → Back Up Now first if you want a copy."
            )
            + "\n\nRemove the signature and save?"
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
            self._start_prescan(dest)
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

    def confirm_revert_all(self) -> None:
        if self.draft is None:
            return
        self._flush_all_editors()
        n = dirty_count(self.draft)
        if not n:
            self.status.config(text="Nothing to revert.")
            return
        if messagebox.askyesno(
            APP_NAME,
            f"Discard all {n} unsaved change{'s' if n != 1 else ''} to {self.draft.baseline.path.name}?\n\n"
            "The file on disk is not touched.",
            icon="warning",
        ):
            self.revert_all()
            self.status.config(text="All changes reverted.")

    def revert_all(self) -> None:
        if self.draft is None:
            return
        from vba_addin_editor.domain.document import revert_all

        self._flush_all_editors()
        self.draft = revert_all(self.draft)
        if self.session is not None:
            self.session.draft = self.draft
        self.current_module_id = None
        self.current_xml_part_path = None
        self._refresh_tree()
        self._refresh_xml_tree()
        self.editor.set_text("")
        self.xml_editor.set_text("")
        self.xml_editor.set_part_info(None)
        self.validate_xml_btn.config(state="disabled")
        self._refresh_ribbon_views()
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
                "The layout cannot be exported or re-imported by this editor; "
                + (
                    "the pre-save backup is the only way to get it back."
                    if self.save_service.auto_backup
                    else "automatic backup is off, so use File → Back Up Now before "
                    "saving if you may want it back."
                )
                + "\n\nIt is not written to the file until Save File."
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

    def _backup_policy_text(self) -> str:
        if self.save_service.auto_backup:
            return "A same-folder backup is created before the original is replaced."
        return "Automatic backup is off: the original is replaced without a backup file."

    def _on_auto_backup_toggled(self) -> None:
        enabled = bool(self.auto_backup_var.get())
        self.settings.auto_backup = enabled
        self.save_service.auto_backup = enabled
        saved = save_settings(self.settings, self._settings_path)
        self._update_status_cells()
        self.status.config(
            text=(
                "Automatic backup on save is on."
                if enabled
                else "Automatic backup on save is off. Use File → Back Up Now when you want one."
            )
            + ("" if saved else " (This choice could not be saved for next time.)")
        )

    def backup_now(self) -> None:
        """Copy the add-in as it is on disk to a same-folder backup file."""
        if self.draft is None:
            messagebox.showinfo(APP_NAME, "Open an add-in first.")
            return
        try:
            backup = self.backup_service.create_manual_backup(self.draft.baseline)
        except OSError as exc:
            messagebox.showerror(APP_NAME, f"The backup could not be created.\n\n{exc}")
            return
        self._flush_all_editors()
        note = "  Unsaved edits are not in it." if self.draft.is_dirty() else ""
        self.status.config(text=f"Backup created: {backup.name}.{note}")

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
        self._prescan.settle()  # restore needs the file free of other handles
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
        """The header's "N changes": the review step of Save, where Save, Revert or Cancel follow."""
        if self.draft is None:
            return
        self.save()

    def show_about(self) -> None:
        identity = build_identity()
        messagebox.showinfo(
            APP_NAME,
            f"{APP_NAME} {identity.get('version', VERSION)}\n"
            f"Commit: {identity.get('source_commit', 'unbuilt')}\n"
            f"Mode: {identity.get('packaged_mode', 'development')}\n"
            f"pyOpenVBA: {identity.get('pyopenvba_installed') or identity.get('pyopenvba_pin')}\n\n"
            "Edits VBA source inside installed .xlam / .ppam add-ins and .pptm\n"
            "presentations, in place, with verification and optional backups.\n\n"
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
        self._schedule_outline()

    def _on_xml_edited(self) -> None:
        self._note_text_edit(kind="xml")

    def _refresh_ribbon_views(self) -> None:
        """After the draft or the open XML part changed: the Design button, and the Ribbon view."""
        ribbon = self.draft is not None and rs.is_ribbon_part(self.current_xml_part_path or "")
        if ribbon and not self.design_btn.winfo_manager():
            self.design_btn.pack(side="right", after=self.validate_xml_btn)
        elif not ribbon and self.design_btn.winfo_manager():
            self.design_btn.pack_forget()
        if self._designer_shown():
            self.ribbon_designer.refresh()

    def show_ribbon_design(self) -> None:
        """The Ribbon view, on the control the XML cursor is in when ribbon XML is showing."""
        draft = self.draft
        if draft is None:
            return
        self._flush_all_editors()
        path = self.current_xml_part_path
        from_xml = self._current_view() == str(self.xml_tab) and path is not None and rs.is_ribbon_part(path)
        offset = len(self.xml_editor.text.get("1.0", "insert")) if from_xml else None
        self._show_view(self.editor_notebook.index(self.ribbon_tab))
        if from_xml and path is not None and offset is not None:
            self.ribbon_designer.select_offset(path, offset)
        self.ribbon_designer.tree.focus_set()

    def _reveal_ribbon(self, part_path: str, offset: int, length: int) -> None:
        """Show a ribbon control after an edit, in the ribbon view used last (Design or XML)."""
        if self._ribbon_surface == "xml" or not rs.is_ribbon_part(part_path):
            self._select_xml_range(part_path, offset, length)
            return
        self._show_view(self.editor_notebook.index(self.ribbon_tab))
        self.ribbon_designer.select_offset(part_path, offset)

    # -- change bars ---------------------------------------------------------------

    def _module_baseline(self) -> tuple[bool, str | None]:
        mod = self.draft.module_by_id(self.current_module_id) if self.draft and self.current_module_id else None
        if mod is None:
            return False, None
        return True, mod.original_body

    def _xml_baseline(self) -> tuple[bool, str | None]:
        path = self.current_xml_part_path
        part = self.draft.xml_part_by_path(path) if self.draft and path else None
        if part is None or not part.editable:
            return False, None
        return True, None if part.is_new else part.original_text

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
        self._refresh_ribbon_views()
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
        if rs.is_ribbon_part(part_path):
            self._ribbon_surface = "xml"
        if self.current_xml_part_path != part_path:
            self.current_xml_part_path = part_path
            self.xml_editor.text.config(state="normal")
            self.xml_editor.set_text(part.text)
            if not part.editable:
                self.xml_editor.text.config(state="disabled")
            self.xml_editor.set_part_info(part)
            self.validate_xml_btn.config(state="normal")
            self._refresh_ribbon_views()
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
        if self._designer_shown():
            if not self.ribbon_designer.goto_macro():
                self.status.config(text="Select a button that runs a macro, then press F12 to open it.")
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
        """Show the control that has ``callback``: in the Ribbon view, or in its XML if that was used last."""
        if self._ribbon_surface == "xml":
            self._open_xml_at(callback.part_path, callback.line, callback.column, callback.length)
        else:
            self._reveal_ribbon(callback.part_path, callback.tag_start, 0)

    def _batch_edit(self, mutate) -> None:
        """Apply a programmatic multi-target edit as one undoable step."""
        draft = self.draft
        if draft is None:
            return
        self._flush_all_editors()
        before = {
            "modules": snapshot_modules(draft),
            "xml": snapshot_xml(draft),
            "package": snapshot_package(draft),
        }
        mutate()
        # Reload before checkpointing: a checkpoint flushes the (stale) editors.
        self._reload_editors_from_draft()
        if self.session is not None:
            after = {
                "modules": snapshot_modules(draft),
                "xml": snapshot_xml(draft),
                "package": snapshot_package(draft),
            }
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
        if not self._ensure_ribbon("the button"):
            return
        parts = [p for p in rs.ribbon_parts(draft) if p.editable]
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

    def add_ribbon_control(self, kind: str) -> None:
        """Ribbon → Add …: a new button, menu, group or separator at the Ribbon view's selection."""
        if self.draft is None or self._saving:
            return
        if not self._designer_shown():
            self._show_view(self.editor_notebook.index(self.ribbon_tab))
        self.ribbon_designer.add_control(kind)

    def add_ribbon_menu(self) -> None:
        """A new ribbon menu: in the Ribbon view directly, from the XML editor through a form."""
        draft = self.draft
        if draft is None:
            return
        if self._current_view() != str(self.xml_tab):
            self.add_ribbon_control("menu")
            return
        self._flush_all_editors()
        if not self._ensure_ribbon("the menu"):
            return
        parts = [p for p in rs.ribbon_parts(draft) if p.editable]
        containers = [c for p in parts for c in rs.find_containers(p.text or "", p.path)]
        if not containers:
            messagebox.showinfo(
                APP_NAME, "The ribbon XML has no custom group or menu to add a menu to."
            )
            return
        # Start in the group or menu the XML cursor is in, when there is one.
        initial = 0
        if self._active_text_editor() is self.xml_editor and self.current_xml_part_path:
            offset = len(self.xml_editor.text.get("1.0", "insert"))
            here = rs.innermost_container(containers, self.current_xml_part_path, offset)
            if here is not None:
                initial = containers.index(here)
        dialog = AddMenuDialog(
            self.root,
            containers=containers,
            existing_ids={
                value for p in rs.ribbon_parts(draft) for value, _o in rs.control_ids(p.text or "")
            },
            image_suggestions=sorted(
                {v for p in parts for v in re.findall(r'imageMso="([^"]+)"', p.text or "")},
                key=str.casefold,
            ),
            macros=rs.ribbon_ready_procedures(draft),
            show_part=len(parts) > 1,
            initial_container=initial,
        )
        spec = dialog.wait()
        if spec is not None:
            self.apply_ribbon_menu(spec)

    def apply_ribbon_menu(self, spec) -> None:
        draft = self.draft
        if draft is None:
            return
        part = draft.xml_part_by_path(spec.container.part_path)
        if part is None or part.text is None:
            return
        taken = {
            value for p in rs.ribbon_parts(draft) for value, _o in rs.control_ids(p.text or "")
        } | {spec.control_id}
        children = []
        for macro in spec.macros:
            control_id = rs.unique_control_id(taken, macro.name)
            taken.add(control_id)
            children.append(
                rs.build_button_xml(
                    control_id=control_id,
                    label=rs.suggested_label(macro.name),
                    on_action=macro.name,
                )
            )
        element = rs.build_menu_xml(
            control_id=spec.control_id,
            label=spec.label,
            image_mso=spec.image_mso,
            size=spec.size,
            item_size=spec.item_size,
            screentip=spec.screentip,
            children=children,
            indent=spec.container.child_indent,
            unit=rs.child_unit(spec.container),
        )
        new_text, offset, _length = rs.insert_into_container(part.text, spec.container, element)

        def mutate() -> None:
            part.text = new_text

        self._batch_edit(mutate)
        self._reveal_ribbon(part.path, offset, len(element.split("\n", 1)[0]))
        count = f" with {len(children)} button(s)" if children else ""
        self.status.config(
            text=f'Added menu "{spec.label}"{count}. Nothing is written until Save File.'
        )

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
        icon = getattr(spec, "icon", None)
        icon_problems: list[str] = []

        def mutate() -> None:
            part.text = new_text
            mod.body = new_body
            if icon is not None:
                try:
                    icon_service.set_custom_icon(draft, part.path, offset, icon)
                except icon_service.IconError as exc:
                    icon_problems.append(str(exc))

        self._batch_edit(mutate)
        if icon_problems:
            messagebox.showwarning(
                APP_NAME, "The button was added without its icon:\n\n" + icon_problems[0]
            )
        self._reveal_ribbon(part.path, offset, length)
        extra = f" and {callback_name} in {mod.current_name}" if wrapper else ""
        self.status.config(
            text=f'Added button "{spec.label}"{extra}. Nothing is written until Save File.'
        )

    # -- import a button (and its macros) from another add-in ------------------------

    def import_ribbon_control(self) -> None:
        """Copy a ribbon control and the VBA it runs from another add-in into this one."""
        draft = self.draft
        if draft is None or self._saving:
            return
        self._flush_all_editors()
        if draft.baseline.safety.password_protected:
            messagebox.showwarning(APP_NAME, "This project is password-protected and read-only.")
            return
        if draft.baseline.package_safety.opc_signature_present:
            messagebox.showinfo(
                APP_NAME,
                "This file has a package digital signature; editing its ribbon would break it.",
            )
            return
        if not self._ensure_ribbon("the imported button"):
            return
        path = filedialog.askopenfilename(
            title="Import Button from Another Add-in", filetypes=_FILETYPES
        )
        if not path:
            return
        source = self.load_port_source(Path(path))
        if source is None:
            return
        plan = self.choose_port(source)
        if plan is not None:
            self.apply_port_plan(plan)

    def load_port_source(self, path: Path) -> DocumentDraft | None:
        """Read another add-in (read-only) to import from; None after an error message."""
        if self.draft is not None and path.resolve() == Path(self.draft.baseline.path).resolve():
            messagebox.showinfo(
                APP_NAME, "Choose a different add-in: this is the file that is open."
            )
            return None
        self._saving = True  # keeps autosave and close away while the worker runs
        try:
            return self._run_in_background(
                lambda: self.doc_service.open(path), f"Reading {path.name}…"
            )
        except (AdapterError, OSError, ValueError) as exc:
            messagebox.showerror(
                APP_NAME,
                str(exc)
                if isinstance(exc, AdapterError)
                else f"{path.name} could not be read.\n\n{exc}",
            )
            return None
        finally:
            self._saving = False
            self._refresh_state()

    def choose_port(self, source: DocumentDraft) -> ps.PortPlan | None:
        """Ask which control of ``source`` to import, and where; returns its plan."""
        draft = self.draft
        if draft is None:
            return None
        controls = ps.portable_controls(source)
        if not controls:
            messagebox.showinfo(
                APP_NAME, f"{source.baseline.path.name} has no ribbon controls that run macros."
            )
            return None
        dialog = ImportControlDialog(
            self.root,
            source_name=str(source.baseline.path),
            controls=controls,
            targets_for=lambda control: ps.port_targets(draft, control),
            plan=lambda control, where, suffix, label: ps.plan_port(
                source, draft, control, where, suffix=suffix, group_label=label
            ),
            suffix=ps.default_suffix(draft),
            show_part=len([p for p in rs.ribbon_parts(draft) if p.editable]) > 1,
        )
        return dialog.wait()

    def apply_port_plan(self, plan: ps.PortPlan) -> bool:
        """Apply a computed import as one undoable edit."""
        draft = self.draft
        if draft is None or not plan.ok:
            return False

        def mutate() -> None:
            ps.apply_port(draft, plan)

        self._batch_edit(mutate)
        if plan.part_path:
            self._reveal_ribbon(plan.part_path, plan.element_offset, plan.element_length)
        added = len(plan.new_module_ids)
        modules = f" and {added} module(s)" if added else ""
        self.status.config(
            text=f"Imported the control{modules}. Nothing is written until Save File."
        )
        if plan.warnings:
            messagebox.showinfo(
                APP_NAME,
                "Imported. Nothing is written until Save File. Please check:\n\n"
                + "\n".join(f"• {w}" for w in plan.warnings),
            )
        return True

    # -- ribbon structure: move, group into menu, custom icons ----------------------

    def _ribbon_edit_enabled(self) -> bool:
        path = self.current_xml_part_path
        if self.draft is None or path is None or not rs.is_ribbon_part(path):
            return False
        part = self.draft.xml_part_by_path(path)
        return part is not None and part.editable

    def _ribbon_xml_context(self, *, dialog: bool = False):
        """(draft, part, part text, cursor offset) for a structural ribbon edit, or None.

        ``dialog`` explains a refusal in a message box instead of the status
        bar, for menu commands whose silence would look like a success.
        """
        path = self.current_xml_part_path
        if path is None or not self._ribbon_edit_enabled():
            self._ribbon_hint(
                "Open the ribbon XML (customUI) on the XML tab and put the cursor "
                "on a control first.",
                dialog=dialog,
            )
            return None
        self._flush_all_editors()
        draft = self.draft
        part = draft.xml_part_by_path(path) if draft else None
        if draft is None or part is None or part.text is None:
            return None
        return draft, part, part.text, len(self.xml_editor.text.get("1.0", "insert"))

    def _ribbon_hint(self, text: str, *, dialog: bool) -> None:
        self.status.config(text=text)
        if dialog:
            messagebox.showinfo(APP_NAME, text)

    def _xml_selection_offsets(self, fallback: int) -> tuple[int, int]:
        text = self.xml_editor.text
        try:
            return len(text.get("1.0", "sel.first")), len(text.get("1.0", "sel.last"))
        except tk.TclError:
            return fallback, fallback

    def _select_xml_range(self, part_path: str, offset: int, length: int) -> None:
        part = self.draft.xml_part_by_path(part_path) if self.draft else None
        if part is None or part.text is None:
            return
        text = part.text
        line = text.count("\n", 0, offset) + 1
        column = offset - (text.rfind("\n", 0, offset) + 1) + 1
        self._open_xml_at(part_path, line, column, length)

    def _ribbon_ids(self, draft) -> set[str]:
        return {value for p in rs.ribbon_parts(draft) for value, _o in rs.control_ids(p.text or "")}

    def _image_suggestions(self, draft) -> list[str]:
        return sorted(
            {
                v
                for p in rs.ribbon_parts(draft)
                for v in re.findall(r'imageMso="([^"]+)"', p.text or "")
            },
            key=str.casefold,
        )

    def move_ribbon_control(self, direction: int) -> bool:
        """Swap the selected control (Ribbon view) or the one at the XML cursor with its neighbour."""
        if self._designer_shown():
            return self.ribbon_designer.move_selected(direction)
        context = self._ribbon_xml_context()
        if context is None:
            return False
        _draft, part, text, offset = context
        try:
            new_text, moved = rs.move_element(text, offset, direction)
        except rs.RibbonEditError as exc:
            self.status.config(text=str(exc))
            return False

        def mutate() -> None:
            part.text = new_text

        self._batch_edit(mutate)
        self._select_xml_range(part.path, moved.start, moved.end - moved.start)
        self.status.config(
            text=f"Moved {moved.describe()} {'up' if direction < 0 else 'down'}. "
            "Nothing is written until Save File."
        )
        return True

    def group_selection_into_menu(self) -> None:
        """Ask for a menu label, then wrap the selected controls in a new <menu>."""
        if self._designer_shown():
            self.ribbon_designer.group_selected_into_menu()
            return
        context = self._ribbon_xml_context(dialog=True)
        if context is None:
            return
        _draft, part, _text, offset = context
        start, end = self._xml_selection_offsets(offset)
        self.group_range_into_menu(part.path, start, end)

    def group_range_into_menu(self, part_path: str, start: int, end: int) -> None:
        """Ask for the menu, then wrap the sibling controls between two offsets of a ribbon part in it."""
        draft = self.draft
        part = draft.xml_part_by_path(part_path) if draft is not None else None
        if draft is None or part is None or part.text is None:
            return
        text = part.text
        try:
            parent, run = rs.selected_siblings(text, start, end)
            rs.group_into_menu(text, start, end, control_id="mnuCheck", label="Check")
        except rs.RibbonEditError as exc:
            messagebox.showinfo(APP_NAME, str(exc))
            return
        dialog = GroupMenuDialog(
            self.root,
            controls=[element.describe() for element in run],
            existing_ids=self._ribbon_ids(draft),
            image_suggestions=self._image_suggestions(draft),
            in_group=parent.name.rsplit(":", 1)[-1] == "group",
        )
        spec = dialog.wait()
        if spec is not None:
            self.apply_group_into_menu(part.path, start, end, spec)

    def apply_group_into_menu(self, part_path: str, start: int, end: int, spec) -> bool:
        draft = self.draft
        if draft is None:
            return False
        self._flush_all_editors()
        part = draft.xml_part_by_path(part_path)
        if part is None or part.text is None or not part.editable:
            return False
        try:
            new_text, offset, length = rs.group_into_menu(
                part.text,
                start,
                end,
                control_id=spec.control_id,
                label=spec.label,
                image_mso=spec.image_mso,
                size=spec.size,
                item_size=spec.item_size,
            )
        except rs.RibbonEditError as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return False
        icon_problems: list[str] = []

        def mutate() -> None:
            part.text = new_text
            if spec.icon is not None:
                try:
                    icon_service.set_custom_icon(draft, part.path, offset, spec.icon)
                except icon_service.IconError as exc:
                    icon_problems.append(str(exc))

        self._batch_edit(mutate)
        self._reveal_ribbon(part.path, offset, length)
        self.status.config(
            text=f'Grouped into menu "{spec.label}". Nothing is written until Save File.'
        )
        if icon_problems:
            messagebox.showwarning(
                APP_NAME, "The menu was created without its icon:\n\n" + icon_problems[0]
            )
        return True

    def set_ribbon_icon_from_clipboard(self) -> None:
        if self._designer_shown():
            self.ribbon_designer.set_icon_from_clipboard()
            return
        self._set_ribbon_icon(icon_service.icon_from_clipboard)

    def set_ribbon_icon_from_file(self) -> None:
        if self._designer_shown():
            self.ribbon_designer.set_icon_from_file()
            return
        if not self._ribbon_edit_enabled():
            self._ribbon_xml_context(dialog=True)  # explains what to do
            return
        name = ask_image_file(self.root)
        if name:
            self._set_ribbon_icon(lambda: icon_service.icon_from_file(Path(name)))

    def _set_ribbon_icon(self, make_icon) -> None:
        context = self._ribbon_xml_context(dialog=True)
        if context is None:
            return
        _draft, part, text, offset = context
        element = rs.image_control_at(text, offset)
        if element is None:
            self._ribbon_hint(
                "Put the cursor inside the <button …> (or menu or other control) "
                "that should show the icon, then try again.",
                dialog=True,
            )
            return
        try:
            icon = make_icon()
        except icon_service.IconError as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return
        self.apply_ribbon_icon(part.path, element.start, icon)

    def apply_ribbon_icon(self, part_path: str, control_start: int, icon) -> bool:
        """Give the control whose tag starts at ``control_start`` a custom picture."""
        draft = self.draft
        if draft is None:
            return False
        self._flush_all_editors()

        def mutate() -> None:
            icon_service.set_custom_icon(draft, part_path, control_start, icon)

        try:
            self._batch_edit(mutate)
        except icon_service.IconError as exc:  # raised before the draft changed
            messagebox.showerror(APP_NAME, str(exc))
            return False
        part = draft.xml_part_by_path(part_path)
        element = rs.element_at(part.text or "", control_start) if part is not None else None
        if element is not None:
            self._reveal_ribbon(part_path, element.start, element.open_end - element.start)
        what = element.describe() if element is not None else "the control"
        self.status.config(
            text=f"Set a custom icon on {what} ({icon.describe()}). "
            "Nothing is written until Save File."
        )
        return True

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
        for tab in (self.xml_tab, self.ribbon_tab):
            self.editor_notebook.tab(tab, state="normal" if xml_supported else "disabled")
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
        self._release_prescan_gc()
        for attr in (
            "_idle_after", "_max_after", "_poll_after", "_recovery_after", "_prescan_after",
            "_queue_after", "_outline_after", "_outline_idle",
        ):
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
    # Hidden while it is built: Windows 10 paints the dark title bar only if
    # it is set before the window first appears.
    root.withdraw()
    MainWindow(root)
    root.deiconify()
    root.mainloop()
