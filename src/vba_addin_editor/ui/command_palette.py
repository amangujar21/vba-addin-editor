"""Ctrl+K command palette: every menu command, module and procedure in one box.

Type to filter. A leading ">" lists commands only, "@" procedures only.
Up/Down move, Enter runs, Esc (or clicking elsewhere) closes.
"""

from __future__ import annotations

import tkinter as tk
from collections.abc import Callable
from dataclasses import dataclass
from tkinter import ttk

from vba_addin_editor.ui import theme
from vba_addin_editor.ui.widgets import icon

COMMAND, MODULE, PROCEDURE = "command", "module", "procedure"
_SECTIONS = ((MODULE, "MODULES"), (PROCEDURE, "PROCEDURES"), (COMMAND, "COMMANDS"))
_LIMITS = {MODULE: 6, PROCEDURE: 8, COMMAND: 200}
WIDTH = 600
MAX_ROWS = 12


@dataclass(frozen=True)
class PaletteItem:
    kind: str  # COMMAND, MODULE or PROCEDURE
    label: str
    detail: str  # menu name, module name, module kind
    shortcut: str
    action: Callable[[], object]


def filter_items(items: list[PaletteItem], raw: str) -> list[tuple[str, list[PaletteItem]]]:
    """Sections (title, items) for the query ``raw``, best matches first."""
    only = {">": COMMAND, "@": PROCEDURE}.get(raw[:1])
    query = (raw[1:] if only else raw).strip().casefold()
    sections: list[tuple[str, list[PaletteItem]]] = []
    for kind, title in _SECTIONS:
        if only is not None and kind != only:
            continue
        if not query and only is None and kind != COMMAND:
            continue  # an empty box suggests commands, not every module and procedure
        pool = [i for i in items if i.kind == kind]
        if query:
            with_detail = kind == COMMAND
            pool = [
                i
                for i in pool
                if query in (f"{i.detail} {i.label}" if with_detail else i.label).casefold()
            ]
            pool.sort(key=lambda i: (not i.label.casefold().startswith(query), i.label.casefold()))
        limit = _LIMITS[kind] if only is None else 200
        if pool:
            sections.append((title, pool[:limit]))
    return sections


class CommandPalette:
    """A borderless box near the top of the window; built once, shown on demand."""

    def __init__(self, root: tk.Tk, provider: Callable[[], list[PaletteItem]]) -> None:
        self.root = root
        self.provider = provider
        self._items: list[PaletteItem] = []
        self._rows: dict[str, PaletteItem] = {}
        self.window: tk.Toplevel | None = None

    # -- building -----------------------------------------------------------

    def _build(self) -> tk.Toplevel:
        colors = theme.palette()
        window = tk.Toplevel(self.root)
        window.withdraw()
        window.overrideredirect(True)
        window.transient(self.root)
        window.configure(background=colors.border_strong)
        body = ttk.Frame(window, style="Raised.TFrame")
        body.pack(fill="both", expand=True, padx=1, pady=1)

        query_row = ttk.Frame(body, style="Raised.TFrame", padding=(12, 4, 12, 4))
        query_row.pack(fill="x")
        ttk.Label(
            query_row, text=icon("search"), font=theme.ICON_FONT, foreground=colors.faint, background=colors.bg3
        ).pack(side="left", padx=(0, 4))
        self.query = tk.StringVar(master=window)
        self.entry = ttk.Entry(query_row, textvariable=self.query, style="Palette.TEntry", font=theme.TITLE_FONT)
        self.entry.pack(side="left", fill="x", expand=True)
        ttk.Label(query_row, text="Esc", style="Faint.TLabel", background=colors.bg3).pack(side="right")
        ttk.Frame(body, style="Line.TFrame", height=1).pack(fill="x")

        results = ttk.Frame(body, style="Raised.TFrame", padding=(6, 6))
        results.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(
            results,
            style="Palette.Treeview",
            show="tree",
            columns=("detail", "key"),
            selectmode="browse",
            height=MAX_ROWS,
        )
        self.tree.column("#0", width=WIDTH - 260, stretch=True)
        self.tree.column("detail", width=120, stretch=False, anchor="w")
        self.tree.column("key", width=120, stretch=False, anchor="e")
        self.tree.tag_configure("section", foreground=colors.faint, font=theme.CAPTION_FONT)
        self.tree.tag_configure("item", foreground=colors.text)
        self.tree.tag_configure("empty", foreground=colors.faint)
        self.tree.pack(fill="both", expand=True)

        ttk.Frame(body, style="Line.TFrame", height=1).pack(fill="x")
        foot = ttk.Frame(body, style="Raised.TFrame", padding=(14, 6))
        foot.pack(fill="x")
        ttk.Label(
            foot,
            text="↑ ↓  move      Enter  run      >  commands only      @  procedures",
            style="Faint.TLabel",
            background=colors.bg3,
        ).pack(side="left")

        self.query.trace_add("write", lambda *_a: self.refresh())
        self.entry.bind("<Down>", lambda _e: self._move(1))
        self.entry.bind("<Up>", lambda _e: self._move(-1))
        self.entry.bind("<Next>", lambda _e: self._move(8))
        self.entry.bind("<Prior>", lambda _e: self._move(-8))
        self.entry.bind("<Return>", lambda _e: self.run_selected())
        window.bind("<Escape>", lambda _e: self.close())
        self.tree.bind("<ButtonRelease-1>", self._on_click)
        self.tree.bind("<Return>", lambda _e: self.run_selected())
        window.bind("<FocusOut>", lambda _e: window.after(10, self._close_if_unfocused))
        return window

    # -- showing ------------------------------------------------------------

    def open(self, prefix: str = "") -> None:
        if self.window is None or not self.window.winfo_exists():
            self.window = self._build()
        self._items = self.provider()
        self.query.set(prefix)
        self.refresh()
        root = self.root
        root.update_idletasks()
        width = min(WIDTH, max(root.winfo_width() - 40, 360))
        x = root.winfo_rootx() + (root.winfo_width() - width) // 2
        y = root.winfo_rooty() + 56
        self.window.geometry(f"{width}x{self.window.winfo_reqheight()}+{x}+{y}")
        self.window.deiconify()
        self.window.lift()
        self.entry.focus_force()
        self.entry.icursor("end")

    def close(self) -> str:
        if self.window is not None and self.window.winfo_exists():
            self.window.withdraw()
        return "break"

    @property
    def visible(self) -> bool:
        return self.window is not None and self.window.winfo_exists() and self.window.winfo_viewable()

    def _close_if_unfocused(self) -> None:
        if self.window is None or not self.window.winfo_exists():
            return
        focus = self.window.focus_get() if self.window.winfo_viewable() else None
        if focus is None or focus.winfo_toplevel() is not self.window:
            self.close()

    # -- results ------------------------------------------------------------

    def refresh(self) -> None:
        tree = self.tree
        tree.delete(*tree.get_children())
        self._rows.clear()
        first = None
        for title, items in filter_items(self._items, self.query.get()):
            tree.insert("", "end", text=title, tags=("section",))
            for item in items:
                iid = tree.insert(
                    "", "end", text="   " + item.label, values=(item.detail, item.shortcut), tags=("item",)
                )
                self._rows[iid] = item
                first = first or iid
        if first is None:
            tree.insert("", "end", text="   No matches", tags=("empty",))
        else:
            tree.selection_set(first)
            tree.see(first)

    def _move(self, step: int) -> str:
        rows = [iid for iid in self.tree.get_children() if iid in self._rows]
        if not rows:
            return "break"
        selection = self.tree.selection()
        index = rows.index(selection[0]) if selection and selection[0] in rows else -1
        index = min(max(index + step, 0), len(rows) - 1)
        self.tree.selection_set(rows[index])
        self.tree.see(rows[index])
        return "break"

    def _on_click(self, event) -> None:
        iid = self.tree.identify_row(event.y)
        if iid in self._rows:
            self.tree.selection_set(iid)
            self.run_selected()

    def run_selected(self) -> str:
        selection = self.tree.selection()
        item = self._rows.get(selection[0]) if selection else None
        self.close()
        if item is not None:
            # After the box is gone, so dialogs the command opens get the focus.
            self.root.after_idle(item.action)
        return "break"
