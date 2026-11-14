"""Small widgets for the night redesign: icon glyphs, rail buttons, a spinner and a toast.

Icons are glyphs of the Windows icon font (Segoe Fluent Icons / Segoe MDL2
Assets), drawn as text, so the app still ships no image files. Without that
font every icon falls back to a plain character.
"""

from __future__ import annotations

import tkinter as tk
from collections.abc import Callable
from tkinter import ttk
from typing import Any

from vba_addin_editor.ui import theme

# name -> (icon-font glyph, plain fallback)
_ICONS = {
    "code": ("", "{}"),
    "ribbon": ("", "▤"),
    "form": ("", "▭"),
    "search": ("", "⌕"),
    "history": ("", "↺"),
    "gear": ("", "⚙"),
    "plus": ("", "+"),
    "lock": ("", "🔒"),
    "x": ("", "✕"),
    "up": ("", "▲"),
    "down": ("", "▼"),
    "right": ("", "›"),
    "check": ("", "✓"),
    "alert": ("", "⚠"),
    "shield": ("", "◈"),
    "bolt": ("", "ϟ"),
    "image": ("", "▣"),
    "menu": ("", "≡"),
    "button": ("", "▭"),
    "folder": ("", "▸"),
    "replace": ("", "⇄"),
}


_ICONS["xml"] = ("", "≣")  # a document page: the package's XML parts


def icon(name: str) -> str:
    glyph, fallback = _ICONS[name]
    return glyph if theme.icons_available() else fallback


class RailButton(tk.Frame):
    """A square icon button for the activity rail, with an accent bar when active."""

    def __init__(self, master, glyph: str, tip: str, command: Callable[[], object]) -> None:
        colors = theme.palette()
        super().__init__(master, background=colors.bg0)
        self.command = command
        self.tip = tip
        self._active = False
        self._hover = False
        self.bar = tk.Frame(self, width=2, background=colors.bg0)
        self.bar.pack(side="left", fill="y", pady=9)
        self.label = tk.Label(
            self,
            text=glyph,
            font=theme.ICON_FONT,
            width=3,
            height=1,
            pady=7,
            background=colors.bg0,
            foreground=colors.faint,
            cursor="hand2",
        )
        self.label.pack(side="left", padx=(4, 6))
        for widget in (self, self.label):
            widget.bind("<Enter>", self._on_enter)
            widget.bind("<Leave>", self._on_leave)
            widget.bind("<Button-1>", lambda _e: self.command())
        Tooltip(self.label, tip)

    def set_active(self, active: bool) -> None:
        self._active = active
        self.apply_palette()

    def _on_enter(self, _event=None) -> None:
        self._hover = True
        self.apply_palette()

    def _on_leave(self, _event=None) -> None:
        self._hover = False
        self.apply_palette()

    def apply_palette(self) -> None:
        colors = theme.palette()
        self.configure(background=colors.bg0)
        self.bar.configure(background=colors.accent if self._active else colors.bg0)
        self.label.configure(
            background=colors.bg4 if self._hover else colors.bg0,
            foreground=colors.text if (self._active or self._hover) else colors.faint,
        )


class Tooltip:
    """A one-line hint shown beside a widget while the pointer rests on it."""

    DELAY_MS = 450

    def __init__(self, widget: tk.Misc, text: str) -> None:
        self.widget = widget
        self.text = text
        self._after = None
        self._window: tk.Toplevel | None = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<Button-1>", self._hide, add="+")

    def _schedule(self, _event=None) -> None:
        self._cancel()
        self._after = self.widget.after(self.DELAY_MS, self._show)

    def _cancel(self) -> None:
        if self._after is not None:
            try:
                self.widget.after_cancel(self._after)
            except tk.TclError:
                pass
            self._after = None

    def _show(self) -> None:
        self._after = None
        if self._window is not None or not self.widget.winfo_exists():
            return
        colors = theme.palette()
        window = self._window = tk.Toplevel(self.widget)
        window.overrideredirect(True)
        window.configure(background=colors.border_strong)
        tk.Label(
            window,
            text=self.text,
            background=colors.bg4,
            foreground=colors.text,
            font=theme.SMALL_FONT,
            padx=8,
            pady=3,
        ).pack(padx=1, pady=1)
        x = self.widget.winfo_rootx() + self.widget.winfo_width() + 6
        y = self.widget.winfo_rooty() + max((self.widget.winfo_height() - 24) // 2, 0)
        window.geometry(f"+{x}+{y}")

    def _hide(self, _event=None) -> None:
        self._cancel()
        if self._window is not None:
            try:
                self._window.destroy()
            except tk.TclError:
                pass
            self._window = None


class Spinner(tk.Canvas):
    """A small turning arc, for "something is running" without a progress figure."""

    def __init__(self, master, size: int = 14, background: str | None = None, color: str | None = None) -> None:
        colors = theme.palette()
        self._bg = background or colors.bg0
        self._color = color or colors.warn
        super().__init__(master, width=size, height=size, background=self._bg, highlightthickness=0)
        self._size = size
        self._angle = 0
        self._after = None
        pad = 2
        self.create_oval(pad, pad, size - pad, size - pad, outline=colors.bg4, width=2)
        self._arc = self.create_arc(
            pad, pad, size - pad, size - pad, start=0, extent=100, style="arc", outline=self._color, width=2
        )
        self.bind("<Map>", lambda _e: self._start(), add=True)
        self.bind("<Unmap>", lambda _e: self._stop(), add=True)
        self.bind("<Destroy>", lambda _e: self._stop(), add=True)

    def _start(self) -> None:
        if self._after is None:
            self._tick()

    def _stop(self) -> None:
        if self._after is not None:
            try:
                self.after_cancel(self._after)
            except tk.TclError:
                pass
            self._after = None

    def _tick(self) -> None:
        self._angle = (self._angle - 30) % 360
        try:
            self.itemconfigure(self._arc, start=self._angle)
            self._after = self.after(80, self._tick)
        except tk.TclError:
            self._after = None


class Toast:
    """A short confirmation in the bottom-right corner that goes away by itself."""

    SHOW_MS = 4200

    def __init__(self, root: tk.Misc) -> None:
        self.root = root
        self._after = None
        self.frame = tk.Frame(root, background=theme.palette().border_strong)
        inner = self.inner = ttk.Frame(self.frame, style="Toast.TFrame", padding=(12, 10, 16, 12))
        inner.pack(padx=1, pady=1)
        self.icon = ttk.Label(inner, text=icon("check"), style="ToastIcon.TLabel")
        self.icon.grid(row=0, column=0, rowspan=2, sticky="n", padx=(0, 10), pady=(1, 0))
        self.title = ttk.Label(inner, text="", style="Toast.TLabel")
        self.title.grid(row=0, column=1, sticky="w")
        self.sub = ttk.Label(inner, text="", style="ToastSub.TLabel", wraplength=300, justify="left")
        self.sub.grid(row=1, column=1, sticky="w", pady=(2, 0))
        self.frame.bind("<Button-1>", lambda _e: self.hide())
        for widget in (inner, self.title, self.sub, self.icon):
            widget.bind("<Button-1>", lambda _e: self.hide())

    def show(self, title: str, sub: str = "") -> None:
        self.title.config(text=title)
        self.sub.config(text=sub)
        if sub:
            self.sub.grid()
        else:
            self.sub.grid_remove()
        self.frame.configure(background=theme.palette().border_strong)
        self.frame.place(relx=1.0, rely=1.0, anchor="se", x=-18, y=-40)
        self.frame.lift()
        if self._after is not None:
            self.root.after_cancel(self._after)
        self._after = self.root.after(self.SHOW_MS, self.hide)

    def hide(self) -> None:
        if self._after is not None:
            try:
                self.root.after_cancel(self._after)
            except tk.TclError:
                pass
            self._after = None
        self.frame.place_forget()

    @property
    def visible(self) -> bool:
        return bool(self.frame.winfo_manager())


class Sidebar(ttk.Frame):
    """A view's left column: caption, optional + button and filter box, a tree, an optional footer."""

    def __init__(
        self,
        master,
        title: str,
        *,
        add: Callable[[], object] | None = None,
        add_tip: str = "",
        filter_var: tk.StringVar | None = None,
        footer: bool = False,
        width: int = 264,
    ) -> None:
        super().__init__(master, width=width)
        self.pack_propagate(False)
        self.head = ttk.Frame(self, padding=(14, 8, 6, 4))
        self.head.pack(fill="x", side="top")
        ttk.Label(self.head, text=title, style="Caption.TLabel").pack(side="left", pady=3)
        if add is not None:
            plus = ttk.Button(self.head, text=icon("plus"), style="Icon.TButton", command=add, takefocus=False)
            plus.pack(side="right")
            Tooltip(plus, add_tip)
        self.filter_entry: ttk.Entry | None = None
        self.hint_labels: list[ttk.Label] = []
        if filter_var is not None:
            self._build_filter(filter_var)
        # Before the tree in packing order, so a short window shrinks the tree, not the footer.
        self.foot: ttk.Frame | None = None
        if footer:
            self.foot = ttk.Frame(self, padding=(14, 7))
            self.foot.pack(fill="x", side="bottom")
            ttk.Frame(self, style="Line.TFrame", height=1).pack(fill="x", side="bottom")
        self.body = ttk.Frame(self, padding=(6, 0, 0, 6))
        self.body.pack(fill="both", expand=True, side="top")
        self.tree = ttk.Treeview(self.body, show="tree", columns=("badge",), selectmode="browse")
        self.tree.column("#0", width=200, stretch=True)
        self.tree.column("badge", width=30, stretch=False, anchor="center")
        scrollbar = ttk.Scrollbar(self.body, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=theme.autohide(scrollbar, side="right", fill="y"))
        self.tree.pack(side="left", fill="both", expand=True)

    def _build_filter(self, filter_var: tk.StringVar) -> None:
        box = ttk.Frame(self, padding=(10, 0, 10, 8))
        box.pack(fill="x", side="top")
        entry = self.filter_entry = ttk.Entry(box, textvariable=filter_var)
        entry.pack(fill="x")
        hint = ttk.Label(entry, text=icon("search"), style="PanelIcon.TLabel", background=theme.palette().bg3)
        hint_text = ttk.Label(entry, text="Filter", style="Faint.TLabel", background=theme.palette().bg3)
        self.hint_labels = [hint, hint_text]

        def update_hint(*_args) -> None:
            if filter_var.get() or entry.focus_get() is entry:
                hint.place_forget()
                hint_text.place_forget()
            else:
                hint.place(x=8, rely=0.5, anchor="w")
                hint_text.place(x=28, rely=0.5, anchor="w")

        for widget in (hint, hint_text):
            widget.bind("<Button-1>", lambda _e: entry.focus_set())
        entry.bind("<FocusIn>", update_hint, add=True)
        entry.bind("<FocusOut>", update_hint, add=True)
        entry.bind("<Escape>", lambda _e: filter_var.set(""))
        filter_var.trace_add("write", update_hint)
        update_hint()

    def apply_palette(self) -> None:
        """Recolour what styles do not reach: the filter hint sits on the entry's field colour."""
        background = theme.palette().bg3
        for label in self.hint_labels:
            label.configure(background=background)


def hover_row(widgets: list[Any], normal: str, hover: str, command: Callable[[], object]) -> None:
    """Make a group of classic Tk widgets act as one clickable row with a hover fill."""

    def paint(color: str) -> None:
        for widget in widgets:
            try:
                widget.configure(background=color)
            except tk.TclError:
                pass

    for widget in widgets:
        widget.bind("<Enter>", lambda _e: paint(hover), add="+")
        widget.bind("<Leave>", lambda _e: paint(normal), add="+")
        widget.bind("<Button-1>", lambda _e: command(), add="+")
        try:
            widget.configure(cursor="hand2")
        except tk.TclError:
            pass
