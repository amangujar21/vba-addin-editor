"""Browser-style tabs: one per open add-in, under the header.

The strip only draws and reports: ``on_select(key)``, ``on_close(key)``,
``on_move(key, index)`` and ``on_new()`` are the window's. A tab is dragged
along the strip to reorder it. ``tab_at(x_root, y_root)`` and
``show_drop(key)`` let other views drop things (a ribbon control, a macro)
onto a tab. Event handlers take event-like objects (``x_root``/``y_root``),
which tests pass.
"""

from __future__ import annotations

import tkinter as tk
from collections.abc import Callable
from dataclasses import dataclass

from vba_addin_editor.ui import theme
from vba_addin_editor.ui.widgets import Tooltip

DRAG_START_PX = 6
MAX_TITLE = 28


@dataclass(frozen=True)
class TabInfo:
    key: str
    title: str
    tip: str = ""
    dirty: bool = False
    extension: str = ""
    waiting: bool = False  # a save is queued until Office lets go of the file


class _Chip(tk.Frame):
    def __init__(self, strip: TabStrip, info: TabInfo) -> None:
        super().__init__(strip.row, highlightthickness=0, borderwidth=0)
        self.strip = strip
        self.info = info
        self.active = False
        self.hover = False
        self.drop = False
        self.top_bar = tk.Frame(self, height=2)
        self.top_bar.pack(side="top", fill="x")
        body = self.body = tk.Frame(self)
        body.pack(side="top", fill="both", expand=True)
        self.tag = tk.Label(body, font=theme.TAG_FONT, width=5, pady=1)
        self.tag.pack(side="left", padx=(10, 7), pady=6)
        title = info.title if len(info.title) <= MAX_TITLE else info.title[: MAX_TITLE - 1] + "…"
        self.title = tk.Label(body, text=title, font="TkDefaultFont", anchor="w")
        self.title.pack(side="left")
        # A dot while there are unsaved changes; the close button on hover (as browsers do).
        self.close_btn = tk.Label(body, text="", font="TkDefaultFont", width=2, cursor="hand2")
        self.close_btn.pack(side="left", padx=(6, 6))
        self.tooltip = Tooltip(self.title, info.tip)
        for widget in (self, body, self.tag, self.title, self.close_btn, self.top_bar):
            widget.bind("<Enter>", self._enter, add=True)
            widget.bind("<Leave>", self._leave, add=True)
            widget.bind("<Button-2>", lambda _e: strip.on_close(self.info.key))
            if widget is not self.close_btn:
                widget.bind("<ButtonPress-1>", lambda e: strip.press(self.info.key, e))
                widget.bind("<B1-Motion>", strip.motion)
                widget.bind("<ButtonRelease-1>", strip.release)
        self.close_btn.bind("<Button-1>", lambda _e: strip.on_close(self.info.key))
        self.apply_palette()

    def _enter(self, _event=None) -> None:
        self.hover = True
        self.apply_palette()

    def _leave(self, _event=None) -> None:
        self.hover = False
        self.apply_palette()

    def mark(self) -> str:
        if self.hover:
            return "✕"
        if self.info.waiting:
            return "⏳"
        if self.info.dirty:
            return "●"
        return "✕" if self.active else ""

    def apply_palette(self) -> None:
        colors = theme.palette()
        bg = colors.bg2 if self.active else colors.bg3 if (self.hover or self.drop) else colors.bg1
        for widget in (self, self.body):
            widget.configure(background=bg)
        bar = colors.accent if (self.active or self.drop) else bg
        self.top_bar.configure(background=bar)
        tag_bg, tag_fg = colors.file_tag(self.info.extension)
        self.tag.configure(
            text=self.info.extension.lstrip(".").upper() or "FILE", background=tag_bg, foreground=tag_fg
        )
        self.title.configure(background=bg, foreground=colors.text if self.active else colors.muted)
        mark = self.mark()
        dirty_mark = mark in ("●", "⏳")
        self.close_btn.configure(
            text=mark,
            background=bg,
            foreground=colors.warn if dirty_mark else colors.faint,
        )


class TabStrip(tk.Frame):
    def __init__(
        self,
        master,
        *,
        on_select: Callable[[str], object],
        on_close: Callable[[str], object],
        on_move: Callable[[str, int], object],
        on_new: Callable[[], object],
    ) -> None:
        colors = theme.palette()
        super().__init__(master, background=colors.bg0)
        self.on_select = on_select
        self.on_close = on_close
        self.on_move = on_move
        self.on_new = on_new
        self.row = tk.Frame(self, background=colors.bg0)
        self.row.pack(side="left", fill="y")
        self.new_btn = tk.Label(
            self, text="+", font=theme.BOLD_FONT, width=3, cursor="hand2", background=colors.bg0,
            foreground=colors.faint,
        )
        self.new_btn.pack(side="left", padx=(4, 0))
        self.new_btn.bind("<Button-1>", lambda _e: self.on_new())
        Tooltip(self.new_btn, "Open another add-in   Ctrl+O")
        self.line = tk.Frame(self, height=1, background=colors.border)
        self.chips: dict[str, _Chip] = {}
        self.order: list[str] = []
        self.active: str | None = None
        self._drag: dict | None = None
        self._drop: str | None = None

    # -- contents ------------------------------------------------------------------------

    def set_tabs(self, tabs: list[TabInfo], active: str | None) -> None:
        keys = [t.key for t in tabs]
        same = keys == self.order and all(self.chips[t.key].info == t for t in tabs)
        if not same:
            for chip in self.chips.values():
                chip.destroy()
            self.chips = {}
            for info in tabs:
                chip = _Chip(self, info)
                chip.pack(side="left", fill="y", padx=(0, 1))
                self.chips[info.key] = chip
            self.order = keys
        self.active = active
        for key, chip in self.chips.items():
            chip.active = key == active
            chip.drop = key == self._drop
            chip.apply_palette()

    def apply_palette(self) -> None:
        colors = theme.palette()
        self.configure(background=colors.bg0)
        self.row.configure(background=colors.bg0)
        self.new_btn.configure(background=colors.bg0, foreground=colors.faint)
        for chip in self.chips.values():
            chip.apply_palette()

    # -- finding a tab under the pointer -------------------------------------------------------

    def tab_at(self, x_root: int, y_root: int) -> str | None:
        """The tab under screen point (x_root, y_root), if any."""
        for key, chip in self.chips.items():
            try:
                x, y = chip.winfo_rootx(), chip.winfo_rooty()
                w, h = chip.winfo_width(), chip.winfo_height()
            except tk.TclError:
                continue
            if x <= x_root < x + w and y <= y_root < y + h:
                return key
        return None

    def _index_at(self, x_root: int) -> int:
        """Where a tab dragged to ``x_root`` goes: before the first tab whose middle is right of it."""
        for index, key in enumerate(self.order):
            chip = self.chips[key]
            if x_root < chip.winfo_rootx() + chip.winfo_width() // 2:
                return index
        return len(self.order)

    def show_drop(self, key: str | None) -> None:
        """Highlight ``key`` as the tab something is about to be dropped on (None: none)."""
        if key == self._drop:
            return
        self._drop = key
        for k, chip in self.chips.items():
            chip.drop = k == key
            chip.apply_palette()

    # -- clicking and dragging ---------------------------------------------------------------

    def press(self, key: str, event) -> str:
        self._drag = {"key": key, "x": event.x_root, "moving": False}
        if key != self.active:
            self.on_select(key)
        return "break"

    def motion(self, event) -> None:
        drag = self._drag
        if drag is None:
            return
        if not drag["moving"] and abs(event.x_root - drag["x"]) < DRAG_START_PX:
            return
        drag["moving"] = True
        self.configure(cursor="sb_h_double_arrow")

    def release(self, event) -> None:
        drag, self._drag = self._drag, None
        self.configure(cursor="")
        if drag is None or not drag["moving"] or drag["key"] not in self.order:
            return
        index = self._index_at(event.x_root)
        old = self.order.index(drag["key"])
        if index > old:
            index -= 1  # the tab itself leaves its old place
        if index != old:
            self.on_move(drag["key"], index)
