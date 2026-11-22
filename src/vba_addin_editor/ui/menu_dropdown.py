"""The drop-down of a ribbon menu, drawn under it in the Ribbon view as Office shows it.

Clicking a menu (or split button) in the ribbon drawing opens this list of what
the menu holds. Clicking a row opens the macro it runs (a row with no macro
yet asks for one); a submenu opens beside its row. Hovering the line between
two rows offers to add a macro, a submenu or a separator right there, and the
row at the bottom adds them at the end. Right-click a row for the rest.

Like the rest of the designer this edits nothing itself: every change goes
through ``RibbonDesigner`` (one ``ribbon_service`` text edit, one undo step),
and the panels are redrawn from the draft afterwards.
"""

from __future__ import annotations

import re
import tkinter as tk
from collections.abc import Callable
from dataclasses import dataclass
from tkinter import ttk
from typing import Any

from vba_addin_editor.services import ribbon_service as rs
from vba_addin_editor.ui import theme
from vba_addin_editor.ui.ribbon_preview import element_label

ROW_H = 26
SEP_H = 9
PAD = 4  # above the first row
FOOT_H = 32
HOT = 5  # px either side of a line between two rows that offers to add there
MIN_W = 280
MAX_W = 480
HOVER_OPEN_MS = 280
_ACTION = {"button", "toggleButton", "checkBox"}
_SUBMENUS = {"menu", "splitButton"}
# (kind passed to the designer, chip text)
CHIPS = (("button", "+ Macro"), ("menu", "+ Submenu"), ("separator", "+ Separator"))


def menu_of(elements: list[rs.RibbonElement], element: rs.RibbonElement) -> rs.RibbonElement | None:
    """The <menu> whose children the drop-down of ``element`` lists (None: it has no drop-down)."""
    local = rs.local_name(element.name)
    if local == "menu":
        return element if element.complete else None
    if local == "splitButton":
        return next(
            (
                e
                for e in elements
                if e.parent == element.index and rs.local_name(e.name) == "menu" and e.complete
            ),
            None,
        )
    return None


def menu_items(elements: list[rs.RibbonElement], menu: rs.RibbonElement) -> list[rs.RibbonElement]:
    """The rows of ``menu``'s drop-down, in order."""
    return [e for e in elements if e.parent == menu.index and e.complete]


def new_macro_name(typed: str) -> str:
    """A Sub name for what was typed: "format table" -> "FormatTable"; "" when none fits."""
    text = typed.strip()
    if " " in text:
        text = "".join(word[:1].upper() + word[1:] for word in text.split())
    text = re.sub(r"[^A-Za-z0-9_]", "", text)
    text = re.sub(r"^[^A-Za-z]+", "", text)
    return text if rs.is_valid_procedure_name(text) else ""


@dataclass
class Row:
    element: rs.RibbonElement | None  # None: the "empty menu" note
    kind: str  # "item" | "submenu" | "separator" | "note"
    top: int
    bottom: int
    label: str = ""
    macro: str = ""  # what onAction names
    state: str = ""  # "ok" | "missing" | "none" (an action with no macro) | ""


class DropPanel(tk.Frame):
    """One level of the drop-down: the rows of one menu."""

    def __init__(self, dropdown: MenuDropdown, level: int) -> None:
        super().__init__(dropdown.host, highlightthickness=1, bd=0)
        self.dropdown = dropdown
        self.level = level
        self.canvas = tk.Canvas(self, highlightthickness=0, bd=0, cursor="arrow", takefocus=True)
        self.canvas.pack(fill="both", expand=True)
        self.owner: rs.RibbonElement | None = None  # the control this panel drops from
        self.menu: rs.RibbonElement | None = None
        self.rows: list[Row] = []
        self.width = MIN_W
        self.height = FOOT_H
        self.x = self.y = 0
        self.hover: tuple[str, int] | None = None  # ("row", i) | ("slot", k)
        self.chips: list[tuple[int, int, int, int, str, int]] = []  # x1, y1, x2, y2, kind, slot
        self._after: str | None = None
        self._rename: ttk.Entry | None = None
        self.finish_rename: Callable[[bool], str] | None = None  # Enter (True) / Esc (False) in the rename box
        canvas = self.canvas
        canvas.bind("<Motion>", lambda e: self.hover_at(e.x, e.y))
        canvas.bind("<Leave>", lambda _e: self._set_hover(None))
        canvas.bind("<Button-1>", lambda e: self.press(e.x, e.y))
        canvas.bind("<Button-3>", self._on_right_click)
        canvas.bind("<Escape>", lambda _e: self.dropdown.close())
        self.apply_palette()

    def apply_palette(self) -> None:
        colors = theme.palette()
        self.configure(
            background=colors.bg3, highlightbackground=colors.border_strong, highlightcolor=colors.border_strong
        )
        self.canvas.configure(background=colors.bg3)

    def destroy(self) -> None:
        self._cancel_timer()
        super().destroy()

    # -- content ---------------------------------------------------------------

    def _measure(self, text: str, font: str = "TkDefaultFont") -> int:
        return int(self.tk.call("font", "measure", font, text))

    def fill(self, owner: rs.RibbonElement, menu: rs.RibbonElement) -> None:
        """Lay out the rows of ``menu`` (dropped from ``owner``)."""
        designer = self.dropdown.designer
        self.owner, self.menu = owner, menu
        rows: list[Row] = []
        y = PAD
        label_w = macro_w = 0
        for element in menu_items(designer.elements, menu):
            local = rs.local_name(element.name)
            if local in ("menuSeparator", "separator"):
                rows.append(Row(element, "separator", y, y + SEP_H))
                y += SEP_H
                continue
            row = Row(element, "submenu" if local in _SUBMENUS else "item", y, y + ROW_H, element_label(element) or local)
            if row.kind == "item":
                row.macro = rs.unescape(element.attr("onAction") or "").strip()
                if row.macro:
                    resolution = designer.resolve(row.macro)
                    row.state = "ok" if resolution.status == rs.STATUS_OK else "missing"
                elif local in _ACTION:
                    row.state = "none"
            rows.append(row)
            label_w = max(label_w, self._measure(row.label))
            macro_w = max(macro_w, self._measure(self._macro_text(row), theme.SMALL_FONT))
            y += ROW_H
        if not rows:
            rows.append(Row(None, "note", y, y + ROW_H, "Empty menu: add a macro below."))
            y += ROW_H
        self.rows = rows
        self.width = max(MIN_W, min(MAX_W, 34 + label_w + 28 + macro_w + 26))
        self.height = y + 4 + FOOT_H
        if self.hover is not None and self.hover[0] == "row" and self.hover[1] >= len(rows):
            self.hover = None

    @staticmethod
    def _macro_text(row: Row) -> str:
        if row.state == "none":
            return "no macro"
        if row.state == "missing":
            return f"⚠ {row.macro}"
        return row.macro

    def items(self) -> list[Row]:
        """The rows that are controls (not the empty note)."""
        return [row for row in self.rows if row.element is not None]

    def row_for(self, element: rs.RibbonElement) -> int | None:
        return next((i for i, row in enumerate(self.rows) if row.element is not None and row.element.index == element.index), None)

    # -- drawing -----------------------------------------------------------------

    def draw(self) -> None:
        colors = theme.palette()
        canvas = self.canvas
        canvas.delete("all")
        self.chips.clear()
        width = self.width
        selected = self.dropdown.designer.selected_indices()
        opened = self.dropdown.opened_index(self.level)
        for i, row in enumerate(self.rows):
            if row.kind == "separator":
                middle = (row.top + row.bottom) // 2
                canvas.create_line(30, middle, width - 8, middle, fill=colors.border_strong)
                continue
            if row.kind == "note":
                canvas.create_text(
                    12, (row.top + row.bottom) // 2, anchor="w", text=row.label, font=theme.SMALL_FONT, fill=colors.faint
                )
                continue
            element = row.element
            assert element is not None
            fill = None
            if element.index in selected or element.index == opened:
                fill = colors.accent_bg
            if self.hover == ("row", i):
                fill = colors.bg4
            if fill is not None:
                canvas.create_rectangle(2, row.top, width - 3, row.bottom - 1, fill=fill, outline="")
            middle = (row.top + row.bottom) // 2
            self._icon(element, 10, middle)
            canvas.create_text(34, middle, anchor="w", text=row.label, font="TkDefaultFont", fill=colors.text)
            if row.kind == "submenu":
                canvas.create_text(width - 12, middle, anchor="e", text="›", font=theme.BOLD_FONT, fill=colors.muted)
            elif row.state:
                tone = {"ok": colors.faint, "missing": colors.warn, "none": colors.warn}[row.state]
                canvas.create_text(
                    width - 12, middle, anchor="e", text=self._macro_text(row), font=theme.SMALL_FONT, fill=tone
                )
        if self.hover is not None and self.hover[0] == "slot":
            self._draw_slot(self.hover[1])
        self._draw_footer()

    def _icon(self, element: rs.RibbonElement, x: int, middle: int) -> None:
        colors = theme.palette()
        photo = self.dropdown.designer.photo_for(element)
        if photo is not None:
            factor = max(photo.width(), photo.height()) // 16
            shown = photo.subsample(factor) if factor > 1 else photo
            self.dropdown.keep(shown)
            self.canvas.create_image(x + 7, middle, image=shown)
            return
        local = rs.local_name(element.name)
        if local in ("checkBox", "toggleButton"):
            self.canvas.create_rectangle(x + 1, middle - 6, x + 13, middle + 6, outline=colors.faint)
            return
        self.canvas.create_rectangle(x, middle - 7, x + 14, middle + 7, outline=colors.border_strong, fill=colors.bg4)

    def slot_y(self, slot: int) -> int:
        items = self.items()
        if slot < len(items):
            return items[slot].top
        return items[-1].bottom if items else self.rows[0].bottom

    def _chip(self, x: int, middle: int, text: str, kind: str, slot: int, *, solid: bool, right: bool = False) -> int:
        """Draw a chip at ``x`` (its right edge when ``right``); returns the x past it."""
        colors = theme.palette()
        width = self._measure(text, theme.SMALL_FONT) + 14
        x1 = x - width if right else x
        x2 = x1 + width
        y1, y2 = middle - 9, middle + 9
        if solid:
            self.canvas.create_rectangle(x1, y1, x2, y2, fill=colors.accent, outline=colors.accent)
            self.canvas.create_text((x1 + x2) // 2, middle, text=text, font=theme.SMALL_FONT, fill=colors.on_accent)
        else:
            self.canvas.create_rectangle(x1, y1, x2, y2, fill=colors.bg3, outline=colors.border_strong)
            self.canvas.create_text((x1 + x2) // 2, middle, text=text, font=theme.SMALL_FONT, fill=colors.accent)
        self.chips.append((x1, y1, x2, y2, kind, slot))
        return x1 if right else x2

    def _draw_slot(self, slot: int) -> None:
        """The line between two rows, with what can be added there."""
        colors = theme.palette()
        y = self.slot_y(slot)
        self.canvas.create_line(6, y, self.width - 6, y, fill=colors.accent, width=2)
        x = self.width - 8
        for kind, text in reversed(CHIPS):
            x = self._chip(x, y, text, kind, slot, solid=True, right=True) - 4

    def _draw_footer(self) -> None:
        colors = theme.palette()
        top = self.height - FOOT_H
        self.canvas.create_line(0, top, self.width, top, fill=colors.border)
        x = 10
        slot = len(self.items())
        for kind, text in CHIPS:
            x = self._chip(x, top + FOOT_H // 2, text, kind, slot, solid=False) + 6

    # -- pointer -----------------------------------------------------------------

    def chip_at(self, x: int, y: int) -> tuple[str, int] | None:
        return next(((kind, slot) for x1, y1, x2, y2, kind, slot in self.chips if x1 <= x < x2 and y1 <= y < y2), None)

    def target_at(self, x: int, y: int) -> tuple[str, int] | None:
        """("slot", k) near the line above the k-th row, ("row", i) on a row, else None."""
        items = self.items()
        if 0 <= x < self.width and len(items) > 1:
            for slot in range(len(items)):
                if abs(y - self.slot_y(slot)) <= HOT and not (slot == 0 and y < self.slot_y(0) - 1):
                    return ("slot", slot)
        return next(
            (("row", i) for i, row in enumerate(self.rows) if row.top <= y < row.bottom and row.element is not None),
            None,
        )

    def hover_at(self, x: int, y: int) -> None:
        if self.chip_at(x, y) is not None:
            return  # keep the chips of the line being pointed at
        self._set_hover(self.target_at(x, y))

    def _set_hover(self, target: tuple[str, int] | None) -> None:
        if target == self.hover:
            return
        self.hover = target
        self.draw()
        self._cancel_timer()
        if target is not None and target[0] == "row":
            row = self.rows[target[1]]
            self._after = self.after(HOVER_OPEN_MS, lambda: self._hovered(row))

    def _hovered(self, row: Row) -> None:
        self._after = None
        if row.kind == "submenu" and row.element is not None:
            self.dropdown.open_submenu(self.level, row.element)
        else:
            self.dropdown.close_below(self.level)

    def _cancel_timer(self) -> None:
        if self._after is not None:
            try:
                self.after_cancel(self._after)
            except tk.TclError:
                pass
            self._after = None

    def press(self, x: int, y: int) -> None:
        """A click at canvas (x, y)."""
        self.canvas.focus_set()
        chip = self.chip_at(x, y)
        if chip is not None:
            self.dropdown.add(self.level, chip[1], chip[0], y)
            return
        target = self.target_at(x, y)
        if target is not None and target[0] == "row":
            self.dropdown.activate(self.level, self.rows[target[1]])

    def _on_right_click(self, event) -> str:
        row = next(
            (r for r in self.rows if r.top <= event.y < r.bottom and r.element is not None), None
        )
        if row is not None:
            menu = self.dropdown.prepare_context_menu(self.level, row)
            if menu is not None:
                menu.tk_popup(event.x_root, event.y_root)
        return "break"

    # -- renaming in place ------------------------------------------------------

    def begin_rename(self, element: rs.RibbonElement) -> ttk.Entry | None:
        """An entry over the row's label; Enter or leaving it renames, Esc does not."""
        index = self.row_for(element)
        if index is None or self.rows[index].kind not in ("item", "submenu"):
            return None
        self.end_rename()
        row = self.rows[index]
        var = tk.StringVar(master=self, value=row.label)
        entry = ttk.Entry(self.canvas, textvariable=var)
        self.canvas.create_window(30, row.top + 2, anchor="nw", window=entry, width=self.width - 60, height=ROW_H - 4)
        self._rename = entry
        done = {"finished": False}

        def finish(save: bool) -> str:
            if done["finished"]:
                return "break"
            done["finished"] = True
            value = var.get().strip()
            self._rename = self.finish_rename = None
            self.after_idle(entry.destroy)
            if save and value and value != row.label:
                self.dropdown.designer.rename_control(element, value)
            else:
                self.canvas.focus_set()
            return "break"

        self.finish_rename = finish
        entry.bind("<Return>", lambda _e: finish(True))
        entry.bind("<Escape>", lambda _e: finish(False))
        entry.bind("<FocusOut>", lambda _e: finish(True))
        entry.focus_set()
        entry.select_range(0, "end")
        return entry

    def end_rename(self) -> None:
        if self._rename is not None:
            entry, self._rename = self._rename, None
            entry.destroy()


class MacroPicker(tk.Frame):
    """A filterable list of the macros a button can run, plus "new macro" for a name that is not one."""

    def __init__(self, dropdown: MenuDropdown) -> None:
        super().__init__(dropdown.host, highlightthickness=1, bd=0, padx=8, pady=8)
        self.dropdown = dropdown
        self.title = tk.Label(self, anchor="w", font=theme.SMALL_BOLD_FONT)
        self.title.pack(fill="x", side="top")
        self.var = tk.StringVar(master=self)
        self.entry = ttk.Entry(self, textvariable=self.var, width=36)
        self.entry.pack(fill="x", side="top", pady=(6, 6))
        self.hint = tk.Label(self, anchor="w", font=theme.SMALL_FONT, text="Enter adds  ·  Esc cancels")
        self.hint.pack(fill="x", side="bottom", pady=(6, 0))
        self.list = tk.Listbox(
            self, height=10, activestyle="none", exportselection=False, highlightthickness=0, bd=0,
            font="TkDefaultFont",
        )
        self.list.pack(fill="both", expand=True, side="top")
        self.choices: list[tuple[str, str]] = []  # (name to use, label typed for a new one)
        self._on_pick: Callable[[str, str], None] | None = None
        self.var.trace_add("write", lambda *_a: self.fill())
        self.entry.bind("<Down>", lambda _e: self._step(1))
        self.entry.bind("<Up>", lambda _e: self._step(-1))
        for widget in (self.entry, self.list):
            widget.bind("<Return>", lambda _e: self._break(self.pick()))
            widget.bind("<Escape>", lambda _e: self._break(self.close()))
        self.list.bind("<ButtonRelease-1>", lambda _e: self.pick())
        self.apply_palette()

    @staticmethod
    def _break(_result=None) -> str:
        return "break"

    def apply_palette(self) -> None:
        colors = theme.palette()
        self.configure(background=colors.bg3, highlightbackground=colors.border_strong, highlightcolor=colors.border_strong)
        self.title.configure(background=colors.bg3, foreground=colors.muted)
        self.hint.configure(background=colors.bg3, foreground=colors.faint)
        self.list.configure(
            background=colors.bg2, foreground=colors.text, selectbackground=colors.select,
            selectforeground=colors.select_text,
        )

    @property
    def shown(self) -> bool:
        return bool(self.winfo_manager())

    def open(self, title: str, x: int, y: int, on_pick: Callable[[str, str], None], initial: str = "") -> None:
        self._on_pick = on_pick
        self.title.configure(text=title.upper())
        self.var.set(initial)
        self.fill()
        host = self.dropdown.host
        width, height = 320, 300
        x = max(4, min(x, host.winfo_width() - width - 4)) if host.winfo_width() > width else 4
        y = max(4, min(y, host.winfo_height() - height - 4)) if host.winfo_height() > height else 4
        self.place(x=x, y=y, width=width, height=height)
        self.lift()
        self.entry.focus_set()
        self.entry.select_range(0, "end")

    def close(self) -> None:
        self._on_pick = None
        self.place_forget()

    def fill(self) -> None:
        typed = self.var.get().strip()
        query = typed.casefold()
        macros = self.dropdown.designer.macro_choices()
        choices: list[tuple[str, str]] = []
        shown: list[str] = []
        for name, module in macros:
            if query and query not in name.casefold():
                continue
            choices.append((name, ""))
            shown.append(f"{name}      {module}")
        # A name that is no macro yet can become one: after the matches, so Enter takes a match.
        fresh = new_macro_name(typed)
        if fresh and fresh.casefold() not in {name.casefold() for name, _m in macros}:
            choices.append((fresh, typed if " " in typed else ""))
            shown.append(f"+ New macro  {fresh}")
        self.choices = choices
        self.list.delete(0, "end")
        for text in shown:
            self.list.insert("end", text)
        if choices:
            self.list.selection_set(0)
            self.list.see(0)

    def _step(self, delta: int) -> str:
        if self.choices:
            current = self.list.curselection()
            index = max(0, min(len(self.choices) - 1, (current[0] if current else -1) + delta))
            self.list.selection_clear(0, "end")
            self.list.selection_set(index)
            self.list.see(index)
        return "break"

    def pick(self, index: int | None = None) -> None:
        if index is None:
            current = self.list.curselection()
            index = current[0] if current else (0 if self.choices else None)
        if index is None or not 0 <= index < len(self.choices) or self._on_pick is None:
            return
        on_pick = self._on_pick
        name, label = self.choices[index]
        self.close()
        on_pick(name, label)


class MenuDropdown:
    """The open drop-down: a chain of panels, the menu clicked and the submenus opened from it."""

    def __init__(self, designer: Any) -> None:
        self.designer = designer
        self.host: tk.Misc = designer
        self.chain: list[str] = []  # keys of the controls each open panel drops from
        self.panels: list[DropPanel] = []
        self.picker = MacroPicker(self)
        self._photos: list[tk.PhotoImage] = []
        self._pending_rename: str | None = None
        self.context_menu = tk.Menu(designer, tearoff=0)
        self._context: tuple[int, rs.RibbonElement] | None = None
        menu = self.context_menu
        menu.add_command(label="Open Macro", command=lambda: self._on_context("open"))
        menu.add_command(label="Change Macro…", command=lambda: self._on_context("change"))
        menu.add_command(label="Rename", command=lambda: self._on_context("rename"))
        menu.add_separator()
        menu.add_command(label="Insert Macro Above…", command=lambda: self._on_context("above"))
        menu.add_command(label="Insert Macro Below…", command=lambda: self._on_context("below"))
        menu.add_command(label="Insert Submenu Below", command=lambda: self._on_context("submenu"))
        menu.add_command(label="Insert Separator Below", command=lambda: self._on_context("separator"))
        menu.add_separator()
        menu.add_command(label="Move Up", command=lambda: self._on_context("up"))
        menu.add_command(label="Move Down", command=lambda: self._on_context("down"))
        menu.add_command(label="Show in Designer", command=lambda: self._on_context("show"))
        menu.add_separator()
        menu.add_command(label="Delete", command=lambda: self._on_context("delete"))

    # -- state -------------------------------------------------------------------

    @property
    def is_open(self) -> bool:
        return bool(self.chain)

    @property
    def root_key(self) -> str | None:
        return self.chain[0] if self.chain else None

    def keep(self, photo: tk.PhotoImage) -> None:
        self._photos.append(photo)

    def apply_palette(self) -> None:
        for panel in self.panels:
            panel.apply_palette()
        self.picker.apply_palette()
        self.sync()

    def set_root(self, key: str | None) -> None:
        """Drop down from the control with ``key`` (None closes); takes effect on the next sync."""
        self.chain = [key] if key else []

    def opened_index(self, level: int) -> int | None:
        """The element index of the submenu open beside ``level``."""
        if level + 1 < len(self.panels) and level + 1 < len(self.chain):
            owner = self.panels[level + 1].owner
            return owner.index if owner is not None else None
        return None

    def close(self) -> None:
        self.chain = []
        self.picker.close()
        for panel in self.panels:
            panel.end_rename()
            panel._cancel_timer()
            panel.hover = None
            panel.place_forget()

    def close_below(self, level: int) -> None:
        if len(self.chain) > level + 1:
            del self.chain[level + 1 :]
            self.sync()

    def contains(self, widget: Any) -> bool:
        """Whether ``widget`` is part of the drop-down (a panel, the picker, or inside them)."""
        name = str(widget)
        return any(name == str(w) or name.startswith(str(w) + ".") for w in (*self.panels, self.picker))

    def focus(self) -> None:
        if self.panels and self.is_open:
            self.panels[0].canvas.focus_set()

    # -- layout ------------------------------------------------------------------

    def _panel(self, level: int) -> DropPanel:
        while len(self.panels) <= level:
            self.panels.append(DropPanel(self, len(self.panels)))
        return self.panels[level]

    def sync(self) -> None:
        """Redraw the open panels from the designer's elements (after any edit or redraw)."""
        designer = self.designer
        self._photos.clear()
        by_key = {designer.key_of(e): e for e in designer.elements}
        placed = 0
        for level, key in enumerate(list(self.chain)):
            owner = by_key.get(key)
            menu = menu_of(designer.elements, owner) if owner is not None else None
            if owner is None or menu is None:
                break
            panel = self._panel(level)
            panel.fill(owner, menu)
            where = self._anchor(level, owner)
            if where is None:
                break
            panel.x, panel.y = where
            placed = level + 1
        if placed == 0:
            if self.chain or any(p.winfo_manager() for p in self.panels):
                self.close()
            return
        del self.chain[placed:]
        for level, panel in enumerate(self.panels):
            if level < placed:
                panel.place(x=panel.x, y=panel.y, width=panel.width + 2, height=panel.height + 2)
                panel.lift()
                panel.draw()
            else:
                panel.end_rename()
                panel.hover = None
                panel.place_forget()
        if self.picker.shown:
            self.picker.lift()
        if self._pending_rename is not None:
            key, self._pending_rename = self._pending_rename, None
            target = by_key.get(key)
            panel = next((p for p in self.panels[:placed] if target is not None and p.row_for(target) is not None), None)
            if panel is not None and target is not None:
                panel.begin_rename(target)

    def _anchor(self, level: int, owner: rs.RibbonElement) -> tuple[int, int] | None:
        """Where the panel of ``level`` goes, in the designer's coordinates."""
        host = self.host
        panel = self.panels[level]
        host_w = max(host.winfo_width(), panel.width + 8)
        host_h = max(host.winfo_height(), 1)
        if level == 0:
            box = self.designer.preview.cell_box(owner.index)
            if box is None:
                return None
            canvas = self.designer.preview.canvas
            ox = canvas.winfo_rootx() - host.winfo_rootx()
            oy = canvas.winfo_rooty() - host.winfo_rooty()
            x, y = ox + box[0], oy + box[3] + 2
        else:
            parent = self.panels[level - 1]
            index = parent.row_for(owner)
            if index is None:
                return None
            x = parent.x + parent.width - 2
            if x + panel.width + 4 > host_w:
                x = parent.x - panel.width + 2
            y = parent.y + parent.rows[index].top - PAD
        x = max(2, min(x, host_w - panel.width - 4))
        if y + panel.height + 4 > host_h:
            y = max(2, host_h - panel.height - 4)
        return x, y

    # -- what the panels ask for ----------------------------------------------------

    def open_submenu(self, level: int, element: rs.RibbonElement) -> None:
        key = self.designer.key_of(element)
        if self.chain[level + 1 : level + 2] == [key] and len(self.chain) == level + 2:
            return
        self.chain = self.chain[: level + 1] + [key]
        self.sync()

    def activate(self, level: int, row: Row) -> None:
        """A click on a row: a submenu opens; anything else runs ``designer.open_item``."""
        element = row.element
        if element is None or row.kind == "separator":
            return
        if row.kind == "submenu":
            self.open_submenu(level, element)
            return
        if row.state == "none":
            self.change_macro(level, element)
            return
        self.designer.open_item(element)

    def add(self, level: int, slot: int, kind: str, y: int | None = None) -> None:
        """A chip: add a macro (through the picker), a submenu or a separator at ``slot``."""
        panel = self.panels[level]
        menu = panel.menu
        if menu is None:
            return
        if kind == "button":
            anchor_y = panel.y + (y if y is not None else panel.slot_y(slot))
            self.pick_macro(
                "Add a macro",
                self._beside(panel),
                anchor_y,
                lambda name, label: self.designer.insert_at(menu, slot, "button", macro=name, label=label),
            )
            return
        ident = self.designer.insert_at(menu, slot, kind)
        if ident and kind == "menu":
            self.rename_later(ident)

    def rename_later(self, key: str) -> None:
        """Start renaming the control with ``key`` once the panels show it."""
        self._pending_rename = key
        self.sync()

    def change_macro(self, level: int, element: rs.RibbonElement) -> None:
        panel = self.panels[level]
        index = panel.row_for(element)
        top = panel.rows[index].top if index is not None else 0
        self.pick_macro(
            f"Macro for {element_label(element) or 'this control'}",
            self._beside(panel),
            panel.y + top,
            lambda name, _label: self.designer.change_macro(element, name),
            initial=rs.unescape(element.attr("onAction") or ""),
        )

    def _beside(self, panel: DropPanel) -> int:
        right = panel.x + panel.width + 6
        if right + 320 > self.host.winfo_width():
            return max(4, panel.x - 326)
        return right

    def pick_macro(self, title: str, x: int, y: int, on_pick: Callable[[str, str], Any], initial: str = "") -> None:
        self.picker.open(title, x, y, on_pick, initial)

    # -- right-click ---------------------------------------------------------------

    def prepare_context_menu(self, level: int, row: Row) -> tk.Menu | None:
        element = row.element
        if element is None:
            return None
        self._context = (level, element)
        local = rs.local_name(element.name)
        named = row.kind in ("item", "submenu")
        free = rs.is_free(self.designer.elements, element)
        states = {
            "Open Macro": bool(row.macro),
            "Change Macro…": local in _ACTION,
            "Rename": named,
            "Move Up": free,
            "Move Down": free,
            "Delete": free,
        }
        menu = self.context_menu
        for index in range((menu.index("end") or 0) + 1):
            if menu.type(index) == "command":
                label = str(menu.entrycget(index, "label"))
                menu.entryconfigure(index, state="normal" if states.get(label, True) else "disabled")
        return menu

    def _on_context(self, action: str) -> None:
        if self._context is None:
            return
        level, element = self._context
        self.context_action(level, element, action)

    def context_action(self, level: int, element: rs.RibbonElement, action: str) -> None:
        """What a context-menu entry does to ``element`` (a row of panel ``level``)."""
        designer = self.designer
        panel = self.panels[level]
        items = [row.element for row in panel.items()]
        slot = next((i for i, e in enumerate(items) if e is not None and e.index == element.index), len(items))
        if action == "open":
            designer.open_item(element)
        elif action == "change":
            self.change_macro(level, element)
        elif action == "rename":
            panel.begin_rename(element)
        elif action in ("above", "below"):
            self.add(level, slot if action == "above" else slot + 1, "button")
        elif action in ("submenu", "separator"):
            self.add(level, slot + 1, "menu" if action == "submenu" else "separator")
        elif action in ("up", "down"):
            designer.move_control(element, -1 if action == "up" else 1)
        elif action == "show":
            self.close()
            designer.reveal(element)
        elif action == "delete":
            designer.delete_control(element)
