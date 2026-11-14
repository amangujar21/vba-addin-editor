"""A drawing of one ribbon tab, laid out roughly as Office does, for the ribbon designer.

It shows what the XML cannot: which controls are large, which stack three to a
column, where the separators fall and how wide each group gets. Built-in icons
(imageMso) are not available offline, so they are drawn as a neutral square;
custom pictures are shown as they are. Clicking a control picks it.
"""

from __future__ import annotations

import tkinter as tk
from collections.abc import Callable
from dataclasses import dataclass, field
from tkinter import ttk

from vba_addin_editor.services import ribbon_service as rs
from vba_addin_editor.ui import theme

TAB_ROW = 28  # the strip of tab names
ROW = 22  # one of the three rows a group body holds
ROWS = 3
GROUP_LABEL = 20
PAD = 10
LARGE_ICON = 28
SMALL_ICON = 14
MAX_LARGE_TEXT = 76  # a large button's label wraps past this width
HEIGHT = TAB_ROW + 8 + ROW * ROWS + GROUP_LABEL + 8

_DROPDOWNS = {"menu", "splitButton", "gallery", "dynamicMenu"}
_INPUTS = {"editBox", "comboBox", "dropDown"}
_SIZED = {"button", "toggleButton", "menu", "splitButton", "gallery", "dynamicMenu"}
_INLINE = {"box", "buttonGroup"}
_SKIPPED = {"dialogBoxLauncher"}


@dataclass
class _Cell:
    element: rs.RibbonElement
    lines: list[str]
    width: int
    large: bool = False
    separator: bool = False


@dataclass
class _Group:
    element: rs.RibbonElement
    columns: list[list[_Cell]] = field(default_factory=list)
    width: int = 0


def element_label(element: rs.RibbonElement) -> str:
    """What a control shows: its label, else its id."""
    raw = element.attr("label") or element.attr("id") or element.attr("idMso") or element.attr("idQ") or ""
    return rs.unescape(raw)


class RibbonPreview(ttk.Frame):
    """The strip above the inspector: tab names, then the groups of one tab."""

    def __init__(
        self,
        master,
        on_pick: Callable[[int], None],
        on_open: Callable[[int], None] | None = None,
    ) -> None:
        super().__init__(master, style="Panel.TFrame")
        self.on_pick = on_pick
        self.on_open = on_open
        self.canvas = tk.Canvas(self, height=HEIGHT, highlightthickness=0, cursor="hand2")
        self.hsb = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(xscrollcommand=theme.autohide(self.hsb, side="bottom", fill="x"))
        self.canvas.pack(side="top", fill="x")
        self.canvas.bind("<Button-1>", self._on_click)
        self.canvas.bind("<Double-Button-1>", self._on_double_click)
        self.canvas.bind("<MouseWheel>", self._on_wheel)
        self.canvas.bind("<Configure>", lambda _e: self._center(), add=True)
        self._hits: list[tuple[int, int, int, int, int]] = []  # x1, y1, x2, y2, element index
        self._photos: list[tk.PhotoImage] = []
        self._focus_box: tuple[int, int] | None = None
        self._state: tuple | None = None
        self.apply_palette()

    def apply_palette(self) -> None:
        self.canvas.configure(background=theme.palette().bg1)
        if self._state is not None:
            self.show(*self._state)

    # -- drawing ---------------------------------------------------------------

    def _measure(self, text: str, font: str = theme.SMALL_FONT) -> int:
        return int(self.tk.call("font", "measure", font, text))

    def show(
        self,
        elements: list[rs.RibbonElement],
        tab: rs.RibbonElement | None,
        selected: frozenset[int] = frozenset(),
        image_for: Callable[[rs.RibbonElement], tk.PhotoImage | None] | None = None,
        message: str = "",
    ) -> None:
        """Draw ``tab`` (one of ``elements``); ``selected`` = indices to outline."""
        self._state = (elements, tab, selected, image_for, message)
        colors = theme.palette()
        canvas = self.canvas
        canvas.delete("all")
        self._hits.clear()
        self._photos.clear()
        self._focus_box = None
        tabs = [e for e in elements if rs.local_name(e.name) == "tab"]
        x = PAD
        for item in tabs:
            label = element_label(item) or "tab"
            width = self._measure(label, theme.SMALL_BOLD_FONT)
            current = tab is not None and item.index == tab.index
            canvas.create_text(
                x, TAB_ROW // 2 + 1, anchor="w", text=label, font=theme.SMALL_BOLD_FONT,
                fill=colors.text if current else colors.faint,
            )
            if current:
                canvas.create_line(x, TAB_ROW - 2, x + width, TAB_ROW - 2, fill=colors.accent, width=2)
            self._hits.append((x - 6, 0, x + width + 6, TAB_ROW, item.index))
            x += width + 22
        canvas.create_line(0, TAB_ROW, 20000, TAB_ROW, fill=colors.border)
        total = x
        if tab is None:
            canvas.create_text(
                PAD, TAB_ROW + 8 + ROW * ROWS // 2, anchor="w", font=theme.SMALL_FONT, fill=colors.faint,
                text=message or "Select a tab, group or control to see it here.",
            )
        else:
            total = max(total, self._draw_tab(elements, tab, selected, image_for))
        canvas.configure(scrollregion=(0, 0, total + PAD, HEIGHT))
        self._center()

    def _draw_tab(self, elements, tab, selected, image_for) -> int:
        colors = theme.palette()
        canvas = self.canvas
        children: dict[int, list[rs.RibbonElement]] = {}
        for element in elements:
            children.setdefault(element.parent, []).append(element)
        top = TAB_ROW + 8
        bottom = top + ROW * ROWS
        groups = [self._layout_group(g, children) for g in children.get(tab.index, []) if rs.local_name(g.name) == "group"]
        if not groups:
            canvas.create_text(
                PAD, top + ROW * ROWS // 2, anchor="w", font=theme.SMALL_FONT, fill=colors.faint,
                text="This tab has no groups yet. Add one with + in the list on the left.",
            )
            return PAD
        ancestors = self._with_ancestors(elements, selected)
        x = PAD
        for group in groups:
            left = x
            if group.element.index in selected:
                canvas.create_rectangle(
                    left - 5, top - 4, left + group.width + 5, bottom + GROUP_LABEL,
                    outline=colors.accent, fill=colors.accent_bg,
                )
                self._focus_box = (left - 5, left + group.width + 5)
            for column in group.columns:
                width = max(cell.width for cell in column)
                y = top
                for cell in column:
                    height = ROW * ROWS if (cell.large or cell.separator) else ROW
                    if cell.element.index in ancestors and not cell.separator:
                        canvas.create_rectangle(
                            x - 2, y, x + width + 2, y + height - 1, outline=colors.accent, fill=colors.accent_bg
                        )
                        self._focus_box = (x - 2, x + width + 2)
                    self._draw_cell(cell, x, y, width, image_for, cell.element.index in ancestors)
                    self._hits.append((x - 2, y, x + width + 2, y + height, cell.element.index))
                    y += height
                x += width + 6
            x = max(x, left + group.width)
            label = element_label(group.element)
            canvas.create_text(
                (left + x) // 2, bottom + GROUP_LABEL // 2 + 1, text=label, font=theme.SMALL_FONT,
                fill=colors.muted if group.element.index in selected else colors.faint,
            )
            self._hits.append((left - 4, bottom, x + 4, bottom + GROUP_LABEL, group.element.index))
            self._hits.append((left - 4, top - 4, x + 4, bottom, group.element.index))
            x += 8
            canvas.create_line(x, top - 2, x, bottom + GROUP_LABEL - 4, fill=colors.border_strong)
            x += 10
        # Controls were added after their group's fallback box: put the specific ones first.
        self._hits.sort(key=lambda hit: (hit[2] - hit[0]) * (hit[3] - hit[1]))
        return x

    @staticmethod
    def _with_ancestors(elements: list[rs.RibbonElement], selected: frozenset[int]) -> set[int]:
        found: set[int] = set()
        for index in selected:
            current = index
            while 0 <= current < len(elements) and current not in found:
                found.add(current)
                current = elements[current].parent
        return found

    def _layout_group(self, group: rs.RibbonElement, children: dict[int, list[rs.RibbonElement]]) -> _Group:
        cells: list[_Cell] = []

        def collect(parent: rs.RibbonElement) -> None:
            for child in children.get(parent.index, []):
                name = rs.local_name(child.name)
                if name in _SKIPPED:
                    continue
                if name in _INLINE:
                    collect(child)
                    continue
                cells.append(self._cell(child, name, rs.local_name(parent.name) == "group"))

        collect(group)
        result = _Group(group)
        column: list[_Cell] = []
        for cell in cells:
            if cell.large or cell.separator:
                if column:
                    result.columns.append(column)
                    column = []
                result.columns.append([cell])
                continue
            column.append(cell)
            if len(column) == ROWS:
                result.columns.append(column)
                column = []
        if column:
            result.columns.append(column)
        body = sum(max(c.width for c in col) + 6 for col in result.columns)
        result.width = max(body, self._measure(element_label(group)) + 12, 44)
        return result

    def _cell(self, element: rs.RibbonElement, name: str, in_group: bool) -> _Cell:
        if name == "separator":
            return _Cell(element, [], 5, separator=True)
        label = element_label(element)
        arrow = " ▾" if name in _DROPDOWNS else ""
        if in_group and name in _SIZED and element.attr("size") == "large":
            lines = self._wrap(label)
            lines[-1] += arrow
            width = max([LARGE_ICON] + [self._measure(line) for line in lines]) + 12
            return _Cell(element, lines, width, large=True)
        text = label + arrow
        width = SMALL_ICON + 8 + self._measure(text) + 6
        if name in _INPUTS:
            width += 58
        elif name == "labelControl":
            width -= SMALL_ICON + 4
        return _Cell(element, [text], width)

    def _wrap(self, label: str) -> list[str]:
        """One or two lines, split at the space nearest the middle, as Office does."""
        if self._measure(label) <= MAX_LARGE_TEXT or " " not in label.strip():
            return [label]
        spaces = [i for i, ch in enumerate(label) if ch == " "]
        split = min(spaces, key=lambda i: abs(i - len(label) / 2))
        return [label[:split], label[split + 1 :]]

    def _draw_cell(self, cell: _Cell, x: int, y: int, width: int, image_for, active: bool) -> None:
        colors = theme.palette()
        canvas = self.canvas
        element = cell.element
        if cell.separator:
            canvas.create_line(x + 2, y + 2, x + 2, y + ROW * ROWS - 2, fill=colors.border_strong)
            return
        name = rs.local_name(element.name)
        fill = colors.text if active else colors.muted
        photo = image_for(element) if image_for is not None else None
        if cell.large:
            center = x + width // 2
            self._icon(center - LARGE_ICON // 2, y + 3, LARGE_ICON, photo, name)
            for row, line in enumerate(cell.lines[:2]):
                canvas.create_text(center, y + LARGE_ICON + 13 + row * 13, text=line, font=theme.SMALL_FONT, fill=fill)
            return
        middle = y + ROW // 2
        text_x = x + 2
        if name != "labelControl":
            self._icon(x + 2, middle - SMALL_ICON // 2, SMALL_ICON, photo, name)
            text_x = x + SMALL_ICON + 8
        canvas.create_text(text_x, middle, anchor="w", text=cell.lines[0], font=theme.SMALL_FONT, fill=fill)
        if name in _INPUTS:
            right = x + width - 2
            canvas.create_rectangle(right - 54, middle - 8, right, middle + 8, outline=colors.border_strong, fill=colors.bg3)

    def _icon(self, x: int, y: int, size: int, photo: tk.PhotoImage | None, name: str) -> None:
        colors = theme.palette()
        canvas = self.canvas
        if photo is not None:
            shown = photo
            factor = max(photo.width(), photo.height()) // size
            if factor > 1:
                shown = photo.subsample(factor)
            self._photos.append(shown)
            canvas.create_image(x + size // 2, y + size // 2, image=shown)
            return
        if name == "checkBox":
            canvas.create_rectangle(x + 1, y + 1, x + size - 2, y + size - 2, outline=colors.faint)
            return
        canvas.create_rectangle(x, y, x + size - 1, y + size - 1, outline=colors.border_strong, fill=colors.bg4)

    def _center(self) -> None:
        """Scroll sideways just enough to bring the selected control into view."""
        if self._focus_box is None:
            return
        region = str(self.canvas.cget("scrollregion")).split()
        total = float(region[2]) if len(region) == 4 else 0.0
        view = self.canvas.winfo_width()
        if total <= view or view <= 1:
            return
        left, right = self._focus_box
        first = self.canvas.canvasx(0)
        if left < first + PAD or right > first + view - PAD:
            self.canvas.xview_moveto(max(left - view / 3, 0) / total)

    # -- picking ---------------------------------------------------------------

    def element_at(self, x: int, y: int) -> int | None:
        """Index of the element drawn at widget coordinates (x, y)."""
        cx, cy = self.canvas.canvasx(x), self.canvas.canvasy(y)
        return next((i for x1, y1, x2, y2, i in self._hits if x1 <= cx < x2 and y1 <= cy < y2), None)

    def _on_click(self, event) -> None:
        index = self.element_at(event.x, event.y)
        if index is not None:
            self.on_pick(index)

    def _on_double_click(self, event) -> None:
        index = self.element_at(event.x, event.y)
        if index is not None and self.on_open is not None:
            self.on_open(index)

    def _on_wheel(self, event) -> str:
        self.canvas.xview_scroll(-3 if event.delta > 0 else 3, "units")
        return "break"
