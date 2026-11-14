"""The ribbon designer: the ribbon as a tree, a drawing of it, and a form for one control.

This is the Ribbon view. The tree on the left is the ribbon (tab > group > menu
> button); the strip on top draws the tab the selection is in; the form under
it edits the selected control in place; the bottom shows the macro it runs.
Nothing here parses or writes XML itself: every change is a pure text edit from
``ribbon_service`` applied through ``MainWindow._batch_edit``, so it is one
undo step and the XML view shows exactly what was done.
"""

from __future__ import annotations

import base64
import re
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Any

from vba_addin_editor.services import icon_service
from vba_addin_editor.services import ribbon_service as rs
from vba_addin_editor.ui import theme
from vba_addin_editor.ui.highlight import TAGS, Highlighter
from vba_addin_editor.ui.ribbon_dialogs import ask_image_file
from vba_addin_editor.ui.ribbon_preview import RibbonPreview, element_label
from vba_addin_editor.ui.widgets import Sidebar, icon
from vba_addin_editor.version import APP_NAME

# Elements that only hold the controls below them: left out of the tree.
_WRAPPERS = {
    "customUI",
    "ribbon",
    "tabs",
    "contextualTabs",
    "tabSet",
    "qat",
    "sharedControls",
    "documentControls",
    "officeMenu",
    "commands",
}
_GLYPHS = {
    "tab": "ribbon",
    "group": "form",
    "menu": "menu",
    "dynamicMenu": "menu",
    "splitButton": "menu",
    "gallery": "menu",
    "customUI": "bolt",
    "toggleButton": "check",
    "checkBox": "check",
}
_KIND_NAMES = {
    "tab": "Tab",
    "group": "Group",
    "button": "Button",
    "toggleButton": "Toggle button",
    "checkBox": "Check box",
    "menu": "Menu",
    "splitButton": "Split button",
    "dynamicMenu": "Dynamic menu",
    "gallery": "Gallery",
    "separator": "Separator",
    "menuSeparator": "Separator",
    "customUI": "Ribbon events",
    "editBox": "Edit box",
    "comboBox": "Combo box",
    "dropDown": "Drop-down",
    "labelControl": "Label",
    "buttonGroup": "Button group",
    "box": "Box",
    "control": "Built-in control",
}
_SEPARATORS = {"separator", "menuSeparator"}
_CONTAINERS = {"tab", "group", "menu", "box", "buttonGroup", "splitButton", "dynamicMenu"}
_NO_LABEL = _SEPARATORS | {"box", "buttonGroup", "dialogBoxLauncher", "customUI", "splitButton"}
_ACTION = {"button", "toggleButton", "checkBox"}
_ICON = rs.IMAGE_ELEMENTS - {"group", "item", "labelControl"}
_SIZED = {"button", "toggleButton", "menu", "splitButton", "gallery", "dynamicMenu"}
_TIPPED = _ACTION | {"menu", "gallery", "dynamicMenu", "comboBox", "dropDown", "editBox", "labelControl"}
_ID_PREFIX = {"button": "btn", "menu": "mnu", "group": "grp", "tab": "tab", "toggleButton": "tgl", "checkBox": "chk"}
_ID_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_.-]*$")
_END_PROC_RE = re.compile(r"\s*End\s+(Sub|Function|Property)\b", re.IGNORECASE)
# Attributes the form edits (or that say what the element is): not repeated under "Also set".
_KNOWN_ATTRS = {
    "id", "idMso", "idQ", "label", "onAction", "imageMso", "image", "size", "screentip", "supertip", "xmlns",
}
MAX_PEEK_LINES = 400
DRAG_THRESHOLD = 6


def procedure_source(body: str, line: int) -> list[str]:
    """The lines of the procedure declared on 1-based ``line`` of ``body``."""
    lines = body.split("\n")
    start = max(line - 1, 0)
    end = start
    for index in range(start, min(len(lines), start + MAX_PEEK_LINES)):
        end = index
        if index > start and _END_PROC_RE.match(lines[index]):
            break
    return lines[start : end + 1]


def drop_zone(fraction: float, *, into: bool, beside: bool, opened: bool) -> str | None:
    """Where a drop lands on a row: "before", "after", "into", "first", or None.

    ``fraction`` is how far down the row the pointer is (0 top, 1 bottom);
    ``into`` / ``beside`` say whether the row may hold the dragged control and
    whether its parent may; ``opened`` = the row shows children under it.
    """
    if into and beside:
        if fraction < 0.28:
            return "before"
        if fraction > 0.72:
            return "first" if opened else "after"
        return "into"
    if into:
        return "into"
    if beside:
        return "before" if fraction < 0.5 else "after"
    return None


class _Field:
    """One row of the form: a caption and a box that edits one attribute."""

    def __init__(self, inspector: _Inspector, parent, row: int, caption: str, attr: str, *, choices=None) -> None:
        self.inspector = inspector
        self.attr = attr
        self.row = row
        self.loaded = ""
        self.choices = choices
        self.var = tk.StringVar(master=parent)
        self.caption = ttk.Label(parent, text=caption, style="Field.TLabel")
        self.box = ttk.Frame(parent, style="Editor.TFrame")
        self.combo: ttk.Combobox | None = None
        if choices is None:
            self.entry: ttk.Entry = ttk.Entry(self.box, textvariable=self.var, width=46)
        else:
            self.entry = self.combo = ttk.Combobox(self.box, textvariable=self.var, width=43, postcommand=self._fill)
            self.entry.bind("<<ComboboxSelected>>", lambda _e: self.commit())
        self.entry.pack(side="left")
        self.entry.bind("<Return>", lambda _e: self.commit())
        self.entry.bind("<FocusOut>", self._on_focus_out)
        self.entry.bind("<Escape>", lambda _e: self.var.set(self.loaded))
        self.shown = False

    def show(self, value: str) -> None:
        if not self.shown:
            self.caption.grid(row=self.row, column=0, sticky="w", padx=(0, 14), pady=3)
            self.box.grid(row=self.row, column=1, sticky="w", pady=3)
            self.shown = True
        self.load(value)

    def hide(self) -> None:
        if self.shown:
            self.caption.grid_remove()
            self.box.grid_remove()
            self.shown = False

    def load(self, value: str) -> None:
        self.loaded = value
        if self.var.get() != value:
            self.var.set(value)

    def pending(self) -> bool:
        return self.shown and self.var.get() != self.loaded and self.var.get().strip() != self.loaded

    def commit(self) -> None:
        if not self.pending():
            if self.shown and self.var.get() != self.loaded:
                self.var.set(self.loaded)  # only spaces were typed around it
            return
        if not self.inspector.commit(self, self.var.get().strip()):
            self.var.set(self.loaded)

    def _on_focus_out(self, _event=None) -> None:
        try:
            focused = str(self.entry.tk.call("focus"))
        except tk.TclError:
            focused = ""
        if "popdown" in focused:
            return  # the combobox's own list took the focus
        self.commit()

    def _fill(self) -> None:
        """The drop-down list: everything, or what contains the text typed so far."""
        names = list(self.choices()) if self.choices is not None else []
        typed = self.var.get().strip().casefold()
        if typed and self.var.get() != self.loaded:
            names = [n for n in names if typed in n.casefold()] or names
        if self.combo is not None:
            self.combo.configure(values=names)


class _MacroField(_Field):
    """A callback attribute: the macro's name, whether it exists, and a way to it."""

    def __init__(self, inspector: _Inspector, parent, row: int, caption: str, attr: str) -> None:
        super().__init__(inspector, parent, row, caption, attr, choices=inspector.designer.macro_names)
        self.status = ttk.Label(self.box, text="", style="EditorOkIcon.TLabel")
        self.status.pack(side="left", padx=(8, 0))
        self.link = ttk.Label(self.box, text="", style="EditorLink.TLabel", cursor="hand2")
        self.link.pack(side="left", padx=(6, 0))
        self.link.bind("<Button-1>", lambda _e: self._follow())
        self.hint = ttk.Label(parent, text="", style="CrumbWarn.TLabel", wraplength=520, justify="left")
        self._missing = False

    def show(self, value: str) -> None:
        super().show(value)
        resolution = self.inspector.designer.resolve(value) if value else None
        self._missing = resolution is not None and resolution.status == rs.STATUS_MISSING
        if resolution is None:
            self.status.config(text="")
            self.link.config(text="")
            self.hint.grid_remove()
        elif resolution.status == rs.STATUS_OK:
            self.status.config(text=icon("check"), style="EditorOkIcon.TLabel")
            self.link.config(text="Open")
            self.hint.grid_remove()
        else:
            self.status.config(text=icon("alert"), style="EditorWarnIcon.TLabel")
            self.link.config(text="Create Sub" if self._missing else "Open")
            self.hint.config(text=resolution.message or "")
            self.hint.grid(row=self.row + 1, column=1, sticky="w", pady=(0, 4))

    def hide(self) -> None:
        super().hide()
        self.hint.grid_remove()

    def _follow(self) -> None:
        self.commit()
        element = self.inspector.element
        if element is not None:
            self.inspector.designer.open_callback(element, self.attr, create=self._missing)


CAPTION_WIDTH = 96  # inspector caption column, px


class _Inspector(ttk.Frame):
    """The form for the selected control. Fields are made once and shown as needed."""

    def __init__(self, master, designer: RibbonDesigner) -> None:
        super().__init__(master, style="Editor.TFrame", padding=(18, 12, 18, 12))
        # Fixed caption column, so the fields don't shift sideways between control kinds.
        self.columnconfigure(0, minsize=CAPTION_WIDTH)
        self.designer = designer
        self.element: rs.RibbonElement | None = None
        self.text = ""
        self._committing = False
        self._photo: tk.PhotoImage | None = None
        self.heading = ttk.Label(self, text="", style="EditorCaption.TLabel")
        self.heading.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))
        self.hint = ttk.Label(self, text="", style="EditorMuted.TLabel", justify="left", wraplength=560)
        self.label = _Field(self, self, 2, "Label", "label")
        self.macro = _MacroField(self, self, 4, "Macro", "onAction")
        self.image = _Field(self, self, 6, "Icon", "imageMso", choices=designer.image_names)
        self._build_icon_buttons()
        self._build_size_row(8)
        self.screentip = _Field(self, self, 10, "Screentip", "screentip")
        self.supertip = _Field(self, self, 12, "Supertip", "supertip")
        self.ident = _Field(self, self, 14, "Id", "id")
        self.extra = ttk.Frame(self, style="Editor.TFrame")
        self._extra_fields: list[_MacroField] = []
        self.also = ttk.Label(self, text="", style="EditorFaint.TLabel", wraplength=560, justify="left")
        self.multi = ttk.Frame(self, style="Editor.TFrame")
        ttk.Button(
            self.multi, text="Group into menu…", style="EditorGhost.TButton",
            command=designer.group_selected_into_menu, takefocus=False,
        ).pack(side="left")
        ttk.Button(
            self.multi, text="Delete", style="EditorGhost.TButton", command=designer.delete_selected, takefocus=False
        ).pack(side="left", padx=(8, 0))
        self._fields = (self.label, self.macro, self.image, self.screentip, self.supertip, self.ident)

    def _build_icon_buttons(self) -> None:
        box = self.image.box
        designer = self.designer
        ttk.Button(
            box, text="Paste", style="EditorText.TButton", command=designer.set_icon_from_clipboard, takefocus=False
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            box, text="File…", style="EditorText.TButton", command=designer.set_icon_from_file, takefocus=False
        ).pack(side="left")
        self.thumb = ttk.Label(box, style="EditorFaint.TLabel", compound="left")
        self.remove_picture = ttk.Button(
            box, text="Remove picture", style="EditorText.TButton", command=designer.remove_picture, takefocus=False
        )

    def _build_size_row(self, row: int) -> None:
        self.size_caption = ttk.Label(self, text="Size", style="Field.TLabel")
        self.size_box = ttk.Frame(self, style="Editor.TFrame")
        self.size_row = row
        self.size_buttons: dict[str, ttk.Button] = {}
        for value, text in (("large", "Large"), ("normal", "Normal")):
            button = ttk.Button(
                self.size_box, text=text, style="Seg.TButton", takefocus=False,
                command=lambda v=value: self._set_size(v),
            )
            button.pack(side="left")
            self.size_buttons[value] = button
        self._size_shown = False

    def _set_size(self, value: str) -> None:
        if self.element is not None:
            self.commit_pending()
            self.designer.set_attrs(self.element, self.text, {"size": "large" if value == "large" else None})

    # -- showing ---------------------------------------------------------------

    def show(self, element: rs.RibbonElement | None, count: int = 0) -> None:
        designer = self.designer
        self.element = element
        self.text = designer.text
        if element is None:
            for item in self._fields:
                item.hide()
            self._show_size(None)
            self._show_extra([])
            self.also.grid_remove()
            self.heading.config(text=f"{count} CONTROLS SELECTED" if count > 1 else "")
            self.hint.config(
                text="Group them into a menu, or delete them."
                if count > 1
                else "Select a tab, group or control on the left to edit it here."
            )
            self.hint.grid(row=1, column=0, columnspan=2, sticky="w")
            if count > 1:
                self.multi.grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 0))
            else:
                self.multi.grid_remove()
            return
        self.hint.grid_remove()
        self.multi.grid_remove()
        local = rs.local_name(element.name)
        parent = designer.parent_of(element)
        builtin = element.attr("idMso")
        kind = _KIND_NAMES.get(local, local).upper()
        self.heading.config(text=f"{kind}  ·  BUILT-IN {builtin}" if builtin else kind)
        value = lambda name: rs.unescape(element.attr(name) or "")
        self._toggle(self.label, local not in _NO_LABEL, value("label"))
        self._toggle(self.macro, local in _ACTION, value("onAction"))
        self._toggle(self.image, local in _ICON, value("imageMso"))
        self._show_picture(element if local in _ICON else None)
        sized = local in _SIZED and parent is not None and rs.local_name(parent.name) == "group"
        self._show_size(element if sized else None)
        self._toggle(self.screentip, local in _TIPPED, value("screentip"))
        self._toggle(self.supertip, local in _TIPPED, value("supertip"))
        self._toggle(self.ident, element.attr("id") is not None, value("id"))
        shown_action = local in _ACTION
        self._show_extra([(a, v) for a, v in rs.callback_attrs(element) if not (shown_action and a == "onAction")])
        others = [name for name, _v, _o in element.attrs if name not in _KNOWN_ATTRS and not rs.is_callback_attribute(name)]
        if others:
            self.also.config(text="Also set in the XML: " + ", ".join(others))
            self.also.grid(row=18, column=0, columnspan=2, sticky="w", pady=(8, 0))
        else:
            self.also.grid_remove()

    @staticmethod
    def _toggle(item: _Field, wanted: bool, value: str) -> None:
        if wanted:
            item.show(value)
        else:
            item.hide()

    def _show_size(self, element: rs.RibbonElement | None) -> None:
        if element is None:
            if self._size_shown:
                self.size_caption.grid_remove()
                self.size_box.grid_remove()
                self._size_shown = False
            return
        if not self._size_shown:
            self.size_caption.grid(row=self.size_row, column=0, sticky="w", padx=(0, 14), pady=3)
            self.size_box.grid(row=self.size_row, column=1, sticky="w", pady=3)
            self._size_shown = True
        large = element.attr("size") == "large"
        self.size_buttons["large"].configure(style="SegOn.TButton" if large else "Seg.TButton")
        self.size_buttons["normal"].configure(style="Seg.TButton" if large else "SegOn.TButton")

    def _show_picture(self, element: rs.RibbonElement | None) -> None:
        photo = self.designer.photo_for(element) if element is not None else None
        custom = element is not None and bool(element.attr("image"))
        self._photo = photo
        if custom:
            self.thumb.config(image=photo if photo is not None else "", text=" Custom picture")
            self.thumb.pack(side="left", padx=(12, 0))
            self.remove_picture.pack(side="left", padx=(4, 0))
        else:
            self.thumb.pack_forget()
            self.remove_picture.pack_forget()

    def _show_extra(self, callbacks: list[tuple[str, str]]) -> None:
        names = [attr for attr, _value in callbacks]
        if names != [item.attr for item in self._extra_fields]:
            for child in self.extra.winfo_children():
                child.destroy()
            self._extra_fields = [
                _MacroField(self, self.extra, 2 * row, attr, attr) for row, attr in enumerate(names)
            ]
        if not callbacks:
            self.extra.grid_remove()
            return
        self.extra.columnconfigure(0, minsize=CAPTION_WIDTH)
        self.extra.grid(row=16, column=0, columnspan=2, sticky="w", pady=(4, 0))
        for item, (_attr, value) in zip(self._extra_fields, callbacks, strict=True):
            item.show(value)

    # -- committing ------------------------------------------------------------

    def all_fields(self) -> list[_Field]:
        return [*self._fields, *self._extra_fields]

    def commit_pending(self) -> bool:
        """Apply what was typed but not yet confirmed. True = something was applied."""
        waiting = [item for item in self.all_fields() if item.pending()]
        for item in waiting:
            item.commit()
        return bool(waiting)

    def commit(self, item: _Field, value: str) -> bool:
        element, text, designer = self.element, self.text, self.designer
        if element is None or self._committing:
            return False
        self._committing = True
        try:
            if item.attr == "label":
                return designer.set_label(element, text, value)
            if isinstance(item, _MacroField):
                return designer.set_macro(element, text, item.attr, value)
            if item.attr == "imageMso":
                return designer.set_image_mso(element, text, value)
            if item.attr == "id":
                return designer.set_id(element, text, value)
            return designer.set_attrs(element, text, {item.attr: value or None})
        finally:
            self._committing = False

    def focus_label(self) -> None:
        if self.label.shown:
            self.label.entry.focus_set()
            self.label.entry.select_range(0, "end")
            self.label.entry.icursor("end")


class _MacroPeek(ttk.Frame):
    """The code of the macro the selected control runs, read-only; F12 opens it for editing."""

    def __init__(self, master, designer: RibbonDesigner) -> None:
        super().__init__(master, style="Editor.TFrame")
        self.designer = designer
        self.location: rs.ProcedureLocation | None = None
        head = ttk.Frame(self, style="Editor.TFrame", padding=(18, 8, 10, 4))
        head.pack(fill="x", side="top")
        ttk.Label(head, text="MACRO", style="EditorCaption.TLabel").pack(side="left")
        self.where = ttk.Label(head, text="", style="Crumb.TLabel")
        self.where.pack(side="left", padx=(12, 0))
        self.open_button = ttk.Button(
            head, text="Open in editor   F12", style="EditorText.TButton", command=designer.goto_macro, takefocus=False
        )
        self.message = ttk.Label(self, text="", style="EditorFaint.TLabel", padding=(18, 2, 18, 8), wraplength=620)
        self.body = ttk.Frame(self, style="Editor.TFrame")
        self.text = tk.Text(
            self.body, wrap="none", font=theme.code_font(self), height=4, spacing1=2, padx=18, pady=4,
            cursor="arrow", takefocus=False, state="disabled",
        )
        scrollbar = ttk.Scrollbar(self.body, orient="vertical", command=self.text.yview)
        hide = theme.autohide(scrollbar, side="right", fill="y")

        def on_scroll(first, last) -> None:
            hide(first, last)
            self._schedule_colors()

        self.text.configure(yscrollcommand=on_scroll)
        self.text.pack(side="left", fill="both", expand=True)
        self.highlighter = Highlighter(self.text, "vba")
        self._colors_after = None
        self.text.bind("<Configure>", lambda _e: self._schedule_colors(), add=True)
        self.text.bind("<Double-Button-1>", self._open_at)
        self.apply_palette()

    def apply_palette(self) -> None:
        colors = theme.palette()
        self.text.configure(
            background=colors.bg2, foreground=colors.text, selectbackground=colors.select,
            selectforeground=colors.select_text, inactiveselectbackground=colors.select,
        )
        for tag in TAGS:
            self.text.tag_configure(tag, foreground=getattr(colors, tag))

    def show(self, location: rs.ProcedureLocation | None, source: list[str], message: str) -> None:
        self.location = location
        if location is None:
            self.where.config(text="")
            self.open_button.pack_forget()
            self.body.pack_forget()
            self.message.config(text=message)
            self.message.pack(fill="x", side="top")
            return
        self.where.config(text=f"{location.module_name}  ›  {location.name}   line {location.line}")
        self.open_button.pack(side="right")
        self.body.pack_forget()
        if message:
            self.message.config(text=message)
            self.message.pack(fill="x", side="top")
        else:
            self.message.pack_forget()
        self.body.pack(fill="both", expand=True, side="top")
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.insert("1.0", "\n".join(source))
        self.text.configure(state="disabled")
        self._schedule_colors()

    def _schedule_colors(self) -> None:
        if self._colors_after is None:
            self._colors_after = self.after_idle(self._color)

    def destroy(self) -> None:
        if self._colors_after is not None:
            try:
                self.after_cancel(self._colors_after)
            except tk.TclError:
                pass
            self._colors_after = None
        super().destroy()

    def _color(self) -> None:
        self._colors_after = None
        if self.winfo_exists() and self.body.winfo_manager():
            self.highlighter.refresh()

    def _open_at(self, event) -> str:
        location = self.location
        if location is not None:
            line = int(self.text.index(f"@{event.x},{event.y}").split(".")[0])
            self.designer.app._open_module_at(location.module_id, location.line + line - 1)
        return "break"


class RibbonDesigner(ttk.Frame):
    """The Ribbon page: tree, preview, form and macro peek for one ribbon part."""

    def __init__(self, master, app: Any) -> None:
        super().__init__(master)
        self.app = app
        self.part_path: str | None = None
        self.text = ""
        self.elements: list[rs.RibbonElement] = []
        self._problems: dict[int, list[rs.RibbonEntry]] = {}
        self._linked = 0
        self._collapsed: set[str] = set()
        self._fresh: set[str] = set()
        self._problems_only = False
        self._refreshing = False
        self._images: dict[str, bytes] = {}
        self._photos: dict[tuple[str, int], tk.PhotoImage | None] = {}
        self._drag: dict | None = None
        self._drop: tuple[str, str] | None = None

        panes = ttk.Panedwindow(self, orient="horizontal")
        panes.pack(fill="both", expand=True)
        self.filter_var = tk.StringVar(master=app.root)
        side = self.sidebar = Sidebar(
            panes, "RIBBON", add=self._post_add_menu, add_tip="Add a button, menu or group  (Ins = button)",
            filter_var=self.filter_var, footer=True, width=300,
        )
        self.tree = side.tree
        self.tree.configure(selectmode="extended")
        self.part_button = ttk.Menubutton(side.head, text="", style="Side.TMenubutton", takefocus=False)
        self.part_menu = tk.Menu(self.part_button, tearoff=0)
        self.part_button.configure(menu=self.part_menu)
        self.health = ttk.Label(side.foot, text="", style="Ok.TLabel", cursor="hand2")
        self.health.pack(side="left")
        self.health.bind("<Button-1>", lambda _e: self.toggle_problems())
        self.filter_var.trace_add("write", lambda *_a: self._on_filter())
        panes.add(side, weight=0)

        main = ttk.Frame(panes, style="Editor.TFrame")
        bar = app._crumb_bar(main)
        self.crumb = ttk.Label(bar, text="", style="Crumb.TLabel")
        self.crumb.pack(side="left")
        self.crumb_leaf = ttk.Label(bar, text="", style="CrumbStrong.TLabel")
        self.crumb_leaf.pack(side="left")
        ttk.Button(
            bar, text="Edit XML", style="EditorText.TButton", command=self.show_xml, takefocus=False
        ).pack(side="right")
        self.note = ttk.Label(bar, text="", style="CrumbWarn.TLabel")
        self.note.pack(side="right", padx=(8, 8))
        self.preview = RibbonPreview(main, self._pick, self._open_index)
        self.preview.pack(fill="x", side="top")
        ttk.Frame(main, style="Line.TFrame", height=1).pack(fill="x", side="top")
        self.inspector = _Inspector(main, self)
        self.inspector.pack(fill="x", side="top")
        ttk.Frame(main, style="Line.TFrame", height=1).pack(fill="x", side="top")
        self.peek = _MacroPeek(main, self)
        self.peek.pack(fill="both", expand=True, side="top")
        panes.add(main, weight=1)

        self.empty = ttk.Frame(self, style="Editor.TFrame")
        inner = ttk.Frame(self.empty, style="Editor.TFrame")
        inner.place(relx=0.5, rely=0.42, anchor="center")
        ttk.Label(inner, text="No ribbon yet", style="EditorHero.TLabel").pack(anchor="w")
        self.empty_note = ttk.Label(inner, text="", style="EditorMuted.TLabel", justify="left", wraplength=460)
        self.empty_note.pack(anchor="w", pady=(6, 16))
        self.empty_button = ttk.Button(
            inner, text="Create a ribbon tab", style="Accent.TButton", padding=(16, 6), command=app.create_ribbon
        )
        self.empty_button.pack(anchor="w")

        self._build_menus()
        self._drop_line = tk.Frame(self.tree, height=2)
        tree = self.tree
        tree.bind("<<TreeviewSelect>>", self._on_select)
        tree.bind("<<TreeviewOpen>>", lambda _e: self._note_open(True))
        tree.bind("<<TreeviewClose>>", lambda _e: self._note_open(False))
        tree.bind("<Double-Button-1>", self._on_double_click)
        tree.bind("<Return>", lambda _e: self._break(self.goto_macro() or self.inspector.focus_label()))
        tree.bind("<F2>", lambda _e: self._break(self.inspector.focus_label()))
        tree.bind("<Delete>", lambda _e: self._break(self.delete_selected()))
        tree.bind("<Insert>", lambda _e: self._break(self.add_control("button")))
        tree.bind("<Alt-Up>", lambda _e: self._break(self.move_selected(-1)))
        tree.bind("<Alt-Down>", lambda _e: self._break(self.move_selected(1)))
        tree.bind("<Button-3>", self._on_right_click)
        tree.bind("<ButtonPress-1>", self._drag_start, add=True)
        tree.bind("<B1-Motion>", self._drag_motion, add=True)
        tree.bind("<ButtonRelease-1>", self._drag_end, add=True)
        self.apply_palette()

    @staticmethod
    def _break(_result=None) -> str:
        return "break"

    def _build_menus(self) -> None:
        app = self.app
        add = self.add_menu = tk.Menu(self, tearoff=0)
        add.add_command(label="Button", accelerator="Ins", command=lambda: self.add_control("button"))
        add.add_command(label="Menu", command=lambda: self.add_control("menu"))
        add.add_command(label="Group", command=lambda: self.add_control("group"))
        add.add_command(label="Separator", command=lambda: self.add_control("separator"))
        add.add_separator()
        add.add_command(label="Import from Another Add-in…", command=app.import_ribbon_control)
        menu = self.context_menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Add Button", accelerator="Ins", command=lambda: self.add_control("button"))
        menu.add_command(label="Add Menu", command=lambda: self.add_control("menu"))
        menu.add_command(label="Add Group", command=lambda: self.add_control("group"))
        menu.add_command(label="Add Separator", command=lambda: self.add_control("separator"))
        menu.add_separator()
        menu.add_command(label="Go to Macro", accelerator="F12", command=self.goto_macro)
        menu.add_command(label="Rename", accelerator="F2", command=self.inspector.focus_label)
        menu.add_command(label="Move Up", accelerator="Alt+Up", command=lambda: self.move_selected(-1))
        menu.add_command(label="Move Down", accelerator="Alt+Down", command=lambda: self.move_selected(1))
        menu.add_command(label="Group into Menu…", command=self.group_selected_into_menu)
        menu.add_command(label="Paste Icon from Clipboard", command=self.set_icon_from_clipboard)
        menu.add_command(label="Icon from File…", command=self.set_icon_from_file)
        menu.add_separator()
        menu.add_command(label="Edit XML", command=self.show_xml)
        menu.add_command(label="Delete", accelerator="Del", command=self.delete_selected)
        app._menus.extend([add, menu, self.part_menu])

    def apply_palette(self) -> None:
        colors = theme.palette()
        tree = self.tree
        tree.tag_configure("container", foreground=colors.muted)
        tree.tag_configure("control", foreground=colors.text)
        tree.tag_configure("separator", foreground=colors.faint)
        tree.tag_configure("problem", foreground=colors.warn)
        tree.tag_configure("drop", background=colors.accent_bg)
        self._drop_line.configure(background=colors.accent)
        self._photos.clear()
        self.preview.apply_palette()
        self.peek.apply_palette()
        self.sidebar.apply_palette()

    # -- reading the draft -------------------------------------------------------

    def _part(self):
        draft = self.app.draft
        return draft.xml_part_by_path(self.part_path) if draft is not None and self.part_path else None

    def refresh(self) -> None:
        """Rebuild from the draft (after any edit, undo, or a switch to this view)."""
        draft = self.app.draft
        paths = [p.path for p in rs.ribbon_parts(draft)] if draft is not None else []
        if self.part_path not in paths:
            newest = [p for p in paths if p.casefold().endswith("customui14.xml")]
            self.part_path = (newest or paths or [None])[0]
            self._collapsed.clear()
            self._fresh.clear()
        self._show_part_switcher(paths)
        part = self._part()
        if part is None:
            self.text, self.elements = "", []
            self._refreshing = True
            try:
                self.tree.delete(*self.tree.get_children())
            finally:
                self._refreshing = False
            self._show_empty(draft is not None)
            self._show_selection()
            return
        self.empty.place_forget()
        keys = self._selection_keys()
        self.text = part.text or ""
        self.elements = rs.element_tree(self.text)
        self._read_problems()
        self._read_images()
        self._fill_tree()
        self._restore_selection(keys)
        self._update_health()
        self._show_selection()

    def _show_empty(self, has_file: bool) -> None:
        if not has_file:
            self.empty.place_forget()
            return
        draft = self.app.draft
        name = draft.baseline.path.name if draft is not None else "This file"
        self.empty_note.config(
            text=f"{name} has no custom ribbon. Create a tab with one group, then add buttons "
            "for your macros here: no XML needed."
        )
        self.empty.place(x=0, y=0, relwidth=1, relheight=1)
        self.empty.lift()

    def _show_part_switcher(self, paths: list[str]) -> None:
        if len(paths) < 2:
            self.part_button.pack_forget()
            return
        self.part_menu.delete(0, "end")
        for path in paths:
            self.part_menu.add_command(label=path, command=lambda p=path: self.show_part(p))
        self.part_button.config(text=(self.part_path or "").rsplit("/", 1)[-1] + "  ▾")
        if not self.part_button.winfo_manager():
            self.part_button.pack(side="left", padx=(8, 0))

    def show_part(self, path: str) -> None:
        self.inspector.commit_pending()
        self.part_path = path
        self._collapsed.clear()
        self.tree.selection_set(())
        self.refresh()

    def _read_problems(self) -> None:
        draft = self.app.draft
        self._problems.clear()
        self._linked = 0
        if draft is None:
            return
        by_start = {e.start: e.index for e in self.elements}
        for entry in rs.ribbon_entries(draft):
            if entry.callback.part_path != self.part_path:
                continue
            if entry.resolution.status == rs.STATUS_OK:
                self._linked += 1
            elif entry.callback.tag_start in by_start:
                self._problems.setdefault(by_start[entry.callback.tag_start], []).append(entry)

    def _read_images(self) -> None:
        draft = self.app.draft
        self._images.clear()
        if draft is None or self.part_path is None:
            return
        for rid in {e.attr("image") for e in self.elements if e.attr("image")}:
            data = icon_service.image_bytes(draft, self.part_path, rid or "")
            if data:
                self._images[rid or ""] = data

    def photo_for(self, element: rs.RibbonElement | None) -> tk.PhotoImage | None:
        """The control's custom picture, when it has one Tk can show."""
        rid = element.attr("image") if element is not None else None
        data = self._images.get(rid) if rid else None
        if not rid or data is None:
            return None
        key = (rid, hash(data))
        if key not in self._photos:
            try:
                self._photos[key] = tk.PhotoImage(master=self, data=base64.b64encode(data).decode("ascii"))
            except tk.TclError:
                self._photos[key] = None  # not a PNG or GIF
        return self._photos[key]

    # -- the tree ----------------------------------------------------------------

    @staticmethod
    def _key(element: rs.RibbonElement) -> str:
        return element.attr("id") or element.attr("idMso") or f"#{element.index}"

    def _filtering(self) -> bool:
        return bool(self.filter_var.get().strip()) or self._problems_only

    def _visible(self) -> tuple[set[int], set[int]] | None:
        """(indices to show, indices that match) for the filter, or None for everything."""
        query = self.filter_var.get().strip().casefold()
        if not query and not self._problems_only:
            return None
        elements = self.elements
        matching: set[int] = set()
        for element in elements:
            if self._problems_only and element.index not in self._problems:
                continue
            words = [element_label(element), element.attr("id") or "", *(name for _a, name in rs.callback_attrs(element))]
            if not query or any(query in word.casefold() for word in words):
                matching.add(element.index)
        visible = set(matching)
        for index in matching:
            parent = elements[index].parent
            while parent >= 0 and parent not in visible:
                visible.add(parent)
                parent = elements[parent].parent
        if query and not self._problems_only:
            inside = {i for i in matching if rs.local_name(elements[i].name) in _CONTAINERS}
            for element in elements:
                if element.parent in inside:
                    inside.add(element.index)
                    visible.add(element.index)
        return visible, matching

    def _fill_tree(self) -> None:
        tree = self.tree
        found = self._visible()
        visible, matching = found if found is not None else (None, None)
        self._refreshing = True
        try:
            tree.delete(*tree.get_children())
            shown: dict[int, str] = {}
            for element in self.elements:
                local = rs.local_name(element.name)
                events = local == "customUI" and bool(rs.callback_attrs(element))
                if local in _WRAPPERS and not events:
                    continue
                if visible is not None and element.index not in visible:
                    continue
                # customUI is everyone's ancestor: under a filter, show its events row only if it matches itself.
                if events and matching is not None and element.index not in matching:
                    continue
                parent = element.parent
                while parent >= 0 and parent not in shown:
                    parent = self.elements[parent].parent
                if local in _SEPARATORS:
                    text, tags = "      ───────", ["separator"]
                else:
                    label = "Ribbon events" if events else element_label(element) or local
                    text = f"{icon(_GLYPHS.get(local, 'button'))}  {label}"
                    tags = ["container" if local in _CONTAINERS else "control"]
                problem = element.index in self._problems
                if problem:
                    tags.append("problem")
                iid = tree.insert(
                    "" if events else shown.get(parent, ""),
                    0 if events else "end",
                    iid=f"e{element.index}",
                    text=text,
                    values=("⚠" if problem else "",),
                    open=visible is not None or self._key(element) not in self._collapsed,
                    tags=tuple(tags),
                )
                if not events:
                    shown[element.index] = iid
        finally:
            self._refreshing = False

    def _note_open(self, opened: bool) -> None:
        element = self._element_of(self.tree.focus())
        if element is None or self._refreshing or self._filtering():
            return
        if opened:
            self._collapsed.discard(self._key(element))
        else:
            self._collapsed.add(self._key(element))

    def _element_of(self, iid: str) -> rs.RibbonElement | None:
        if not iid.startswith("e"):
            return None
        index = int(iid[1:])
        return self.elements[index] if index < len(self.elements) else None

    def selected_elements(self) -> list[rs.RibbonElement]:
        found = [self._element_of(iid) for iid in self.tree.selection()]
        return sorted((e for e in found if e is not None), key=lambda e: e.start)

    def selected_element(self) -> rs.RibbonElement | None:
        elements = self.selected_elements()
        return elements[0] if len(elements) == 1 else None

    def parent_of(self, element: rs.RibbonElement) -> rs.RibbonElement | None:
        return self.elements[element.parent] if 0 <= element.parent < len(self.elements) else None

    def _selection_keys(self) -> list[tuple[str | None, str, int]]:
        return [(e.attr("id"), e.name, e.index) for e in self.selected_elements()]

    def _restore_selection(self, keys: list[tuple[str | None, str, int]]) -> None:
        wanted: list[str] = []
        for ident, name, index in keys:
            match = None
            if ident:
                match = next((e for e in self.elements if e.attr("id") == ident and e.name == name), None)
            elif index < len(self.elements) and self.elements[index].name == name:
                match = self.elements[index]
            if match is not None and self.tree.exists(f"e{match.index}"):
                wanted.append(f"e{match.index}")
        self._select_iids(wanted)

    def _select_iids(self, iids: list[str]) -> None:
        self._refreshing = True
        try:
            self.tree.selection_set(iids)
            if iids:
                self.tree.focus(iids[0])
                self.tree.see(iids[0])
        finally:
            self._refreshing = False

    def select_start(self, start: int) -> bool:
        """Select the element whose tag begins at ``start`` (clearing a filter that hides it)."""
        element = next((e for e in self.elements if e.start == start), None)
        while element is not None and rs.local_name(element.name) in _WRAPPERS:
            element = self.parent_of(element)
        if element is None:
            return False
        if not self.tree.exists(f"e{element.index}"):
            self._problems_only = False
            self.filter_var.set("")
            self._fill_tree()
            self._update_health()
        iid = f"e{element.index}"
        if not self.tree.exists(iid):
            return False
        parent = self.tree.parent(iid)
        while parent:
            self.tree.item(parent, open=True)
            parent = self.tree.parent(parent)
        self._select_iids([iid])
        self._show_selection()
        return True

    def select_offset(self, part_path: str, offset: int) -> bool:
        """Show ``part_path`` and select the control around ``offset`` of its text."""
        if part_path != self.part_path:
            self.part_path = part_path
            self._collapsed.clear()
        self.refresh()
        around = [e for e in self.elements if e.start <= offset < (e.end if e.complete else e.open_end)]
        target = max(around, key=lambda e: e.start, default=None)
        return target is not None and self.select_start(target.start)

    def _on_filter(self) -> None:
        if self.part_path is None:
            return
        keys = self._selection_keys()
        self._fill_tree()
        self._restore_selection(keys)
        self._show_selection()

    def toggle_problems(self) -> None:
        if not self._problems and not self._problems_only:
            return
        self._problems_only = not self._problems_only
        self._on_filter()
        self._update_health()

    def _update_health(self) -> None:
        problems = sum(len(entries) for entries in self._problems.values())
        if self._problems_only:
            self.health.config(text="Showing problems only  ·  show all", style="Link.TLabel")
        elif problems:
            noun = "callback needs" if problems == 1 else "callbacks need"
            self.health.config(text=f"{icon('alert')}  {problems} {noun} a macro", style="WarnLink.TLabel")
        else:
            noun = "macro" if self._linked == 1 else "macros"
            self.health.config(text=f"{icon('check')}  {self._linked} {noun} linked", style="OkNote.TLabel")

    # -- showing the selection -----------------------------------------------------

    def _on_select(self, _event=None) -> None:
        if self._refreshing:
            return
        if self.inspector.commit_pending():
            return  # the edit refreshed everything, selection included
        self._show_selection()

    def _show_selection(self) -> None:
        elements = self.selected_elements()
        single = elements[0] if len(elements) == 1 else None
        self.inspector.show(single, len(elements))
        anchor = elements[0] if elements else None
        tab = self._tab_of(anchor) if anchor is not None else None
        if tab is None and anchor is None:
            tab = next((e for e in self.elements if rs.local_name(e.name) == "tab"), None)
        message = "" if anchor is None else "This part of the ribbon is not drawn here; the list on the left has it."
        self.preview.show(
            self.elements, tab, frozenset(e.index for e in elements), self.photo_for,
            message if self.elements else "",
        )
        self._show_crumb(single)
        self._show_peek(single)
        part = self._part()
        self.note.config(text="Read-only: malformed when opened" if part is not None and not part.editable else "")

    def _tab_of(self, element: rs.RibbonElement | None) -> rs.RibbonElement | None:
        while element is not None and rs.local_name(element.name) != "tab":
            element = self.parent_of(element)
        return element

    def _show_crumb(self, element: rs.RibbonElement | None) -> None:
        if element is None:
            self.crumb.config(text=(self.part_path or "").rsplit("/", 1)[-1])
            self.crumb_leaf.config(text="")
            return
        names = rs.element_path(self.elements, element).split(" › ")
        if rs.local_name(element.name) in _SEPARATORS:
            names.append("Separator")
        self.crumb.config(text="".join(f"{name}   ›   " for name in names[:-1]))
        self.crumb_leaf.config(text=names[-1] if names else "")

    def _show_peek(self, element: rs.RibbonElement | None) -> None:
        draft = self.app.draft
        if element is None or draft is None:
            self.peek.show(None, [], "Select a button to see the macro it runs.")
            return
        callbacks = rs.callback_attrs(element)
        name = next((v for a, v in callbacks if a == "onAction"), callbacks[0][1] if callbacks else "")
        if not name:
            local = rs.local_name(element.name)
            self.peek.show(
                None, [],
                "No macro yet. Type or pick one under Macro; a name that does not exist can be created from there."
                if local in _ACTION
                else "Select a button to see the macro it runs.",
            )
            return
        resolution = self.resolve(name)
        if not resolution.matches:
            self.peek.show(None, [], resolution.message or f"{name} was not found.")
            return
        location = resolution.matches[0]
        module = draft.module_by_id(location.module_id)
        source = procedure_source(module.body, location.line) if module is not None else []
        self.peek.show(location, source, resolution.message or "" if resolution.status != rs.STATUS_OK else "")

    # -- lookups the form uses ------------------------------------------------------

    def resolve(self, name: str) -> rs.Resolution:
        draft = self.app.draft
        if draft is None:
            return rs.Resolution(rs.STATUS_MISSING, ())
        return rs.resolve(rs.procedure_index(draft), name)

    def macro_names(self) -> list[str]:
        """Subs in standard modules, by name: what a button can run."""
        draft = self.app.draft
        if draft is None:
            return []
        names = {
            loc.name
            for locations in rs.procedure_index(draft).values()
            for loc in locations
            if loc.is_standard and loc.kind == "Sub"
        }
        return sorted(names, key=str.casefold)

    def image_names(self) -> list[str]:
        draft = self.app.draft
        return self.app._image_suggestions(draft) if draft is not None else []

    # -- editing ---------------------------------------------------------------------

    def _editable_part(self):
        """The part to edit, with the editors flushed; None (and a status line) when it cannot be."""
        app = self.app
        if app.draft is None or app._saving:
            return None
        app._flush_all_editors()
        part = self._part()
        if part is None or part.text is None:
            return None
        if not part.editable:
            app.status.config(text="This ribbon XML was malformed when the file was opened, so it is read-only.")
            return None
        return part

    def _resolve(self, text: str, element: rs.RibbonElement, text_then: str) -> rs.RibbonElement | None:
        """``element`` (found in ``text_then``) as it is in ``text`` now."""
        elements = rs.element_tree(text)
        if element.index < len(elements):
            same = elements[element.index]
            if same.name == element.name and (text == text_then or same.attr("id") == element.attr("id")):
                return same
        ident = element.attr("id")
        if not ident:
            return None
        return next((e for e in elements if e.attr("id") == ident and e.name == element.name), None)

    def _apply(self, element: rs.RibbonElement, text_then: str, transform, *, also=None, status: str = "") -> bool:
        """Edit the part with ``transform(text, element) -> new text | (new text, start to select)``."""
        part = self._editable_part()
        if part is None:
            return False
        target = self._resolve(part.text or "", element, text_then)
        if target is None:
            self.app.status.config(text="That control is no longer in the ribbon XML.")
            self.refresh()
            return False
        try:
            result = transform(part.text or "", target)
        except rs.RibbonEditError as exc:
            self.app.status.config(text=str(exc))
            return False
        new_text, select = result if isinstance(result, tuple) else (result, None)
        if new_text == part.text and also is None:
            return False
        if select is None:
            # An attribute edit leaves the tag where it was. Keep the control selected even when
            # its id changed (the selection is restored by id), unless another row was picked
            # meanwhile: a focus-out commit arrives after the click that moved the selection.
            chosen = self.selected_elements()
            if len(chosen) == 1 and chosen[0].index == target.index and chosen[0].name == target.name:
                select = target.start

        def mutate() -> None:
            part.text = new_text
            if also is not None:
                also()

        self.app._batch_edit(mutate)
        if select is not None:
            self.select_start(select)
        if status:
            self.app.status.config(text=status)
        return True

    def set_attrs(self, element: rs.RibbonElement, text_then: str, changes: dict[str, str | None]) -> bool:
        return self._apply(element, text_then, lambda text, target: rs.set_element_attrs(text, target, changes))

    def _all_ids(self) -> set[str]:
        draft = self.app.draft
        return self.app._ribbon_ids(draft) if draft is not None else set()

    def set_label(self, element: rs.RibbonElement, text_then: str, value: str) -> bool:
        changes: dict[str, str | None] = {"label": value or None}
        ident = element.attr("id")
        if value and ident and ident in self._fresh:
            # A control just added here still has its placeholder id: name it after the label.
            prefix = _ID_PREFIX.get(rs.local_name(element.name), "ctl")
            renamed = rs.unique_control_id(self._all_ids() - {ident}, value, prefix)
            changes["id"] = renamed
            self._fresh.discard(ident)
            self._fresh.add(renamed)
        return self.set_attrs(element, text_then, changes)

    def set_id(self, element: rs.RibbonElement, text_then: str, value: str) -> bool:
        old = element.attr("id") or ""
        if not _ID_RE.match(value):
            self.app.status.config(text="An id starts with a letter or underscore and uses letters, digits, _ . -")
            return False
        if value.casefold() in {item.casefold() for item in self._all_ids() - {old}}:
            self.app.status.config(text=f'Another control already has the id "{value}".')
            return False
        if not self.set_attrs(element, text_then, {"id": value}):
            return False
        self._fresh.discard(old)
        draft = self.app.draft
        uses = sum(m.body.count(f'"{old}"') for m in draft.modules if not m.is_deleted) if draft and old else 0
        if uses:
            self.app.status.config(text=f'Id changed to {value}. The VBA code still says "{old}" in {uses} place(s).')
        return True

    def set_macro(self, element: rs.RibbonElement, text_then: str, attr: str, value: str) -> bool:
        """Point a callback attribute at a macro; a plain Sub gets its (control) wrapper."""
        app = self.app
        draft = app.draft
        if draft is None:
            return False
        if not value:
            return self.set_attrs(element, text_then, {attr: None})
        callback, wrapper, module = value, None, None
        resolution = self.resolve(value)
        plain_button = attr == "onAction" and rs.local_name(element.name) == "button"
        if plain_button and resolution.status == rs.STATUS_OK:
            match = resolution.matches[0]
            if not rs.is_ribbon_ready(match.signature) and not draft.baseline.safety.password_protected:
                module = draft.module_by_id(match.module_id)
                plan = app._plan_button_callback(module, match) if module is not None else None
                if plan is None:
                    return False
                callback, wrapper, _note = plan
        new_body = rs.append_procedure(module.body, wrapper)[0] if module is not None and wrapper else None

        def also() -> None:
            if module is not None and new_body is not None:
                module.body = new_body

        label = element_label(element)
        status = f"{label} runs {callback}."
        if wrapper and module is not None:
            status = f"{label} runs {callback}, a new Sub in {module.current_name} that calls {value}."
        return self._apply(
            element,
            text_then,
            lambda text, target: rs.set_element_attrs(text, target, {attr: callback}),
            also=also if new_body is not None else None,
            status=status,
        )

    def set_image_mso(self, element: rs.RibbonElement, text_then: str, value: str) -> bool:
        if not element.attr("image"):
            return self.set_attrs(element, text_then, {"imageMso": value or None})
        return self._replace_picture(element, text_then, value) if value else False

    def remove_picture(self) -> None:
        element = self.selected_element()
        if element is not None and element.attr("image"):
            self._replace_picture(element, self.text, "")

    def _replace_picture(self, element: rs.RibbonElement, text_then: str, image_mso: str) -> bool:
        part = self._editable_part()
        draft = self.app.draft
        if part is None or draft is None:
            return False
        target = self._resolve(part.text or "", element, text_then)
        if target is None:
            return False
        failures: list[str] = []

        def mutate() -> None:
            try:
                icon_service.clear_custom_icon(draft, part.path, target.start, image_mso)
            except icon_service.IconError as exc:
                failures.append(str(exc))

        self.app._batch_edit(mutate)
        if failures:
            self.app.status.config(text=failures[0])
        return not failures

    # -- adding, deleting, moving ------------------------------------------------------

    def _post_add_menu(self) -> None:
        if self.app.draft is not None:
            x, y = self.winfo_pointerxy()
            self.add_menu.tk_popup(x, y)

    def _place_for(self, names: tuple[str, ...]) -> tuple[rs.RibbonElement, rs.RibbonElement | None, str] | None:
        """(parent, sibling to go after, element name) for a new control, from the selection."""
        elements = self.elements
        selected = self.selected_elements()
        current: rs.RibbonElement | None = selected[-1] if selected else None
        if current is None:
            return next(
                ((e, None, n) for e in elements for n in names if e.complete and rs.can_contain(e, n)), None
            )
        for name in names:
            if current.complete and rs.can_contain(current, name):
                return current, None, name
        while current is not None:
            parent = self.parent_of(current)
            for name in names:
                if parent is not None and parent.complete and rs.can_contain(parent, name):
                    return parent, current if current.complete else None, name
            current = parent
        # A group goes into the tab the selection is in even from deep inside it.
        return next(((e, None, n) for e in elements for n in names if e.complete and rs.can_contain(e, n)), None)

    def add_control(self, kind: str) -> bool:
        """Add a button, menu, group or separator at the selection and start naming it."""
        app = self.app
        if app.draft is None:
            return False
        self.inspector.commit_pending()
        if not app._ensure_ribbon(f"the {kind}"):
            return False
        if self._editable_part() is None:
            return False
        self.refresh()
        names = ("separator", "menuSeparator") if kind == "separator" else (kind,)
        place = self._place_for(names)
        if place is None:
            where = {"group": "a tab", "button": "a group or menu", "menu": "a group or menu"}.get(kind, "a group")
            app.status.config(text=f"Select {where} to put the {kind} in.")
            return False
        parent, after, name = place
        ids = self._all_ids()
        sized = rs.local_name(parent.name) == "group" and kind in ("button", "menu")
        size = ""
        if sized:
            siblings = [e for e in self.elements if e.parent == parent.index and rs.local_name(e.name) in _SIZED]
            model = after if after is not None and after in siblings else (siblings[-1] if siblings else None)
            size = "large" if model is None or model.attr("size") == "large" else ""
        if kind == "button":
            ident = rs.unique_control_id(ids, "NewButton")
            xml = rs.build_button_xml(control_id=ident, label="New Button", on_action="", size=size)
        elif kind == "menu":
            ident = rs.unique_control_id(ids, "NewMenu", "mnu")
            xml = rs.build_menu_xml(control_id=ident, label="New Menu", size=size, indent="", unit="")
        elif kind == "group":
            ident = rs.unique_control_id(ids, "NewGroup", "grp")
            xml = rs.build_group_xml(control_id=ident, label="New Group")
        else:
            taken = {item.casefold() for item in ids}
            ident = next(f"sep{n}" for n in range(1, len(taken) + 2) if f"sep{n}" not in taken)
            xml = rs.build_separator_xml(control_id=ident, in_menu=name == "menuSeparator")

        def transform(text: str, target: rs.RibbonElement):
            sibling = self._resolve(text, after, self.text) if after is not None else None
            new_text, offset, _length = rs.insert_child(text, target, xml, after=sibling)
            return new_text, offset

        where = element_label(parent) or rs.local_name(parent.name)
        if not self._apply(
            parent, self.text, transform,
            status=f"Added a {kind} to {where}. Nothing is written until you save.",
        ):
            return False
        if kind != "separator":
            self._fresh.add(ident)
            self.after_idle(self.inspector.focus_label)
        return True

    def delete_selected(self) -> bool:
        self.inspector.commit_pending()
        part = self._editable_part()
        if part is None:
            return False
        self.refresh()
        chosen = self.selected_elements()
        elements = self.elements
        doomed = [e for e in chosen if rs.is_free(elements, e)]
        if not doomed:
            if chosen:
                self.app.status.config(text=f"{chosen[0].describe()} is part of the ribbon's skeleton and stays.")
            return False
        # A control inside a selected container goes with it.
        outer = [e for e in doomed if not any(o.start < e.start and e.end <= o.end for o in doomed)]
        inside = sum(1 for e in elements for o in outer if o.start < e.start and e.end <= o.end)
        if inside:
            what = outer[0].describe() if len(outer) == 1 else f"{len(outer)} controls"
            if not messagebox.askyesno(
                APP_NAME, f"Delete {what} and the {inside} control(s) inside?\n\nUndo (Ctrl+Z) brings them back."
            ):
                return False
        neighbour = self._neighbour(outer[0])
        text = part.text or ""
        for element in sorted(outer, key=lambda e: e.start, reverse=True):
            text = rs.delete_element(text, element)

        def mutate() -> None:
            part.text = text

        self.app._batch_edit(mutate)
        if neighbour is not None:
            start = neighbour.start if neighbour.start < outer[0].start else None
            ident = neighbour.attr("id")
            found = next(
                (e for e in self.elements if (e.start == start) or (ident and e.attr("id") == ident)), None
            )
            if found is not None:
                self.select_start(found.start)
        count = len(outer) + inside
        self.app.status.config(
            text=f"Deleted {outer[0].describe() if count == 1 else f'{count} controls'}. Ctrl+Z brings it back."
        )
        return True

    def _neighbour(self, element: rs.RibbonElement) -> rs.RibbonElement | None:
        """What to select once ``element`` is gone: the row above it, else its parent."""
        iid = f"e{element.index}"
        if not self.tree.exists(iid):
            return self.parent_of(element)
        previous = self.tree.prev(iid)
        return self._element_of(previous or self.tree.parent(iid) or "")

    def move_selected(self, direction: int) -> bool:
        """Swap the selected control with the one above (-1) or below (+1) it."""
        self.inspector.commit_pending()
        element = self.selected_element()
        if element is None:
            self.app.status.config(text="Select one control to move.")
            return False

        def transform(text: str, target: rs.RibbonElement):
            new_text, moved = rs.move_element(text, target.start, direction)
            return new_text, moved.start

        return self._apply(
            element, self.text, transform,
            status=f"Moved {element.describe()} {'up' if direction < 0 else 'down'}.",
        )

    def move_to(self, source: rs.RibbonElement, target: rs.RibbonElement, where: str) -> bool:
        """Drop ``source`` "into" ``target`` (last), as its "first" child, or "before"/"after" it."""
        self.inspector.commit_pending()

        def transform(text: str, element: rs.RibbonElement):
            there = self._resolve(text, target, self.text)
            if there is None:
                raise rs.RibbonEditError("That place is no longer in the ribbon XML.")
            elements = rs.element_tree(text)
            if where == "into":
                new_text, moved = rs.relocate_element(text, element, there)
            elif where == "first":
                first = next((e for e in elements if e.parent == there.index and e.complete), None)
                new_text, moved = rs.relocate_element(text, element, there, before=first)
            else:
                parent = elements[there.parent] if there.parent >= 0 else None
                if parent is None:
                    raise rs.RibbonEditError("Nothing can be put there.")
                anchor = {"before" if where == "before" else "after": there}
                new_text, moved = rs.relocate_element(text, element, parent, **anchor)
            return new_text, moved.start

        return self._apply(source, self.text, transform, status=f"Moved {source.describe()}.")

    def group_selected_into_menu(self) -> None:
        self.inspector.commit_pending()
        chosen = self.selected_elements()
        if not chosen or self.part_path is None:
            self.app.status.config(text="Select the buttons to put in a menu.")
            return
        if len({e.parent for e in chosen}) > 1:
            self.app.status.config(text="Select buttons that sit next to each other in one group or menu.")
            return
        self.app.group_range_into_menu(self.part_path, chosen[0].start, chosen[-1].end)

    def set_icon_from_clipboard(self) -> None:
        self._set_icon(icon_service.icon_from_clipboard)

    def set_icon_from_file(self) -> None:
        if self._icon_target() is None:
            return
        name = ask_image_file(self.winfo_toplevel())
        if name:
            self._set_icon(lambda: icon_service.icon_from_file(Path(name)))

    def _icon_target(self) -> rs.RibbonElement | None:
        element = self.selected_element()
        if element is None or rs.local_name(element.name) not in rs.IMAGE_ELEMENTS:
            self.app.status.config(text="Select the button or menu that should show the icon.")
            return None
        return element

    def _set_icon(self, make_icon) -> None:
        self.inspector.commit_pending()
        element = self._icon_target()
        if element is None or self.part_path is None or self._editable_part() is None:
            return
        try:
            picture = make_icon()
        except icon_service.IconError as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return
        self.app.apply_ribbon_icon(self.part_path, element.start, picture)

    # -- going places --------------------------------------------------------------------

    def _callback(self, element: rs.RibbonElement, attr: str) -> rs.RibbonCallback | None:
        return next(
            (
                cb
                for cb in rs.find_callbacks(self.text, self.part_path or "")
                if cb.tag_start == element.start and cb.attribute == attr
            ),
            None,
        )

    def open_callback(self, element: rs.RibbonElement, attr: str, *, create: bool = False) -> bool:
        """Open the macro behind ``attr`` of ``element``; ``create`` makes a missing one."""
        self.refresh()
        current = self._resolve(self.text, element, self.text) or self.selected_element()
        callback = self._callback(current, attr) if current is not None else None
        if callback is None:
            return False
        if create:
            return bool(self.app.create_callback(callback))
        self.app.goto_callback(callback)
        return True

    def goto_macro(self) -> bool:
        """F12 / Enter: open the macro the selected control runs. False = it has none."""
        self.inspector.commit_pending()
        element = self.selected_element()
        if element is None:
            return False
        callbacks = rs.callback_attrs(element)
        attr = next((a for a, _v in callbacks if a == "onAction"), callbacks[0][0] if callbacks else None)
        return attr is not None and self.open_callback(element, attr)

    def show_xml(self) -> None:
        """The same control in the XML editor."""
        self.inspector.commit_pending()
        if self.part_path is None:
            return
        element = (self.selected_elements() or [None])[0]
        if element is None:
            self.app._open_xml_at(self.part_path, 1)
        else:
            self.app._select_xml_range(self.part_path, element.start, element.open_end - element.start)

    def focus_filter(self) -> None:
        if self.sidebar.filter_entry is not None:
            self.sidebar.filter_entry.focus_set()

    def _pick(self, index: int) -> None:
        if index < len(self.elements):
            self.inspector.commit_pending()
            self.select_start(self.elements[index].start)
            self.tree.focus_set()

    def _open_index(self, index: int) -> None:
        self._pick(index)
        self.goto_macro()

    def _on_double_click(self, event) -> str | None:
        if "indicator" in str(self.tree.identify_element(event.x, event.y)):
            return None
        if self.tree.identify_row(event.y) and self.goto_macro():
            return "break"
        return None

    # -- right-click ------------------------------------------------------------------------

    def prepare_context_menu(self, iid: str = "") -> tk.Menu:
        """Select the row under the pointer and set which entries apply."""
        if iid and iid not in self.tree.selection():
            self.tree.selection_set(iid)
            self.tree.focus(iid)
            self._on_select()
        chosen = self.selected_elements()
        single = chosen[0] if len(chosen) == 1 else None
        states = {
            "Go to Macro": single is not None and bool(rs.callback_attrs(single)),
            "Rename": single is not None and rs.local_name(single.name) not in _NO_LABEL,
            "Move Up": single is not None,
            "Move Down": single is not None,
            "Group into Menu…": bool(chosen),
            "Paste Icon from Clipboard": single is not None and rs.local_name(single.name) in rs.IMAGE_ELEMENTS,
            "Icon from File…": single is not None and rs.local_name(single.name) in rs.IMAGE_ELEMENTS,
            "Delete": bool(chosen),
        }
        menu = self.context_menu
        for index in range((menu.index("end") or 0) + 1):
            if menu.type(index) == "command":
                label = str(menu.entrycget(index, "label"))
                if label in states:
                    menu.entryconfigure(index, state="normal" if states[label] else "disabled")
        return menu

    def _on_right_click(self, event) -> str:
        menu = self.prepare_context_menu(self.tree.identify_row(event.y))
        menu.tk_popup(event.x_root, event.y_root)
        return "break"

    # -- drag and drop ------------------------------------------------------------------------

    def _drag_start(self, event) -> None:
        self._drag = None
        if "indicator" in str(self.tree.identify_element(event.x, event.y)):
            return
        iid = self.tree.identify_row(event.y)
        if iid and self._element_of(iid) is not None:
            self._drag = {"iid": iid, "y": event.y, "moving": False}

    def drop_target(self, source: rs.RibbonElement, iid: str, fraction: float) -> tuple[rs.RibbonElement, str] | None:
        """(element, where) for dropping ``source`` on row ``iid``, or None when it cannot go there."""
        target = self._element_of(iid)
        if target is None or target.index == source.index or not rs.is_free(self.elements, source):
            return None
        current: rs.RibbonElement | None = target
        while current is not None:  # not into itself
            if current.index == source.index:
                return None
            current = self.parent_of(current)
        parent = self.parent_of(target)
        into = target.complete and rs.can_contain(target, source.name)
        beside = (
            parent is not None and rs.can_contain(parent, source.name)
            and rs.local_name(parent.name) not in ("splitButton",)
        )
        opened = bool(self.tree.item(iid, "open")) and bool(self.tree.get_children(iid))
        where = drop_zone(fraction, into=into, beside=beside, opened=opened)
        return (target, where) if where is not None else None

    def _drag_motion(self, event) -> None:
        drag = self._drag
        if drag is None:
            return
        if not drag["moving"]:
            if abs(event.y - drag["y"]) < DRAG_THRESHOLD:
                return
            drag["moving"] = True
            self.tree.selection_set(drag["iid"])
            self.tree.configure(cursor="fleur")
        height = self.tree.winfo_height()
        if event.y < 14:
            self.tree.yview_scroll(-1, "units")
        elif event.y > height - 14:
            self.tree.yview_scroll(1, "units")
        self._show_drop(None)
        source = self._element_of(drag["iid"])
        iid = self.tree.identify_row(event.y)
        box = self.tree.bbox(iid) if iid else ""
        if source is None or not box:
            return
        _x, top, _w, row_height = box
        found = self.drop_target(source, iid, (event.y - top) / max(row_height, 1))
        if found is not None:
            self._show_drop((iid, found[1]))

    def _show_drop(self, drop: tuple[str, str] | None) -> None:
        tree = self.tree
        if self._drop is not None and tree.exists(self._drop[0]):
            old = self._drop[0]
            tree.item(old, tags=tuple(t for t in tree.item(old, "tags") if t != "drop"))
        self._drop_line.place_forget()
        self._drop = drop
        if drop is None:
            return
        iid, where = drop
        if where == "into":
            tree.item(iid, tags=(*tree.item(iid, "tags"), "drop"))
            return
        box = tree.bbox(iid)
        if not box:
            return
        x, top, _width, row_height = box
        indent = 16 if where == "first" else 0
        y = top if where == "before" else top + row_height - 1
        self._drop_line.place(x=x + indent, y=y, width=max(tree.winfo_width() - x - indent - 8, 20), height=2)

    def _drag_end(self, _event=None) -> None:
        drag, drop = self._drag, self._drop
        self._drag = None
        self._show_drop(None)
        if drag is None or not drag["moving"]:
            return
        self.tree.configure(cursor="")
        source = self._element_of(drag["iid"])
        target = self._element_of(drop[0]) if drop is not None else None
        if source is not None and target is not None and drop is not None:
            self.move_to(source, target, drop[1])
