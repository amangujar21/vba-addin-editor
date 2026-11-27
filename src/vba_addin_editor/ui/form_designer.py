"""The UserForms view: a designer for the forms of the open add-in.

Left: the forms and their controls. Middle: the form drawn to scale (points to
pixels at 96 DPI); click to select, drag to move, drag the corner square to
resize, double-click to open the control's event code. Right: the selected
control's (or form's) properties.

The designer keeps no layout of its own. Every change is a ``form_service``
edit that the window applies with ``apply_form_edit`` as one undo step; the
window's refresh redraws this view from the draft.
"""

from __future__ import annotations

import tkinter as tk
from dataclasses import dataclass
from tkinter import font as tkfont
from tkinter import ttk
from typing import TYPE_CHECKING

from vba_addin_editor.domain.document import FormControlSnapshot, FormDesignSnapshot, ModuleDraft
from vba_addin_editor.services import form_service as fs
from vba_addin_editor.ui import theme
from vba_addin_editor.ui.widgets import Tooltip, icon

if TYPE_CHECKING:
    from vba_addin_editor.ui.main_window import MainWindow

SCALE = 4 / 3  # pixels per point at 96 DPI
PAD = 24
TITLE = 24  # the drawn title bar, in pixels
HANDLE = 7
FRAME_INSET = (1.0, 7.0)  # where a Frame's inside starts, in points (its caption line)
CAPTION_KINDS = frozenset({"Label", "CommandButton", "CheckBox", "OptionButton", "ToggleButton", "Frame"})
FONTLESS_KINDS = frozenset({"Image", "SpinButton", "ScrollBar", "MultiPage", "Form"})


def _num(value: float | None) -> str:
    if value is None:
        return ""
    text = f"{value:.2f}".rstrip("0").rstrip(".")
    return text or "0"


@dataclass
class _Drawn:
    """Where a control was drawn: its box in canvas pixels and its snapshot."""

    control: FormControlSnapshot
    box: tuple[float, float, float, float]


class _Row:
    """One inspector row: a caption and an entry that commits on Return / focus out."""

    def __init__(self, designer: FormDesigner, parent, row: int, caption: str, key: str) -> None:
        self.designer = designer
        self.key = key
        self.loaded = ""
        self.var = tk.StringVar(master=parent)
        self.label = ttk.Label(parent, text=caption, style="Field.TLabel")
        self.entry = ttk.Entry(parent, textvariable=self.var, width=24)
        self.label.grid(row=row, column=0, sticky="w", padx=(0, 10), pady=2)
        self.entry.grid(row=row, column=1, sticky="we", pady=2)
        self.entry.bind("<Return>", lambda _e: self.commit())
        self.entry.bind("<FocusOut>", lambda _e: self.commit())
        self.entry.bind("<Escape>", lambda _e: self.var.set(self.loaded))
        self.shown = True

    def show(self, value: str, *, enabled: bool = True) -> None:
        if not self.shown:
            self.label.grid()
            self.entry.grid()
            self.shown = True
        self.loaded = value
        self.var.set(value)
        self.entry.state(["!disabled"] if enabled else ["disabled"])

    def hide(self) -> None:
        if self.shown:
            self.label.grid_remove()
            self.entry.grid_remove()
            self.shown = False

    def commit(self) -> None:
        value = self.var.get().strip()
        if not self.shown or value == self.loaded or "disabled" in self.entry.state():
            return
        if not self.designer.commit_field(self.key, value):
            self.var.set(self.loaded)


class FormDesigner(ttk.Frame):
    def __init__(self, parent, window: MainWindow) -> None:
        super().__init__(parent, style="Editor.TFrame")
        self.window = window
        self.module_id: str | None = None
        self.control_name: str | None = None  # None: the form itself
        self.page: dict[str, str] = {}  # MultiPage name -> page shown
        self.drawn: dict[str, _Drawn] = {}
        self.form_box: tuple[float, float, float, float] = (0, 0, 0, 0)
        self._drag: dict | None = None
        self._fonts: dict[tuple[str, int, bool], tkfont.Font] = {}
        self._items: dict[str, tuple[str, str | None]] = {}  # tree iid -> (module id, control)

        panes = ttk.Panedwindow(self, orient="horizontal")
        panes.pack(fill="both", expand=True)
        side, self.tree = window._build_sidebar(
            panes, "USERFORMS", add=window.new_userform, add_tip="New UserForm"
        )
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        panes.add(side, weight=0)

        main = ttk.Frame(panes, style="Editor.TFrame")
        bar = window._crumb_bar(main)
        ttk.Label(bar, text=icon("form"), style="CrumbSep.TLabel").pack(side="left", padx=(0, 6))
        self.heading = ttk.Label(bar, text="", style="CrumbStrong.TLabel", anchor="w")
        self.heading.pack(side="left", fill="x", expand=True)
        self.code_btn = ttk.Button(
            bar, text="View code", style="EditorText.TButton", command=self.open_code, takefocus=False
        )
        self.code_btn.pack(side="right")
        Tooltip(self.code_btn, "The form's code; for a control, its event procedure (double-click does this too)")
        self.delete_btn = ttk.Button(
            bar, text="Delete", style="EditorText.TButton", command=self.delete_selected, takefocus=False
        )
        self.delete_btn.pack(side="right", padx=(0, 4))
        self.add_btn = ttk.Menubutton(bar, text="Add control", style="Crumb.TMenubutton", takefocus=False)
        self.add_menu = tk.Menu(self.add_btn, tearoff=0)
        for kind in fs.CONTROL_KINDS:
            self.add_menu.add_command(label=kind, command=lambda k=kind: self.add_control(k))
        self.add_btn.configure(menu=self.add_menu)
        self.add_btn.pack(side="right", padx=(0, 4))
        window._menus.append(self.add_menu)
        self.import_btn = ttk.Button(
            bar, text="Import…", style="EditorText.TButton", command=window.import_userform, takefocus=False
        )
        self.import_btn.pack(side="right", padx=(0, 4))
        Tooltip(self.import_btn, "Copy a UserForm, layout and code, from another add-in")

        self.hint = ttk.Label(
            main,
            text=(
                "Click to select · drag to move · drag the corner square to resize · double-click "
                "for the event code. Layouts are copied and saved byte for byte; nothing is written "
                "until Save File."
            ),
            style="EditorInfo.TLabel",
            padding=(14, 6),
        )
        self.hint.pack(fill="x", side="bottom")

        body = ttk.Panedwindow(main, orient="horizontal")
        body.pack(fill="both", expand=True)
        holder = ttk.Frame(body, style="Editor.TFrame")
        self.canvas = tk.Canvas(holder, highlightthickness=0, borderwidth=0)
        xs = ttk.Scrollbar(holder, orient="horizontal", command=self.canvas.xview)
        ys = ttk.Scrollbar(holder, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(
            xscrollcommand=theme.autohide(xs, side="bottom", fill="x"),
            yscrollcommand=theme.autohide(ys, side="right", fill="y"),
        )
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.bind("<ButtonPress-1>", self.press)
        self.canvas.bind("<B1-Motion>", self.motion)
        self.canvas.bind("<ButtonRelease-1>", self.release)
        self.canvas.bind("<Double-Button-1>", self.double_click)
        self.canvas.bind("<Delete>", lambda _e: self.delete_selected())
        for key, dx, dy in (("Left", -1, 0), ("Right", 1, 0), ("Up", 0, -1), ("Down", 0, 1)):
            self.canvas.bind(f"<{key}>", lambda _e, x=dx, y=dy: self.nudge(x, y))
        body.add(holder, weight=1)

        inspector = ttk.Frame(body, style="Editor.TFrame", padding=(12, 10, 10, 6))
        body.add(inspector, weight=0)
        self.title = ttk.Label(inspector, text="", style="CrumbStrong.TLabel")
        self.title.pack(anchor="w", pady=(0, 6))
        grid = ttk.Frame(inspector, style="Editor.TFrame")
        grid.pack(fill="x")
        grid.columnconfigure(1, weight=1)
        rows = (
            ("Name", "name"),
            ("Caption", "caption"),
            ("Left", "left"),
            ("Top", "top"),
            ("Width", "width"),
            ("Height", "height"),
            ("Tip text", "tip"),
            ("Font", "font"),
            ("Font size", "size"),
        )
        self.rows = {key: _Row(self, grid, i, caption, key) for i, (caption, key) in enumerate(rows)}
        self.bold_var = tk.BooleanVar(master=grid, value=False)
        self.bold = ttk.Checkbutton(grid, text="Bold", variable=self.bold_var, command=self._bold_toggled)
        self.bold.grid(row=len(rows), column=1, sticky="w", pady=2)
        ttk.Label(inspector, text="STORED PROPERTIES", style="Caption.TLabel").pack(anchor="w", pady=(12, 2))
        props = ttk.Frame(inspector, style="Editor.TFrame")
        props.pack(fill="both", expand=True)
        self.props = ttk.Treeview(
            props, columns=("value",), show="tree headings", selectmode="browse", style="Editor.Treeview", height=8
        )
        self.props.heading("#0", text="Property", anchor="w")
        self.props.heading("value", text="Value", anchor="w")
        self.props.column("#0", width=150, stretch=False)
        self.props.column("value", width=140)
        sb = ttk.Scrollbar(props, orient="vertical", command=self.props.yview)
        self.props.configure(yscrollcommand=theme.autohide(sb, side="right", fill="y"))
        self.props.pack(side="left", fill="both", expand=True)
        panes.add(main, weight=1)
        self.apply_palette()

    # -- state ------------------------------------------------------------------------

    def apply_palette(self) -> None:
        self.canvas.configure(background=theme.palette().bg1)
        self._fonts.clear()
        self.redraw()

    def module(self) -> ModuleDraft | None:
        draft = self.window.draft
        if draft is None or self.module_id is None:
            return None
        mod = draft.module_by_id(self.module_id)
        return mod if mod is not None and draft.is_form(mod) else None

    def design(self) -> FormDesignSnapshot | None:
        draft, mod = self.window.draft, self.module()
        return fs.design_of(draft, mod) if draft is not None and mod is not None else None

    def control(self) -> FormControlSnapshot | None:
        design = self.design()
        if design is None or self.control_name is None:
            return None
        return design.control(self.control_name)

    def editable(self) -> bool:
        mod = self.module()
        return mod is not None and not mod.is_deleted and not self.window._saving

    # -- the tree ---------------------------------------------------------------------

    def refresh(self) -> None:
        """Rebuild from the draft, keeping the selection where it still exists."""
        tree = self.tree
        tree.delete(*tree.get_children())
        self._items.clear()
        draft = self.window.draft
        if draft is None:
            self.module_id, self.control_name = None, None
            self.heading.config(text="")
            self._show_inspector()
            self.redraw()
            return
        deleted = {m.id for m in draft.deleted_original_modules() if m.project_item_kind == "userform"}
        orphans = draft.baseline.orphan_form_names()
        shown: list[tuple[str, ModuleDraft | None, str]] = []
        for module in draft.modules:
            if draft.is_form(module) and (not module.is_deleted or module.id in deleted):
                shown.append((module.current_name, module, module.id))
        names = {name.casefold() for name, _m, _i in shown}
        for form in draft.baseline.forms:
            if form.name.casefold() not in names and form.name.casefold() in orphans:
                shown.append((form.name, None, ""))
        for name, module, _mid in shown:
            iid = f"form::{name}"
            design = draft.baseline.form(name) if module is None else fs.design_of(draft, module)
            label, badge, tags = name, "", ()
            if module is None:
                label += " [layout only — no code module]"
                tags = ("locked",)
            elif module.id in deleted:
                label += " [deleted — pending save]"
                tags = ("locked",)
            elif module.is_new:
                badge, tags = "A", ("new",)
            elif module in draft.changed_form_layouts() or module.body != module.original_body:
                badge, tags = "M", ("dirty",)
            if design is not None and design.problem:
                label += " [unreadable]"
            tree.insert("", "end", iid=iid, text=label, values=(badge,), tags=tags, open=True)
            self._items[iid] = (module.id if module is not None else "", None)
            if design is not None:
                self._insert_controls(iid, module.id if module is not None else "", design.controls)
        if not shown:
            self.module_id, self.control_name = None, None
            self.heading.config(text="This project has no UserForms. Use + or Import… to add one.")
        else:
            current = self._iid(self.module_id, self.control_name)
            if current is None or not tree.exists(current):
                current = self._iid(self.module_id, None)
            if current is None or not tree.exists(current):
                first = shown[0][1]
                self.module_id = first.id if first is not None else None
                self.control_name = None
                current = f"form::{shown[0][0]}"
            tree.selection_set(current)
            tree.see(current)
            self._apply_selection(current)
        self.redraw()

    def _insert_controls(self, parent: str, module_id: str, controls) -> None:
        for control in controls:
            if not control.name:  # a MultiPage's own TabStrip
                continue
            iid = f"{parent}/{control.name}"
            self.tree.insert(parent, "end", iid=iid, text=f"{control.name}  ({control.short_kind})", open=True)
            self._items[iid] = (module_id, control.name)
            self._insert_controls(iid, module_id, control.children)

    def _iid(self, module_id: str | None, control: str | None) -> str | None:
        for iid, (mid, name) in self._items.items():
            if mid == module_id and (name or "").casefold() == (control or "").casefold():
                return iid
        return None

    def _on_tree_select(self, _event=None) -> None:
        selection = self.tree.selection()
        if selection:
            self._apply_selection(selection[0])
            self.redraw()

    def _apply_selection(self, iid: str) -> None:
        module_id, control = self._items.get(iid, ("", None))
        self.module_id = module_id or None
        self.control_name = control
        if iid not in self._items:
            return
        # A control on a MultiPage page shows that page.
        design = self.design()
        if design is not None and control is not None:
            self._reveal_page(design, control)
        form_name = iid.split("/", 1)[0].removeprefix("form::")
        design_text = ""
        if design is not None:
            if control is None:
                design_text = f"{form_name}  ·  {design.control_count()} control(s)"
                if design.problem:
                    design_text += f"  ·  the layout could not be read: {design.problem}"
            else:
                item = design.control(control)
                kind = item.short_kind if item is not None else ""
                design_text = f"{form_name}  ›  {control}  ({kind})"
        elif not self.module_id:
            design_text = f"{form_name}  ·  a layout with no code module (kept as it is)"
        self.heading.config(text=design_text)
        self._show_inspector()

    def _reveal_page(self, design: FormDesignSnapshot, name: str) -> None:
        def visit(items, multipage: str | None, page: str | None) -> bool:
            for item in items:
                if item.name.casefold() == name.casefold():
                    if multipage is not None and page is not None:
                        self.page[multipage] = page
                    if item.short_kind == "Form" and multipage is not None:
                        self.page[multipage] = item.name
                    return True
                mp = item.name if item.short_kind == "MultiPage" else multipage
                pg = item.name if item.short_kind == "Form" else page
                if visit(item.children, mp, pg):
                    return True
            return False

        visit(design.controls, None, None)

    def select(self, form: str, control: str | None = None) -> None:
        """Select a form (and one of its controls) by name, as a click in the tree does."""
        draft = self.window.draft
        if draft is None:
            return
        module = draft.find_current(form)
        if module is None:
            return
        self.module_id, self.control_name = module.id, control
        iid = self._iid(module.id, control)
        if iid is not None:
            self.tree.selection_set(iid)
            self.tree.see(iid)
            self._apply_selection(iid)
        self.redraw()

    # -- the inspector ------------------------------------------------------------------

    def _show_inspector(self) -> None:
        props = self.props
        props.delete(*props.get_children())
        design, control = self.design(), self.control()
        rows = self.rows
        editable = self.editable() and design is not None and not design.problem
        self.code_btn.config(state="normal" if self.module() is not None else "disabled")
        self.add_btn.config(state="normal" if editable else "disabled")
        module = self.module()
        self.delete_btn.config(state="normal" if module is not None and not module.is_deleted else "disabled")
        if design is None:
            self.title.config(text="")
            for row in rows.values():
                row.hide()
            self.bold.grid_remove()
            return
        if control is None:
            self.title.config(text=f"UserForm  {design.name}")
            rows["name"].show(design.name, enabled=False)
            rows["caption"].show(design.caption, enabled=editable)
            rows["width"].show(_num(design.width), enabled=editable)
            rows["height"].show(_num(design.height), enabled=editable)
            for key in ("left", "top", "tip", "font", "size"):
                rows[key].hide()
            self.bold.grid_remove()
            items = design.properties
        else:
            kind = control.short_kind
            can = editable and control.editable
            self.title.config(text=f"{kind}  {control.name}")
            rows["name"].show(control.name, enabled=can)
            if kind in CAPTION_KINDS:
                rows["caption"].show(control.caption or "", enabled=can)
            else:
                rows["caption"].hide()
            rows["left"].show(_num(control.left), enabled=can)
            rows["top"].show(_num(control.top), enabled=can)
            sized = can and kind not in fs.FIXED_SIZE_KINDS
            rows["width"].show(_num(control.width), enabled=sized)
            rows["height"].show(_num(control.height), enabled=sized)
            rows["tip"].show(control.tip or "", enabled=can)
            if kind in FONTLESS_KINDS:
                rows["font"].hide()
                rows["size"].hide()
                self.bold.grid_remove()
            else:
                face, size, bold = control.font or ("Tahoma", 8.25, False)
                rows["font"].show(face, enabled=can)
                rows["size"].show(_num(size), enabled=can)
                self.bold_var.set(bold)
                self.bold.grid()
                self.bold.state(["!disabled"] if can else ["disabled"])
            items = control.properties
        for name, value in items:
            props.insert("", "end", text=name, values=(value,))

    def commit_field(self, key: str, value: str) -> bool:
        """Apply one inspector field; False puts the old value back."""
        mod, design = self.module(), self.design()
        if mod is None or design is None or not self.editable():
            return False
        control = self.control()
        name = control.name if control is not None else None
        try:
            if key in ("left", "top", "width", "height", "size"):
                number = float(value.replace(",", "."))
            else:
                number = 0.0
        except ValueError:
            self.bell()
            return False
        if control is None:
            if key == "caption":
                return self.window.apply_form_edit(mod, lambda e: e.set_form_caption(value), "Caption")
            if key in ("width", "height"):
                width = number if key == "width" else design.width or 240
                height = number if key == "height" else design.height or 180
                return self.window.apply_form_edit(
                    mod, lambda e: e.set_form_size(width, height), "Size"
                )
            return False
        assert name is not None
        if key == "name":
            problem = fs.control_name_problem(value, design, current=name)
            if problem:
                self.window.show_error(problem)
                return False
            return self.window.apply_form_edit(
                mod, lambda e: e.rename_control(name, value), "Name", rename=(name, value), select=value
            )
        if key == "caption":
            return self.window.apply_form_edit(mod, lambda e: e.set_caption(name, value), "Caption")
        if key in ("left", "top", "width", "height"):
            return self.window.apply_form_edit(
                mod, lambda e: e.set_geometry(name, **{key: number}), key.title()
            )
        if key == "tip":
            return self.window.apply_form_edit(mod, lambda e: e.set_tip(name, value or None), "Tip text")
        if key in ("font", "size"):
            face, size, bold = control.font or ("Tahoma", 8.25, False)
            face = value if key == "font" else face
            size = number if key == "size" else size
            if not face or size <= 0:
                return False
            return self.window.apply_form_edit(mod, lambda e: e.set_font(name, face, size, bold), "Font")
        return False

    def _bold_toggled(self) -> None:
        mod, control = self.module(), self.control()
        if mod is None or control is None or not self.editable():
            return
        face, size, _bold = control.font or ("Tahoma", 8.25, False)
        bold = bool(self.bold_var.get())
        name = control.name
        if not self.window.apply_form_edit(mod, lambda e: e.set_font(name, face, size, bold), "Font"):
            self.bold_var.set(not bold)

    # -- commands ------------------------------------------------------------------------

    def add_control(self, kind: str) -> bool:
        """Add a ``kind`` control to the selected container (a Frame or page), else to the form."""
        mod, design = self.module(), self.design()
        if mod is None or design is None or not self.editable():
            return False
        container: str | None = None
        control = self.control()
        if control is not None and control.short_kind in fs.CONTAINER_KINDS:
            container = control.name
        elif control is not None and control.short_kind == "MultiPage":
            page = self.page.get(control.name) or next(
                (c.name for c in control.children if c.short_kind == "Form"), None
            )
            container = page
        name = fs.next_control_name(design, kind)
        left, top = self._free_spot(design, container)
        ok = self.window.apply_form_edit(
            mod,
            lambda e: e.add_control(kind, name, container=container, left=left, top=top),
            f"Add {kind}",
            select=name,
        )
        if ok:
            self.control_name = name
            self.select(mod.current_name, name)
        return ok

    def _free_spot(self, design: FormDesignSnapshot, container: str | None) -> tuple[float, float]:
        """Below the lowest control of the container, so a new control does not cover one."""
        items = design.controls
        if container is not None:
            holder = design.control(container)
            items = holder.children if holder is not None else ()
        bottom = max((c.top + (c.height or 0) for c in items if c.name), default=0.0)
        return 6.0, round(bottom + 6.0 if items else 6.0)

    def delete_selected(self) -> None:
        mod, control = self.module(), self.control()
        if mod is None:
            return
        if control is None:
            self.window.delete_userform(mod)
            return
        if not self.editable() or not control.editable:
            self.window.show_error(f"{control.name} belongs to its MultiPage and cannot be deleted here.")
            return
        name = control.name
        if self.window.apply_form_edit(mod, lambda e: e.remove_control(name), f"Delete {name}"):
            self.control_name = None
            self.select(mod.current_name, None)

    def nudge(self, dx: float, dy: float) -> None:
        control, mod = self.control(), self.module()
        if control is None or mod is None or not control.editable or not self.editable():
            return
        name, left, top = control.name, control.left + dx, control.top + dy
        self.window.apply_form_edit(mod, lambda e: e.set_geometry(name, left=left, top=top), "Move")

    def open_code(self) -> None:
        """The form's code, at the selected control's event procedure when it has one."""
        mod = self.module()
        if mod is not None:
            self.window.open_form_code(mod, self.control_name, create=False)

    # -- drawing -------------------------------------------------------------------------

    def _font(self, face: str, points: float, bold: bool) -> tkfont.Font:
        px = max(6, round(points * SCALE))
        key = (face, px, bold)
        found = self._fonts.get(key)
        if found is None:
            found = tkfont.Font(root=self.canvas, family=face, size=-px, weight="bold" if bold else "normal")
            self._fonts[key] = found
        return found

    def redraw(self) -> None:
        canvas = self.canvas
        canvas.delete("all")
        self.drawn.clear()
        design = self.design()
        colors = theme.palette()
        if design is None or design.problem or design.width is None or design.height is None:
            text = "" if design is None else "The layout could not be read; it is kept as it is."
            canvas.create_text(PAD, PAD, anchor="nw", text=text, fill=colors.faint, font=theme.SMALL_FONT)
            return
        w, h = design.width * SCALE, design.height * SCALE
        x0, y0 = PAD, PAD
        canvas.create_rectangle(x0, y0, x0 + w + 2, y0 + TITLE + h + 2, outline=colors.border_strong, fill=colors.bg3)
        canvas.create_text(x0 + 8, y0 + TITLE / 2, anchor="w", text=design.caption, fill=colors.text, font=theme.SMALL_FONT)
        canvas.create_text(x0 + w - 8, y0 + TITLE / 2, anchor="e", text="✕", fill=colors.faint, font=theme.SMALL_FONT)
        cx, cy = x0 + 1, y0 + TITLE + 1
        self.form_box = (cx, cy, cx + w, cy + h)
        canvas.create_rectangle(cx, cy, cx + w, cy + h, outline="", fill=colors.bg2, tags=("form",))
        self._draw_items(design.controls, cx, cy)
        selected = self.control()
        if selected is not None and selected.name in self.drawn:
            bx0, by0, bx1, by1 = self.drawn[selected.name].box
            canvas.create_rectangle(bx0 - 2, by0 - 2, bx1 + 2, by1 + 2, outline=colors.accent, dash=(3, 2), tags=("sel",))
            if selected.editable and selected.short_kind not in fs.FIXED_SIZE_KINDS:
                canvas.create_rectangle(
                    bx1 - HANDLE / 2, by1 - HANDLE / 2, bx1 + HANDLE / 2, by1 + HANDLE / 2,
                    outline=colors.accent, fill=colors.accent, tags=("handle",),
                )
        elif selected is None:
            fx1, fy1 = cx + w, cy + h
            canvas.create_rectangle(
                fx1 - HANDLE / 2, fy1 - HANDLE / 2, fx1 + HANDLE / 2, fy1 + HANDLE / 2,
                outline=colors.accent, fill=colors.accent, tags=("handle", "formhandle"),
            )
        canvas.configure(scrollregion=(0, 0, x0 + w + PAD * 2, y0 + TITLE + h + PAD * 2))

    def _draw_items(self, items, ox: float, oy: float) -> None:
        for item in items:
            if not item.name:
                continue
            x, y = ox + item.left * SCALE, oy + item.top * SCALE
            w = (item.width or 72) * SCALE
            h = (item.height or 18) * SCALE
            box = (x, y, x + w, y + h)
            self.drawn[item.name] = _Drawn(item, box)
            self._draw_control(item, box)
            kind = item.short_kind
            if kind == "Frame":
                self._draw_items(item.children, x + FRAME_INSET[0] * SCALE, y + FRAME_INSET[1] * SCALE)
            elif kind == "MultiPage":
                pages = [c for c in item.children if c.short_kind == "Form"]
                shown = self.page.get(item.name) or (pages[0].name if pages else None)
                for page in pages:
                    if page.name == shown:
                        px, py = x + page.left * SCALE, y + page.top * SCALE
                        self.drawn[page.name] = _Drawn(page, (px, py, px + (page.width or 0) * SCALE, py + (page.height or 0) * SCALE))
                        self._draw_items(page.children, px, py)
                    else:
                        self.drawn.pop(page.name, None)

    def _draw_control(self, item: FormControlSnapshot, box) -> None:
        canvas, colors = self.canvas, theme.palette()
        x0, y0, x1, y1 = box
        kind = item.short_kind
        tag = ("control", f"c:{item.name}")
        face, size, bold = item.font or ("Tahoma", 8.25, False)
        font = self._font(face, size, bold)
        text = item.caption if item.caption is not None else ""
        if kind == "Label":
            canvas.create_rectangle(x0, y0, x1, y1, outline="", fill="", tags=tag)
            canvas.create_text(x0 + 1, y0 + 1, anchor="nw", text=text, fill=colors.text, font=font,
                               width=max(x1 - x0, 1), tags=tag)
        elif kind in ("TextBox", "ListBox", "ComboBox"):
            canvas.create_rectangle(x0, y0, x1, y1, outline=colors.border_strong, fill=colors.bg0, tags=tag)
            if kind == "ComboBox":
                bw = min(14, x1 - x0)
                canvas.create_rectangle(x1 - bw, y0, x1, y1, outline=colors.border_strong, fill=colors.bg3, tags=tag)
                canvas.create_text(x1 - bw / 2, (y0 + y1) / 2, text="▾", fill=colors.muted, font=theme.SMALL_FONT, tags=tag)
        elif kind in ("CheckBox", "OptionButton"):
            canvas.create_rectangle(x0, y0, x1, y1, outline="", fill="", tags=tag)
            mid, s = (y0 + y1) / 2, 11
            if kind == "CheckBox":
                canvas.create_rectangle(x0 + 1, mid - s / 2, x0 + 1 + s, mid + s / 2, outline=colors.muted, fill=colors.bg0, tags=tag)
            else:
                canvas.create_oval(x0 + 1, mid - s / 2, x0 + 1 + s, mid + s / 2, outline=colors.muted, fill=colors.bg0, tags=tag)
            canvas.create_text(x0 + s + 6, mid, anchor="w", text=text, fill=colors.text, font=font, tags=tag)
        elif kind in ("CommandButton", "ToggleButton"):
            canvas.create_rectangle(x0, y0, x1, y1, outline=colors.border_strong, fill=colors.bg3, tags=tag)
            canvas.create_text((x0 + x1) / 2, (y0 + y1) / 2, text=text, fill=colors.text, font=font,
                               width=max(x1 - x0 - 4, 1), tags=tag)
        elif kind == "Frame":
            canvas.create_rectangle(x0, y0, x1, y1, outline="", fill="", tags=tag)
            canvas.create_rectangle(x0, y0 + 6, x1, y1, outline=colors.border_strong, tags=tag)
            if text:
                label = canvas.create_text(x0 + 8, y0 + 6, anchor="w", text=f" {text} ", fill=colors.text, font=font, tags=tag)
                bx = canvas.bbox(label)
                if bx:
                    canvas.create_rectangle(*bx, outline="", fill=colors.bg2, tags=tag)
                    canvas.tag_raise(label)
        elif kind in ("MultiPage", "TabStrip"):
            canvas.create_rectangle(x0, y0, x1, y1, outline="", fill="", tags=tag)
            tabs = [c for c in item.children if c.short_kind == "Form"] if kind == "MultiPage" else []
            shown = self.page.get(item.name) or (tabs[0].name if tabs else None)
            tx = x0 + 2
            band = 18
            for page in tabs or [FormControlSnapshot("Tab1", "Tab", ())]:
                label = page.caption or page.name
                width = self._font("Tahoma", 8.25, False).measure(label) + 14
                active = page.name == shown
                canvas.create_rectangle(tx, y0, tx + width, y0 + band, outline=colors.border_strong,
                                        fill=colors.bg2 if active else colors.bg3, tags=(*tag, f"tab:{item.name}:{page.name}"))
                canvas.create_text(tx + 7, y0 + band / 2, anchor="w", text=label, fill=colors.text if active else colors.muted,
                                   font=theme.SMALL_FONT, tags=(*tag, f"tab:{item.name}:{page.name}"))
                tx += width
            canvas.create_rectangle(x0, y0 + band, x1, y1, outline=colors.border_strong, tags=tag)
        elif kind == "Image":
            canvas.create_rectangle(x0, y0, x1, y1, outline=colors.border_strong, fill=colors.bg3, tags=tag)
            has_picture = any(name == "Picture" for name, _v in item.properties)
            canvas.create_text((x0 + x1) / 2, (y0 + y1) / 2, text="🖼" if has_picture else "Image",
                               fill=colors.faint, font=theme.SMALL_FONT, tags=tag)
        elif kind in ("ScrollBar", "SpinButton"):
            canvas.create_rectangle(x0, y0, x1, y1, outline=colors.border_strong, fill=colors.bg3, tags=tag)
            vertical = (y1 - y0) >= (x1 - x0)
            a, b = ("▴", "▾") if vertical else ("◂", "▸")
            if vertical:
                canvas.create_text((x0 + x1) / 2, y0 + 7, text=a, fill=colors.muted, font=theme.SMALL_FONT, tags=tag)
                canvas.create_text((x0 + x1) / 2, y1 - 7, text=b, fill=colors.muted, font=theme.SMALL_FONT, tags=tag)
            else:
                canvas.create_text(x0 + 7, (y0 + y1) / 2, text=a, fill=colors.muted, font=theme.SMALL_FONT, tags=tag)
                canvas.create_text(x1 - 7, (y0 + y1) / 2, text=b, fill=colors.muted, font=theme.SMALL_FONT, tags=tag)
        else:
            canvas.create_rectangle(x0, y0, x1, y1, outline=colors.border_strong, dash=(2, 2), tags=tag)
            canvas.create_text(x0 + 3, y0 + 2, anchor="nw", text=kind, fill=colors.faint, font=theme.SMALL_FONT, tags=tag)

    # -- mouse ---------------------------------------------------------------------------

    def control_box(self, name: str) -> tuple[float, float, float, float] | None:
        drawn = self.drawn.get(name)
        return drawn.box if drawn is not None else None

    def _hit(self, x: float, y: float) -> tuple[str, str | None]:
        """("handle"|"tab"|"control"|"form"|"", name) under canvas point (x, y)."""
        canvas = self.canvas
        for item in reversed(canvas.find_overlapping(x - 1, y - 1, x + 1, y + 1)):
            tags = canvas.gettags(item)
            if "handle" in tags:
                return "handle", None
            for tag in tags:
                if tag.startswith("tab:"):
                    return "tab", tag[4:]
            for tag in tags:
                if tag.startswith("c:"):
                    return "control", tag[2:]
            if "form" in tags:
                return "form", None
        return "", None

    def _point(self, event) -> tuple[float, float]:
        return self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)

    def press(self, event) -> None:
        self.canvas.focus_set()
        x, y = self._point(event)
        what, name = self._hit(x, y)
        self._drag = None
        if what == "handle":
            control = self.control()
            box = self.drawn[control.name].box if control is not None and control.name in self.drawn else self.form_box
            self._drag = {"mode": "resize", "start": (x, y), "box": box, "name": control.name if control else None}
            return
        if what == "tab" and name is not None:
            multipage, page = name.split(":", 1)
            self.page[multipage] = page
            self._select_control(page)
            return
        if what == "control" and name is not None:
            self._select_control(name)
            drawn = self.drawn.get(name)
            if drawn is not None and drawn.control.editable:
                self._drag = {"mode": "move", "start": (x, y), "box": drawn.box, "name": name}
            return
        if what == "form":
            self._select_control(None)

    def _select_control(self, name: str | None) -> None:
        mod = self.module()
        if mod is None:
            return
        if name == self.control_name and self.tree.selection():
            return
        self.select(mod.current_name, name)

    def motion(self, event) -> None:
        drag = self._drag
        if drag is None:
            return
        x, y = self._point(event)
        sx, sy = drag["start"]
        x0, y0, x1, y1 = drag["box"]
        dx, dy = x - sx, y - sy
        if drag["mode"] == "move":
            box = (x0 + dx, y0 + dy, x1 + dx, y1 + dy)
        else:
            box = (x0, y0, max(x0 + 4, x1 + dx), max(y0 + 4, y1 + dy))
        drag["now"] = (dx, dy)
        self.canvas.delete("ghost")
        self.canvas.create_rectangle(*box, outline=theme.palette().accent, dash=(4, 2), tags=("ghost",))

    def release(self, event) -> None:
        drag, self._drag = self._drag, None
        self.canvas.delete("ghost")
        if drag is None:
            return
        x, y = self._point(event)
        sx, sy = drag["start"]
        dx, dy = (x - sx) / SCALE, (y - sy) / SCALE
        if abs(dx) < 1.5 and abs(dy) < 1.5:
            return
        mod, design = self.module(), self.design()
        if mod is None or design is None or not self.editable():
            return
        name = drag["name"]
        if drag["mode"] == "move":
            control = design.control(name)
            if control is None:
                return
            left, top = round(control.left + dx), round(control.top + dy)
            self.window.apply_form_edit(mod, lambda e: e.set_geometry(name, left=left, top=top), "Move")
        elif name is None:
            width = round((design.width or 0) + dx)
            height = round((design.height or 0) + dy)
            self.window.apply_form_edit(mod, lambda e: e.set_form_size(width, height), "Size")
        else:
            control = design.control(name)
            if control is None:
                return
            width = max(2, round((control.width or 0) + dx))
            height = max(2, round((control.height or 0) + dy))
            self.window.apply_form_edit(
                mod, lambda e: e.set_geometry(name, width=width, height=height), "Size"
            )

    def double_click(self, event) -> None:
        x, y = self._point(event)
        what, name = self._hit(x, y)
        mod = self.module()
        if mod is None or mod.is_deleted:
            return
        if what == "control" and name is not None:
            drawn = self.drawn.get(name)
            if drawn is not None and drawn.control.short_kind == "Form":
                return  # a MultiPage page: its events belong to the MultiPage
            self.window.open_form_code(mod, name, create=True)
        elif what == "form":
            self.window.open_form_code(mod, None, create=True)
