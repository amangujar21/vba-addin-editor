"""Ribbon dialogs: button list, Add Ribbon Button form, and a simple chooser."""

from __future__ import annotations

import re
import tkinter as tk
from dataclasses import dataclass
from tkinter import messagebox, ttk

from vba_addin_editor.services import ribbon_service as rs
from vba_addin_editor.version import APP_NAME

_ID_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_.-]*$")


def choose(root, title: str, prompt: str, options: list[str], initial: int = 0) -> int | None:
    """Modal single-choice list. Returns the chosen index or None."""
    win = tk.Toplevel(root)
    win.title(title)
    win.transient(root)
    win.grab_set()
    ttk.Label(win, text=prompt, padding=(8, 8, 8, 4), wraplength=420).pack(anchor="w")
    listbox = tk.Listbox(win, height=min(max(len(options), 3), 14), width=60, exportselection=False)
    listbox.pack(fill="both", expand=True, padx=8)
    for option in options:
        listbox.insert("end", option)
    if options:
        listbox.selection_set(min(initial, len(options) - 1))
        listbox.see(min(initial, len(options) - 1))
    result: list[int | None] = [None]

    def confirm(_event=None) -> None:
        sel = listbox.curselection()
        if sel:
            result[0] = int(sel[0])
        win.destroy()

    buttons = ttk.Frame(win, padding=8)
    buttons.pack(fill="x")
    ttk.Button(buttons, text="Cancel", command=win.destroy).pack(side="right", padx=4)
    ttk.Button(buttons, text="OK", command=confirm).pack(side="right")
    listbox.bind("<Double-Button-1>", confirm)
    win.bind("<Return>", confirm)
    win.bind("<Escape>", lambda _e: win.destroy())
    listbox.focus_set()
    root.wait_window(win)
    return result[0]


@dataclass(frozen=True)
class ButtonSpec:
    container: rs.RibbonContainer
    control_id: str
    label: str
    image_mso: str
    size: str
    screentip: str
    supertip: str


class AddButtonDialog:
    """Form for a new ribbon button bound to one callback."""

    def __init__(
        self,
        root,
        *,
        callback_name: str,
        note: str,
        containers: list[rs.RibbonContainer],
        existing_ids: set[str],
        image_suggestions: list[str],
        default_label: str,
        default_id: str,
        show_part: bool,
    ) -> None:
        self.root = root
        self.containers = containers
        self.existing_ids = {item.casefold() for item in existing_ids}
        self.result: ButtonSpec | None = None
        win = self.win = tk.Toplevel(root)
        win.title("Add Ribbon Button")
        win.transient(root)
        win.resizable(True, False)
        frame = ttk.Frame(win, padding=10)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)

        row = 0
        ttk.Label(frame, text="Runs macro:").grid(row=row, column=0, sticky="e", padx=4, pady=3)
        ttk.Label(frame, text=callback_name, font=("Consolas", 10, "bold")).grid(
            row=row, column=1, sticky="w", padx=4
        )
        if note:
            row += 1
            ttk.Label(frame, text=note, foreground="#555", wraplength=460).grid(
                row=row, column=1, sticky="w", padx=4
            )

        row += 1
        ttk.Label(frame, text="Label:").grid(row=row, column=0, sticky="e", padx=4, pady=3)
        self.label_var = tk.StringVar(value=default_label)
        self.label_entry = ttk.Entry(frame, textvariable=self.label_var, width=50)
        self.label_entry.grid(row=row, column=1, sticky="ew", padx=4)

        row += 1
        ttk.Label(frame, text="Put it in:").grid(row=row, column=0, sticky="e", padx=4, pady=3)
        names = [
            (f"{c.display}   [{c.part_path.rsplit('/', 1)[-1]}]" if show_part else c.display)
            + ("" if c.element == "group" else f"   ({c.element})")
            for c in containers
        ]
        self.container_box = ttk.Combobox(frame, values=names, state="readonly", width=60)
        self.container_box.grid(row=row, column=1, sticky="ew", padx=4)
        if names:
            self.container_box.current(0)
        self.container_box.bind("<<ComboboxSelected>>", lambda _e: self._sync_size())

        row += 1
        ttk.Label(frame, text="Icon (imageMso):").grid(row=row, column=0, sticky="e", padx=4, pady=3)
        self.image_var = tk.StringVar()
        ttk.Combobox(frame, textvariable=self.image_var, values=image_suggestions, width=40).grid(
            row=row, column=1, sticky="w", padx=4
        )

        row += 1
        ttk.Label(frame, text="Size:").grid(row=row, column=0, sticky="e", padx=4, pady=3)
        self.size_var = tk.StringVar(value="large")
        self.size_box = ttk.Combobox(
            frame, textvariable=self.size_var, values=["large", "normal"], state="readonly", width=10
        )
        self.size_box.grid(row=row, column=1, sticky="w", padx=4)

        row += 1
        ttk.Label(frame, text="Screentip:").grid(row=row, column=0, sticky="e", padx=4, pady=3)
        self.screentip_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.screentip_var).grid(row=row, column=1, sticky="ew", padx=4)

        row += 1
        ttk.Label(frame, text="Supertip:").grid(row=row, column=0, sticky="e", padx=4, pady=3)
        self.supertip_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.supertip_var).grid(row=row, column=1, sticky="ew", padx=4)

        row += 1
        ttk.Label(frame, text="Control id:").grid(row=row, column=0, sticky="e", padx=4, pady=3)
        self.id_var = tk.StringVar(value=default_id)
        ttk.Entry(frame, textvariable=self.id_var, width=40).grid(row=row, column=1, sticky="w", padx=4)

        row += 1
        buttons = ttk.Frame(frame)
        buttons.grid(row=row, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(buttons, text="Cancel", command=win.destroy).pack(side="right", padx=4)
        ttk.Button(buttons, text="Add Button", command=self.confirm).pack(side="right")
        win.bind("<Return>", lambda _e: self.confirm())
        win.bind("<Escape>", lambda _e: win.destroy())
        self._sync_size()
        self.label_entry.focus_set()
        self.label_entry.select_range(0, "end")

    def _selected_container(self) -> rs.RibbonContainer | None:
        index = self.container_box.current()
        return self.containers[index] if 0 <= index < len(self.containers) else None

    def _sync_size(self) -> None:
        # Only buttons directly in a group accept size; elsewhere it breaks the ribbon.
        container = self._selected_container()
        in_group = container is not None and container.element == "group"
        self.size_box.config(state="readonly" if in_group else "disabled")

    def validate(self) -> str | None:
        if not self.label_var.get().strip():
            return "Enter a label for the button."
        if self._selected_container() is None:
            return "Choose where to put the button."
        control_id = self.id_var.get().strip()
        if not _ID_RE.match(control_id):
            return "Control id must start with a letter or underscore and use letters, digits, _ . -"
        if control_id.casefold() in self.existing_ids:
            return f'A control with id "{control_id}" already exists.'
        return None

    def confirm(self) -> None:
        problem = self.validate()
        if problem:
            messagebox.showerror(APP_NAME, problem, parent=self.win)
            return
        container = self._selected_container()
        assert container is not None
        self.result = ButtonSpec(
            container=container,
            control_id=self.id_var.get().strip(),
            label=self.label_var.get().strip(),
            image_mso=self.image_var.get().strip(),
            size=self.size_var.get() if container.element == "group" else "",
            screentip=self.screentip_var.get().strip(),
            supertip=self.supertip_var.get().strip(),
        )
        self.win.destroy()

    def wait(self) -> ButtonSpec | None:
        self.win.grab_set()
        self.root.wait_window(self.win)
        return self.result


class RibbonButtonsDialog:
    """Non-modal list of every ribbon callback with its resolved macro."""

    def __init__(self, root, window, initial_filter: str = "") -> None:
        self.window = window
        self.entries: list[rs.RibbonEntry] = []
        self.visible: list[rs.RibbonEntry] = []
        win = self.win = tk.Toplevel(root)
        win.title("Ribbon Buttons")
        win.geometry("900x440")
        top = ttk.Frame(win, padding=(8, 8, 8, 4))
        top.pack(fill="x")
        ttk.Label(top, text="Filter:").pack(side="left")
        self.filter_var = tk.StringVar(value=initial_filter)
        entry = ttk.Entry(top, textvariable=self.filter_var, width=40)
        entry.pack(side="left", padx=4)
        self.filter_var.trace_add("write", lambda *_a: self._populate())
        self.problems_only = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            top, text="Problems only", variable=self.problems_only, command=self._populate
        ).pack(side="left", padx=8)
        self.summary = ttk.Label(top, text="")
        self.summary.pack(side="right")

        columns = ("label", "callback", "module", "status")
        body = ttk.Frame(win, padding=(8, 0))
        body.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(body, columns=columns, show="headings", selectmode="browse")
        for col, title, width in (
            ("label", "Button", 260),
            ("callback", "Macro", 280),
            ("module", "Module", 180),
            ("status", "Status", 110),
        ):
            self.tree.heading(col, text=title)
            self.tree.column(col, width=width, anchor="w")
        sb = ttk.Scrollbar(body, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure("problem", foreground="#b00020")
        self.tree.bind("<Double-Button-1>", lambda _e: self.goto_macro())
        self.tree.bind("<Return>", lambda _e: self.goto_macro())
        self.tree.bind("<<TreeviewSelect>>", lambda _e: self._sync_buttons())

        buttons = ttk.Frame(win, padding=8)
        buttons.pack(fill="x")
        self.macro_btn = ttk.Button(buttons, text="Go to Macro", command=self.goto_macro)
        self.macro_btn.pack(side="left", padx=2)
        self.xml_btn = ttk.Button(buttons, text="Go to XML", command=self.goto_xml)
        self.xml_btn.pack(side="left", padx=2)
        self.create_btn = ttk.Button(buttons, text="Create Callback…", command=self.create_callback)
        self.create_btn.pack(side="left", padx=2)
        ttk.Button(buttons, text="Refresh", command=self.refresh).pack(side="right", padx=2)
        win.bind("<Escape>", lambda _e: win.destroy())
        self.refresh()
        entry.focus_set()

    def refresh(self) -> None:
        self.window.flush_editors()
        draft = self.window.draft
        self.entries = rs.ribbon_entries(draft) if draft is not None else []
        self._populate()

    def _populate(self) -> None:
        needle = self.filter_var.get().strip().casefold()
        self.tree.delete(*self.tree.get_children())
        self.visible = []
        for entry in self.entries:
            cb = entry.callback
            module = ", ".join(sorted({m.module_name for m in entry.resolution.matches}))
            label = cb.label or cb.control_id or cb.element
            if cb.attribute != "onAction":
                label = f"{label}  ({cb.attribute})"
            haystack = f"{label} {cb.name} {module} {cb.control_id}".casefold()
            if needle and needle not in haystack:
                continue
            problem = entry.resolution.status != rs.STATUS_OK
            if self.problems_only.get() and not problem:
                continue
            self.tree.insert(
                "",
                "end",
                iid=str(len(self.visible)),
                values=(label, cb.name, module or "—", rs.STATUS_LABELS[entry.resolution.status]),
                tags=("problem",) if problem else (),
            )
            self.visible.append(entry)
        bad = sum(1 for e in self.entries if e.resolution.status != rs.STATUS_OK)
        self.summary.config(
            text=f"{len(self.entries)} callbacks" + (f", {bad} with problems" if bad else "")
        )
        if self.visible:
            self.tree.selection_set("0")
        self._sync_buttons()

    def _selected(self) -> rs.RibbonEntry | None:
        sel = self.tree.selection()
        return self.visible[int(sel[0])] if sel else None

    def _sync_buttons(self) -> None:
        entry = self._selected()
        has = entry is not None
        self.xml_btn.config(state="normal" if has else "disabled")
        self.macro_btn.config(state="normal" if has and entry.resolution.matches else "disabled")
        missing = has and entry.resolution.status == rs.STATUS_MISSING
        self.create_btn.config(state="normal" if missing else "disabled")

    def goto_macro(self) -> None:
        entry = self._selected()
        if entry is not None:
            self.window.goto_callback(entry.callback)

    def goto_xml(self) -> None:
        entry = self._selected()
        if entry is not None:
            self.window.goto_callback_xml(entry.callback)

    def create_callback(self) -> None:
        entry = self._selected()
        if entry is not None and self.window.create_callback(entry.callback):
            self.refresh()
