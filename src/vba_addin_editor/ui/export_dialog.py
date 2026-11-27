"""Copy ribbon controls, macros, modules and UserForms from one add-in into another.

One dialog for both directions (Export to… from the add-in you are in, and
Import from… into it). It lists what the source has: its ribbon (tabs, groups
and controls, nested), its macros by module, its modules and its UserForms.
Tick what to copy; each ticked ribbon control gets a place in the target. The
plan is recomputed on every change and shown, renames included; the dialog
never edits anything. ``wait()`` returns the confirmed plan, or None.

Rows are keyed: ``c:<part>@<start>`` a control, ``m:<module>.<proc>`` a macro,
``mod:<module>`` a module, ``f:<form>`` a UserForm. Tests drive it with
``check(key)``, ``set_where(key, index)``, ``refresh()`` and ``confirm()``.
"""

from __future__ import annotations

import tkinter as tk
from collections.abc import Callable
from tkinter import ttk

from vba_addin_editor.services import port_service as ps
from vba_addin_editor.ui import theme

CHECKED, UNCHECKED, INCLUDED = "☑", "☐", "▣"


def default_target(control: ps.PortableControl, targets: list[ps.PortTarget]) -> int:
    """Where a control goes unless the user says: controls share one new group on the first tab."""
    if not targets:
        return -1
    if control.element not in ("tab", "group"):
        for index, target in enumerate(targets):
            if target.new_group:
                return index
    return 0


class ExportDialog:
    def __init__(
        self,
        root,
        *,
        source: ps.DocumentDraft,
        target_name: str,
        targets_for: Callable[[ps.PortableControl], list[ps.PortTarget]],
        plan: Callable[[list[ps.ExportItem], str, str], ps.PortPlan],
        suffix: str,
        show_part: bool = False,
        preselect: tuple[str, ...] | list[str] = (),
        title: str = "Copy to Another Add-in",
    ) -> None:
        self.root = root
        self.source = source
        self.targets_for = targets_for
        self.planner = plan
        self.show_part = show_part
        self.controls = {c.key: c for c in ps.portable_controls(source)}
        self.macros = {f"m:{m.module}.{m.name}": m for m in ps.portable_macros(source)}
        self.modules = {
            f"mod:{m.name}": m for m in ps.portable_modules(source) if m.kind != "form"
        }
        self.forms = {f"f:{f.name}": f for f in ps.portable_forms(source)}
        self.checked: list[str] = []  # in the order ticked
        self.where: dict[str, int] = {}  # control key -> index into its targets
        self._targets: dict[str, list[ps.PortTarget]] = {}
        self.plan: ps.PortPlan | None = None
        self.result: ps.PortPlan | None = None
        self._pending: str | None = None
        self._focused_control: str | None = None

        win = self.win = tk.Toplevel(root)
        win.title(title)
        win.transient(root)
        win.geometry("980x700")
        buttons = ttk.Frame(win, padding=(10, 0, 10, 10))
        buttons.pack(fill="x", side="bottom")  # before the form: a short window keeps these
        ttk.Button(buttons, text="Cancel", command=win.destroy).pack(side="right", padx=4)
        self.import_btn = ttk.Button(buttons, text="Copy", style="Accent.TButton", command=self.confirm)
        self.import_btn.pack(side="right")
        self.count_label = ttk.Label(buttons, text="", style="Muted.TLabel")
        self.count_label.pack(side="left")

        head = ttk.Frame(win, padding=(10, 10, 10, 4))
        head.pack(fill="x", side="top")
        ttk.Label(head, text=source.baseline.path.name, font=theme.BOLD_FONT).pack(side="left")
        ttk.Label(head, text="  →  ").pack(side="left")
        ttk.Label(head, text=target_name, font=theme.BOLD_FONT).pack(side="left")

        panes = ttk.Panedwindow(win, orient="horizontal")
        panes.pack(fill="both", expand=True, padx=10)
        left = ttk.Frame(panes)
        panes.add(left, weight=1)
        ttk.Label(left, text="Tick what to copy (what it needs comes along):").pack(anchor="w", pady=(0, 4))
        tree_box = ttk.Frame(left)
        tree_box.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tree_box, show="tree", selectmode="browse")
        sb = ttk.Scrollbar(tree_box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<Button-1>", self._on_click)
        self.tree.bind("<space>", lambda _e: self._toggle_focused())
        self.tree.bind("<<TreeviewSelect>>", lambda _e: self._show_place())
        self._fill_tree()

        right = ttk.Frame(panes, padding=(10, 0, 0, 0))
        panes.add(right, weight=1)
        right.columnconfigure(1, weight=1)
        self.place_caption = ttk.Label(right, text="Put it in:")
        self.place_caption.grid(row=0, column=0, sticky="e", padx=4, pady=3)
        self.target_box = ttk.Combobox(right, state="readonly", width=52)
        self.target_box.grid(row=0, column=1, sticky="ew", padx=4)
        self.target_box.bind("<<ComboboxSelected>>", lambda _e: self._place_changed())
        ttk.Label(right, text="New group label:").grid(row=1, column=0, sticky="e", padx=4, pady=3)
        self.group_var = tk.StringVar(master=win)
        ttk.Entry(right, textvariable=self.group_var, width=30).grid(row=1, column=1, sticky="w", padx=4)
        self.group_var.trace_add("write", lambda *_a: self._schedule())
        ttk.Label(right, text="Name suffix:").grid(row=2, column=0, sticky="e", padx=4, pady=3)
        self.suffix_var = tk.StringVar(master=win, value=suffix)
        ttk.Entry(right, textvariable=self.suffix_var, width=24).grid(row=2, column=1, sticky="w", padx=4)
        self.suffix_var.trace_add("write", lambda *_a: self._schedule())
        ttk.Label(
            right,
            text="Macros always get the suffix, so both add-ins can be loaded together; other "
            "names only when this add-in already uses them.",
            style="Muted.TLabel",
            wraplength=420,
            justify="left",
        ).grid(row=3, column=1, sticky="w", padx=4, pady=(0, 6))
        preview = ttk.Frame(right)
        preview.grid(row=4, column=0, columnspan=2, sticky="nsew", pady=(3, 0))
        right.rowconfigure(4, weight=1)
        self.preview = tk.Text(preview, wrap="word", height=18, state="disabled")
        psb = ttk.Scrollbar(preview, orient="vertical", command=self.preview.yview)
        self.preview.configure(yscrollcommand=psb.set)
        self.preview.pack(side="left", fill="both", expand=True)
        psb.pack(side="right", fill="y")

        win.bind("<Escape>", lambda _e: win.destroy())
        for key in preselect:
            if self.tree.exists(key):
                self.check(key, True, refresh=False)
        first = next((k for k in preselect if self.tree.exists(k)), None)
        if first is not None:
            self.tree.see(first)
            self.tree.selection_set(first)
            self.tree.focus(first)
        self._show_place()
        self.refresh()
        self.tree.focus_set()

    # -- the list ------------------------------------------------------------------------

    def _fill_tree(self) -> None:
        tree = self.tree
        if self.controls:
            tree.insert("", "end", iid="sec:ribbon", text="Ribbon", open=True)
            # Nest each control under the nearest chosen-able control around it.
            stack: list[ps.PortableControl] = []
            for control in sorted(self.controls.values(), key=lambda c: (c.part_path, c.start)):
                while stack and not (
                    stack[-1].part_path == control.part_path and control.end <= stack[-1].end
                ):
                    stack.pop()
                parent = stack[-1].key if stack else "sec:ribbon"
                label = control.label or control.control_id
                text = f"{label}   ({control.element})"
                if control.callbacks:
                    text += "   → " + ", ".join(dict.fromkeys(control.callbacks))
                tree.insert(f"c:{parent}" if stack else parent, "end", iid=f"c:{control.key}", text=text, open=True)
                stack.append(control)
        if self.macros:
            tree.insert("", "end", iid="sec:macros", text="Macros", open=True)
            for key, macro in self.macros.items():
                module_iid = f"sec:macros/{macro.module}"
                if not tree.exists(module_iid):
                    tree.insert("sec:macros", "end", iid=module_iid, text=macro.module, open=False)
                tree.insert(module_iid, "end", iid=key, text=macro.name + ("   (ribbon)" if macro.ribbon else ""))
        if self.modules:
            tree.insert("", "end", iid="sec:modules", text="Whole modules", open=False)
            for key, module in self.modules.items():
                tree.insert("sec:modules", "end", iid=key, text=module.describe())
        if self.forms:
            tree.insert("", "end", iid="sec:forms", text="UserForms", open=bool(self.forms) and not self.controls)
            for key, form in self.forms.items():
                tree.insert("sec:forms", "end", iid=key, text=form.describe())
        self._labels = {iid: tree.item(iid, "text") for iid in self._all_items()}
        self._paint()

    def _all_items(self, parent: str = "") -> list[str]:
        out: list[str] = []
        for iid in self.tree.get_children(parent):
            out.append(iid)
            out += self._all_items(iid)
        return out

    @staticmethod
    def _checkable(iid: str) -> bool:
        return not iid.startswith("sec:")

    def _included(self, iid: str) -> bool:
        """A control inside a ticked control: copied with it."""
        parent = self.tree.parent(iid)
        while parent and parent.startswith("c:"):
            if parent in self.checked:
                return True
            parent = self.tree.parent(parent)
        return False

    def _paint(self) -> None:
        for iid in self._all_items():
            if not self._checkable(iid):
                continue
            mark = INCLUDED if self._included(iid) else CHECKED if iid in self.checked else UNCHECKED
            self.tree.item(iid, text=f"{mark}  {self._labels[iid]}")

    def check(self, iid: str, on: bool = True, *, refresh: bool = True) -> None:
        if not self._checkable(iid):
            return
        if on and iid not in self.checked:
            self.checked.append(iid)
            if iid.startswith("c:") and iid not in self.where:
                targets = self._targets_of(iid)
                self.where[iid] = default_target(self.controls[iid[2:]], targets)
        elif not on and iid in self.checked:
            self.checked.remove(iid)
        self._paint()
        self._show_place()
        if refresh:
            self.refresh()

    def _on_click(self, event) -> str | None:
        iid = self.tree.identify_row(event.y)
        if not iid or not self._checkable(iid):
            return None
        if self.tree.identify_element(event.x, event.y) == "Treeitem.indicator":
            return None
        self.tree.selection_set(iid)
        self.tree.focus(iid)
        self.check(iid, iid not in self.checked)
        return "break"

    def _toggle_focused(self) -> str:
        iid = self.tree.focus()
        if iid:
            self.check(iid, iid not in self.checked)
        return "break"

    # -- where ribbon controls go ------------------------------------------------------------

    def _targets_of(self, iid: str) -> list[ps.PortTarget]:
        if iid not in self._targets:
            self._targets[iid] = self.targets_for(self.controls[iid[2:]])
        return self._targets[iid]

    def set_where(self, iid: str, index: int) -> None:
        self.where[iid] = index
        self._show_place()
        self.refresh()

    def _show_place(self) -> None:
        """The place box follows the selected ribbon control (or the first ticked one)."""
        selected = self.tree.selection()
        iid = selected[0] if selected else ""
        if not (iid.startswith("c:") and iid in self.checked):
            iid = next((k for k in self.checked if k.startswith("c:")), "")
        self._focused_control = iid or None
        if not iid:
            self.place_caption.config(text="Put it in:")
            self.target_box.config(values=[], state="disabled")
            self.target_box.set("")
            return
        control = self.controls[iid[2:]]
        targets = self._targets_of(iid)
        self.place_caption.config(text=f"Put “{control.label or control.control_id}” in:")
        self.target_box.config(values=[t.describe(self.show_part) for t in targets], state="readonly")
        index = self.where.get(iid, -1)
        if 0 <= index < len(targets):
            self.target_box.current(index)
        else:
            self.target_box.set("(this add-in's ribbon has no place for it)")

    def _place_changed(self) -> None:
        if self._focused_control is not None:
            self.where[self._focused_control] = self.target_box.current()
            self.refresh()

    # -- the plan ---------------------------------------------------------------------------

    def items(self) -> list[ps.ExportItem] | str:
        """The ticked rows as export items, or why they cannot be."""
        items: list[ps.ExportItem] = []
        for iid in self.checked:
            if iid.startswith("c:"):
                if self._included(iid):
                    continue
                control = self.controls[iid[2:]]
                targets = self._targets_of(iid)
                index = self.where.get(iid, -1)
                if not 0 <= index < len(targets):
                    return (
                        f"This add-in's ribbon has no place a {control.element} can go. Create a "
                        "ribbon in it first (Ribbon → Create Ribbon), or untick "
                        f"“{control.label or control.control_id}”."
                    )
                items.append(ps.ExportItem.of_control(control, targets[index]))
            elif iid.startswith("m:"):
                macro = self.macros[iid]
                items.append(ps.ExportItem.of_macro(macro.module, macro.name))
            elif iid.startswith("mod:"):
                items.append(ps.ExportItem.of_module(self.modules[iid].name))
            elif iid.startswith("f:"):
                items.append(ps.ExportItem.of_form(self.forms[iid].name))
        return items

    def _schedule(self) -> None:
        if self._pending is not None:
            self.win.after_cancel(self._pending)
        self._pending = self.win.after(300, self.refresh)

    def refresh(self) -> None:
        self._pending = None
        items = self.items()
        if isinstance(items, str):
            self.plan, text = None, items
        elif not items:
            self.plan, text = None, "Tick a ribbon control, a macro, a module or a UserForm to copy."
        else:
            self.plan = self.planner(items, self.suffix_var.get().strip(), self.group_var.get())
            text = self.plan.describe()
        self.preview.config(state="normal")
        self.preview.delete("1.0", "end")
        self.preview.insert("1.0", text)
        self.preview.config(state="disabled")
        n = len(self.checked)
        self.count_label.config(text=f"{n} chosen" if n else "")
        ok = self.plan is not None and self.plan.ok
        self.import_btn.config(state="normal" if ok else "disabled")

    def confirm(self) -> None:
        if self._pending is not None:
            self.win.after_cancel(self._pending)
            self.refresh()
        if self.plan is None or not self.plan.ok:
            return
        self.result = self.plan
        self.win.destroy()

    def wait(self) -> ps.PortPlan | None:
        self.win.grab_set()
        self.root.wait_window(self.win)
        return self.result
