"""Ribbon structure panel beside the XML editor: outline, callback check, quick actions.

It reads the editor's current text (lexically, so half-typed XML still shows
what it can) and edits nothing itself: every action goes through MainWindow,
which owns the draft and its undo history.
"""

from __future__ import annotations

from tkinter import ttk
from typing import Any

from vba_addin_editor.services import ribbon_service as rs
from vba_addin_editor.ui import theme
from vba_addin_editor.ui.widgets import Tooltip, icon

# Wrappers that only hold the controls below them: left out of the outline.
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
_ICONS = {"tab": "ribbon", "group": "form", "menu": "menu", "dynamicMenu": "menu", "splitButton": "menu"}
MAX_LISTED_PROBLEMS = 4


class RibbonPanel(ttk.Frame):
    def __init__(self, master, app: Any) -> None:
        super().__init__(master, width=320)
        self.app = app
        self.part_path: str | None = None
        self._rows: dict[str, rs.RibbonElement] = {}
        self._by_index: dict[int, str] = {}

        head = ttk.Frame(self, padding=(14, 6, 6, 6))
        head.pack(fill="x", side="top")
        ttk.Label(head, text="RIBBON STRUCTURE", style="Caption.TLabel").pack(side="left")
        for glyph, tip, command in (
            ("plus", "Add ribbon menu", app.add_ribbon_menu),
            ("down", "Move control down  (Alt+Down)", lambda: app.move_ribbon_control(1)),
            ("up", "Move control up  (Alt+Up)", lambda: app.move_ribbon_control(-1)),
        ):
            button = ttk.Button(head, text=icon(glyph), style="Icon.TButton", command=command, takefocus=False)
            button.pack(side="right")
            Tooltip(button, tip)
        ttk.Frame(self, style="Line.TFrame", height=1).pack(fill="x", side="top")

        foot = ttk.Frame(self, padding=(10, 4, 10, 12))
        foot.pack(fill="x", side="bottom")
        for col, (glyph, label, command) in enumerate(
            (
                ("button", "Button", app.add_ribbon_button_from_panel),
                ("menu", "Menu", app.add_ribbon_menu),
                ("image", "Icon", app.set_ribbon_icon_from_file),
            )
        ):
            ttk.Button(
                foot, text=label, style="Panel.TButton", command=command, takefocus=False, padding=(4, 3)
            ).grid(row=0, column=col, sticky="ew", padx=(0 if col == 0 else 6, 0))
            foot.columnconfigure(col, weight=1, uniform="actions")

        self.check = ttk.Frame(self, padding=(12, 8, 12, 8))
        self.check.pack(fill="x", side="bottom", padx=10, pady=(0, 8))

        body = ttk.Frame(self, padding=(6, 6, 0, 0))
        body.pack(fill="both", expand=True, side="top")
        self.tree = ttk.Treeview(body, show="tree", columns=("kind",), selectmode="browse")
        self.tree.column("#0", width=200, stretch=True)
        self.tree.column("kind", width=84, stretch=False, anchor="e")
        sb = ttk.Scrollbar(body, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=theme.autohide(sb, side="right", fill="y"))
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self._selecting = False
        self.apply_palette()

    def apply_palette(self) -> None:
        colors = theme.palette()
        self.tree.tag_configure("container", foreground=colors.muted)
        self.tree.tag_configure("control", foreground=colors.text)
        self.tree.tag_configure("problem", foreground=colors.warn)

    # -- contents -------------------------------------------------------------

    def refresh(self, part_path: str | None, text: str) -> None:
        """Rebuild from ``text`` (the editor's current XML for ``part_path``)."""
        self.part_path = part_path
        tree = self.tree
        tree.delete(*tree.get_children())
        self._rows.clear()
        self._by_index.clear()
        if part_path is None:
            self._show_check([])
            return
        elements = rs.element_tree(text)
        problems = self._problem_entries(part_path)
        problem_ids = {e.callback.control_id for e in problems}
        shown: dict[int, str] = {}
        for element in elements:
            local = element.name.rsplit(":", 1)[-1]
            parent = element.parent
            while parent >= 0 and parent not in shown:
                parent = elements[parent].parent
            if local in _WRAPPERS:
                continue
            label = rs.unescape(
                element.attr("label") or element.attr("id") or element.attr("idMso") or element.attr("idQ") or ""
            )
            kind = "large" if element.attr("size") == "large" else local
            container = local in {"tab", "group", "menu", "box", "buttonGroup", "splitButton", "dynamicMenu"}
            tags = ["container" if container else "control"]
            if element.attr("id") in problem_ids:
                tags.append("problem")
            glyph = icon(_ICONS.get(local, "button"))
            iid = tree.insert(
                shown.get(parent, ""),
                "end",
                text=f"{glyph}  {label or local}",
                values=(kind,),
                open=True,
                tags=tuple(tags),
            )
            shown[element.index] = iid
            self._rows[iid] = element
            self._by_index[element.index] = iid
        self._show_check(problems, total=len(rs.find_callbacks(text, part_path)))

    def _problem_entries(self, part_path: str) -> list:
        draft = self.app.draft
        if draft is None:
            return []
        return [
            e
            for e in rs.ribbon_entries(draft)
            if e.callback.part_path == part_path and e.resolution.status != rs.STATUS_OK
        ]

    def _show_check(self, problems: list, total: int = 0) -> None:
        for child in self.check.winfo_children():
            child.destroy()
        if self.part_path is None:
            return
        linked = max(total - len(problems), 0)
        row = ttk.Frame(self.check)
        row.pack(fill="x", anchor="w")
        ttk.Label(row, text=icon("check"), style="OkIcon.TLabel").pack(side="left")
        ttk.Label(
            row,
            text=f"{linked} callback{'s' if linked != 1 else ''} linked to macros",
            style="Ok.TLabel",
        ).pack(side="left", padx=(6, 0))
        for entry in problems[:MAX_LISTED_PROBLEMS]:
            cb = entry.callback
            row = ttk.Frame(self.check)
            row.pack(fill="x", anchor="w", pady=(5, 0))
            ttk.Label(row, text=icon("alert"), style="WarnIcon.TLabel").pack(side="left")
            what = "no macro" if entry.resolution.status == rs.STATUS_MISSING else "check it"
            name = cb.name if len(cb.name) <= 16 else cb.name[:15] + "…"
            ttk.Label(row, text=f"{name}: {what}", style="Warn.TLabel").pack(side="left", padx=(6, 0))
            if entry.resolution.status == rs.STATUS_MISSING:
                link = ttk.Label(row, text="Create stub", style="Link.TLabel", cursor="hand2")
                link.bind("<Button-1>", lambda _e, c=cb: self.app.create_callback(c))
            else:
                link = ttk.Label(row, text="Show", style="Link.TLabel", cursor="hand2")
                link.bind("<Button-1>", lambda _e, c=cb: self.app.goto_callback_xml(c))
            link.pack(side="right")
        if len(problems) > MAX_LISTED_PROBLEMS:
            more = ttk.Label(
                self.check,
                text=f"and {len(problems) - MAX_LISTED_PROBLEMS} more…",
                style="Link.TLabel",
                cursor="hand2",
            )
            more.pack(anchor="w", pady=(5, 0))
            more.bind("<Button-1>", lambda _e: self.app.check_ribbon_callbacks())

    # -- selection ------------------------------------------------------------

    def _on_select(self, _event=None) -> None:
        if self._selecting or self.part_path is None:
            return
        selection = self.tree.selection()
        element = self._rows.get(selection[0]) if selection else None
        if element is not None:
            self.app._select_xml_range(self.part_path, element.start, element.open_end - element.start)

    def follow_cursor(self, text: str, offset: int) -> None:
        """Highlight the control the XML cursor is in."""
        element = rs.element_at(text, offset)
        elements = rs.element_tree(text) if element is not None else []
        while element is not None and element.index not in self._by_index:
            element = elements[element.parent] if 0 <= element.parent < len(elements) else None
        iid = self._by_index.get(element.index) if element is not None else None
        self._selecting = True
        try:
            if iid is None:
                self.tree.selection_remove(*self.tree.selection())
            elif self.tree.exists(iid) and self.tree.selection() != (iid,):
                self.tree.selection_set(iid)
                self.tree.see(iid)
        finally:
            self.after_idle(self._end_selecting)

    def _end_selecting(self) -> None:
        self._selecting = False
