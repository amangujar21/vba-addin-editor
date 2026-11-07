"""Review & save: every pending change with its diff, the backup choice and Save, in one place."""

from __future__ import annotations

import re
import tkinter as tk
from collections.abc import Callable
from dataclasses import dataclass, field
from tkinter import ttk

from vba_addin_editor.ui import theme
from vba_addin_editor.ui.widgets import icon

SAVE, CANCEL, REVERT = "save", "cancel", "revert"
_BADGE_ORDER = (("add", "A"), ("delete", "D"), ("remove", "D"), ("rename", "R"), ("edit", "M"))
_HUNK = re.compile(r"@@ -(\d+)(?:,\d+)? \+(\d+)(?:,\d+)? @@")


@dataclass(frozen=True)
class ReviewEntry:
    badge: str  # M edited, A new, R renamed, D deleted
    name: str
    category: str  # "VBA" or "XML"
    note: str  # e.g. "clsLog → Logger"
    diff: str  # unified diff, may be empty
    added: int
    removed: int


@dataclass(frozen=True)
class ReviewView:
    file_name: str
    destination: str
    file_type: str
    entries: tuple[ReviewEntry, ...]
    backup_text: str
    confirm_label: str = "Save"
    warnings: tuple[str, ...] = ()
    in_place: bool = True  # False for Save a Copy: no backup toggle, no revert
    extra: tuple[str, ...] = field(default=())  # safety lines shown under the list


def _shorten(text: str, limit: int = 40) -> str:
    if len(text) <= limit:
        return text
    keep = (limit - 1) // 2
    return text[:keep] + "…" + text[-keep:]


def badge_for(operations: tuple[str, ...]) -> str:
    for op, badge in _BADGE_ORDER:
        if op in operations:
            return badge
    return "M"


def diff_stats(diff: str) -> tuple[int, int]:
    added = removed = 0
    for line in diff.splitlines():
        if line.startswith("+") and not line.startswith("+++"):
            added += 1
        elif line.startswith("-") and not line.startswith("---"):
            removed += 1
    return added, removed


def entry_from_item(item) -> ReviewEntry:
    """A ReviewEntry for one review_service.ReviewItem."""
    ops = tuple(item.operations)
    badge = badge_for(ops)
    name = item.new_name or item.old_name or item.target_id
    note = ""
    if "rename" in ops and item.old_name and item.new_name:
        note = f"{item.old_name} → {item.new_name}"
    elif badge == "D":
        note = "deleted"
    added, removed = diff_stats(item.unified_diff or "")
    return ReviewEntry(badge, name, item.category, note, item.unified_diff or "", added, removed)


def ask_review(
    parent: tk.Misc,
    view: ReviewView,
    *,
    backup_var: tk.BooleanVar | None = None,
    on_backup_toggled: Callable[[], object] | None = None,
) -> str:
    """Show the review; SAVE, CANCEL or REVERT. Tests replace this function."""
    dialog = ReviewDialog(parent, view, backup_var=backup_var, on_backup_toggled=on_backup_toggled)
    return dialog.wait()


class ReviewDialog:
    def __init__(
        self,
        parent: tk.Misc,
        view: ReviewView,
        *,
        backup_var: tk.BooleanVar | None = None,
        on_backup_toggled: Callable[[], object] | None = None,
    ) -> None:
        self.view = view
        self.result = CANCEL
        colors = theme.palette()
        win = self.win = tk.Toplevel(parent)
        win.title("Review changes")
        win.transient(parent.winfo_toplevel())
        win.configure(background=colors.bg1)
        win.minsize(640, 400)

        # Header, footer and notes are packed before the expanding body.
        head = ttk.Frame(win, padding=(20, 12, 12, 12))
        head.pack(fill="x", side="top")
        ttk.Label(head, text="Review changes", font=theme.TITLE_FONT).pack(side="left")
        count = len(view.entries)
        ttk.Label(
            head,
            text=f"{count} change{'s' if count != 1 else ''} · {view.file_type} · {view.file_name}",
            style="Faint.TLabel",
        ).pack(side="left", padx=(10, 0))
        ttk.Button(head, text=icon("x"), style="Icon.TButton", command=self._cancel, takefocus=False).pack(
            side="right"
        )
        ttk.Frame(win, style="Line.TFrame", height=1).pack(fill="x", side="top")

        foot = ttk.Frame(win, padding=(20, 12, 14, 12))
        foot.pack(fill="x", side="bottom")
        ttk.Frame(win, style="Line.TFrame", height=1).pack(fill="x", side="bottom")
        if view.warnings:
            notes = ttk.Frame(win, style="Banner.TFrame", padding=(20, 8))
            notes.pack(fill="x", side="bottom")
            for line in view.warnings:
                ttk.Label(notes, text=f"{icon('alert')}  {line}", style="Banner.TLabel", wraplength=820).pack(
                    anchor="w"
                )

        self.save_btn = ttk.Button(foot, text=view.confirm_label, style="Accent.TButton", command=self._save)
        self.save_btn.pack(side="right")
        ttk.Button(foot, text="Cancel", style="Panel.TButton", command=self._cancel).pack(
            side="right", padx=(0, 8)
        )
        if view.in_place:
            ttk.Button(foot, text="Revert all…", style="Panel.TButton", command=self._revert).pack(
                side="right", padx=(0, 8)
            )
        if view.in_place and backup_var is not None:
            self.backup_check = ttk.Checkbutton(
                foot,
                text="Keep a backup of the current file",
                variable=backup_var,
                command=lambda: self._backup_toggled(on_backup_toggled),
            )
            self.backup_check.pack(side="left")
        self.where = ttk.Label(foot, text="", style="Faint.TLabel")
        self.where.pack(side="left", padx=(14, 8), fill="x", expand=True)
        self.where.config(
            text=("Saves in place to " if view.in_place else "Writes a copy to ") + _shorten(view.destination)
        )

        body = ttk.Frame(win)
        body.pack(fill="both", expand=True, side="top")
        left = ttk.Frame(body, padding=(8, 8, 0, 8), width=290)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)
        self.list = ttk.Treeview(left, show="", columns=("badge", "name", "stats"), selectmode="browse")
        self.list.column("badge", width=26, stretch=False, anchor="center")
        self.list.column("name", width=170, stretch=True)
        self.list.column("stats", width=76, stretch=False, anchor="e")
        self.list.pack(fill="both", expand=True)
        for badge, color in (("M", colors.warn), ("A", colors.ok), ("R", colors.info), ("D", colors.danger)):
            self.list.tag_configure(badge, foreground=color)
        if view.extra:
            ttk.Label(
                left, text="\n".join(view.extra), style="Faint.TLabel", justify="left", wraplength=270
            ).pack(fill="x", pady=(8, 0))
        ttk.Frame(body, style="Line.TFrame", width=1).pack(side="left", fill="y")

        right = ttk.Frame(body, style="Editor.TFrame")
        right.pack(side="left", fill="both", expand=True)
        self.diff = tk.Text(
            right,
            wrap="none",
            font=theme.code_font(win),
            background=colors.bg2,
            foreground=colors.text,
            padx=0,
            pady=8,
            cursor="arrow",
            spacing1=1,
            spacing3=1,
        )
        vsb = ttk.Scrollbar(right, orient="vertical", command=self.diff.yview)
        hsb = ttk.Scrollbar(right, orient="horizontal", command=self.diff.xview)
        self.diff.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        hsb.pack(side="bottom", fill="x")
        vsb.pack(side="right", fill="y")
        self.diff.pack(side="left", fill="both", expand=True)
        self.diff.tag_configure("add", background=colors.diff_add_bg)
        self.diff.tag_configure("del", background=colors.diff_del_bg)
        self.diff.tag_configure("addsign", foreground=colors.ok)
        self.diff.tag_configure("delsign", foreground=colors.danger)
        self.diff.tag_configure("num", foreground=colors.gutter)
        self.diff.tag_configure("hunk", foreground=colors.faint, background=colors.bg1, spacing1=6, spacing3=6)
        self.diff.tag_configure("info", foreground=colors.muted, lmargin1=20, lmargin2=20, spacing1=8)

        for index, entry in enumerate(view.entries):
            stats = "" if not (entry.added or entry.removed) else f"+{entry.added} −{entry.removed}"
            if entry.note and not stats:
                stats = entry.note if len(entry.note) <= 10 else ""
            self.list.insert("", "end", iid=str(index), values=(entry.badge, entry.name, stats), tags=(entry.badge,))
        self.list.bind("<<TreeviewSelect>>", lambda _e: self._show_selected())
        if view.entries:
            self.list.selection_set("0")
        self._show_selected()

        win.bind("<Escape>", lambda _e: self._cancel())
        win.bind("<Return>", lambda _e: self._save())
        win.protocol("WM_DELETE_WINDOW", self._cancel)
        self._place(parent)
        self.save_btn.focus_set()

    def _place(self, parent: tk.Misc) -> None:
        top = parent.winfo_toplevel()
        top.update_idletasks()
        width = min(920, max(top.winfo_width() - 60, 640))
        height = min(580, max(top.winfo_height() - 60, 420))
        x = top.winfo_rootx() + max((top.winfo_width() - width) // 2, 0)
        y = top.winfo_rooty() + max((top.winfo_height() - height) // 3, 0)
        self.win.geometry(f"{width}x{height}+{x}+{y}")

    def _backup_toggled(self, callback) -> None:
        if callback is not None:
            callback()

    def _show_selected(self) -> None:
        text = self.diff
        text.configure(state="normal")
        text.delete("1.0", "end")
        selection = self.list.selection()
        if not selection:
            text.insert("end", "No changes.", ("info",))
            text.configure(state="disabled")
            return
        entry = self.view.entries[int(selection[0])]
        if entry.note and entry.badge == "R":
            text.insert("end", f"Renamed: {entry.note}\n", ("info",))
        if not entry.diff.strip():
            if entry.badge != "R":
                text.insert("end", "No line changes to show.\n", ("info",))
            text.configure(state="disabled")
            return
        old = new = 1
        for line in entry.diff.splitlines():
            # review_service joins difflib's header lines without newlines, so
            # "--- a/X+++ b/X@@ -1,4 +1,5 @@" runs straight into the first diff line.
            match = _HUNK.search(line) if line.startswith(("---", "@@")) else None
            if match:
                old, new = max(int(match.group(1)), 1), max(int(match.group(2)), 1)
                text.insert("end", f"  Line {new}\n", ("hunk",))
                line = line[match.end():]
                if not line:
                    continue
            elif line.startswith(("---", "+++")):
                continue
            sign, body = line[:1], line[1:]
            if sign == "+":
                text.insert("end", f"{new:>6}  ", ("num", "add"))
                text.insert("end", "+ ", ("addsign", "add"))
                text.insert("end", body + "\n", ("add",))
                new += 1
            elif sign == "-":
                text.insert("end", f"{old:>6}  ", ("num", "del"))
                text.insert("end", "− ", ("delsign", "del"))
                text.insert("end", body + "\n", ("del",))
                old += 1
            elif sign == "\\":
                continue  # "\ No newline at end of file"
            else:
                text.insert("end", f"{new:>6}  ", ("num",))
                text.insert("end", "  " + body + "\n")
                old += 1
                new += 1
        text.configure(state="disabled")

    def _finish(self, result: str) -> None:
        self.result = result
        try:
            self.win.grab_release()
        except tk.TclError:
            pass
        self.win.destroy()

    def _save(self) -> None:
        self._finish(SAVE)

    def _cancel(self) -> None:
        self._finish(CANCEL)

    def _revert(self) -> None:
        self._finish(REVERT)

    def wait(self) -> str:
        try:
            self.win.grab_set()
        except tk.TclError:
            pass
        self.win.wait_window()
        return self.result
