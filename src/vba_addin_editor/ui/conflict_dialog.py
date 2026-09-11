"""External-change comparison and resolution surface (IMP-03)."""

from __future__ import annotations

import os
import tkinter as tk
from tkinter import messagebox, ttk

from vba_addin_editor.services.conflict_service import ConflictItem, ConflictProposal
from vba_addin_editor.version import APP_NAME


class ConflictDialog:
    def __init__(self, parent, proposal: ConflictProposal) -> None:
        self.proposal = proposal
        self.result: str | None = None
        self.win = tk.Toplevel(parent)
        self.win.title("Compare External Changes")
        self.win.geometry("920x560")
        self.win.transient(parent)
        if not os.environ.get("PYTEST_CURRENT_TEST"):
            self.win.grab_set()

        ttk.Label(
            self.win,
            text="B is the captured baseline, D is the recovered/local draft, E is the current file.",
        ).pack(anchor="w", padx=8, pady=4)

        columns = ("target", "property", "b_to_e", "b_to_d", "choice")
        self.tree = ttk.Treeview(self.win, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("target", text="Target")
        self.tree.heading("property", text="Property")
        self.tree.heading("b_to_e", text="B → E (disk)")
        self.tree.heading("b_to_d", text="B → D (draft)")
        self.tree.heading("choice", text="Resolution")
        self.tree.pack(fill="both", expand=True, padx=8, pady=4)
        self._item_by_iid: dict[str, ConflictItem] = {}
        for index, item in enumerate(proposal.items):
            iid = str(index)
            self._item_by_iid[iid] = item
            self.tree.insert(
                "",
                "end",
                iid=iid,
                values=(
                    item.target_id,
                    f"{item.category}:{item.property_name}",
                    _preview(item.baseline, item.external),
                    _preview(item.baseline, item.draft),
                    item.resolution or ("required" if item.requires_choice else "auto"),
                ),
            )

        btns = ttk.Frame(self.win)
        btns.pack(fill="x", padx=8, pady=4)
        ttk.Button(btns, text="Keep Disk", command=lambda: self._choose("keep_disk")).pack(side="left", padx=2)
        ttk.Button(btns, text="Keep Draft", command=lambda: self._choose("keep_draft")).pack(side="left", padx=2)
        ttk.Button(btns, text="Edit Merged Text…", command=self._edit_merged).pack(side="left", padx=2)

        self.merge = tk.Text(self.win, height=8, wrap="none")
        self.merge.pack(fill="both", expand=False, padx=8, pady=4)

        action = ttk.Frame(self.win)
        action.pack(fill="x", padx=8, pady=8)
        self.apply_btn = ttk.Button(action, text="Apply", command=self._on_apply)
        self.apply_btn.pack(side="right", padx=2)
        ttk.Button(action, text="Cancel", command=self._on_cancel).pack(side="right", padx=2)
        self._refresh_apply()

    def wait(self) -> str | None:
        if os.environ.get("PYTEST_CURRENT_TEST"):
            return self.result
        self.win.wait_window()
        return self.result

    def _selected_item(self) -> ConflictItem | None:
        sel = self.tree.selection()
        if not sel:
            return None
        return self._item_by_iid.get(sel[0])

    def _choose(self, resolution: str) -> None:
        item = self._selected_item()
        if item is None:
            return
        item.resolution = resolution
        if resolution != "merged":
            item.merged_text = None
        self._refresh_row(item)
        self._refresh_apply()

    def _edit_merged(self) -> None:
        item = self._selected_item()
        if item is None:
            return
        if item.property_name not in {"body", "text", "name"}:
            messagebox.showinfo(APP_NAME, "Merged text is only available for module or XML text.")
            return
        text = self.merge.get("1.0", "end-1c")
        if not text:
            text = item.draft or item.external or ""
            self.merge.delete("1.0", "end")
            self.merge.insert("1.0", text)
            return
        item.resolution = "merged"
        item.merged_text = text
        self._refresh_row(item)
        self._refresh_apply()

    def _refresh_row(self, item: ConflictItem) -> None:
        for iid, mapped in self._item_by_iid.items():
            if mapped is item:
                values = list(self.tree.item(iid, "values"))
                values[4] = item.resolution or ("required" if item.requires_choice else "auto")
                self.tree.item(iid, values=values)
                break

    def _refresh_apply(self) -> None:
        ready = all((not item.requires_choice) or item.resolution for item in self.proposal.items)
        if self.proposal.parse_failed:
            ready = False
        self.apply_btn.config(state="normal" if ready else "disabled")

    def _on_apply(self) -> None:
        if not all((not item.requires_choice) or item.resolution for item in self.proposal.items):
            return
        self.result = "apply"
        self.win.destroy()

    def _on_cancel(self) -> None:
        self.result = "cancel"
        self.win.destroy()


def _preview(left, right) -> str:
    if left == right:
        return "(unchanged)"
    text = (right or "∅") if right != left else ""
    shown = str(text).replace("\n", " ")
    return shown[:80]
