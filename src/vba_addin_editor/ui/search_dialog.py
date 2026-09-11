"""Non-modal project search and replace (IMP-08)."""

from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk

from vba_addin_editor.domain.history import HistoryCommand
from vba_addin_editor.services.search_service import (
    PAGE_SIZE,
    SearchHit,
    SearchResults,
    SearchService,
)
from vba_addin_editor.version import APP_NAME


class ProjectSearchDialog:
    def __init__(self, parent, window) -> None:
        self.parent = parent
        self.window = window
        self.service: SearchService = window.search_service
        self.results: SearchResults | None = None
        self.page = 0
        self.stale = False
        self.win = tk.Toplevel(parent)
        self.win.title("Find in Project")
        self.win.geometry("860x620")
        self.win.transient(parent)

        query_row = ttk.Frame(self.win)
        query_row.pack(fill="x", padx=8, pady=4)
        ttk.Label(query_row, text="Find:").pack(side="left")
        self.query = ttk.Entry(query_row)
        self.query.pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(query_row, text="Find", command=self.search).pack(side="left")

        opts = ttk.Frame(self.win)
        opts.pack(fill="x", padx=8)
        self.match_case = tk.BooleanVar(value=False)
        self.whole_word = tk.BooleanVar(value=False)
        self.include_xml = tk.BooleanVar(value=False)
        ttk.Checkbutton(opts, text="Match Case", variable=self.match_case).pack(side="left")
        ttk.Checkbutton(opts, text="Whole Word", variable=self.whole_word).pack(side="left")
        ttk.Checkbutton(opts, text="Include XML", variable=self.include_xml).pack(side="left")

        self.status = ttk.Label(self.win, text="")
        self.status.pack(fill="x", padx=8)

        panes = ttk.Panedwindow(self.win, orient="horizontal")
        panes.pack(fill="both", expand=True, padx=8, pady=4)

        left = ttk.Frame(panes)
        self.listbox = tk.Listbox(left)
        self.listbox.pack(fill="both", expand=True)
        self.listbox.bind("<<ListboxSelect>>", self._on_select)
        self.listbox.bind("<Double-Button-1>", lambda _e: self.goto_selected())
        page_row = ttk.Frame(left)
        page_row.pack(fill="x")
        ttk.Button(page_row, text="Prev", command=self.prev_page).pack(side="left")
        ttk.Button(page_row, text="Next", command=self.next_page).pack(side="left")
        panes.add(left, weight=3)

        right = ttk.Frame(panes)
        ttk.Label(right, text="Procedure outline").pack(anchor="w")
        self.outline = tk.Listbox(right, height=12)
        self.outline.pack(fill="both", expand=True)
        self.outline.bind("<Double-Button-1>", lambda _e: self.goto_outline())
        panes.add(right, weight=2)

        replace_row = ttk.Frame(self.win)
        replace_row.pack(fill="x", padx=8, pady=4)
        ttk.Label(replace_row, text="Replace:").pack(side="left")
        self.replacement = ttk.Entry(replace_row)
        self.replacement.pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(replace_row, text="Preview Selected", command=self.preview_replace).pack(side="left")
        ttk.Button(replace_row, text="Replace Selected", command=self.apply_replace).pack(side="left")

        self.preview = tk.Text(self.win, height=8, wrap="none")
        self.preview.pack(fill="both", expand=False, padx=8, pady=4)
        self.preview.configure(state="disabled")

        self.selected: set[tuple[str, int]] = set()
        self._page_hits: tuple[SearchHit, ...] = ()
        self._outline_items = []
        self.query.bind("<Return>", lambda _e: self.search())

    def mark_stale(self) -> None:
        self.stale = True
        self.status.config(text="Results are stale because the draft changed. Search again.")

    def search(self) -> None:
        window = self.window
        if window.draft is None:
            return
        window._flush_all_editors()
        query = self.query.get()
        revision = window.session.revision if window.session is not None else 0
        self.results = self.service.search(
            window.draft,
            query,
            match_case=self.match_case.get(),
            whole_word=self.whole_word.get(),
            include_xml=self.include_xml.get(),
            revision=revision,
        )
        self.page = 0
        self.stale = False
        self.selected.clear()
        self._render_page()
        self._render_outline()

    def _render_page(self) -> None:
        self.listbox.delete(0, "end")
        if self.results is None:
            return
        self._page_hits = self.service.page(self.results, self.page)
        start = self.page * PAGE_SIZE
        for hit in self._page_hits:
            snippet = hit.snippet.replace("\n", " ")
            self.listbox.insert(
                "end",
                f"{hit.path}:{hit.line}:{hit.column}  {snippet}",
            )
        total_pages = max((self.results.total + PAGE_SIZE - 1) // PAGE_SIZE, 1)
        extra = "  STALE" if self.stale else ""
        self.status.config(
            text=f"{self.results.total} matches  page {self.page + 1}/{total_pages}  "
            f"showing {start + 1}-{start + len(self._page_hits)}{extra}"
            if self.results.total
            else "No matches."
        )

    def _render_outline(self) -> None:
        self.outline.delete(0, "end")
        self._outline_items = []
        window = self.window
        if window.draft is None or window.current_module_id is None:
            return
        module = window.draft.module_by_id(window.current_module_id)
        if module is None:
            return
        self._outline_items = self.service.outline(module.body, module.id)
        for item in self._outline_items:
            flag = " Static" if item.is_static else ""
            self.outline.insert("end", f"{item.visibility}{flag} {item.kind} {item.name}  (line {item.line})")

    def current_hit(self) -> SearchHit | None:
        sel = self.listbox.curselection()
        if not sel:
            return None
        index = int(sel[0])
        if index >= len(self._page_hits):
            return None
        return self._page_hits[index]

    def _on_select(self, _event=None) -> None:
        hit = self.current_hit()
        if hit is None:
            return
        self.selected.add((hit.target_id, hit.offset))

    def goto_selected(self) -> None:
        if self.stale:
            messagebox.showwarning(APP_NAME, "Search results are stale. Search again.")
            return
        hit = self.current_hit()
        if hit is None:
            return
        self.window.goto_search_hit(hit)

    def goto_outline(self) -> None:
        sel = self.outline.curselection()
        if not sel:
            return
        item = self._outline_items[int(sel[0])]
        self.window.goto_search_hit(
            SearchHit(
                target_id=item.module_id,
                path="",
                kind="vba",
                line=item.line,
                column=1,
                offset=0,
                length=0,
                snippet=item.signature,
            )
        )

    def next_page(self) -> None:
        if self.results is None:
            return
        max_page = max((self.results.total - 1) // PAGE_SIZE, 0)
        if self.page < max_page:
            self.page += 1
            self._render_page()

    def prev_page(self) -> None:
        if self.page > 0:
            self.page -= 1
            self._render_page()

    def preview_replace(self) -> None:
        if self.results is None or self.window.draft is None:
            return
        rows = self.service.preview_replacements(
            self.window.draft, self.results, self.replacement.get(), self.selected
        )
        self.preview.configure(state="normal")
        self.preview.delete("1.0", "end")
        if not rows:
            self.preview.insert("1.0", "No selected occurrences.")
        for hit, before, after in rows:
            self.preview.insert("end", f"{hit.path}:{hit.line}:{hit.column}\n  {before!r} -> {after!r}\n")
        self.preview.configure(state="disabled")

    def apply_replace(self) -> None:
        window = self.window
        if window.session is None or self.results is None:
            return
        if self.stale or window.session.revision != self.results.revision:
            messagebox.showwarning(APP_NAME, "Search results are stale. Search again.")
            return
        result = self.service.apply_replacements(
            window.session,
            self.results,
            self.replacement.get(),
            self.selected,
            package_adapter=window.doc_service.package_adapter,
        )
        if not result.ok:
            messagebox.showerror(APP_NAME, "\n".join(result.problems) or result.reason or "Replace failed.")
            return
        if result.changed == 0 or result.draft is None:
            messagebox.showinfo(APP_NAME, "No replacements made.")
            return
        window.session.draft.modules[:] = result.draft.modules
        window.session.draft.xml_parts[:] = result.draft.xml_parts
        window.draft = window.session.draft
        window.history_service.record_structural(
            window.session,
            HistoryCommand(
                op="replace_all",
                target_id="*",
                before={"modules": result.before_modules, "xml": result.before_xml},
                after={"modules": result.after_modules, "xml": result.after_xml},
            ),
        )
        window._reload_editors_from_draft()
        window._checkpoint_now()
        self.mark_stale()
        messagebox.showinfo(APP_NAME, f"Replaced {result.changed} occurrence(s).")
