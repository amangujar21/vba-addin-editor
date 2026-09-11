"""Monospaced code editor with line numbers, find/replace, undo (plan 9.4)."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from vba_addin_editor.ui.text_context_menu import install_text_context_menu


class CodeEditor(ttk.Frame):
    def __init__(self, master, on_change=None) -> None:
        super().__init__(master)
        self.on_change = on_change
        self._find_token = ""
        self.linenumbers = tk.Canvas(self, width=44, highlightthickness=0, bg="#f0f0f0")
        self.text = tk.Text(
            self,
            wrap="none",
            font=("Consolas", 11),
            undo=False,
            tabs=("4c",),
            blockcursor=False,
        )
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.text.yview)
        self.hsb = ttk.Scrollbar(self, orient="horizontal", command=self.text.xview)
        self.text.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)
        self.linenumbers.grid(row=0, column=0, sticky="ns")
        self.text.grid(row=0, column=1, sticky="nsew")
        self.vsb.grid(row=0, column=2, sticky="ns")
        self.hsb.grid(row=1, column=1, sticky="ew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)
        self.text.bind("<<Modified>>", self._on_modified, add=True)
        self.text.bind("<KeyRelease>", lambda _e: self._redraw_line_numbers(), add=True)
        self.text.bind("<ButtonRelease-1>", lambda _e: self._redraw_line_numbers(), add=True)
        self.text.bind("<MouseWheel>", lambda _e: self.after_idle(self._redraw_line_numbers), add=True)
        self.text.event_add("<<Find>>", "<Control-f>")
        self.text.bind("<<Find>>", lambda _e: self.find_dialog())
        self.text.event_add("<<Replace>>", "<Control-h>")
        self.text.bind("<<Replace>>", lambda _e: self.replace_dialog())
        self.text.bind("<Control-z>", self._on_undo, add=True)
        self.text.bind("<Control-y>", self._on_redo, add=True)
        self.on_undo = None
        self.on_redo = None
        install_text_context_menu(self.text)

        self.header = None  # optional metadata label set by subclasses (XmlEditor)

    # -- content ---------------------------------------------------------

    def get_text(self) -> str:
        return self.text.get("1.0", "end-1c")

    def set_text(self, value: str) -> None:
        self.text.delete("1.0", "end")
        self.text.insert("1.0", value)
        self.text.edit_reset()
        self._redraw_line_numbers()

    def goto_line(self, line: int) -> None:
        self.text.mark_set("insert", f"{line}.0")
        self.text.see("insert")

    # -- internals ---------------------------------------------------------

    def _on_modified(self, event) -> None:
        self.text.edit_modified(False)
        self._redraw_line_numbers()
        if self.on_change is not None:
            self.on_change()

    def _on_undo(self, _event=None):
        if self.on_undo is not None:
            self.on_undo()
            return "break"
        return None

    def _on_redo(self, _event=None):
        if self.on_redo is not None:
            self.on_redo()
            return "break"
        return None

    def _redraw_line_numbers(self) -> None:
        self.linenumbers.delete("all")
        first = int(self.text.index("@0,0").split(".")[0])
        last = first + int(self.text.winfo_height() // 17) + 2
        for line in range(first, last + 1):
            dline = self.text.dlineinfo(f"{line}.0")
            if dline is None:
                continue
            y = dline[1]
            self.linenumbers.create_text(38, y, anchor="ne", text=str(line), font=("Consolas", 9))
        self.event_generate("<<CursorMove>>")

    def cursor_position(self) -> str:
        return self.text.index("insert")

    # -- find / replace ----------------------------------------------------

    def find_dialog(self) -> None:
        win = tk.Toplevel(self)
        win.title("Find")
        var = tk.StringVar(value=self._find_token)
        ttk.Entry(win, textvariable=var, width=30).grid(row=0, column=0, padx=6, pady=6)
        ttk.Button(win, text="Find Next", command=lambda: self.find_next(var.get())).grid(row=0, column=1, padx=6)

        def _next() -> None:
            self.find_next(var.get())

        win.bind("<Return>", lambda _e: _next())

    def replace_dialog(self) -> None:
        win = tk.Toplevel(self)
        win.title("Replace")
        find_var = tk.StringVar(value=self._find_token)
        repl_var = tk.StringVar()
        ttk.Label(win, text="Find:").grid(row=0, column=0, sticky="e", padx=4, pady=4)
        ttk.Entry(win, textvariable=find_var, width=30).grid(row=0, column=1, padx=4)
        ttk.Label(win, text="Replace:").grid(row=1, column=0, sticky="e", padx=4, pady=4)
        ttk.Entry(win, textvariable=repl_var, width=30).grid(row=1, column=1, padx=4)
        ttk.Button(
            win,
            text="Replace All",
            command=lambda: self.replace_all(find_var.get(), repl_var.get()),
        ).grid(row=2, column=1, sticky="e", padx=4, pady=6)

    def find_next(self, token: str) -> bool:
        if not token:
            return False
        self._find_token = token
        start = self.text.index("insert+1c")
        pos = self.text.search(token, start, stopindex="end", nocase=True)
        if not pos:
            pos = self.text.search(token, "1.0", stopindex="start", nocase=True)
        if not pos:
            return False
        self.text.mark_set("insert", pos)
        self.text.see(pos)
        line = int(pos.split(".")[0])
        col = int(pos.split(".")[1])
        self.text.tag_remove("sel", "1.0", "end")
        self.text.tag_add("sel", pos, f"{line}.{col + len(token)}")
        return True

    def replace_all(self, token: str, replacement: str) -> None:
        if not token:
            return
        content = self.get_text()
        import re

        updated = re.sub(re.escape(token), replacement.replace("\\", "\\\\"), content, flags=re.IGNORECASE)
        if updated != content:
            cursor = self.text.index("insert")
            self.set_text(updated)
            self.text.mark_set("insert", cursor)
