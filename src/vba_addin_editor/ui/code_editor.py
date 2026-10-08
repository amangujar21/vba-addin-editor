"""Monospaced code editor with line numbers, find/replace, undo (plan 9.4)."""

from __future__ import annotations

import bisect
import re
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
        # "break": the window binds the same keys for when focus is elsewhere,
        # and Tk's own Text bindings for them move the cursor / delete a character.
        self.text.event_add("<<Find>>", "<Control-f>")
        self.text.bind("<<Find>>", lambda _e: self._break(self.show_find()))
        self.text.event_add("<<Replace>>", "<Control-h>")
        self.text.bind("<<Replace>>", lambda _e: self._break(self.show_find(replace=True)))
        self.text.bind("<F3>", lambda _e: self._break(self.find_next()))
        self.text.bind("<Shift-F3>", lambda _e: self._break(self.find_next(backwards=True)))
        self.text.tag_configure("find_match", background="#ffe58a")
        self.text.tag_raise("sel")
        # Keep the current match visible while the find box has the focus.
        self.text.configure(inactiveselectbackground=self.text.cget("selectbackground"))
        self.text.bind("<Control-z>", self._on_undo, add=True)
        self.text.bind("<Control-y>", self._on_redo, add=True)
        self.on_undo = None
        self.on_redo = None
        self.context_menu = install_text_context_menu(self.text)

        self.header = None  # optional metadata label set by subclasses (XmlEditor)
        self._build_find_bar()

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

    def goto_position(self, line: int, column: int, length: int = 0) -> None:
        start = f"{line}.{max(column - 1, 0)}"
        self.text.mark_set("insert", start)
        self.text.tag_remove("sel", "1.0", "end")
        if length:
            self.text.tag_add("sel", start, f"{start}+{length}c")
        self.text.see("insert")

    # -- internals ---------------------------------------------------------

    def _on_modified(self, event) -> None:
        self.text.edit_modified(False)
        self._redraw_line_numbers()
        if self.find_bar.winfo_manager() and self._highlight_after is None:
            self._highlight_after = self.after_idle(self._highlight_matches)
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

    @staticmethod
    def _break(_result=None) -> str:
        return "break"

    def _build_find_bar(self) -> None:
        """One find/replace bar per editor, shown under the text when asked for."""
        self._highlight_after = None
        self.find_var = tk.StringVar(value=self._find_token)
        self.replace_var = tk.StringVar()
        self.match_case = tk.BooleanVar(value=False)
        self.whole_word = tk.BooleanVar(value=False)
        bar = self.find_bar = ttk.Frame(self, padding=(4, 2))
        ttk.Label(bar, text="Find:").grid(row=0, column=0, sticky="e", padx=(0, 4))
        self.find_entry = ttk.Entry(bar, textvariable=self.find_var, width=32)
        self.find_entry.grid(row=0, column=1, sticky="ew")
        ttk.Button(
            bar, text="Previous", width=9, command=lambda: self.find_next(backwards=True)
        ).grid(row=0, column=2, padx=(4, 0))
        ttk.Button(bar, text="Next", width=6, command=self.find_next).grid(row=0, column=3, padx=2)
        ttk.Checkbutton(
            bar, text="Match case", variable=self.match_case, command=self._highlight_matches
        ).grid(row=0, column=4, padx=4)
        ttk.Checkbutton(
            bar, text="Whole word", variable=self.whole_word, command=self._highlight_matches
        ).grid(row=0, column=5, padx=4)
        self.find_status = ttk.Label(bar, text="", width=16)
        self.find_status.grid(row=0, column=6, sticky="w", padx=4)
        ttk.Button(bar, text="Close", width=6, command=self.hide_find).grid(row=0, column=8, sticky="e")
        self.replace_entry = ttk.Entry(bar, textvariable=self.replace_var, width=32)
        self._replace_widgets = [
            ttk.Label(bar, text="Replace:"),
            self.replace_entry,
            ttk.Button(bar, text="Replace", width=9, command=self.replace_current),
            ttk.Button(bar, text="Replace All", command=self.replace_all),
        ]
        bar.columnconfigure(7, weight=1)
        self.find_var.trace_add("write", lambda *_a: self._highlight_matches())
        self.find_entry.bind("<Return>", lambda _e: self._break(self.find_next()))
        self.find_entry.bind("<Shift-Return>", lambda _e: self._break(self.find_next(backwards=True)))
        self.replace_entry.bind("<Return>", lambda _e: self._break(self.replace_current()))
        for widget in (self.find_entry, self.replace_entry):
            widget.bind("<Escape>", lambda _e: self._break(self.hide_find()))
            widget.bind("<F3>", lambda _e: self._break(self.find_next()))
            widget.bind("<Shift-F3>", lambda _e: self._break(self.find_next(backwards=True)))
        self.text.bind("<Escape>", lambda _e: self.hide_find(), add=True)

    def show_find(self, replace: bool = False) -> None:
        """Show the bar (with the replace row if asked) and focus the find box."""
        try:
            selected = self.text.get("sel.first", "sel.last")
        except tk.TclError:
            selected = ""
        if selected and "\n" not in selected:
            self.find_var.set(selected)
        # Row 3: rows 0-2 hold the text, its scrollbar and the XML part header.
        self.find_bar.grid(row=3, column=0, columnspan=3, sticky="ew")
        label, entry, one, every = self._replace_widgets
        if replace:
            label.grid(row=1, column=0, sticky="e", padx=(0, 4), pady=(2, 0))
            entry.grid(row=1, column=1, sticky="ew", pady=(2, 0))
            one.grid(row=1, column=2, padx=(4, 0), pady=(2, 0))
            every.grid(row=1, column=3, columnspan=2, sticky="w", padx=2, pady=(2, 0))
        else:
            for widget in self._replace_widgets:
                widget.grid_remove()
        self._highlight_matches()
        self.find_entry.focus_set()
        self.find_entry.select_range(0, "end")
        self.find_entry.icursor("end")

    def hide_find(self) -> None:
        if not self.find_bar.winfo_manager():
            return
        self.find_bar.grid_remove()
        self.text.tag_remove("find_match", "1.0", "end")
        self.text.focus_set()

    def _matches(self, token: str, content: str) -> list[tuple[int, int]]:
        if not token:
            return []
        pattern = re.escape(token)
        if self.whole_word.get():
            pattern = rf"(?<!\w){pattern}(?!\w)"
        flags = 0 if self.match_case.get() else re.IGNORECASE
        return [(m.start(), m.end()) for m in re.finditer(pattern, content, flags)]

    @staticmethod
    def _indexer(content: str):
        """offset -> Tk "line.column" index, without Tk counting characters."""
        starts = [0] + [m.end() for m in re.finditer("\n", content)]

        def index(offset: int) -> str:
            line = bisect.bisect_right(starts, offset) - 1
            return f"{line + 1}.{offset - starts[line]}"

        return index

    def _highlight_matches(self) -> list[tuple[int, int]]:
        self._highlight_after = None
        self.text.tag_remove("find_match", "1.0", "end")
        token = self.find_var.get()
        content = self.get_text()
        matches = self._matches(token, content)
        if self.find_bar.winfo_manager():
            index = self._indexer(content)
            for start, end in matches[:2000]:
                self.text.tag_add("find_match", index(start), index(end))
            selection = self._selection_offsets()
            if not token:
                self.find_status.config(text="")
            elif not matches:
                self.find_status.config(text="No matches")
            elif selection in matches:
                self.find_status.config(text=f"{matches.index(selection) + 1} of {len(matches)}")
            else:
                self.find_status.config(
                    text=f"{len(matches)} match" + ("es" if len(matches) > 1 else "")
                )
        return matches

    def _selection_offsets(self) -> tuple[int, int] | None:
        try:
            return (
                len(self.text.get("1.0", "sel.first")),
                len(self.text.get("1.0", "sel.last")),
            )
        except tk.TclError:
            return None

    def find_next(self, token: str | None = None, backwards: bool = False) -> bool:
        """Select the next (or previous) match from the cursor, wrapping around."""
        if token is not None:
            self.find_var.set(token)
        token = self.find_var.get()
        if not token:
            return False
        self._find_token = token
        content = self.get_text()
        matches = self._highlight_matches()
        if not matches:
            return False
        try:
            # From the far edge of the selection, so the current match is skipped.
            cursor = len(self.text.get("1.0", "sel.first" if backwards else "sel.last"))
        except tk.TclError:
            cursor = len(self.text.get("1.0", "insert"))
        if backwards:
            position = bisect.bisect_right([end for _start, end in matches], cursor) - 1
        else:
            position = bisect.bisect_left([start for start, _end in matches], cursor)
        position %= len(matches)
        start, end = matches[position]
        index = self._indexer(content)
        self.text.tag_remove("sel", "1.0", "end")
        self.text.tag_add("sel", index(start), index(end))
        self.text.mark_set("insert", index(end))
        self.text.see(index(start))
        self.after_idle(self._redraw_line_numbers)  # once the view has scrolled
        if self.find_bar.winfo_manager():
            self.find_status.config(text=f"{position + 1} of {len(matches)}")
        return True

    def _editable(self) -> bool:
        return str(self.text.cget("state")) != "disabled"

    def replace_current(self) -> bool:
        """Replace the selected match, then move to the next one."""
        token = self.find_var.get()
        if not token or not self._editable():
            return False
        if self._selection_offsets() not in self._matches(token, self.get_text()):
            return self.find_next()
        self.text.delete("sel.first", "sel.last")
        self.text.insert("insert", self.replace_var.get())
        self.find_next()
        return True

    def replace_all(self, token: str | None = None, replacement: str | None = None) -> int:
        """Replace every match in this editor. Returns how many were replaced."""
        if token is not None:
            self.find_var.set(token)
        if replacement is not None:
            self.replace_var.set(replacement)
        token, replacement = self.find_var.get(), self.replace_var.get()
        if not token or not self._editable():
            return 0
        content = self.get_text()
        matches = self._matches(token, content)
        if not matches:
            self._highlight_matches()
            return 0
        index = self._indexer(content)
        # Back to front, so earlier indices stay valid and the view does not jump.
        for start, end in reversed(matches):
            self.text.delete(index(start), index(end))
            self.text.insert(index(start), replacement)
        self._highlight_matches()
        if self.find_bar.winfo_manager():
            self.find_status.config(text=f"Replaced {len(matches)}")
        return len(matches)
