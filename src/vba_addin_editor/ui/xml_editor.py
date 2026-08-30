"""XML part editor: reuses the generic CodeEditor text widget.

CodeEditor is intentionally generic (line numbers, undo, find/replace); this
subclass adds the XML part path/encoding header. The right-click
Cut/Copy/Paste menu is shared Text behavior installed by text_context_menu.
"""

from __future__ import annotations

from tkinter import ttk

from vba_addin_editor.ui.code_editor import CodeEditor


class XmlEditor(CodeEditor):
    """CodeEditor reuse for OOXML package parts (XML plan section 16)."""

    def __init__(self, master, on_change=None) -> None:
        super().__init__(master, on_change=on_change)
        self.header = ttk.Label(self, text="", padding=(2, 2))
        self.header.grid(row=2, column=0, columnspan=3, sticky="w")
        self._validate_action = None

    def set_part_info(self, part) -> None:
        """Show path/encoding metadata; part may be None to clear the header."""
        if part is None:
            self.header.config(text="")
            return
        if not part.editable:
            self.header.config(
                text=(
                    f"{part.path}  ·  READ-ONLY — This XML part was already malformed "
                    "when the file was opened. It will be preserved exactly; VBA editing "
                    f"and saving remain available.  {part.open_problem or ''}"
                ).rstrip()
            )
            return
        newline = "CRLF" if part.newline == "\r\n" else "LF"
        warning = (
            "  — package-critical XML: invalid relationships/content types can "
            "make PowerPoint repair or reject the file"
            if (part.is_relationships_part or part.is_content_types_part)
            else ""
        )
        self.header.config(
            text=f"{part.path}  ·  {part.encoding.upper()} · {newline}{warning}"
        )

    def set_validate_action(self, action) -> None:
        """Register the Validate XML action for the selected part."""
        self._validate_action = action

    def validate_selected(self) -> None:
        if self._validate_action is not None:
            self._validate_action()
