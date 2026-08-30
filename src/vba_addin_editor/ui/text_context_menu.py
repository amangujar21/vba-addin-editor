"""Shared Windows-style right-click Cut/Copy/Paste for Tk Text widgets
(XML plan section 17). One instance per widget; no domain/save coupling."""

from __future__ import annotations

import tkinter as tk


def _has_selection(text: tk.Text) -> bool:
    try:
        text.index("sel.first")
        text.index("sel.last")
        return True
    except tk.TclError:
        return False


class TextContextMenu:
    """Installs <Button-3> Cut/Copy/Paste on a Text widget.

    Selection behavior mirrors normal Windows editors: right-click inside the
    selection keeps it; outside moves the insertion cursor and clears it.
    Commands use Tk virtual events so undo and clipboard semantics stay native.
    """

    def __init__(self, text: tk.Text) -> None:
        self.text = text
        self.menu = tk.Menu(text, tearoff=0)
        self.menu.add_command(label="Cut", command=self._cut)
        self.menu.add_command(label="Copy", command=self._copy)
        self.menu.add_command(label="Paste", command=self._paste)
        text.bind("<Button-3>", self._show_menu, add=True)

    # -- menu state -------------------------------------------------------

    def menu_states(self) -> tuple[str, str, str]:
        """(cut, copy, paste) states; exposed for tests and future reuse."""
        editable = str(self.text.cget("state")) != "disabled"
        selected = _has_selection(self.text)
        cut = "normal" if editable and selected else "disabled"
        copy = "normal" if selected else "disabled"
        try:
            self.text.clipboard_get()
            has_clipboard = True
        except tk.TclError:
            has_clipboard = False
        paste = "normal" if editable and has_clipboard else "disabled"
        return cut, copy, paste

    # -- handlers -----------------------------------------------------------

    def _show_menu(self, event) -> str:
        self.prepare_menu(event)
        try:
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu.grab_release()
        return "break"

    def prepare_menu(self, event) -> str:
        """Selection/cursor/state logic without opening the popup (testable)."""
        text = self.text
        try:
            clicked = text.index(f"@{event.x},{event.y}")
            in_selection = _has_selection(text) and (
                text.compare("sel.first", "<=", clicked)
                and text.compare(clicked, "<=", "sel.last")
            )
            if not in_selection:
                text.tag_remove("sel", "1.0", "end")
                text.mark_set("insert", clicked)
        except tk.TclError:
            pass
        text.focus_set()
        cut, copy, paste = self.menu_states()
        self.menu.entryconfig(0, state=cut)
        self.menu.entryconfig(1, state=copy)
        self.menu.entryconfig(2, state=paste)
        return "break"

    def _cut(self) -> None:
        self.text.event_generate("<<Cut>>")

    def _copy(self) -> None:
        self.text.event_generate("<<Copy>>")

    def _paste(self) -> None:
        text = self.text
        if _has_selection(text):
            # Windows semantics: paste replaces the selection. Tk's own
            # <<Paste>> binding inserts at the insert mark without removing a
            # programmatic selection, so do the replacement explicitly first.
            try:
                text.mark_set("insert", "sel.first")
                text.delete("sel.first", "sel.last")
            except tk.TclError:
                pass
        text.event_generate("<<Paste>>")


def install_text_context_menu(text: tk.Text) -> TextContextMenu:
    """Install the shared context menu on a Text widget and return the controller."""
    controller = TextContextMenu(text)
    # Keep a reference on the widget so Tk/Tcl objects are not collected.
    text._vbaae_context_menu = controller  # type: ignore[attr-defined]
    return controller
