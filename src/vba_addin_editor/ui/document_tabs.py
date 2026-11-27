"""Several add-ins open at once, one per tab, in one window.

The window has one set of views (Code, Ribbon, UserForms, Package XML). Each
open add-in is a ``DocumentTab`` holding everything that belongs to that file:
its draft and session (undo history and recovery live on the session), what is
shown in each view, a queued save. Switching tabs flushes the editors into the
outgoing draft, checkpoints it for recovery, and loads the incoming tab into
the views.

``MainWindow`` reads ``self.draft``, ``self.session``, ``self.current_module_id``
... everywhere; those are properties of the active tab here (``_TabField``), so
the single-document code keeps working on whichever tab is showing. While a
worker runs (``_saving``) tabs cannot be switched: the worker reads the active
tab's draft.
"""

from __future__ import annotations

import itertools
import os
import tkinter as tk
from dataclasses import dataclass
from pathlib import Path
from tkinter import messagebox
from typing import Any

from vba_addin_editor.adapters.ooxml_package_adapter import XML_EDITABLE_EXTENSIONS
from vba_addin_editor.platform.settings import save_settings
from vba_addin_editor.ui.tab_strip import TabInfo, TabStrip
from vba_addin_editor.version import APP_NAME

_keys = itertools.count(1)


@dataclass(eq=False)
class DocumentTab:
    draft: Any  # DocumentDraft; None only for the window's blank stand-in
    session: Any = None
    key: str = ""
    current_module_id: str | None = None
    current_xml_part_path: str | None = None
    ribbon_surface: str = "design"
    queued_save: str | None = None  # draft signature a confirmed save waits to write
    recovery_warning: str = ""
    view: int = 0
    vba_pos: tuple[str, float] | None = None  # (insert index, first visible fraction)
    xml_pos: tuple[str, float] | None = None

    def __post_init__(self) -> None:
        if not self.key:
            self.key = f"tab{next(_keys)}"

    @property
    def path(self) -> Path:
        return Path(self.draft.baseline.path)

    def is_dirty(self) -> bool:
        return self.draft is not None and self.draft.is_dirty()


class _TabField:
    """``window.<name>`` reads and writes the active tab's field (the blank one when none)."""

    def __init__(self, field_name: str) -> None:
        self.field_name = field_name

    def __get__(self, window: Any, owner: Any = None) -> Any:
        if window is None:
            return self
        return getattr(window._tab_state(), self.field_name)

    def __set__(self, window: Any, value: Any) -> None:
        tab = window._tab_state()
        if self.field_name == "draft" and tab is window._blank_tab and value is not None:
            # Code that opens a draft without a tab (older paths): it becomes a tab.
            tab = window._blank_tab
            window._blank_tab = DocumentTab(draft=None)
            window._tabs.append(tab)
            window._tab = tab
        setattr(tab, self.field_name, value)


def _same_file(a: Path, b: Path) -> bool:
    try:
        return os.path.normcase(str(a.resolve())) == os.path.normcase(str(b.resolve()))
    except OSError:
        return os.path.normcase(str(a)) == os.path.normcase(str(b))


class TabsMixin:
    """Tab management for MainWindow (which provides the views and the save flow).

    Methods take ``self: Any``: the views and the save flow they use are MainWindow's.
    """

    draft = _TabField("draft")
    session = _TabField("session")
    current_module_id = _TabField("current_module_id")
    current_xml_part_path = _TabField("current_xml_part_path")
    _ribbon_surface = _TabField("ribbon_surface")
    _queued_save = _TabField("queued_save")
    _recovery_warning = _TabField("recovery_warning")

    # Set by MainWindow (declared for type checkers).
    root: tk.Tk
    _saving: bool

    def _init_tabs(self: Any) -> None:
        self._tabs: list[DocumentTab] = []
        self._tab: DocumentTab | None = None
        self._blank_tab = DocumentTab(draft=None)

    def _tab_state(self: Any) -> DocumentTab:
        return self._tab if self._tab is not None else self._blank_tab

    def _build_tab_strip(self: Any) -> None:
        self.tab_strip = TabStrip(
            self.root,
            on_select=self.select_tab,
            on_close=self.close_tab,
            on_move=self.move_tab,
            on_new=self.open_file,
        )
        self.tab_strip_line = tk.Frame(self.root, height=1)

    # -- what is open --------------------------------------------------------------------

    @property
    def tabs(self: Any) -> list[DocumentTab]:
        return list(self._tabs)

    @property
    def active_tab(self: Any) -> DocumentTab | None:
        return self._tab

    def tab_for_path(self: Any, path: Path) -> DocumentTab | None:
        return next((t for t in self._tabs if _same_file(t.path, Path(path))), None)

    def tab_by_key(self: Any, key: str) -> DocumentTab | None:
        return next((t for t in self._tabs if t.key == key), None)

    def other_tabs(self: Any) -> list[DocumentTab]:
        return [t for t in self._tabs if t is not self._tab]

    # -- opening, switching, closing --------------------------------------------------------

    def _add_tab(self: Any, session, draft) -> DocumentTab:
        """A new tab for an opened file, right of the active one, and show it."""
        self._stash_tab()
        tab = DocumentTab(draft=draft, session=session)
        at = self._tabs.index(self._tab) + 1 if self._tab in self._tabs else len(self._tabs)
        self._tabs.insert(at, tab)
        self._tab = tab
        self._show_tab(first_open=True)
        self._remember_open_tabs()
        return tab

    def _replace_tab(self: Any, tab: DocumentTab, session, draft) -> None:
        """``tab`` (no unsaved changes) shows its file as just opened again."""
        if tab is not self._tab:
            self._stash_tab()
        if tab.session is not None and tab.session is not session:
            self.session_service.close(tab.session)
        tab.draft, tab.session = draft, session
        tab.current_module_id = tab.current_xml_part_path = None
        tab.queued_save, tab.recovery_warning = None, ""
        tab.vba_pos = tab.xml_pos = None
        tab.ribbon_surface = "design"
        self._tab = tab
        self._show_tab(first_open=True)
        self._remember_open_tabs()

    def select_tab(self: Any, key: str) -> bool:
        tab = self.tab_by_key(key)
        if tab is None or tab is self._tab:
            return tab is not None
        if self._saving:
            return False
        self._stash_tab()
        self._tab = tab
        self._show_tab()
        self._remember_open_tabs()
        return True

    def activate_tab(self: Any, tab: DocumentTab) -> bool:
        return self.select_tab(tab.key)

    def next_tab(self: Any, step: int = 1) -> str:
        if len(self._tabs) > 1 and self._tab in self._tabs:
            index = (self._tabs.index(self._tab) + step) % len(self._tabs)
            self.select_tab(self._tabs[index].key)
        return "break"

    def move_tab(self: Any, key: str, index: int) -> None:
        tab = self.tab_by_key(key)
        if tab is None:
            return
        self._tabs.remove(tab)
        self._tabs.insert(max(0, min(index, len(self._tabs))), tab)
        self._update_tab_strip()
        self._remember_open_tabs()

    def close_active_tab(self: Any) -> str:
        if self._tab is not None:
            self.close_tab(self._tab.key)
        return "break"

    def close_tab(self: Any, key: str, *, ask: bool = True) -> bool:
        """Close a tab, offering to save its changes first; False when it stays open."""
        tab = self.tab_by_key(key)
        if tab is None or self._saving:
            return False
        self._flush_all_editors()  # a background tab's text was flushed when it was left
        if ask and tab.is_dirty():
            if tab is not self._tab:
                self.select_tab(tab.key)  # its changes are reviewed where they can be seen
            choice = messagebox.askyesnocancel(
                APP_NAME, f"{tab.path.name} has unsaved changes. Save them before closing it?"
            )
            if choice is None:
                return False
            if choice:
                self.save()
                if tab.is_dirty():
                    return False  # the save failed, was blocked or was cancelled
            elif tab.session is not None:
                self.recovery_service.mark_discard(tab.session)
        self._drop_tab(tab)
        return True

    def _drop_tab(self: Any, tab: DocumentTab) -> None:
        if tab.session is not None:
            self.session_service.close(tab.session)
        index = self._tabs.index(tab)
        self._tabs.remove(tab)
        if tab is self._tab:
            self._cancel_autosave()
            self._tab = None
            self._close_document_dialogs()
            if self._tabs:
                self._tab = self._tabs[min(index, len(self._tabs) - 1)]
                self._show_tab()
            else:
                self._show_no_tab()
        if not any(t.queued_save is not None for t in self._tabs):
            self._stop_queued_save_poll()
        self._update_tab_strip()
        self._remember_open_tabs()

    # -- moving state between the views and a tab ---------------------------------------------

    def _stash_tab(self: Any) -> None:
        """The outgoing tab keeps its editors' text, a recovery checkpoint, and where it was."""
        tab = self._tab
        if tab is None:
            return
        self._checkpoint_now()  # flushes the editors into the draft first
        try:
            tab.view = int(self.editor_notebook.index("current"))
        except tk.TclError:
            tab.view = 0
        tab.vba_pos = (self.editor.text.index("insert"), self.editor.text.yview()[0])
        tab.xml_pos = (self.xml_editor.text.index("insert"), self.xml_editor.text.yview()[0])
        self._close_document_dialogs()

    def _show_tab(self: Any, *, first_open: bool = False) -> None:
        """Load the active tab into every view."""
        tab = self._tab
        assert tab is not None and tab.draft is not None
        draft = tab.draft
        self._refresh_host_status()
        self.ribbon_designer.reset()
        xml_supported = tab.path.suffix.lower() in XML_EDITABLE_EXTENSIONS
        for page in (self.xml_tab, self.ribbon_tab):
            self.editor_notebook.tab(page, state="normal" if xml_supported else "disabled")
        # The ids are set before the trees select them, so the selection events do not reload.
        module_id, part_path = tab.current_module_id, tab.current_xml_part_path
        module = draft.module_by_id(module_id) if module_id else None
        if module is None or module.is_deleted:
            module = next((m for m in draft.modules if not m.is_deleted), None)
        tab.current_module_id = module.id if module is not None else None
        part = draft.xml_part_by_path(part_path) if part_path and xml_supported else None
        tab.current_xml_part_path = part.path if part is not None else None
        self.editor.text.config(state="normal")
        self.editor.set_text(module.body if module is not None else "")
        self.xml_editor.text.config(state="normal")
        self.xml_editor.set_text((part.text or "") if part is not None else "")
        self.xml_editor.set_part_info(part)
        if part is not None and not part.editable:
            self.xml_editor.text.config(state="disabled")
        self.validate_xml_btn.config(state="normal" if part is not None else "disabled")
        self._refresh_tree()
        self._refresh_xml_tree()
        if module is not None and self.tree.exists(module.id):
            self.tree.selection_set(module.id)
            self.tree.see(module.id)
        for editor, pos in ((self.editor, tab.vba_pos), (self.xml_editor, tab.xml_pos)):
            if pos is not None:
                editor.text.mark_set("insert", pos[0])
                editor.text.yview_moveto(pos[1])
        view = tab.view if not first_open else 0
        if view in (1, 3) and not xml_supported:
            view = 0
        self._shown_view = ""  # make _on_view_changed bring the view up to date
        self.editor_notebook.select(view)
        self._on_view_changed()
        self._refresh_ribbon_views()
        self._schedule_outline()
        self._update_tab_strip()
        self._refresh_state()
        if tab.queued_save is not None:
            self._schedule_queued_save_poll()

    def _show_no_tab(self: Any) -> None:
        """Nothing is open: empty views and the welcome screen."""
        self.editor.set_text("")
        self.xml_editor.set_text("")
        self.xml_editor.set_part_info(None)
        self.validate_xml_btn.config(state="disabled")
        self._refresh_host_status()
        self._refresh_tree()
        self._refresh_xml_tree()
        self.ribbon_designer.reset()
        self.ribbon_designer.refresh()
        self.form_designer.refresh()
        self._update_tab_strip()
        self._refresh_state()

    def _close_document_dialogs(self: Any) -> None:
        """Dialogs that show one file's contents do not follow a tab switch."""
        for attr in ("_search_dialog", "_ribbon_dialog"):
            dialog = getattr(self, attr, None)
            if dialog is None:
                continue
            win = getattr(dialog, "win", None)
            try:
                if win is not None and win.winfo_exists():
                    win.destroy()
            except tk.TclError:
                pass
            setattr(self, attr, None)

    def _update_tab_strip(self: Any) -> None:
        strip = getattr(self, "tab_strip", None)
        if strip is None:
            return
        infos = [
            TabInfo(
                key=t.key,
                title=t.path.name,
                tip=str(t.path),
                dirty=t.is_dirty(),
                extension=t.path.suffix,
                waiting=t.queued_save is not None,
            )
            for t in self._tabs
        ]
        strip.set_tabs(infos, self._tab.key if self._tab is not None else None)
        if self._tabs and not strip.winfo_manager():
            strip.pack(fill="x", side="top", after=self.header_line)
            self.tab_strip_line.pack(fill="x", side="top", after=strip)
        elif not self._tabs and strip.winfo_manager():
            strip.pack_forget()
            self.tab_strip_line.pack_forget()

    # -- Save All, closing the window, reopening ----------------------------------------------

    def save_all(self: Any) -> None:
        """Save every tab with changes, one review each."""
        if self._saving:
            return
        self._flush_all_editors()
        start = self._tab
        for tab in [t for t in self._tabs if t.is_dirty()]:
            if tab not in self._tabs or not self.select_tab(tab.key):
                continue
            self.save()
        if start is not None and start in self._tabs:
            self.select_tab(start.key)

    def dirty_tabs(self: Any) -> list[DocumentTab]:
        self._flush_all_editors()
        return [t for t in self._tabs if t.is_dirty()]

    def _confirm_close_all(self: Any) -> bool:
        """Before the window closes: every tab's changes saved, or discarded on purpose."""
        dirty = self.dirty_tabs()
        if not dirty:
            return True
        if len(dirty) == 1:
            tab = dirty[0]
            self.select_tab(tab.key)
            self._refresh_host_status()
            extra = ""
            if self._live_host_running:
                extra = (
                    f"\n\nClose {self._host_label} before saving this add-in. "
                    "Saving now will be blocked until that application is closed."
                )
            choice = messagebox.askyesnocancel(
                APP_NAME, f"{tab.path.name} has unsaved changes. Save before exiting?" + extra
            )
        else:
            names = "\n".join(f"  • {t.path.name}" for t in dirty)
            choice = messagebox.askyesnocancel(
                APP_NAME,
                f"{len(dirty)} add-ins have unsaved changes:\n\n{names}\n\n"
                "Save them all before exiting? (Each is reviewed before it is saved.)",
            )
        if choice is None:
            return False
        if choice:
            for tab in dirty:
                self.select_tab(tab.key)
                self.save()
                if tab.is_dirty():
                    return False  # a save failed or was blocked: keep the window open
            return True
        for tab in dirty:
            if tab.session is not None:
                self.recovery_service.mark_discard(tab.session)
        return True

    def _remember_open_tabs(self: Any) -> None:
        settings = self.settings
        paths = [str(t.path) for t in self._tabs]
        active = str(self._tab.path) if self._tab is not None else ""
        if paths == settings.open_tabs and active == settings.active_tab:
            return
        settings.open_tabs, settings.active_tab = paths, active
        save_settings(settings, self._settings_path)

    def reopen_tabs(self: Any) -> int:
        """Open the add-ins that were open last time (those that still exist); how many opened."""
        if not self.settings.reopen_tabs:
            return 0
        wanted = [Path(p) for p in self.settings.open_tabs if isinstance(p, str) and p]
        active = self.settings.active_tab
        opened = 0
        for path in wanted:
            if self.tab_for_path(path) is not None:
                continue
            if not path.exists():
                continue
            if self.load_path(path, remember=False):
                opened += 1
        target = self.tab_for_path(Path(active)) if active else None
        if target is not None:
            self.select_tab(target.key)
        self._remember_open_tabs()
        return opened

    def _on_reopen_tabs_toggled(self: Any) -> None:
        self.settings.reopen_tabs = bool(self.reopen_tabs_var.get())
        save_settings(self.settings, self._settings_path)
