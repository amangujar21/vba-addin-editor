"""Dark and light colour themes for the whole window (ttk "clam" + Tk option database).

Everything is drawn by Tk itself: no images, no extra Tcl packages. Colours live
only in the two palettes below; widgets ask for a named ttk style or read
``palette()`` instead of hardcoding a colour.
"""

from __future__ import annotations

import ctypes
import sys
import tkinter as tk
from dataclasses import dataclass
from tkinter import ttk
from typing import Any


@dataclass(frozen=True)
class Palette:
    name: str
    bg0: str  # chrome: header, tab strip, status bar
    bg1: str  # sidebars, panels, dialogs
    bg2: str  # editor text
    bg3: str  # buttons, entries
    bg4: str  # hover
    border: str
    text: str
    muted: str
    faint: str
    accent: str
    accent_hover: str
    on_accent: str
    select: str
    select_text: str
    tree_select: str
    current_line: str
    find_match: str
    warn: str
    ok: str
    info: str
    danger: str
    banner_bg: str
    banner_text: str
    keyword: str
    string: str
    comment: str
    number: str
    xml_tag: str
    xml_attr: str


DARK = Palette(
    name="dark",
    bg0="#16181c",
    bg1="#1b1e23",
    bg2="#1f2228",
    bg3="#272b32",
    bg4="#30353d",
    border="#2a2e35",
    text="#d4d7dd",
    muted="#8a909a",
    faint="#575d68",
    accent="#4d8dff",
    accent_hover="#6aa1ff",
    on_accent="#ffffff",
    select="#264f78",
    select_text="#ffffff",
    tree_select="#2a3a52",
    current_line="#23272e",
    find_match="#5b4a1c",
    warn="#e5b45a",
    ok="#6cc489",
    info="#6fa8ff",
    danger="#ef6b73",
    banner_bg="#2e2817",
    banner_text="#f0d48a",
    keyword="#7aa2f7",
    string="#9ece6a",
    comment="#6a717c",
    number="#e0af68",
    xml_tag="#7aa2f7",
    xml_attr="#bb9af7",
)

LIGHT = Palette(
    name="light",
    bg0="#eceef1",
    bg1="#f6f7f9",
    bg2="#ffffff",
    bg3="#e3e6eb",
    bg4="#d5d9e0",
    border="#d9dde3",
    text="#1f2328",
    muted="#5f6673",
    faint="#9aa1ac",
    accent="#2f6fe0",
    accent_hover="#4682f0",
    on_accent="#ffffff",
    select="#cfe0fb",
    select_text="#1f2328",
    tree_select="#dbe6f8",
    current_line="#f3f6fa",
    find_match="#ffe7a3",
    warn="#a86a00",
    ok="#2f8f52",
    info="#2f6fe0",
    danger="#c62f3a",
    banner_bg="#fff4d6",
    banner_text="#5c4300",
    keyword="#1f5fbf",
    string="#3c7d22",
    comment="#7a828e",
    number="#a35d00",
    xml_tag="#1f5fbf",
    xml_attr="#8a3fb0",
)

PALETTES = {"dark": DARK, "light": LIGHT}

UI_FAMILY = "Segoe UI"
UI_SIZE = 10
CODE_FAMILY = "Consolas"
CODE_SIZE = 11
CODE_FONT = "VbaaeCode"  # named font shared by every code editor (zoom changes it once)
BOLD_FONT = "VbaaeBold"
SMALL_FONT = "VbaaeSmall"
HERO_FONT = "VbaaeHero"

_active: Palette = DARK


def palette() -> Palette:
    """The palette the window was themed with."""
    return _active


def apply_theme(root: tk.Misc, name: str = "dark", code_size: int = CODE_SIZE) -> Palette:
    """Theme ``root`` and everything created under it afterwards."""
    global _active
    c = _active = PALETTES.get(name, DARK)
    _configure_fonts(root, code_size)
    _configure_option_database(root, c)
    _configure_styles(root, c)
    root.configure(background=c.bg0)  # type: ignore[call-arg]
    if not getattr(root, "_vbaae_themed", False):
        root._vbaae_themed = True  # type: ignore[attr-defined]
        # Dialogs get the matching title bar when they first appear.
        root.bind_class("Toplevel", "<Map>", _on_toplevel_map, add="+")
    return c


def menu_options(c: Palette) -> dict:
    """Colours for an existing tk.Menu (new menus take them from the option database)."""
    return {
        "background": c.bg3,
        "foreground": c.text,
        "activebackground": c.select,
        "activeforeground": c.select_text,
        "disabledforeground": c.faint,
        "selectcolor": c.text,
    }


def _named_font(root: tk.Misc, name: str, **options) -> None:
    # Plain Tcl, not tkinter.font.Font: a Font object deletes its Tcl font when
    # it is garbage collected, possibly on a worker thread.
    args: list = []
    for key, value in options.items():
        args += [f"-{key}", value]
    if name in root.tk.splitlist(root.tk.call("font", "names")):
        root.tk.call("font", "configure", name, *args)
    else:
        root.tk.call("font", "create", name, *args)


def _configure_fonts(root: tk.Misc, code_size: int) -> None:
    for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkCaptionFont"):
        _named_font(root, name, family=UI_FAMILY, size=UI_SIZE)
    _named_font(root, BOLD_FONT, family=UI_FAMILY + " Semibold", size=UI_SIZE)
    _named_font(root, SMALL_FONT, family=UI_FAMILY, size=UI_SIZE - 1)
    _named_font(root, HERO_FONT, family=UI_FAMILY + " Semibold", size=UI_SIZE + 8)
    _named_font(root, CODE_FONT, family=CODE_FAMILY, size=code_size)


def code_font(widget: tk.Misc) -> str:
    """Name of the shared code font, created with defaults if the window is unthemed."""
    if CODE_FONT not in widget.tk.splitlist(widget.tk.call("font", "names")):
        _named_font(widget, CODE_FONT, family=CODE_FAMILY, size=CODE_SIZE)
    return CODE_FONT


def _configure_option_database(root: tk.Misc, c: Palette) -> None:
    """Classic Tk widgets (Text, Listbox, Menu, Canvas) read their colours from here."""
    options = {
        "*Menu.background": c.bg3,
        "*Menu.foreground": c.text,
        "*Menu.activeBackground": c.select,
        "*Menu.activeForeground": c.select_text,
        "*Menu.disabledForeground": c.faint,
        "*Menu.selectColor": c.text,
        "*Menu.relief": "flat",
        "*Menu.borderWidth": 0,
        "*Menu.activeBorderWidth": 0,
        "*Listbox.background": c.bg2,
        "*Listbox.foreground": c.text,
        "*Listbox.selectBackground": c.select,
        "*Listbox.selectForeground": c.select_text,
        "*Listbox.disabledForeground": c.faint,
        "*Listbox.highlightThickness": 0,
        "*Listbox.borderWidth": 0,
        "*Listbox.activeStyle": "none",
        "*Text.background": c.bg2,
        "*Text.foreground": c.text,
        "*Text.insertBackground": c.text,
        "*Text.selectBackground": c.select,
        "*Text.selectForeground": c.select_text,
        "*Text.inactiveSelectBackground": c.select,
        "*Text.highlightThickness": 0,
        "*Text.borderWidth": 0,
        "*Text.relief": "flat",
        "*Text.padX": 6,
        "*Text.padY": 4,
        "*Toplevel.background": c.bg1,
        "*Canvas.background": c.bg2,
        "*Canvas.highlightThickness": 0,
        "*TCombobox*Listbox.background": c.bg3,
        "*TCombobox*Listbox.foreground": c.text,
        "*TCombobox*Listbox.selectBackground": c.select,
        "*TCombobox*Listbox.selectForeground": c.select_text,
    }
    for pattern, value in options.items():
        root.option_add(pattern, value)


def _flat(color: str) -> dict:
    """Border colours that make a clam element read as a flat block of ``color``."""
    return {"lightcolor": color, "darkcolor": color, "bordercolor": color}


def _flat_map(*pairs: tuple[str, str]) -> dict:
    states = list(pairs)
    return {"background": states, "lightcolor": states, "darkcolor": states, "bordercolor": states}


def _configure_styles(root: tk.Misc, c: Palette) -> None:
    s = ttk.Style(root)
    s.theme_use("clam")
    # theme_use queues <<ThemeChanged>> for idle time; deliver it now, while
    # there are few widgets and the window surely still exists.
    root.update_idletasks()
    s.configure(
        ".",
        background=c.bg1,
        foreground=c.text,
        fieldbackground=c.bg3,
        troughcolor=c.bg1,
        selectbackground=c.select,
        selectforeground=c.select_text,
        insertcolor=c.text,
        arrowcolor=c.muted,
        focuscolor=c.accent,
        relief="flat",
        **_flat(c.bg1),
    )
    s.map(".", foreground=[("disabled", c.faint)])

    s.configure("Chrome.TFrame", background=c.bg0)
    s.configure("Editor.TFrame", background=c.bg2)
    s.configure("Chrome.TLabel", background=c.bg0, foreground=c.muted)
    s.configure("Title.TLabel", background=c.bg0, foreground=c.text, font=BOLD_FONT)
    s.configure("Dirty.TLabel", background=c.bg0, foreground=c.warn)
    s.configure("Status.TLabel", background=c.bg0, foreground=c.muted, font=SMALL_FONT)
    s.configure("Logo.TLabel", background=c.bg0, foreground=c.accent)
    s.configure("Banner.TFrame", background=c.banner_bg)
    s.configure("Banner.TLabel", background=c.banner_bg, foreground=c.banner_text)
    s.configure("Muted.TLabel", foreground=c.muted)
    s.configure("Faint.TLabel", foreground=c.faint, font=SMALL_FONT)
    s.configure("Caption.TLabel", foreground=c.faint, font=SMALL_FONT)
    s.configure("Hero.TLabel", foreground=c.text, font=HERO_FONT)
    s.configure("Link.TLabel", foreground=c.info)
    s.configure("EditorInfo.TLabel", background=c.bg2, foreground=c.muted, font=SMALL_FONT)
    s.configure("Heading.TLabel", font=BOLD_FONT)

    # Buttons: flat blocks; Accent is the one primary action, Ghost sits on chrome.
    s.configure("TButton", background=c.bg3, padding=(12, 4), focusthickness=0, **_flat(c.bg3))
    s.map(
        "TButton",
        foreground=[("disabled", c.faint)],
        **_flat_map(("disabled", c.bg1), ("pressed", c.bg4), ("active", c.bg4)),
    )
    s.configure("Toolbar.TButton", width=0, padding=(14, 4))
    s.configure(
        "Accent.TButton", background=c.accent, foreground=c.on_accent, font=BOLD_FONT, **_flat(c.accent)
    )
    s.map(
        "Accent.TButton",
        foreground=[("disabled", c.faint)],
        **_flat_map(("disabled", c.bg3), ("pressed", c.accent), ("active", c.accent_hover)),
    )
    for style, bg in (("Ghost.TButton", c.bg0), ("Panel.TButton", c.bg1)):
        s.configure(style, background=bg, padding=(10, 3), width=0, focusthickness=0, **_flat(bg))
        s.map(
            style,
            foreground=[("disabled", c.faint)],
            **_flat_map(("disabled", bg), ("pressed", c.bg4), ("active", c.bg3)),
        )

    # Header menus: text only until hovered or open.
    s.layout(
        "Header.TMenubutton",
        [("Menubutton.border", {"sticky": "nswe", "children": [
            ("Menubutton.padding", {"sticky": "nswe", "children": [
                ("Menubutton.label", {"sticky": ""})]})]})],
    )
    s.configure(
        "Header.TMenubutton", background=c.bg0, foreground=c.muted, padding=(9, 4), width=0, **_flat(c.bg0)
    )
    s.map(
        "Header.TMenubutton",
        foreground=[("disabled", c.faint), ("active", c.text), ("pressed", c.text)],
        **_flat_map(("pressed", c.bg3), ("active", c.bg3)),
    )

    s.layout(
        "TNotebook.Tab",
        [("Notebook.tab", {"sticky": "nswe", "children": [
            ("Notebook.padding", {"side": "top", "sticky": "nswe", "children": [
                ("Notebook.label", {"side": "top", "sticky": ""})]})]})],
    )
    s.configure("TNotebook", background=c.bg0, borderwidth=0, tabmargins=(8, 2, 0, 0), **_flat(c.bg0))
    s.configure(
        "TNotebook.Tab", background=c.bg0, foreground=c.muted, padding=(16, 6), borderwidth=0, **_flat(c.bg0)
    )
    s.map(
        "TNotebook.Tab",
        foreground=[("disabled", c.faint), ("selected", c.text), ("active", c.text)],
        expand=[("selected", (0, 0, 0, 0))],
        **_flat_map(("selected", c.bg1)),
    )

    s.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
    linespace = int(root.tk.call("font", "metrics", "TkDefaultFont", "-linespace"))
    s.configure(
        "Treeview",
        background=c.bg1,
        fieldbackground=c.bg1,
        foreground=c.text,
        borderwidth=0,
        rowheight=linespace + 9,
        indent=14,
    )
    s.map("Treeview", background=[("selected", c.tree_select)], foreground=[("selected", c.select_text)])
    s.configure("Treeview.Heading", background=c.bg0, foreground=c.muted, padding=(6, 4), **_flat(c.bg0))
    s.map("Treeview.Heading", background=[("active", c.bg3)])

    # Thin scrollbars: just a thumb in a trough, no arrows.
    for orient, sticky in (("Vertical", "ns"), ("Horizontal", "ew")):
        thumb: Any = [(f"{orient}.Scrollbar.thumb", {"expand": True, "sticky": "nswe"})]
        s.layout(f"{orient}.TScrollbar", [(f"{orient}.Scrollbar.trough", {"sticky": sticky, "children": thumb})])
        s.configure(
            f"{orient}.TScrollbar",
            troughcolor=c.bg2,
            background=c.bg4,
            borderwidth=0,
            arrowsize=10,
            gripcount=0,
            lightcolor=c.bg4,
            darkcolor=c.bg4,
            bordercolor=c.bg2,
        )
        s.map(f"{orient}.TScrollbar", background=[("active", c.faint)])

    s.configure("TPanedwindow", background=c.bg0)
    s.configure("Sash", sashthickness=3, gripcount=0, background=c.bg0, **_flat(c.bg0))

    entry = {"fieldbackground": c.bg3, "insertcolor": c.text, "padding": (6, 4), **_flat(c.bg3)}
    s.configure("TEntry", **entry)
    s.map("TEntry", bordercolor=[("focus", c.accent)], lightcolor=[("focus", c.accent)])
    s.configure("TCombobox", background=c.bg3, arrowcolor=c.muted, **entry)
    s.map(
        "TCombobox",
        fieldbackground=[("readonly", c.bg3)],
        foreground=[("disabled", c.faint)],
        bordercolor=[("focus", c.accent)],
        lightcolor=[("focus", c.accent)],
    )
    s.configure("TSpinbox", background=c.bg3, arrowcolor=c.muted, **entry)

    for style, bg in (("TCheckbutton", c.bg1), ("Chrome.TCheckbutton", c.bg0), ("Editor.TCheckbutton", c.bg2)):
        s.configure(
            style,
            background=bg,
            foreground=c.text if bg == c.bg1 else c.muted,
            indicatorbackground=c.bg3,
            indicatorforeground=c.on_accent,
            indicatorrelief="flat",
            focusthickness=0,
            **_flat(bg),
        )
        s.map(
            style,
            background=[("active", bg)],
            foreground=[("disabled", c.faint), ("active", c.text)],
            indicatorbackground=[("selected", c.accent), ("active", c.bg4)],
        )
    s.configure("TRadiobutton", indicatorbackground=c.bg3, indicatorforeground=c.on_accent, focusthickness=0)
    s.map("TRadiobutton", background=[("active", c.bg1)], indicatorbackground=[("selected", c.accent)])
    s.configure("TLabelframe", background=c.bg1, bordercolor=c.border, lightcolor=c.bg1, darkcolor=c.bg1)
    s.configure("TLabelframe.Label", background=c.bg1, foreground=c.muted)
    s.configure("TSeparator", background=c.border)
    s.configure("TProgressbar", background=c.accent, troughcolor=c.bg3, **_flat(c.bg3))


def autohide(scrollbar: ttk.Scrollbar, **pack_options):
    """A yscrollcommand that packs ``scrollbar`` only while there is something to scroll."""

    def update(first, last) -> None:
        scrollbar.set(first, last)
        needed = float(first) > 0.0 or float(last) < 1.0
        if needed and not scrollbar.winfo_manager():
            scrollbar.pack(**pack_options)
        elif not needed and scrollbar.winfo_manager():
            scrollbar.pack_forget()

    return update


# -- Windows title bar ----------------------------------------------------------

_DWMWA_USE_IMMERSIVE_DARK_MODE = (20, 19)  # Windows 10 20H1+, then 1809-1909
_WM_NCACTIVATE = 0x0086


def set_title_bar(window: tk.Misc, dark: bool | None = None) -> None:
    """Ask Windows for a dark (or light) title bar on ``window``; a no-op elsewhere.

    Set before the window is first shown: Windows 10 only repaints the caption
    on the next activation change.
    """
    if sys.platform != "win32":
        return
    if dark is None:
        dark = _active.name == "dark"
    try:
        window.update_idletasks()
        hwnd = int(window.wm_frame(), 16)  # type: ignore[attr-defined]
        value = ctypes.c_int(1 if dark else 0)
        for attribute in _DWMWA_USE_IMMERSIVE_DARK_MODE:
            result = ctypes.windll.dwmapi.DwmSetWindowAttribute(  # type: ignore[attr-defined]
                hwnd, attribute, ctypes.byref(value), ctypes.sizeof(value)
            )
            if result == 0:
                break
        if window.winfo_ismapped():
            # Already on screen: make Windows repaint the caption now.
            send = ctypes.windll.user32.SendMessageW  # type: ignore[attr-defined]
            active = window.focus_get() is not None
            send(hwnd, _WM_NCACTIVATE, int(not active), 0)
            send(hwnd, _WM_NCACTIVATE, int(active), 0)
    except (AttributeError, OSError, ValueError, KeyError, tk.TclError):
        pass


def _on_toplevel_map(event) -> None:
    window = event.widget
    if isinstance(window, tk.Toplevel) and not getattr(window, "_vbaae_titled", False):
        window._vbaae_titled = True  # type: ignore[attr-defined]
        set_title_bar(window)
