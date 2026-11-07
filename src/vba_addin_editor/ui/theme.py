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
    bg0: str  # chrome: title bar, activity rail, status bar
    bg1: str  # sidebars, panels, dialogs
    bg2: str  # editor text
    bg3: str  # inputs, raised surfaces
    bg4: str  # hover
    border: str  # hairlines between areas
    border_strong: str  # outlines of inputs, pills and ghost buttons
    text: str
    muted: str
    faint: str
    gutter: str  # line numbers away from the cursor
    accent: str
    accent_hover: str
    accent_bg: str  # tinted fill behind accent text (toggles, chips)
    on_accent: str
    select: str
    select_text: str
    tree_select: str
    current_line: str
    find_match: str
    warn: str  # also "edited"
    warn_bg: str
    ok: str  # also "new"
    info: str  # also "renamed"
    danger: str
    banner_bg: str
    banner_text: str
    banner_sub: str
    diff_add_bg: str
    diff_del_bg: str
    keyword: str
    string: str
    comment: str
    number: str
    xml_tag: str
    xml_attr: str
    file_tags: tuple[tuple[str, str, str], ...]  # (extension, background, foreground)

    def file_tag(self, extension: str) -> tuple[str, str]:
        for ext, bg, fg in self.file_tags:
            if ext == extension.lower():
                return bg, fg
        return self.bg3, self.muted


# One night palette: flat fills only, taken from the redesign mockup.
DARK = Palette(
    name="dark",
    bg0="#0b0c0f",
    bg1="#101115",
    bg2="#131418",
    bg3="#1a1c21",
    bg4="#22252b",
    border="#1c1e24",
    border_strong="#272a31",
    text="#e3e4e8",
    muted="#8c909a",
    faint="#555963",
    gutter="#3b3f48",
    accent="#8b9dff",
    accent_hover="#a3b2ff",
    accent_bg="#1b2040",
    on_accent="#0b0c0f",
    select="#2a3358",
    select_text="#ffffff",
    tree_select="#1d2233",
    current_line="#171920",
    find_match="#3a3218",
    warn="#e3b35c",
    warn_bg="#2f2716",
    ok="#6fcf97",
    info="#7aa8ff",
    danger="#f07178",
    banner_bg="#17140d",
    banner_text="#ecd9a8",
    banner_sub="#9a8c68",
    diff_add_bg="#12201a",
    diff_del_bg="#24161a",
    keyword="#8b9dff",
    string="#a8d58a",
    comment="#5c616c",
    number="#e8b46a",
    xml_tag="#8b9dff",
    xml_attr="#76cfc4",
    file_tags=((".xlam", "#14261b", "#6fcf97"), (".ppam", "#2a1a17", "#f0a078"), (".pptm", "#251b2c", "#c9a0f0")),
)

LIGHT = Palette(
    name="light",
    bg0="#eceef1",
    bg1="#f6f7f9",
    bg2="#ffffff",
    bg3="#e9ecf0",
    bg4="#dde1e7",
    border="#e1e4e9",
    border_strong="#cfd4db",
    text="#1f2328",
    muted="#5f6673",
    faint="#9aa1ac",
    gutter="#b4bac3",
    accent="#4459d6",
    accent_hover="#5a6ee6",
    accent_bg="#e3e8ff",
    on_accent="#ffffff",
    select="#cfd8fb",
    select_text="#1f2328",
    tree_select="#dfe5fb",
    current_line="#f4f6fa",
    find_match="#ffe7a3",
    warn="#a86a00",
    warn_bg="#fbefd6",
    ok="#2f8f52",
    info="#2f6fe0",
    danger="#c62f3a",
    banner_bg="#fff4d6",
    banner_text="#5c4300",
    banner_sub="#8a7440",
    diff_add_bg="#e6f6ea",
    diff_del_bg="#fbe9ea",
    keyword="#3346c2",
    string="#3c7d22",
    comment="#7a828e",
    number="#a35d00",
    xml_tag="#3346c2",
    xml_attr="#1d7f74",
    file_tags=((".xlam", "#dff3e5", "#1f7a41"), (".ppam", "#fbe6dd", "#b14a1c"), (".pptm", "#efe3f8", "#7a3fa8")),
)

PALETTES = {"dark": DARK, "light": LIGHT}

UI_FAMILY = "Segoe UI"
UI_SIZE = 10
CODE_FAMILY = "Consolas"  # last resort; the first installed of CODE_FAMILIES wins
CODE_FAMILIES = ("Cascadia Code", "Cascadia Mono", "Consolas")
UI_FAMILIES = ("Segoe UI Variable Text", "Segoe UI")
ICON_FAMILIES = ("Segoe Fluent Icons", "Segoe MDL2 Assets")
CODE_SIZE = 11
CODE_FONT = "VbaaeCode"  # named font shared by every code editor (zoom changes it once)
BOLD_FONT = "VbaaeBold"
SMALL_FONT = "VbaaeSmall"
SMALL_BOLD_FONT = "VbaaeSmallBold"
CAPTION_FONT = "VbaaeCaption"  # upper-case section headings
HERO_FONT = "VbaaeHero"
TITLE_FONT = "VbaaeTitle"
ICON_FONT = "VbaaeIcon"
ICON_SMALL_FONT = "VbaaeIconSmall"
LOGO_FONT = "VbaaeLogo"
LOGO_BIG_FONT = "VbaaeLogoBig"
TAG_FONT = "VbaaeTag"

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


def _first_family(root: tk.Misc, choices: tuple[str, ...], fallback: str) -> str:
    try:
        installed = {f.casefold() for f in root.tk.splitlist(root.tk.call("font", "families"))}
    except tk.TclError:
        return fallback
    return next((f for f in choices if f.casefold() in installed), fallback)


_icons_available = False


def icons_available() -> bool:
    """Whether a Windows icon font was found (otherwise widgets show text fallbacks)."""
    return _icons_available


def _configure_fonts(root: tk.Misc, code_size: int) -> None:
    global _icons_available
    ui = _first_family(root, UI_FAMILIES, UI_FAMILY)
    code = _first_family(root, CODE_FAMILIES, CODE_FAMILY)
    icon = _first_family(root, ICON_FAMILIES, "")
    _icons_available = bool(icon)
    for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkCaptionFont"):
        _named_font(root, name, family=ui, size=UI_SIZE)
    semibold = ui + " Semibold" if ui == UI_FAMILY else ui
    weight = "normal" if ui == UI_FAMILY else "bold"
    _named_font(root, BOLD_FONT, family=semibold, size=UI_SIZE, weight=weight)
    _named_font(root, SMALL_FONT, family=ui, size=UI_SIZE - 1)
    _named_font(root, SMALL_BOLD_FONT, family=semibold, size=UI_SIZE - 1, weight=weight)
    _named_font(root, CAPTION_FONT, family=semibold, size=UI_SIZE - 2, weight=weight)
    _named_font(root, HERO_FONT, family=semibold, size=UI_SIZE + 7, weight=weight)
    _named_font(root, TITLE_FONT, family=semibold, size=UI_SIZE + 1, weight=weight)
    _named_font(root, ICON_FONT, family=icon or ui, size=UI_SIZE + 2)
    _named_font(root, ICON_SMALL_FONT, family=icon or ui, size=UI_SIZE - 1)
    _named_font(root, LOGO_FONT, family=code, size=UI_SIZE, weight="bold")
    _named_font(root, LOGO_BIG_FONT, family=code, size=UI_SIZE + 6, weight="bold")
    _named_font(root, TAG_FONT, family=code, size=UI_SIZE - 3, weight="bold")
    _named_font(root, CODE_FONT, family=code, size=code_size)


def code_font(widget: tk.Misc) -> str:
    """Name of the shared code font, created with defaults if the window is unthemed."""
    if CODE_FONT not in widget.tk.splitlist(widget.tk.call("font", "names")):
        _named_font(widget, CODE_FONT, family=_first_family(widget, CODE_FAMILIES, CODE_FAMILY), size=CODE_SIZE)
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

    for style, bg in (("Chrome", c.bg0), ("Editor", c.bg2), ("Panel", c.bg1), ("Raised", c.bg3)):
        s.configure(f"{style}.TFrame", background=bg)
    s.configure("Line.TFrame", background=c.border)
    s.configure("Chrome.TLabel", background=c.bg0, foreground=c.muted)
    s.configure("Title.TLabel", background=c.bg0, foreground=c.text, font=TITLE_FONT)
    s.configure("Path.TLabel", background=c.bg0, foreground=c.faint, font=SMALL_FONT)
    s.configure("Dirty.TLabel", background=c.bg0, foreground=c.warn)
    s.configure("Status.TLabel", background=c.bg0, foreground=c.faint, font=SMALL_FONT)
    s.configure("StatusOk.TLabel", background=c.bg0, foreground=c.ok, font=ICON_SMALL_FONT)
    s.configure("StatusIcon.TLabel", background=c.bg0, foreground=c.faint, font=ICON_SMALL_FONT)
    s.configure("Logo.TLabel", background=c.accent, foreground=c.on_accent, font=LOGO_FONT, padding=(6, 1))
    s.configure("Banner.TFrame", background=c.banner_bg)
    s.configure("Banner.TLabel", background=c.banner_bg, foreground=c.banner_text)
    s.configure("BannerBold.TLabel", background=c.banner_bg, foreground=c.banner_text, font=BOLD_FONT)
    s.configure("BannerSub.TLabel", background=c.banner_bg, foreground=c.banner_sub, font=SMALL_FONT)
    s.configure("Muted.TLabel", foreground=c.muted)
    s.configure("Faint.TLabel", foreground=c.faint, font=SMALL_FONT)
    s.configure("Caption.TLabel", foreground=c.faint, font=CAPTION_FONT)
    s.configure("Hero.TLabel", foreground=c.text, font=HERO_FONT)
    s.configure("Link.TLabel", foreground=c.accent)
    s.configure("Heading.TLabel", font=BOLD_FONT)
    s.configure("Ok.TLabel", foreground=c.ok)
    s.configure("Warn.TLabel", foreground=c.warn)
    s.configure("PanelIcon.TLabel", foreground=c.faint, font=ICON_SMALL_FONT)
    s.configure("OkIcon.TLabel", foreground=c.ok, font=ICON_SMALL_FONT)
    s.configure("WarnIcon.TLabel", foreground=c.warn, font=ICON_SMALL_FONT)
    s.configure("EditorInfo.TLabel", background=c.bg2, foreground=c.faint, font=SMALL_FONT)
    # Breadcrumb bar above each editor.
    s.configure("Crumb.TLabel", background=c.bg2, foreground=c.faint, font=SMALL_FONT)
    s.configure("CrumbStrong.TLabel", background=c.bg2, foreground=c.muted, font=SMALL_BOLD_FONT)
    s.configure("CrumbSep.TLabel", background=c.bg2, foreground=c.faint, font=ICON_SMALL_FONT)
    for name, color in (("Warn", c.warn), ("Ok", c.ok), ("Info", c.info), ("Muted", c.faint)):
        s.configure(f"Crumb{name}.TLabel", background=c.bg2, foreground=color, font=SMALL_FONT)
    s.configure("Toast.TFrame", background=c.bg3)
    s.configure("Toast.TLabel", background=c.bg3, foreground=c.text, font=BOLD_FONT)
    s.configure("ToastSub.TLabel", background=c.bg3, foreground=c.faint, font=SMALL_FONT)
    s.configure("ToastIcon.TLabel", background=c.bg3, foreground=c.ok, font=ICON_FONT)
    s.configure("Raised.TLabel", background=c.bg3, foreground=c.text)
    s.configure("LogoBig.TLabel", background=c.accent, foreground=c.on_accent, font=LOGO_BIG_FONT, padding=(12, 4))
    # Labels that sit on the editor surface (welcome screen, form view).
    s.configure("EditorHero.TLabel", background=c.bg2, foreground=c.text, font=HERO_FONT)
    s.configure("EditorMuted.TLabel", background=c.bg2, foreground=c.muted)
    s.configure("EditorFaint.TLabel", background=c.bg2, foreground=c.faint, font=SMALL_FONT)
    s.configure("EditorCaption.TLabel", background=c.bg2, foreground=c.faint, font=CAPTION_FONT)
    for badge, color in (("M", c.warn), ("A", c.ok), ("R", c.info)):
        s.configure(f"Badge{badge}.TLabel", foreground=color, font=TAG_FONT)

    # Buttons: flat blocks; Accent is the one primary action, Ghost is outlined.
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
    for style, bg, fg, outline, hover in (
        ("Ghost.TButton", c.bg0, c.muted, c.border_strong, c.bg4),
        ("Panel.TButton", c.bg1, c.muted, c.border_strong, c.bg4),
        ("Banner.TButton", c.banner_bg, c.banner_text, c.banner_sub, c.banner_bg),
    ):
        s.configure(
            style,
            background=bg,
            foreground=fg,
            padding=(10, 3),
            width=0,
            focusthickness=0,
            lightcolor=bg,
            darkcolor=bg,
            bordercolor=outline,
        )
        s.map(
            style,
            foreground=[("disabled", c.faint), ("active", c.text if bg != c.banner_bg else fg)],
            background=[("pressed", hover), ("active", hover)],
            lightcolor=[("pressed", hover), ("active", hover)],
            darkcolor=[("pressed", hover), ("active", hover)],
        )
    # Borderless text and icon buttons.
    for style, bg, font in (
        ("Flat.TButton", c.bg0, None),
        ("Icon.TButton", c.bg1, ICON_SMALL_FONT),
        ("ChromeIcon.TButton", c.bg0, ICON_SMALL_FONT),
        ("EditorIcon.TButton", c.bg2, ICON_SMALL_FONT),
        ("RaisedIcon.TButton", c.bg3, ICON_SMALL_FONT),
    ):
        options = {"font": font} if font else {}
        s.configure(
            style,
            background=bg,
            foreground=c.muted if font is None else c.faint,
            padding=(9, 4) if font is None else (5, 3),
            width=0,
            focusthickness=0,
            **_flat(bg),
            **options,
        )
        s.map(
            style,
            foreground=[("disabled", c.faint), ("active", c.text)],
            **_flat_map(("pressed", c.bg4), ("active", c.bg4)),
        )
    s.configure("Changes.TButton", background=c.bg0, foreground=c.warn, padding=(10, 4), width=0, **_flat(c.bg0))
    s.map(
        "Changes.TButton",
        foreground=[("disabled", c.faint), ("active", c.warn)],
        **_flat_map(("pressed", c.bg4), ("active", c.bg4)),
    )
    s.configure(
        "EditorText.TButton", background=c.bg2, foreground=c.muted, font=SMALL_FONT, padding=(8, 2), width=0,
        **_flat(c.bg2),
    )
    s.map(
        "EditorText.TButton",
        foreground=[("disabled", c.faint), ("active", c.text)],
        **_flat_map(("pressed", c.bg4), ("active", c.bg4)),
    )
    s.configure(
        "EditorGhost.TButton", background=c.bg2, foreground=c.muted, padding=(14, 5), width=0, focusthickness=0,
        lightcolor=c.bg2, darkcolor=c.bg2, bordercolor=c.border_strong,
    )
    s.map(
        "EditorGhost.TButton",
        foreground=[("active", c.text)],
        background=[("active", c.bg4)],
        lightcolor=[("active", c.bg4)],
        darkcolor=[("active", c.bg4)],
    )
    s.configure(
        "CommandIcon.TButton", background=c.bg3, foreground=c.faint, font=ICON_SMALL_FONT, padding=(8, 5),
        focusthickness=0, lightcolor=c.bg3, darkcolor=c.bg3, bordercolor=c.border_strong,
    )
    s.map("CommandIcon.TButton", foreground=[("active", c.text)], background=[("active", c.bg4)])
    # The title bar's search / command box.
    s.configure(
        "Command.TButton",
        background=c.bg3,
        foreground=c.faint,
        padding=(10, 4),
        anchor="w",
        focusthickness=0,
        lightcolor=c.bg3,
        darkcolor=c.bg3,
        bordercolor=c.border_strong,
    )
    s.map(
        "Command.TButton",
        foreground=[("active", c.muted)],
        background=[("active", c.bg3)],
        bordercolor=[("active", c.faint)],
        lightcolor=[("active", c.bg3)],
        darkcolor=[("active", c.bg3)],
    )

    # Header menus: text only until hovered or open.
    menubutton_layout: Any = [("Menubutton.border", {"sticky": "nswe", "children": [
        ("Menubutton.padding", {"sticky": "nswe", "children": [
            ("Menubutton.label", {"sticky": ""})]})]})]
    for style, bg, fg in (("Header.TMenubutton", c.bg0, c.muted), ("Crumb.TMenubutton", c.bg2, c.text)):
        s.layout(style, menubutton_layout)
        s.configure(style, background=bg, foreground=fg, padding=(8, 3), width=0, **_flat(bg))
        s.map(
            style,
            foreground=[("disabled", c.faint), ("active", c.text), ("pressed", c.text)],
            **_flat_map(("pressed", c.bg4), ("active", c.bg4)),
        )
    s.configure("Crumb.TMenubutton", font=SMALL_FONT, padding=(6, 1))

    # The activity rail replaces tab strips: notebooks keep their pages, not their tabs.
    s.layout("Hidden.TNotebook.Tab", [])
    s.configure("Hidden.TNotebook", background=c.bg2, borderwidth=0, tabmargins=0, padding=0, **_flat(c.bg2))

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
        rowheight=linespace + 11,
        indent=12,
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

    entry = {
        "fieldbackground": c.bg3,
        "insertcolor": c.text,
        "padding": (7, 4),
        "lightcolor": c.bg3,
        "darkcolor": c.bg3,
        "bordercolor": c.border_strong,
    }
    s.configure("TEntry", **entry)
    s.map("TEntry", bordercolor=[("focus", c.accent)], lightcolor=[("focus", c.bg3)])
    s.configure("Editor.Treeview", background=c.bg2, fieldbackground=c.bg2)
    s.configure("Editor.Treeview.Heading", background=c.bg2, foreground=c.faint, font=SMALL_FONT, **_flat(c.bg2))
    s.configure("Palette.Treeview", background=c.bg3, fieldbackground=c.bg3, rowheight=linespace + 14)
    s.configure("Palette.TEntry", fieldbackground=c.bg3, padding=(10, 8), **_flat(c.bg3))
    s.map("Palette.TEntry", bordercolor=[("focus", c.bg3)], lightcolor=[("focus", c.bg3)])
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
    s.configure("TLabelframe", background=c.bg1, bordercolor=c.border_strong, lightcolor=c.bg1, darkcolor=c.bg1)
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
