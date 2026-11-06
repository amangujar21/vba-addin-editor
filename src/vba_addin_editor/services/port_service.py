"""Port a ribbon control and the VBA behind it from one add-in into another.

The control (a button, toggle, menu, group, ...) is copied from a source
add-in's customUI into the open (target) add-in, together with every piece of
VBA it needs: its callbacks, the procedures they call, module-level
declarations (constants, variables, Declare statements, Types, Enums) and
whole class modules, followed transitively.

Both add-ins may be loaded in Office at the same time, so the copy must not
clash with the target or with the add-in it came from:

* ribbon callbacks, and procedures run by name (``Application.Run``,
  ``OnTime``, ``OnKey`` ...), always get a new name (``<name><suffix>``), so
  no macro name is shared between the two add-ins;
* other public names, module names and control ids are renamed only when the
  target already uses them;
* every reference in the copied code and XML follows the renames, including
  control ids in strings (``InvalidateControl "btnX"``) and macro names in
  strings (``Application.Run "'Source.xlam'!Proc"``);
* code the target already has, identically, is reused instead of duplicated;
* a copied ``IRibbonUI`` variable is set from the target's ribbon onLoad
  (a small hook Sub is added, and called from the target's onLoad).

Each source module that contributes code becomes one new module in the target
(class modules are copied whole), so Private names keep their meaning. Pure
logic: nothing here touches Tk, and the target draft is only changed by
``apply_port``.
"""

from __future__ import annotations

import bisect
import re
from dataclasses import dataclass, field
from itertools import pairwise
from pathlib import Path

from vba_addin_editor.adapters.ooxml_package_adapter import open_package
from vba_addin_editor.adapters.source_codec import replace_vb_name
from vba_addin_editor.domain.document import (
    DocumentDraft,
    ModuleDisplayKind,
    ModuleDraft,
    clone_draft,
    new_module_id,
)
from vba_addin_editor.services import icon_service
from vba_addin_editor.services import ribbon_service as rs
from vba_addin_editor.services.validation_service import VBA_RESERVED, validate_module_name

MAX_MODULE_NAME = 31
MAX_IDENTIFIER = 255

# Controls a user can pick to port, and the containers they may sit in.
CONTROL_ELEMENTS = frozenset(
    {
        "button",
        "toggleButton",
        "checkBox",
        "editBox",
        "comboBox",
        "dropDown",
        "gallery",
        "menu",
        "dynamicMenu",
        "splitButton",
        "labelControl",
        "box",
        "buttonGroup",
    }
)
_CONTROL_PARENTS = frozenset({"group", "menu", "box", "buttonGroup"})
_BUTTON_GROUP_CHILDREN = frozenset(
    {"button", "toggleButton", "gallery", "menu", "splitButton", "dynamicMenu", "control"}
)
# Attributes that place a control relative to its old neighbours.
_POSITION_ATTRS = frozenset({"insertAfterMso", "insertBeforeMso", "insertAfterQ", "insertBeforeQ"})
# Words of a statement whose string arguments name a macro to run.
_MACRO_RUNNERS = frozenset(
    {"run", "ontime", "onkey", "onaction", "macrooptions", "onundo", "onrepeat"}
)
_MACRO_STRING_RE = re.compile(
    r"^(?:(?P<book>'[^']*'|[^'!\s\[\]]+)?!)?"
    r"(?:(?P<mod>[A-Za-z][A-Za-z0-9_]*)\.)?(?P<proc>[A-Za-z][A-Za-z0-9_]*)$"
)
_SUFFIX_RE = re.compile(r"^[A-Za-z0-9_]{1,24}$")
# Library references every Office project has; never worth a warning.
_STANDARD_REFERENCES = frozenset(
    {
        "{000204ef-0000-0000-c000-000000000046}",  # VBA
        "{00020813-0000-0000-c000-000000000046}",  # Excel
        "{91493440-5a91-11cf-8700-00aa0060263b}",  # PowerPoint
        "{00020430-0000-0000-c000-000000000046}",  # stdole (OLE Automation)
        "{2df8d04c-5bfa-101b-bde5-00aa0044de52}",  # Office
    }
)
_DEF_TYPES = frozenset(
    {
        "int", "lng", "bool", "byte", "cur", "sng", "dbl", "dec", "date", "str", "obj", "var",
        "lnglng", "lngptr",
    }
)
_VISIBILITY = frozenset({"public", "private", "friend", "global"})
_WHAT_LABELS = {
    "proc": "macro",
    "var": "variable",
    "const": "constant",
    "declare": "API declaration",
    "type": "type",
    "enum": "enum",
    "member": "enum member",
    "event": "event",
}


class PortError(ValueError):
    """The port cannot be planned; the message is for the user."""


# -- VBA lexing -----------------------------------------------------------------


@dataclass(frozen=True)
class _Tok:
    kind: str  # id | str | num | op | nl
    text: str  # for str: the decoded value
    start: int
    end: int


_IDENT_RE = re.compile(r"[A-Za-zÀ-￿][A-Za-z0-9_À-￿]*")
_NUMBER_RE = re.compile(r"(?:\d+\.?\d*|\.\d+)(?:[eEdD][+-]?\d+)?[%&!#@^]?")
_RADIX_RE = re.compile(r"&[HhOo][0-9A-Fa-f]+[%&^]?")


def _lex(text: str) -> list[_Tok]:
    """VBA tokens of ``text``; comments, continuations and blanks are dropped.

    Each logical line ends with an ``nl`` token (continued lines do not).
    """
    toks: list[_Tok] = []
    i, n = 0, len(text)

    def statement_start() -> bool:
        return not toks or toks[-1].kind == "nl" or _is_op(toks[-1], ":")

    while i < n:
        c = text[i]
        if c == "\n":
            toks.append(_Tok("nl", "\n", i, i + 1))
            i += 1
            continue
        if c in " \t\r\f\v":
            i += 1
            continue
        if c == "_" and (i == 0 or text[i - 1] in " \t"):
            j = i + 1
            while j < n and text[j] in " \t\r":
                j += 1
            if j >= n or text[j] == "\n":
                i = j + 1  # line continuation: the statement goes on
                continue
        if c == "'":
            while i < n and text[i] != "\n":
                i += 1
            continue
        if c == '"':
            j = i + 1
            value: list[str] = []
            while j < n and text[j] != "\n":
                if text[j] == '"':
                    if j + 1 < n and text[j + 1] == '"':
                        value.append('"')
                        j += 2
                        continue
                    break
                value.append(text[j])
                j += 1
            end = j + 1 if j < n and text[j] == '"' else j
            toks.append(_Tok("str", "".join(value), i, end))
            i = end
            continue
        if c == "[":
            j = text.find("]", i)
            line_end = text.find("\n", i)
            if j != -1 and (line_end == -1 or j < line_end):
                toks.append(_Tok("op", text[i : j + 1], i, j + 1))
                i = j + 1
                continue
        match = _IDENT_RE.match(text, i)
        if match:
            word = match.group(0)
            if word.casefold() == "rem" and statement_start():
                while i < n and text[i] != "\n":
                    i += 1
                continue
            toks.append(_Tok("id", word, i, match.end()))
            i = match.end()
            continue
        if c.isdigit() or (c == "." and i + 1 < n and text[i + 1].isdigit()):
            match = _NUMBER_RE.match(text, i)
            if match:
                toks.append(_Tok("num", match.group(0), i, match.end()))
                i = match.end()
                continue
        if c == "&":
            match = _RADIX_RE.match(text, i)
            if match:
                toks.append(_Tok("num", match.group(0), i, match.end()))
                i = match.end()
                continue
        if c == ":" and i + 1 < n and text[i + 1] == "=":
            toks.append(_Tok("op", ":=", i, i + 2))
            i += 2
            continue
        toks.append(_Tok("op", c, i, i + 1))
        i += 1
    return toks


def _is_op(tok: _Tok | None, text: str) -> bool:
    return tok is not None and tok.kind == "op" and tok.text == text


def _word(tok: _Tok | None) -> str:
    return tok.text.casefold() if tok is not None and tok.kind == "id" else ""


# -- module structure ---------------------------------------------------------------


@dataclass
class _Name:
    name: str
    public: bool
    what: str  # proc | var | const | declare | type | enum | member | event


@dataclass
class _Item:
    """One top-level piece of a module: a procedure, a declaration, or a #If block."""

    index: int
    start: int  # body offset, attached comments included
    code_start: int  # start of the first code line
    end: int  # just past the last line
    last_line_start: int
    kind: str = "decl"  # option | decl | proc
    names: dict[str, _Name] = field(default_factory=dict)
    ribbon_vars: list[str] = field(default_factory=list)  # IRibbonUI variables it declares
    signature: str = ""


@dataclass
class _Line:
    start: int
    end: int
    toks: list[_Tok]


def _logical_lines(body: str, toks: list[_Tok]) -> list[_Line]:
    lines: list[_Line] = []
    begin, current = 0, []
    for tok in toks:
        if tok.kind == "nl":
            lines.append(_Line(begin, tok.end, current))
            begin, current = tok.end, []
        else:
            current.append(tok)
    if begin < len(body):
        lines.append(_Line(begin, len(body), current))
    return lines


def _segments(toks: list[_Tok]) -> list[list[_Tok]]:
    """Split a declaration list at its top-level commas."""
    out: list[list[_Tok]] = [[]]
    depth = 0
    for tok in toks:
        if _is_op(tok, "("):
            depth += 1
        elif _is_op(tok, ")"):
            depth -= 1
        elif _is_op(tok, ",") and depth == 0:
            out.append([])
            continue
        out[-1].append(tok)
    return [seg for seg in out if seg]


def _is_ribbon_type(seg: list[_Tok]) -> bool:
    words = [_word(t) for t in seg]
    if "as" not in words:
        return False
    rest = [w for w in words[words.index("as") + 1 :] if w and w != "new"]
    return bool(rest) and rest[-1] == "iribbonui" and (len(rest) == 1 or rest[0] == "office")


def _classify(toks: list[_Tok]) -> tuple[str, ...]:
    """What a top-level logical line is: (kind, *details)."""
    if not toks:
        return ("trivia",)
    if _is_op(toks[0], "#"):
        word = _word(toks[1]) if len(toks) > 1 else ""
        if word == "if":
            return ("pp_if",)
        if word == "end":
            return ("pp_end",)
        if word == "const":
            return ("option",)
        return ("pp_else",)
    words = [_word(t) for t in toks[:6]]
    first = words[0]
    if first == "option" or (first.startswith("def") and first[3:] in _DEF_TYPES):
        return ("option",)
    if first == "attribute":
        return ("attribute",)
    if first == "end" and len(words) > 1:
        if words[1] in ("sub", "function", "property"):
            return ("proc_end",)
        if words[1] in ("type", "enum"):
            return (f"{words[1]}_end",)
    i, vis = 0, ""
    if first in _VISIBILITY:
        i, vis = 1, first
    if words[i : i + 1] == ["static"]:
        i += 1
    kind = words[i] if i < len(words) else ""

    def ident(j: int) -> str | None:
        return toks[j].text if j < len(toks) and toks[j].kind == "id" else None

    if kind in ("sub", "function", "property"):
        name = ident(i + (2 if kind == "property" else 1))
        if name:
            return ("proc_start", name, vis)
    if kind == "declare":
        j = i + 1
        if _word(toks[j] if j < len(toks) else None) == "ptrsafe":
            j += 1
        name = ident(j + 1)
        if name:
            return ("declare", name, vis)
    if kind in ("type", "enum", "event") and ident(i + 1):
        return (f"{kind}_start" if kind != "event" else "event", toks[i + 1].text, vis)
    if kind == "const":
        return ("const", vis, str(i + 1))
    if first == "dim":
        return ("var", "dim", "1")
    if vis:
        return ("var", vis, str(i))
    return ("other",)


def _parse_items(body: str, toks: list[_Tok]) -> list[_Item]:
    items: list[_Item] = []
    cur: _Item | None = None
    pending: list[_Line] = []  # comment lines right above the next item
    depth = 0
    proc_open = False
    block: str | None = None
    block_public = True

    for line in _logical_lines(body, toks):
        cls = _classify(line.toks)
        kind = cls[0]
        if cur is None:
            if kind == "trivia":
                pending = pending + [line] if body[line.start : line.end].strip() else []
                continue
            if kind == "attribute" and items:
                items[-1].end = line.end  # e.g. Attribute gVar.VB_VarHelpID = -1
                items[-1].last_line_start = line.start
                pending = []
                continue
            start = pending[0].start if pending else line.start
            cur = _Item(len(items), start, line.start, line.end, line.start)
            pending = []
            depth, proc_open, block = 0, False, None
        cur.end = line.end
        cur.last_line_start = line.start
        if kind == "pp_if":
            depth += 1
        elif kind == "pp_end":
            depth = max(0, depth - 1)
        elif proc_open:
            if kind == "proc_end":
                proc_open = False
        elif block is not None:
            if kind == f"{block}_end":
                block = None
            elif block == "enum" and line.toks and line.toks[0].kind == "id":
                name = line.toks[0].text
                cur.names.setdefault(name.casefold(), _Name(name, block_public, "member"))
        elif kind == "proc_start":
            proc_open = True
            cur.kind = "proc"
            if not cur.signature:
                cur.signature = body[line.start : line.end].strip()
            cur.names.setdefault(cls[1].casefold(), _Name(cls[1], cls[2] != "private", "proc"))
        elif kind in ("type_start", "enum_start"):
            block = kind.split("_")[0]
            block_public = cls[2] != "private"
            cur.names.setdefault(cls[1].casefold(), _Name(cls[1], block_public, block))
        elif kind in ("declare", "event"):
            public = kind == "event" or cls[2] != "private"
            cur.names.setdefault(cls[1].casefold(), _Name(cls[1], public, kind))
        elif kind in ("const", "var"):
            public = cls[1] in ("public", "global")
            for seg in _segments(line.toks[int(cls[2]) :]):
                if _word(seg[0]) == "withevents":
                    seg = seg[1:]
                if not seg or seg[0].kind != "id":
                    continue
                cur.names.setdefault(seg[0].text.casefold(), _Name(seg[0].text, public, kind))
                if kind == "var" and _is_ribbon_type(seg):
                    cur.ribbon_vars.append(seg[0].text)
        elif kind == "option" and not cur.names and cur.kind == "decl":
            cur.kind = "option"
        if depth == 0 and not proc_open and block is None:
            items.append(cur)
            cur = None
    if cur is not None:
        items.append(cur)
    return items


def _type_member_starts(body: str, toks: list[_Tok]) -> set[int]:
    """Offsets of the member names inside Type ... End Type blocks."""
    found: set[int] = set()
    in_type = False
    for line in _logical_lines(body, toks):
        kind = _classify(line.toks)[0]
        if kind == "type_start":
            in_type = True
        elif kind == "type_end":
            in_type = False
        elif in_type and line.toks and line.toks[0].kind == "id":
            found.add(line.toks[0].start)
    return found


# -- the two projects ---------------------------------------------------------------------


@dataclass
class _Module:
    name: str
    kind: str  # standard | class | document | form
    draft: ModuleDraft
    body: str
    toks: list[_Tok]
    starts: list[int]  # token starts, for bisect
    items: list[_Item]
    by_name: dict[str, list[_Item]]
    type_members: set[int]

    def token_range(self, start: int, end: int) -> range:
        return range(bisect.bisect_left(self.starts, start), bisect.bisect_left(self.starts, end))


def _module_kind(module: ModuleDraft) -> str:
    if module.pyopenvba_kind == "standard":
        return "standard"
    if module.project_item_kind in ("class", "document"):
        return module.project_item_kind
    return "form"


def _model(module: ModuleDraft) -> _Module:
    body = module.body
    toks = _lex(body)
    items = _parse_items(body, toks)
    by_name: dict[str, list[_Item]] = {}
    for item in items:
        for key in item.names:
            by_name.setdefault(key, []).append(item)
    return _Module(
        name=module.current_name,
        kind=_module_kind(module),
        draft=module,
        body=body,
        toks=toks,
        starts=[t.start for t in toks],
        items=items,
        by_name=by_name,
        type_members=_type_member_starts(body, toks),
    )


@dataclass(frozen=True)
class _Ref:
    module: _Module
    item: _Item | None  # None: the module itself
    name: str  # casefolded symbol (or module) name


class _Project:
    def __init__(self, draft: DocumentDraft) -> None:
        self.draft = draft
        self.word = draft.baseline.project_name.casefold()
        self.modules: dict[str, _Module] = {
            m.current_name.casefold(): _model(m) for m in draft.modules if not m.is_deleted
        }
        self.public: dict[str, list[tuple[_Module, _Item]]] = {}
        for mod in self.modules.values():
            if mod.kind != "standard":
                continue
            for item in mod.items:
                for key, name in item.names.items():
                    if name.public:
                        self.public.setdefault(key, []).append((mod, item))

    def resolve(self, mod: _Module, k: int) -> list[_Ref]:
        """What identifier token ``k`` of ``mod`` refers to inside this project."""
        toks = mod.toks
        tok = toks[k]
        key = tok.text.casefold()
        if tok.start in mod.type_members:
            return []
        prev = toks[k - 1] if k > 0 else None
        if _is_op(prev, ".") or _is_op(prev, "!"):
            qual = toks[k - 2] if k > 1 else None
            if qual is None or qual.kind != "id":
                return []  # With-block member, or a member of an expression
            if _is_op(toks[k - 3] if k > 2 else None, "."):
                project = toks[k - 4] if k > 3 else None
                if _word(project) != self.word:
                    return []  # obj.Prop.Member
            qmod = self.modules.get(qual.text.casefold())
            if qmod is not None:
                if qmod.kind in ("standard", "document"):
                    return [_Ref(qmod, it, key) for it in qmod.by_name.get(key, [])]
                return []  # a class or form member: the module itself was resolved
            for ref in self.resolve(mod, k - 2):  # Enum.Member
                item = ref.item
                if item is not None and item.names[ref.name].what == "enum" and key in item.names:
                    return [_Ref(ref.module, item, key)]
            return []
        nxt = toks[k + 1] if k + 1 < len(toks) else None
        if _is_op(nxt, ":="):
            return []  # named argument
        own = mod.by_name.get(key)
        if own:
            return [_Ref(mod, it, key) for it in own]
        if key in self.public:
            return [_Ref(m, it, key) for m, it in self.public[key] if m is not mod]
        other = self.modules.get(key)
        if other is not None:
            return [_Ref(other, None, key)]
        return []


def _normalized(text: str) -> str:
    lines = [line.rstrip() for line in text.replace("\r", "").split("\n")]
    return "\n".join(line for line in lines if line).strip()


def _item_code(mod: _Module, item: _Item) -> str:
    return _normalized(mod.body[item.code_start : item.end])


# -- picking a control ---------------------------------------------------------------------


@dataclass(frozen=True)
class PortableControl:
    part_path: str
    start: int  # offset of the element in the source part
    element: str  # local name
    control_id: str
    label: str
    display: str  # Tab › Group › Control
    callbacks: tuple[str, ...]

    def describe(self) -> str:
        text = f"{self.display}   ({self.element})"
        if self.callbacks:
            text += "   → " + ", ".join(dict.fromkeys(self.callbacks))
        return text


def _local(name: str) -> str:
    return name.rsplit(":", 1)[-1]


def _subtree(elements: list[rs.RibbonElement], top: rs.RibbonElement) -> list[rs.RibbonElement]:
    return [e for e in elements if e.complete and top.start <= e.start and e.end <= top.end]


def _callback_attrs(element: rs.RibbonElement) -> list[tuple[str, str]]:
    return [
        (name, value)
        for name, value, _offset in element.attrs
        if rs.is_callback_attribute(name) and value.strip()
    ]


def portable_controls(draft: DocumentDraft) -> list[PortableControl]:
    """Custom ribbon controls (and groups) of ``draft`` that run VBA."""
    found: list[PortableControl] = []
    for part in rs.ribbon_parts(draft):
        text = part.text or ""
        elements = rs.element_tree(text)
        for e in elements:
            parent = elements[e.parent] if e.parent >= 0 else None
            if parent is None or not e.complete or not (e.attr("id") or e.attr("idQ")):
                continue
            local, parent_local = _local(e.name), _local(parent.name)
            is_group = local == "group" and parent_local == "tab"
            if not is_group and (local not in CONTROL_ELEMENTS or parent_local not in _CONTROL_PARENTS):
                continue
            callbacks = tuple(
                rs.split_callback_name(rs.unescape(value))[1]
                for sub in _subtree(elements, e)
                for _name, value in _callback_attrs(sub)
            )
            if not callbacks:
                continue
            found.append(
                PortableControl(
                    part_path=part.path,
                    start=e.start,
                    element=local,
                    control_id=rs.unescape(e.attr("id") or e.attr("idQ") or ""),
                    label=rs.unescape(e.attr("label") or ""),
                    display=rs.element_path(elements, e),
                    callbacks=callbacks,
                )
            )
    return found


@dataclass(frozen=True)
class PortTarget:
    container: rs.RibbonContainer
    new_group: bool = False  # container is a tab: the control goes in a new group there

    @property
    def element(self) -> str:
        """The element the control will be a direct child of."""
        return "group" if self.new_group else self.container.element

    def describe(self, show_part: bool = False) -> str:
        where = self.container.display
        if self.new_group:
            text = f"New group on tab {where}"
        elif self.container.element == "tab":
            text = f"Tab {where}"
        elif self.container.element == "group":
            text = where
        else:
            text = f"{where}   ({self.container.element})"
        if show_part:
            text += f"   [{self.container.part_path.rsplit('/', 1)[-1]}]"
        return text


def _accepts(container: str, element: str) -> bool:
    if container == "tab":
        return element == "group"
    if element == "group":
        return False
    if container == "menu":
        return element in rs.MENU_CHILD_ELEMENTS
    if container == "buttonGroup":
        return element in _BUTTON_GROUP_CHILDREN
    return True


def port_targets(draft: DocumentDraft, control: PortableControl) -> list[PortTarget]:
    """Where ``control`` can go in ``draft``'s ribbon, in document order."""
    targets: list[PortTarget] = []
    for part in rs.ribbon_parts(draft):
        if not part.editable:
            continue
        text = part.text or ""
        tabs = rs.find_tabs(text, part.path)
        if control.element == "group":
            targets.extend(PortTarget(tab) for tab in tabs)
            continue
        targets.extend(
            PortTarget(c)
            for c in rs.find_containers(text, part.path)
            if _accepts(c.element, control.element)
        )
        targets.extend(PortTarget(tab, new_group=True) for tab in tabs)
    return targets


def default_suffix(target: DocumentDraft) -> str:
    stem = re.sub(r"[^A-Za-z0-9]", "", Path(target.baseline.path).stem)[:20]
    return "_" + (stem or "Ported")


# -- the plan ---------------------------------------------------------------------------------


@dataclass
class PortPlan:
    draft: DocumentDraft | None  # the target with the port applied; None when blocked
    problems: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    summary: list[str] = field(default_factory=list)
    renames: list[tuple[str, str, str]] = field(default_factory=list)  # (what, old, new)
    new_module_ids: list[str] = field(default_factory=list)
    changed_module_ids: list[str] = field(default_factory=list)
    part_path: str = ""
    element_offset: int = 0
    element_length: int = 0

    @property
    def ok(self) -> bool:
        return self.draft is not None and not self.problems

    def describe(self) -> str:
        lines: list[str] = []
        if self.problems:
            lines.append("Cannot import:")
            lines += [f"  • {p}" for p in self.problems]
            lines.append("")
        if self.summary:
            lines.append("Adds:")
            lines += [f"  • {s}" for s in self.summary]
            lines.append("")
        if self.renames:
            lines.append("Renamed so nothing clashes with this add-in or the one it came from:")
            lines += [f"  • {what} {old} → {new}" for what, old, new in self.renames]
            lines.append("")
        if self.warnings:
            lines.append("Check after importing:")
            lines += [f"  • {w}" for w in self.warnings]
        return "\n".join(lines).rstrip()


def _blocked(*problems: str) -> PortPlan:
    return PortPlan(draft=None, problems=list(problems))


def _package_member(draft: DocumentDraft, member: str) -> bytes | None:
    data = draft.baseline.package_bytes
    try:
        with open_package(data if data is not None else draft.baseline.path) as package:
            return package.read(member)
    except (KeyError, OSError, ValueError):
        return None


class _Namer:
    """Hands out names that are unique, case-insensitively, against everything taken."""

    def __init__(self, taken: set[str], suffix: str) -> None:
        self.taken = {t.casefold() for t in taken}
        self.suffix = suffix

    def take(self, name: str) -> None:
        self.taken.add(name.casefold())

    def fresh(
        self,
        name: str,
        *,
        force: bool = False,
        avoid: frozenset[str] | set[str] = frozenset(),
        limit: int = MAX_IDENTIFIER,
    ) -> str:
        """``name`` itself when free (and not ``force``), else ``name<suffix>[n]``."""

        def free(candidate: str) -> bool:
            key = candidate.casefold()
            return key not in self.taken and key not in avoid and key not in VBA_RESERVED

        if not force and free(name):
            self.take(name)
            return name
        base = (name + self.suffix)[:limit]
        candidate, n = base, 2
        while not free(candidate) or candidate.casefold() == name.casefold():
            tail = str(n)
            candidate, n = base[: limit - len(tail)] + tail, n + 1
        self.take(candidate)
        return candidate


_reference_cache: dict[tuple[str, int], list[tuple[str, str]]] = {}


def _references(draft: DocumentDraft) -> list[tuple[str, str]]:
    from vba_addin_editor.adapters.pyopenvba_adapter import PyOpenVBAAdapter

    fp = draft.baseline.file_fingerprint
    key = (fp.sha256, fp.size)
    if key not in _reference_cache:
        _reference_cache[key] = PyOpenVBAAdapter().project_references(
            draft.baseline.path, draft.baseline.package_bytes
        )
    return _reference_cache[key]


def plan_port(
    source: DocumentDraft,
    target: DocumentDraft,
    control: PortableControl,
    where: PortTarget,
    *,
    suffix: str,
    group_label: str = "",
    check_references: bool = True,
) -> PortPlan:
    """Work out the whole port on a copy of ``target``; ``target`` is not changed."""
    if Path(source.baseline.path).resolve() == Path(target.baseline.path).resolve():
        return _blocked("Choose a different add-in: this is the file that is open.")
    if not _SUFFIX_RE.match(suffix):
        return _blocked("The name suffix must be 1-24 letters, digits or underscores.")
    if source.baseline.safety.password_protected:
        return _blocked(
            f"{source.baseline.path.name}'s VBA project is password-protected; unlock it in "
            "the VBA editor and save it first."
        )
    if target.baseline.safety.password_protected:
        return _blocked("This add-in's VBA project is password-protected and read-only.")
    if target.baseline.package_safety.opc_signature_present:
        return _blocked(
            "This file has a package digital signature; editing its ribbon would break it."
        )
    src_part = source.xml_part_by_path(control.part_path)
    tgt_part = target.xml_part_by_path(where.container.part_path)
    if src_part is None or src_part.text is None:
        return _blocked(f"{control.part_path} could not be read in {source.baseline.path.name}.")
    if tgt_part is None or tgt_part.text is None or not tgt_part.editable:
        return _blocked(f"{where.container.part_path} is not editable in this add-in.")
    if not _accepts(where.element, control.element):
        return _blocked(f"A {control.element} cannot go in a {where.element}.")
    try:
        return _Porter(source, target, control, where, suffix, group_label, check_references).run()
    except PortError as exc:
        return _blocked(str(exc))


class _Porter:
    def __init__(
        self,
        source: DocumentDraft,
        target: DocumentDraft,
        control: PortableControl,
        where: PortTarget,
        suffix: str,
        group_label: str,
        check_references: bool,
    ) -> None:
        self.source = source
        self.target = target
        self.control = control
        self.where = where
        self.suffix = suffix
        self.group_label = group_label
        self.check_references = check_references
        self.src = _Project(source)
        self.tgt = _Project(target)
        self.plan = PortPlan(draft=None)
        self.src_file = source.baseline.path.name
        self.tgt_file = target.baseline.path.name
        self.src_ids = {
            value.casefold()
            for part in rs.ribbon_parts(source)
            for value, _offset in rs.control_ids(part.text or "")
        }
        # Filled while collecting.
        self.selected: dict[str, set[int]] = {}  # source module -> item indexes to copy
        self.whole: set[str] = set()  # class modules copied whole
        self.reused: dict[tuple[str, int], str] = {}  # (module, item) -> target module
        self.reused_classes: set[str] = set()
        self.forced: set[tuple[str, str]] = set()  # (module, name) that must get a new name
        self.macro_strings: dict[tuple[str, int], tuple[_Module, str, str | None, str | None]] = {}
        self.id_strings: set[tuple[str, int]] = set()
        self.loose_strings: set[tuple[str, str]] = set()
        self.notes: list[str] = []
        # Filled while naming / writing.
        self.module_names: dict[str, str] = {}
        self.symbol_names: dict[tuple[str, str], str] = {}
        self.id_names: dict[str, str] = {}  # raw source id/idQ -> new id
        self.missing_ids: set[str] = set()
        self.new_modules: dict[str, ModuleDraft] = {}

    def note(self, text: str) -> None:
        if text not in self.notes:
            self.notes.append(text)

    # -- 1. the XML subtree and its callbacks ----------------------------------------------

    def run(self) -> PortPlan:
        part = self.source.xml_part_by_path(self.control.part_path)
        text = part.text or "" if part is not None else ""
        elements = rs.element_tree(text)
        top = next((e for e in elements if e.start == self.control.start and e.complete), None)
        if top is None:
            raise PortError("The chosen control is no longer in the source ribbon.")
        subtree = _subtree(elements, top)
        callback_refs = self._callbacks(subtree)
        roots = [
            (mod, item)
            for ref in callback_refs.values()
            if ref is not None
            for mod, key in [ref]
            for item in mod.by_name.get(key, [])
        ]
        self.forced |= {(mod.name.casefold(), key) for mod, key in filter(None, callback_refs.values())}
        self._collect(roots)
        self._decide_names(subtree)
        draft = clone_draft(self.target)
        self._write_modules(draft)
        self._write_ribbon(draft, text, elements, top, subtree, callback_refs)
        self._hook_ribbon_vars(draft)
        self._final_checks(draft)
        self.plan.draft = None if self.plan.problems else draft
        return self.plan

    def _callbacks(self, subtree) -> dict[str, tuple[_Module, str] | None]:
        """Raw callback attribute value -> (source module, procedure key); None if missing."""
        index = rs.procedure_index(self.source)
        refs: dict[str, tuple[_Module, str] | None] = {}
        for e in subtree:
            for attr, raw in _callback_attrs(e):
                if raw in refs:
                    continue
                value = rs.unescape(raw).strip()
                resolution = rs.resolve(index, value)
                if resolution.status == rs.STATUS_MISSING:
                    refs[raw] = None
                    self.note(f'{attr}="{value}" has no macro in {self.src_file}; copied as it is.')
                    continue
                if resolution.status != rs.STATUS_OK or len(resolution.matches) != 1:
                    raise PortError(f'{attr}="{value}": {resolution.message}')
                match = resolution.matches[0]
                mod = self.src.modules[match.module_name.casefold()]
                if mod.kind != "standard":
                    raise PortError(
                        f'{attr}="{value}" runs a macro in {mod.name}, which is not a standard '
                        "module. Move the macro to a standard module first."
                    )
                refs[raw] = (mod, match.name.casefold())
        return refs

    # -- 2. what VBA to copy ------------------------------------------------------------------

    def _collect(self, roots: list[tuple[_Module, _Item]]) -> None:
        queue: list[tuple[_Module, _Item | None]] = []
        root_keys = {(m.name.casefold(), i.index) for m, i in roots}

        def add_item(mod: _Module, item: _Item) -> None:
            if mod.kind == "class":
                add_class(mod)
                return
            if mod.kind != "standard":
                return
            key = (mod.name.casefold(), item.index)
            chosen = self.selected.setdefault(key[0], set())
            if item.index in chosen or key in self.reused:
                return
            if key not in root_keys:
                reuse = self._reusable(mod, item)
                if reuse is not None:
                    self.reused[key] = reuse
                    return
            chosen.add(item.index)
            queue.append((mod, item))

        def add_class(mod: _Module) -> None:
            key = mod.name.casefold()
            if key in self.whole or key in self.reused_classes:
                return
            existing = self.tgt.modules.get(key)
            if (
                existing is not None
                and existing.kind == "class"
                and _normalized(existing.body) == _normalized(mod.body)
            ):
                self.reused_classes.add(key)
                return
            self.whole.add(key)
            queue.append((mod, None))

        for mod, item in roots:
            add_item(mod, item)
        while queue:
            mod, item = queue.pop(0)
            span = range(len(mod.toks)) if item is None else mod.token_range(item.start, item.end)
            for k in span:
                tok = mod.toks[k]
                if tok.kind == "id":
                    self._sheet_check(mod, k)
                    for ref in self.src.resolve(mod, k):
                        self._follow(ref, add_item, add_class)
                elif tok.kind == "str":
                    self._string_ref(mod, k, add_item)
        # Option / Def* / #Const lines of each module code comes from; drop modules
        # that turned out to contribute nothing (all their code was reused).
        for key in list(self.selected):
            mod = self.src.modules[key]
            if not self.selected[key]:
                del self.selected[key]
                continue
            self.selected[key] |= {it.index for it in mod.items if it.kind == "option"}

    def _follow(self, ref: _Ref, add_item, add_class) -> None:
        target = ref.module
        if ref.item is not None:
            if target.kind == "standard":
                add_item(target, ref.item)
            elif target.kind == "document":
                self.note(
                    f"The code calls {target.name}.{ref.item.names[ref.name].name}, which is "
                    f"in {self.src_file}'s {target.name} and was not copied."
                )
            return
        if target.kind == "class":
            add_class(target)
        elif target.kind == "form":
            if target.name.casefold() in self.tgt.modules:
                self.note(f"The code uses UserForm {target.name}, which this add-in already has.")
            else:
                self.plan.problems.append(
                    f"The code uses UserForm {target.name}, and this editor cannot copy "
                    f"UserForm layouts. Export {target.name} from {self.src_file} in the VBA "
                    "editor (File → Export File) and import it into this add-in first."
                )
        elif target.kind == "document" and target.name.casefold() not in self.tgt.modules:
            self.note(
                f"The code uses {target.name} (a sheet of {self.src_file}), which this add-in "
                "does not have."
            )

    def _sheet_check(self, mod: _Module, k: int) -> None:
        """ThisWorkbook.Worksheets(...) means the source add-in's own sheets."""
        if k < 2 or _word(mod.toks[k]) not in ("worksheets", "sheets", "names"):
            return
        if _is_op(mod.toks[k - 1], ".") and _word(mod.toks[k - 2]) == "thisworkbook":
            self.note(
                f"{mod.name} reads ThisWorkbook.{mod.toks[k].text}: the sheets or names of "
                f"{self.src_file} are not copied, so this add-in must have its own."
            )

    def _statement_words(self, mod: _Module, k: int) -> set[str]:
        words: set[str] = set()
        j = k - 1
        while j >= 0 and mod.toks[j].kind != "nl" and not _is_op(mod.toks[j], ":"):
            if mod.toks[j].kind == "id":
                words.add(mod.toks[j].text.casefold())
            j -= 1
        return words

    def _string_ref(self, mod: _Module, k: int, add_item) -> None:
        tok = mod.toks[k]
        value = tok.text
        if not value:
            return
        if value.casefold() in self.src_ids:
            self.id_strings.add((mod.name, tok.start))
            return
        match = _MACRO_STRING_RE.match(value)
        book = match.group("book") if match else None
        runs = match is not None and (
            book is not None or bool(self._statement_words(mod, k) & _MACRO_RUNNERS)
        )
        if runs and match is not None and self._macro_string(mod, tok, match, add_item):
            return
        if self.src_file.casefold() in value.casefold():
            self.note(f'{mod.name} has the string "{value}", which still names {self.src_file}.')
        elif match is not None and match.group("mod") is None and book is None:
            key = value.casefold()
            candidates = mod.by_name.get(key, []) + [it for _m, it in self.src.public.get(key, [])]
            if any(it.names[key].what == "proc" for it in candidates):
                self.loose_strings.add((mod.name, value))

    def _macro_string(self, mod: _Module, tok: _Tok, match: re.Match, add_item) -> bool:
        """A string that runs a source macro by name; True when it was recognised."""
        proc_key = match.group("proc").casefold()
        mod_name = match.group("mod")
        owner: _Module | None = mod
        if mod_name is not None:
            owner = self.src.modules.get(mod_name.casefold())
            items = owner.by_name.get(proc_key, []) if owner is not None else []
        else:
            items = mod.by_name.get(proc_key, [])
            if not items:
                found = self.src.public.get(proc_key, [])
                owner = found[0][0] if found else None
                items = [it for m, it in found if m is owner]
        items = [it for it in items if it.names[proc_key].what == "proc"]
        if owner is None or not items or owner.kind != "standard":
            return False
        for item in items:
            add_item(owner, item)
        self.forced.add((owner.name.casefold(), proc_key))
        self.macro_strings[(mod.name, tok.start)] = (owner, proc_key, match.group("book"), mod_name)
        return True

    def _reusable(self, mod: _Module, item: _Item) -> str | None:
        """Target module holding an identical public copy of ``item``, if any."""
        if item.kind == "option" or not item.names or not all(n.public for n in item.names.values()):
            return None
        code = _item_code(mod, item)
        key = next(iter(item.names))
        for tmod, titem in self.tgt.public.get(key, []):
            if set(titem.names) == set(item.names) and _item_code(tmod, titem) == code:
                return tmod.name
        return None

    # -- 3. names -----------------------------------------------------------------------------------

    def _decide_names(self, subtree: list[rs.RibbonElement]) -> None:
        tgt_public = set(self.tgt.public)
        tgt_modules = set(self.tgt.modules)
        tgt_callbacks = {
            rs.split_callback_name(entry.callback.name)[1].casefold()
            for entry in rs.ribbon_entries(self.target)
        }
        modules = _Namer(tgt_modules | tgt_public | tgt_callbacks, self.suffix)
        for key in [*self.selected, *sorted(self.whole)]:
            mod = self.src.modules[key]
            new = modules.fresh(mod.name, limit=MAX_MODULE_NAME)
            for _attempt in range(5):
                if validate_module_name(new, self.target) is None:
                    break
                new = modules.fresh(mod.name, force=True, limit=MAX_MODULE_NAME)
            self.module_names[key] = new
            if new != mod.name:
                self.plan.renames.append(("module", mod.name, new))

        # A renamed symbol must not equal anything in the target, any new module, or
        # any name the copied code declares; a forced one not anything in the source either.
        copied = {
            nk for key, chosen in self.selected.items()
            for idx in chosen for nk in self.src.modules[key].items[idx].names
        }
        symbols = _Namer(
            tgt_public | tgt_modules | tgt_callbacks | set(self.module_names.values()), self.suffix
        )
        source_names = set(self.src.public) | set(self.src.modules)
        rows = [
            (key, name_key, name)
            for key, chosen in self.selected.items()
            for idx in sorted(chosen)
            for name_key, name in self.src.modules[key].items[idx].names.items()
        ]
        # Forced names (callbacks, run by name) first, so they get the plain suffix.
        rows.sort(key=lambda row: (row[0], row[1]) not in self.forced)
        for key, name_key, name in rows:
            if (key, name_key) in self.symbol_names:
                continue
            force = (key, name_key) in self.forced
            if not force and not name.public:
                continue
            avoid = (copied | source_names) if force else (copied - {name_key})
            new = symbols.fresh(name.name, force=force, avoid=avoid)
            if new != name.name:
                self.symbol_names[(key, name_key)] = new
                self.plan.renames.append((_WHAT_LABELS.get(name.what, name.what), name.name, new))

        # Ribbon control ids (custom ids are per file, so only clashes are renamed).
        tgt_ids = {
            value
            for part in rs.ribbon_parts(self.target)
            for value, _offset in rs.control_ids(part.text or "")
        }
        ids = _Namer(tgt_ids, self.suffix)
        for e in subtree:
            for attr in ("id", "idQ"):
                raw = e.attr(attr)
                if raw is None or raw in self.id_names:
                    continue
                old = rs.unescape(raw)
                new = ids.fresh(old.rsplit(":", 1)[-1])
                self.id_names[raw] = new
                if new != old:
                    self.plan.renames.append(("control id", old, new))
        if self.where.new_group:
            stem = re.sub(r"[^A-Za-z0-9_]", "", self.control.control_id) or "Imported"
            self.group_id = ids.fresh("grp" + stem[0].upper() + stem[1:])

    # -- 4. the VBA text ---------------------------------------------------------------------------

    def _rewrite(self, mod: _Module, start: int, end: int) -> str:
        """``mod.body[start:end]`` with every renamed reference rewritten."""
        edits: list[tuple[int, int, str]] = []
        for k in mod.token_range(start, end):
            tok = mod.toks[k]
            if tok.kind == "id":
                new = self._new_identifier(mod, k)
                if new is not None and new != tok.text:
                    edits.append((tok.start, tok.end, new))
            elif tok.kind == "str":
                new = self._new_string(mod, tok)
                if new is not None:
                    edits.append((tok.start, tok.end, '"' + new.replace('"', '""') + '"'))
        text = mod.body[start:end]
        for s, e, new in sorted(edits, reverse=True):
            text = text[: s - start] + new + text[e - start :]
        return text

    def _new_identifier(self, mod: _Module, k: int) -> str | None:
        toks = mod.toks
        nxt = toks[k + 1] if k + 1 < len(toks) else None
        if (
            _word(toks[k]) == self.src.word
            and _is_op(nxt, ".")
            and k + 2 < len(toks)
            and toks[k + 2].text.casefold() in self.src.modules
        ):
            return self.target.baseline.project_name  # Project.Module.Member
        refs = self.src.resolve(mod, k)
        if not refs:
            return None
        ref = refs[0]
        if ref.item is not None:
            return self.symbol_names.get((ref.module.name.casefold(), ref.name))
        if _is_op(nxt, ".") and k + 2 < len(toks):
            # Module.Member, where Member is reused from a target module.
            for member in self.src.resolve(mod, k + 2):
                if member.item is not None:
                    reused = self.reused.get((member.module.name.casefold(), member.item.index))
                    if reused is not None:
                        return reused
        return self.module_names.get(ref.module.name.casefold())

    def _new_string(self, mod: _Module, tok: _Tok) -> str | None:
        macro = self.macro_strings.get((mod.name, tok.start))
        if macro is not None:
            owner, proc_key, book, mod_name = macro
            match = _MACRO_STRING_RE.match(tok.text)
            assert match is not None
            text = self.symbol_names.get((owner.name.casefold(), proc_key), match.group("proc"))
            if mod_name is not None:
                text = f"{self.module_names.get(owner.name.casefold(), mod_name)}.{text}"
            if book is not None:
                if book.strip("'").casefold() == self.src_file.casefold():
                    quoted = book.startswith("'") or " " in self.tgt_file
                    book = f"'{self.tgt_file}'" if quoted else self.tgt_file
                text = f"{book}!{text}"
            elif tok.text.startswith("!"):
                text = "!" + text
            return text if text != tok.text else None
        if (mod.name, tok.start) in self.id_strings:
            raw = next(
                (r for r in self.id_names if rs.unescape(r).casefold() == tok.text.casefold()), None
            )
            if raw is None:
                self.missing_ids.add(tok.text)
                return None
            new = self.id_names[raw]
            return new if new != tok.text else None
        return None

    def _module_body(self, mod: _Module, chosen: set[int]) -> str:
        items = [mod.items[i] for i in sorted(chosen)]
        options = [it for it in items if it.kind == "option"]
        decls = [it for it in items if it.kind == "decl"]
        procs = [it for it in items if it.kind == "proc"]
        head = [f"' Imported from {self.src_file} (module {mod.name}) by VBA Add-in Editor."]
        head += [self._rewrite(mod, it.start, it.end).strip("\n") for it in options]
        blocks = ["\n".join(head)]
        if decls:
            texts = [self._rewrite(mod, it.start, it.end).strip("\n") for it in decls]
            joined = texts[0]
            for before, text in pairwise(texts):
                # One-line declarations stay together; blocks get a blank line around them.
                joined += ("\n\n" if "\n" in before or "\n" in text else "\n") + text
            blocks.append(joined)
        blocks += [self._rewrite(mod, it.start, it.end).strip("\n") for it in procs]
        return "\n\n".join(blocks) + "\n"

    def _write_modules(self, draft: DocumentDraft) -> None:
        for key, chosen in self.selected.items():
            mod = self.src.modules[key]
            new = _new_module(self.module_names[key], self._module_body(mod, chosen), standard=True)
            draft.modules.append(new)
            self.new_modules[key] = new
            procs = sum(1 for i in chosen if mod.items[i].kind == "proc")
            decls = sum(1 for i in chosen if mod.items[i].kind == "decl")
            what = f"{procs} procedure(s)" + (f", {decls} declaration(s)" if decls else "")
            self.plan.summary.append(f"Module {new.current_name}: {what} from {mod.name}")
        for key in sorted(self.whole):
            mod = self.src.modules[key]
            name = self.module_names[key]
            new = _new_module(name, self._rewrite(mod, 0, len(mod.body)).lstrip("\n"), standard=False)
            if mod.draft.hidden_header:
                new.hidden_header = replace_vb_name(mod.draft.hidden_header, name)
            draft.modules.append(new)
            self.new_modules[key] = new
            self.plan.summary.append(f"Class module {name}: copied whole from {mod.name}")
        self.plan.new_module_ids = [m.id for m in self.new_modules.values()]
        reused = sorted(
            {
                f"{', '.join(n.name for n in self.src.modules[k].items[i].names.values())} ({v})"
                for (k, i), v in self.reused.items()
            }
        )
        if reused:
            self.plan.summary.append("Reuses identical code already here: " + "; ".join(reused))
        if self.reused_classes:
            self.plan.summary.append(
                "Reuses identical class module(s) already here: "
                + ", ".join(self.src.modules[k].name for k in sorted(self.reused_classes))
            )
        renamed_procs = {nk for (_k, nk), _new in self.symbol_names.items()}
        for mod_name, value in sorted(self.loose_strings):
            if value.casefold() in renamed_procs:
                self.note(
                    f'{mod_name} has the string "{value}", the old name of a renamed macro; if it '
                    "is used to run that macro, change it to the new name."
                )
        for missing in sorted(self.missing_ids):
            self.note(f'The code refers to ribbon control "{missing}", which was not imported.')

    # -- 5. the ribbon XML ----------------------------------------------------------------------------

    def _callback_value(self, raw: str, ref: tuple[_Module, str] | None) -> str:
        value = rs.unescape(raw).strip()
        if ref is None:
            return value
        mod, key = ref
        module_name, proc = rs.split_callback_name(value)
        new_proc = self.symbol_names.get((mod.name.casefold(), key), proc)
        if module_name is not None:
            return f"{self.module_names.get(mod.name.casefold(), module_name)}.{new_proc}"
        return new_proc

    def _copy_images(self, draft, subtree, src_root) -> dict[str, str]:
        """Copy each image the controls show; returns source rId -> target rId."""
        images: dict[str, str] = {}
        rels_part = self.source.xml_part_by_path(rs.relationships_path(self.control.part_path))
        rels = {
            rid: target
            for rid, kind, target in rs.relationships((rels_part.text or "") if rels_part else "")
            if kind in rs.IMAGE_REL_TYPES
        }
        for e in subtree:
            raw = e.attr("image")
            if not raw or raw in images:
                continue
            target = rels.get(rs.unescape(raw))
            member = rs.relationship_target_path(self.control.part_path, target) if target else None
            data = _package_member(self.source, member) if member else None
            if data is None or member is None:
                loader = src_root.attr("loadImage") if src_root is not None else None
                if loader:
                    self.note(
                        f'image="{rs.unescape(raw)}" is supplied by the loadImage callback '
                        f"{rs.unescape(loader)}, which was not imported."
                    )
                else:
                    self.note(f'The image "{rs.unescape(raw)}" is missing from {self.src_file}.')
                continue
            stem = self.id_names.get(e.attr("id") or e.attr("idQ") or "", "icon")
            try:
                path, rid = icon_service.add_image_member(
                    draft, self.where.container.part_path, data, extension=Path(member).suffix, stem=stem
                )
            except icon_service.IconError as exc:
                raise PortError(str(exc)) from exc
            images[raw] = rid
            self.plan.summary.append(f"Icon {path}")
        return images

    def _write_ribbon(self, draft, text, elements, top, subtree, callback_refs) -> None:
        container = self.where.container
        part = draft.xml_part_by_path(container.part_path)
        assert part is not None and part.text is not None
        src_root, tgt_root = rs.root_element(text), rs.root_element(part.text)
        src_prefix = rs.element_prefix(src_root.name) if src_root else ""
        tgt_prefix = rs.element_prefix(tgt_root.name) if tgt_root else ""
        images = self._copy_images(draft, subtree, src_root)

        # Attributes the top element may not carry where it now goes.
        drop_top = set(_POSITION_ATTRS)
        if self.where.element != "group":
            drop_top |= {"size", "getSize"}  # only controls directly in a group have a size
        if self.where.element != "menu":
            drop_top |= {"description", "getDescription"}  # menu items only
        dropped = sorted(name for name, _v, _o in top.attrs if name in drop_top)

        # Rewrite each copied opening tag, last first so offsets hold.
        sub_text = text[top.start : top.end]
        for e in sorted(subtree, key=lambda x: x.start, reverse=True):
            tag = text[e.start : e.open_end]
            if e is top:
                tag = rs.remove_tag_attrs(tag, drop_top)
            idq = e.attr("idQ")
            if idq is not None:
                tag = rs.set_tag_attr(rs.remove_tag_attrs(tag, {"idQ"}), "id", self.id_names[idq])
            for name, raw, _offset in e.attrs:
                if name == "id" and self.id_names.get(raw, rs.unescape(raw)) != rs.unescape(raw):
                    tag = rs.set_tag_attr(tag, "id", self.id_names[raw])
                elif name == "image" and raw in images:
                    tag = rs.set_tag_attr(tag, "image", images[raw])
                elif raw in callback_refs and rs.is_callback_attribute(name):
                    new = self._callback_value(raw, callback_refs[raw])
                    if new != rs.unescape(raw).strip():
                        tag = rs.set_tag_attr(tag, name, new)
            a, b = e.start - top.start, e.open_end - top.start
            sub_text = sub_text[:a] + tag + sub_text[b:]
        if dropped:
            self.note(
                f"Left out {', '.join(dropped)} on the {self.control.element}: not valid in a "
                f"{self.where.element}."
            )
        if any(e.attr("idQ") is not None for e in subtree):
            self.note(
                "The control had a shared (idQ) id; it now has its own id, so the two add-ins "
                "do not compete for it."
            )
        sub_text = _swap_prefix(sub_text, src_prefix, tgt_prefix)

        # Re-indent for the new place.
        src_indent = rs.line_indent(text, top.start)
        child_indent = container.child_indent
        if self.where.new_group:
            unit = container.child_indent[len(container.indent) :] or "  "
            child_indent = container.child_indent + unit
        lines = sub_text.split("\n")
        element_xml = "\n".join(
            [lines[0]]
            + [
                child_indent + (line[len(src_indent) :] if line.startswith(src_indent) else line.lstrip())
                for line in lines[1:]
            ]
        )
        if self.where.new_group:
            label = self.group_label.strip() or self._source_group_label(elements, top)
            p = f"{tgt_prefix}:" if tgt_prefix else ""
            element_xml = (
                f'<{p}group id="{rs.xml_attr(self.group_id)}" label="{rs.xml_attr(label)}">\n'
                f"{child_indent}{element_xml}\n{container.child_indent}</{p}group>"
            )
        new_text, offset, _length = rs.insert_into_container(part.text, container, element_xml)
        part.text = new_text
        self.plan.part_path = part.path
        self.plan.element_offset = offset
        self.plan.element_length = len(element_xml.split("\n", 1)[0])
        what = self.control.label or self.control.control_id
        self.plan.summary.insert(0, f'{self.control.element} "{what}" in {self.where.describe()}')

    def _source_group_label(self, elements, top) -> str:
        current = top
        while current.parent >= 0:
            current = elements[current.parent]
            if _local(current.name) == "group":
                return rs.unescape(current.attr("label") or current.attr("id") or "Imported")
        return "Imported"

    # -- 6. IRibbonUI variables -----------------------------------------------------------------------

    def _hook_ribbon_vars(self, draft: DocumentDraft) -> None:
        """Copied IRibbonUI variables are set by the source's onLoad: set them here too."""
        holders = [
            (key, [v for i in sorted(chosen) for v in self.src.modules[key].items[i].ribbon_vars])
            for key, chosen in self.selected.items()
        ]
        holders = [(key, names) for key, names in holders if names]
        if not holders:
            return
        part = draft.xml_part_by_path(self.where.container.part_path)
        assert part is not None and part.text is not None
        root = rs.root_element(part.text)
        on_load = rs.unescape(root.attr("onLoad") or "").strip() if root else ""
        namer = _Namer(
            set(self.tgt.public)
            | set(self.tgt.modules)
            | set(self.src.public)
            | {n.casefold() for n in self.symbol_names.values()}
            | {n.casefold() for n in self.module_names.values()},
            self.suffix,
        )
        for key, chosen in self.selected.items():
            for idx in chosen:
                for nk in self.src.modules[key].items[idx].names:
                    namer.take(self.symbol_names.get((key, nk), nk))
        param = "loadedRibbon"
        hooks: list[tuple[ModuleDraft, str]] = []
        for key, names in holders:
            module = self.new_modules[key]
            hook = namer.fresh("RibbonLoaded", force=True)
            sets = "".join(
                f"    Set {self.symbol_names.get((key, v.casefold()), v)} = {param}\n" for v in names
            )
            hooks.append((module, hook))
            module.body = module.body.rstrip("\n") + (
                "\n\n' Called from the ribbon's onLoad so the copied ribbon variable is set.\n"
                f"Public Sub {hook}({param} As IRibbonUI)\n{sets}End Sub\n"
            )
        names = ", ".join(h for _m, h in hooks)
        if not on_load:
            first_module, first_hook = hooks[0]
            if len(hooks) > 1:
                header = f"Public Sub {first_hook}({param} As IRibbonUI)\n"
                calls = "".join(f"    {h} {param}\n" for _m, h in hooks[1:])
                first_module.body = first_module.body.replace(header, header + calls, 1)
            assert root is not None
            tag = part.text[root.start : root.open_end]
            new_tag = rs.set_tag_attr(tag, "onLoad", first_hook)
            part.text = part.text[: root.start] + new_tag + part.text[root.open_end :]
            if self.plan.element_offset > root.start:
                self.plan.element_offset += len(new_tag) - len(tag)
            self.plan.summary.append(f'onLoad="{first_hook}" on the ribbon, to set the ribbon variable')
            return
        resolution = rs.resolve(rs.procedure_index(draft), on_load)
        location = resolution.matches[0] if resolution.status == rs.STATUS_OK else None
        module = draft.module_by_id(location.module_id) if location is not None else None
        model = _model(module) if module is not None else None
        item = next(
            (
                it
                for it in (model.items if model else [])
                if location is not None and location.name.casefold() in it.names and it.kind == "proc"
            ),
            None,
        )
        first = _first_parameter(item.signature) if item is not None else ""
        if (
            location is None
            or module is None
            or item is None
            or not first
            or module.pyopenvba_kind != "standard"
        ):
            self.note(
                f'Call {names} from this ribbon\'s onLoad "{on_load}" (pass it the IRibbonUI), '
                "or the copied ribbon variable stays Nothing."
            )
            return
        calls = "".join(f"    {h} {first}\n" for _m, h in hooks)
        module.body = module.body[: item.last_line_start] + calls + module.body[item.last_line_start :]
        self.plan.changed_module_ids.append(module.id)
        self.plan.summary.append(f"{location.name} (the onLoad, in {module.current_name}) now calls {names}")

    # -- 7. checks --------------------------------------------------------------------------------

    def _final_checks(self, draft: DocumentDraft) -> None:
        src_host, tgt_host = self.source.baseline.host_kind, self.target.baseline.host_kind
        if src_host != tgt_host:
            self.note(
                f"{self.src_file} is a {src_host.title()} file and this one is {tgt_host.title()}: "
                f"code that uses {src_host.title()} objects will not run here."
            )
        part_path = self.where.container.part_path
        if part_path.lower().endswith("/customui.xml") and any(
            p.path.lower().endswith("customui14.xml") for p in rs.ribbon_parts(self.target)
        ):
            self.note(
                "This add-in also has customUI14.xml, which Office 2010 and later use instead of "
                "customUI.xml; the control will only show in Office 2007."
            )
        for mod in self.new_modules.values():
            problem = validate_module_name(mod.current_name, draft, exclude_id=mod.id)
            if problem:
                self.plan.problems.append(f"{mod.current_name}: {problem}")
        if self.check_references and self.new_modules:
            target_refs = _references(self.target)
            have = {g.casefold() for _n, g in target_refs} | {n.casefold() for n, _g in target_refs}
            missing = [
                name
                for name, guid in _references(self.source)
                if guid.casefold() not in _STANDARD_REFERENCES
                and guid.casefold() not in have
                and name.casefold() not in have
            ]
            if missing:
                self.note(
                    f"{self.src_file} references {', '.join(missing)} and this add-in does not. If "
                    "the imported code uses it, add it in the VBA editor (Tools → References)."
                )
        self.plan.warnings.extend(self.notes)


def _new_module(name: str, body: str, *, standard: bool) -> ModuleDraft:
    return ModuleDraft(
        id=new_module_id(),
        origin_name=None,
        current_name=name,
        body=body,
        kind=ModuleDisplayKind.STANDARD if standard else ModuleDisplayKind.CLASS,
        pyopenvba_kind="standard" if standard else "other",
        is_new=True,
        is_deleted=False,
        destructive_ops_safe=True,
        can_delete=True,
        can_rename=True,
        project_item_kind="standard" if standard else "class",
    )


def _first_parameter(signature: str) -> str:
    match = re.search(
        r"\(\s*(?:(?:ByVal|ByRef|Optional)\s+)*([A-Za-z][A-Za-z0-9_]*)", signature, re.IGNORECASE
    )
    return match.group(1) if match else ""


def _swap_prefix(xml: str, src: str, tgt: str) -> str:
    """Element names of ``xml`` moved from namespace prefix ``src`` to ``tgt``."""
    if src == tgt:
        return xml

    def fix(m: re.Match) -> str:
        slash, prefix, name = m.group(1), (m.group(2) or "").rstrip(":"), m.group(3)
        if prefix != src:
            return m.group(0)
        return f"<{slash}{tgt + ':' if tgt else ''}{name}"

    return re.sub(r"<(/?)([A-Za-z_][\w.-]*:)?([A-Za-z_][\w.-]*)", fix, xml)


# -- applying ------------------------------------------------------------------------------------


def apply_port(target: DocumentDraft, plan: PortPlan) -> None:
    """Make ``target`` hold what ``plan`` computed (call inside one undo step)."""
    result = plan.draft
    if result is None:
        raise PortError("This import cannot be applied.")
    existing = {m.id: m for m in target.modules}
    for module in result.modules:
        mine = existing.get(module.id)
        if mine is None:
            target.modules.append(module)
        elif mine.body != module.body:
            mine.body = module.body
    for part in result.xml_parts:
        mine = target.xml_part_by_path(part.path)
        if mine is None:
            target.xml_parts.append(part)
        elif mine.text != part.text:
            mine.text = part.text
    for binary in result.binary_parts:
        if target.binary_part_by_path(binary.path) is None:
            target.binary_parts.append(binary)
