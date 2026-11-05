"""Ribbon (customUI) callback linking: XML callbacks <-> VBA procedures.

Pure text logic with no Tk dependency. XML is scanned lexically rather than
parsed so it keeps working while a part is mid-edit and not well-formed.
Offsets are into the LF-normalized part text held by the draft.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from functools import lru_cache
from xml.sax.saxutils import escape

from vba_addin_editor.domain.document import (
    DocumentDraft,
    ModuleDraft,
    XmlPartDraft,
    draft_from_snapshot,
)
from vba_addin_editor.services.search_service import ProcedureInfo, parse_procedures

_RIBBON_PART_RE = re.compile(r"(?:^|/)customui\d*\.xml$", re.IGNORECASE)
_COMMENT_RE = re.compile(r"<!--.*?-->", re.DOTALL)
_TAG_RE = re.compile(
    r"<(?P<close>/?)(?P<name>[A-Za-z_][\w:.-]*)(?P<attrs>(?:[^<>\"']|\"[^\"]*\"|'[^']*')*?)(?P<self>/?)>"
)
_ATTR_RE = re.compile(r"(?P<name>[A-Za-z_][\w:.-]*)\s*=\s*(?:\"(?P<dq>[^\"]*)\"|'(?P<sq>[^']*)')")
_CALLBACK_ATTR_RE = re.compile(r"on[A-Z]\w*|get[A-Z]\w*|loadImage")
_CONTAINER_ELEMENTS = ("group", "menu", "box", "buttonGroup")
_PROC_NAME_RE = re.compile(r"^[A-Za-z][A-Za-z0-9_]*$")

STATUS_OK = "ok"
STATUS_MISSING = "missing"
STATUS_NOT_CALLABLE = "not_callable"
STATUS_AMBIGUOUS = "ambiguous"

STATUS_LABELS = {
    STATUS_OK: "OK",
    STATUS_MISSING: "Missing",
    STATUS_NOT_CALLABLE: "Not callable",
    STATUS_AMBIGUOUS: "Ambiguous",
}


@dataclass(frozen=True)
class RibbonCallback:
    part_path: str
    attribute: str  # onAction, getLabel, ...
    name: str  # attribute value as written
    element: str  # button, toggleButton, ...
    control_id: str
    label: str
    offset: int  # start of the attribute value
    length: int
    line: int
    column: int  # 1-based
    tag_start: int
    tag_end: int

    @property
    def procedure_name(self) -> str:
        return split_callback_name(self.name)[1]


@dataclass(frozen=True)
class ProcedureLocation:
    module_id: str
    module_name: str
    project_item_kind: str
    is_standard: bool
    name: str
    kind: str
    line: int
    signature: str


@dataclass(frozen=True)
class Resolution:
    status: str
    matches: tuple[ProcedureLocation, ...]
    message: str | None = None


@dataclass(frozen=True)
class RibbonContainer:
    part_path: str
    element: str
    control_id: str
    display: str  # "Tab › Group › Menu"
    close_start: int
    indent: str
    child_indent: str
    open_start: int = 0


@dataclass(frozen=True)
class RibbonIssue:
    part_path: str
    line: int
    message: str
    offset: int = 0
    length: int = 0

    @property
    def key(self) -> tuple[str, str]:
        return self.part_path, self.message

    def describe(self) -> str:
        return f"{self.part_path} line {self.line}: {self.message}"


@dataclass(frozen=True)
class _Tag:
    name: str
    start: int
    end: int
    closing: bool
    self_closing: bool
    attrs: dict[str, tuple[str, int]]  # name -> (value, value offset)


# -- parts -------------------------------------------------------------------


def is_ribbon_part(path: str) -> bool:
    return bool(_RIBBON_PART_RE.search(path))


def ribbon_parts(draft: DocumentDraft) -> list[XmlPartDraft]:
    return [
        part
        for part in draft.xml_parts
        if is_ribbon_part(part.path) and part.text is not None
    ]


# -- lexical XML scan ----------------------------------------------------------


def _scan_tags(text: str) -> list[_Tag]:
    # Blank comments (same length) so offsets still line up.
    masked = _COMMENT_RE.sub(lambda m: " " * len(m.group(0)), text)
    tags: list[_Tag] = []
    for match in _TAG_RE.finditer(masked):
        attrs: dict[str, tuple[str, int]] = {}
        base = match.start("attrs")
        for attr in _ATTR_RE.finditer(match.group("attrs")):
            group = "dq" if attr.group("dq") is not None else "sq"
            attrs[attr.group("name")] = (attr.group(group), base + attr.start(group))
        tags.append(
            _Tag(
                name=match.group("name"),
                start=match.start(),
                end=match.end(),
                closing=bool(match.group("close")),
                self_closing=bool(match.group("self")),
                attrs=attrs,
            )
        )
    return tags


def _line_col(text: str, offset: int) -> tuple[int, int]:
    line = text.count("\n", 0, offset) + 1
    return line, offset - (text.rfind("\n", 0, offset) + 1) + 1


def _line_indent(text: str, offset: int) -> tuple[str, bool]:
    """Indent of the line containing offset, and whether offset is its first token."""
    line_start = text.rfind("\n", 0, offset) + 1
    prefix = text[line_start:offset]
    indent = prefix[: len(prefix) - len(prefix.lstrip(" \t"))]
    return indent, prefix.strip() == ""


def _unescape(value: str) -> str:
    return (
        value.replace("&quot;", '"')
        .replace("&apos;", "'")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&amp;", "&")
    )


def find_callbacks(text: str, part_path: str = "") -> list[RibbonCallback]:
    found: list[RibbonCallback] = []
    for tag in _scan_tags(text):
        if tag.closing:
            continue
        control_id = next(
            (tag.attrs[k][0] for k in ("id", "idQ", "idMso") if k in tag.attrs), ""
        )
        label = _unescape(tag.attrs.get("label", ("", 0))[0])
        for attr, (value, offset) in tag.attrs.items():
            if not _CALLBACK_ATTR_RE.fullmatch(attr) or not value.strip():
                continue
            line, column = _line_col(text, offset)
            found.append(
                RibbonCallback(
                    part_path=part_path,
                    attribute=attr,
                    name=_unescape(value).strip(),
                    element=tag.name,
                    control_id=control_id,
                    label=label,
                    offset=offset,
                    length=len(value),
                    line=line,
                    column=column,
                    tag_start=tag.start,
                    tag_end=tag.end,
                )
            )
    return found


def callback_at(text: str, offset: int, part_path: str = "") -> RibbonCallback | None:
    """Callback under the cursor: inside an attribute value, else inside its tag."""
    callbacks = find_callbacks(text, part_path)
    for cb in callbacks:
        if cb.offset <= offset <= cb.offset + cb.length:
            return cb
    in_tag = [cb for cb in callbacks if cb.tag_start <= offset < cb.tag_end]
    return next((cb for cb in in_tag if cb.attribute == "onAction"), in_tag[0] if in_tag else None)


def control_ids(text: str) -> list[tuple[str, int]]:
    return [
        tag.attrs["id"]
        for tag in _scan_tags(text)
        if not tag.closing and "id" in tag.attrs
    ]


def find_containers(text: str, part_path: str = "") -> list[RibbonContainer]:
    """Custom groups/menus/boxes a button can be added to, in document order."""
    stack: list[tuple[_Tag, list[_Tag]]] = []
    found: list[tuple[int, RibbonContainer]] = []
    for tag in _scan_tags(text):
        if tag.closing:
            while stack:
                open_tag, children = stack.pop()
                if open_tag.name == tag.name:
                    container = _container_for(text, part_path, open_tag, children, tag, stack)
                    if container is not None:
                        found.append((open_tag.start, container))
                    break
            continue
        if stack:
            stack[-1][1].append(tag)
        if not tag.self_closing:
            stack.append((tag, []))
    return [container for _start, container in sorted(found, key=lambda item: item[0])]


def _container_for(text, part_path, open_tag, children, close_tag, stack):
    if open_tag.name not in _CONTAINER_ELEMENTS or "id" not in open_tag.attrs:
        return None
    names = [
        _unescape(t.attrs.get("label", t.attrs.get("id", ("", 0)))[0])
        for t, _c in stack
        if t.name in ("tab", *_CONTAINER_ELEMENTS, "splitButton")
    ]
    names.append(_unescape(open_tag.attrs.get("label", open_tag.attrs["id"])[0]))
    indent, _first = _line_indent(text, open_tag.start)
    child_indent = None
    for child in children:
        child_line_indent, first = _line_indent(text, child.start)
        if first:
            child_indent = child_line_indent
            break
    if child_indent is None:
        parent_indent = _line_indent(text, stack[-1][0].start)[0] if stack else ""
        unit = indent[len(parent_indent):] if indent.startswith(parent_indent) else ""
        child_indent = indent + (unit or "    ")
    return RibbonContainer(
        part_path=part_path,
        element=open_tag.name,
        control_id=open_tag.attrs["id"][0],
        display=" › ".join(n for n in names if n),
        close_start=close_tag.start,
        indent=indent,
        child_indent=child_indent,
        open_start=open_tag.start,
    )


def innermost_container(
    containers: list[RibbonContainer], part_path: str, offset: int
) -> RibbonContainer | None:
    """The deepest container of ``part_path`` whose element surrounds ``offset``."""
    around = [
        c for c in containers if c.part_path == part_path and c.open_start <= offset <= c.close_start
    ]
    return max(around, key=lambda c: c.open_start, default=None)


# -- building XML ----------------------------------------------------------------


def _attr(value: str) -> str:
    return escape(value, {'"': "&quot;"})


def build_button_xml(
    *,
    control_id: str,
    label: str,
    on_action: str,
    image_mso: str = "",
    size: str = "",
    screentip: str = "",
    supertip: str = "",
) -> str:
    parts = [f'<button id="{_attr(control_id)}"', f'label="{_attr(label)}"']
    if size:
        parts.append(f'size="{_attr(size)}"')
    if image_mso:
        parts.append(f'imageMso="{_attr(image_mso)}"')
    parts.append(f'onAction="{_attr(on_action)}"')
    if screentip:
        parts.append(f'screentip="{_attr(screentip)}"')
    if supertip:
        parts.append(f'supertip="{_attr(supertip)}"')
    return " ".join(parts) + " />"


def child_unit(container: RibbonContainer) -> str:
    """One level of indentation as the container's own children use it."""
    if container.child_indent.startswith(container.indent):
        return container.child_indent[len(container.indent):] or "    "
    return "    "


def build_menu_xml(
    *,
    control_id: str,
    label: str,
    image_mso: str = "",
    size: str = "",
    item_size: str = "",
    screentip: str = "",
    supertip: str = "",
    children: tuple[str, ...] | list[str] = (),
    indent: str = "",
    unit: str = "    ",
) -> str:
    """A <menu> element holding ``children`` (complete element strings).

    The first line carries no indentation and later lines carry ``indent``, so
    the block goes wherever a one-line element indented by ``indent`` would.
    """
    parts = [f'<menu id="{_attr(control_id)}"', f'label="{_attr(label)}"']
    if size:
        parts.append(f'size="{_attr(size)}"')
    if item_size:
        parts.append(f'itemSize="{_attr(item_size)}"')
    if image_mso:
        parts.append(f'imageMso="{_attr(image_mso)}"')
    if screentip:
        parts.append(f'screentip="{_attr(screentip)}"')
    if supertip:
        parts.append(f'supertip="{_attr(supertip)}"')
    lines = [" ".join(parts) + ">"]
    lines += [indent + unit + child for child in children]
    lines.append(indent + "</menu>")
    return "\n".join(lines)


def insert_into_container(
    text: str, container: RibbonContainer, element_xml: str
) -> tuple[str, int, int]:
    """Insert element_xml as the last child. Returns (new_text, offset, length)."""
    close = container.close_start
    _indent, close_is_first = _line_indent(text, close)
    if close_is_first:
        at = text.rfind("\n", 0, close) + 1
        insert = container.child_indent + element_xml + "\n"
        offset = at + len(container.child_indent)
    else:
        at = close
        insert = "\n" + container.child_indent + element_xml + "\n" + container.indent
        offset = at + 1 + len(container.child_indent)
    return text[:at] + insert + text[at:], offset, len(element_xml)


def unique_control_id(existing: set[str], base: str, prefix: str = "btn") -> str:
    stem = re.sub(r"[^A-Za-z0-9_]", "", base) or ("Menu" if prefix == "mnu" else "Button")
    candidate = prefix + stem[0].upper() + stem[1:]
    taken = {item.casefold() for item in existing}
    if candidate.casefold() not in taken:
        return candidate
    n = 2
    while f"{candidate}{n}".casefold() in taken:
        n += 1
    return f"{candidate}{n}"


def suggested_label(proc_name: str) -> str:
    name = re.sub(r"Callback$", "", proc_name) or proc_name
    words = re.sub(r"([a-z0-9])([A-Z])", r"\1 \2", name.replace("_", " "))
    return re.sub(r"\s+", " ", words).strip()


# -- VBA side ----------------------------------------------------------------------


def split_callback_name(value: str) -> tuple[str | None, str]:
    """'Book.xlam'!Module1.Proc -> ('Module1', 'Proc')."""
    tail = value.rsplit("!", 1)[-1].strip().strip("'")
    if "." in tail:
        module, proc = tail.rsplit(".", 1)
        return module or None, proc
    return None, tail


def procedure_index(draft: DocumentDraft) -> dict[str, list[ProcedureLocation]]:
    index: dict[str, list[ProcedureLocation]] = {}
    for module in draft.modules:
        if module.is_deleted:
            continue
        for proc in _procedures(module.body, module.id):
            if proc.kind not in ("Sub", "Function"):
                continue
            index.setdefault(proc.name.casefold(), []).append(_location(module, proc))
    return index


def ribbon_ready_procedures(draft: DocumentDraft) -> list[ProcedureLocation]:
    """Subs a ribbon control can call as they are: standard module, IRibbonControl argument."""
    ready = [
        location
        for locations in procedure_index(draft).values()
        for location in locations
        if location.is_standard and location.kind == "Sub" and is_ribbon_ready(location.signature)
    ]
    return sorted(ready, key=lambda p: (p.module_name.casefold(), p.line))


@lru_cache(maxsize=256)
def _procedures(body: str, module_id: str) -> tuple[ProcedureInfo, ...]:
    # Baseline and draft share most module bodies; parse each text once.
    return tuple(parse_procedures(body, module_id))


def _location(module: ModuleDraft, proc: ProcedureInfo) -> ProcedureLocation:
    return ProcedureLocation(
        module_id=module.id,
        module_name=module.current_name,
        project_item_kind=module.project_item_kind,
        is_standard=module.pyopenvba_kind == "standard",
        name=proc.name,
        kind=proc.kind,
        line=proc.line,
        signature=proc.signature,
    )


def resolve(index: dict[str, list[ProcedureLocation]], callback_name: str) -> Resolution:
    module_name, proc_name = split_callback_name(callback_name)
    matches = index.get(proc_name.casefold(), [])
    if module_name is not None:
        qualified = [m for m in matches if m.module_name.casefold() == module_name.casefold()]
        if not qualified:
            return Resolution(
                STATUS_MISSING, (), f"'{callback_name}' has no matching Sub in module {module_name}."
            )
        callable_ = [m for m in qualified if m.is_standard or m.project_item_kind == "document"]
        if not callable_:
            return Resolution(
                STATUS_NOT_CALLABLE,
                tuple(qualified),
                f"'{callback_name}' is in class module {module_name}, which the ribbon cannot call.",
            )
        return Resolution(STATUS_OK, tuple(callable_))
    if not matches:
        return Resolution(STATUS_MISSING, (), f"'{callback_name}' does not exist in any module.")
    standard = [m for m in matches if m.is_standard]
    if not standard:
        where = ", ".join(sorted({m.module_name for m in matches}))
        return Resolution(
            STATUS_NOT_CALLABLE,
            tuple(matches),
            f"'{callback_name}' is only defined in {where}; ribbon callbacks must be "
            "in a standard module (or qualified as Module.Proc for document modules).",
        )
    if len(standard) > 1:
        where = ", ".join(m.module_name for m in standard)
        return Resolution(
            STATUS_AMBIGUOUS,
            tuple(standard),
            f"'{callback_name}' is defined in several modules ({where}); "
            "Office cannot tell which one to run.",
        )
    return Resolution(STATUS_OK, tuple(standard))


def procedure_at(body: str, module_id: str, line: int) -> ProcedureInfo | None:
    procs = [p for p in parse_procedures(body, module_id) if p.line <= line]
    return procs[-1] if procs else None


def _parameters(signature: str) -> str:
    match = re.search(r"\((.*)\)", signature)
    return match.group(1).strip() if match else ""


def is_ribbon_ready(signature: str) -> bool:
    return "iribboncontrol" in _parameters(signature).casefold()


def takes_arguments(signature: str) -> bool:
    return bool(_parameters(signature))


def callback_parameters(attribute: str, element: str) -> str:
    if attribute == "onLoad":
        return "ribbon As IRibbonUI"
    if attribute == "loadImage":
        return "imageId As String, ByRef image"
    if attribute in ("onShow", "onHide"):
        return "contextObject As Object"
    if attribute == "onAction":
        if element in ("toggleButton", "checkBox"):
            return "control As IRibbonControl, pressed As Boolean"
        if element in ("dropDown", "gallery"):
            return "control As IRibbonControl, selectedId As String, selectedIndex As Integer"
        return "control As IRibbonControl"
    if attribute == "onChange":
        return "control As IRibbonControl, text As String"
    if attribute.startswith("getItem"):
        return "control As IRibbonControl, index As Integer, ByRef returnedVal"
    if attribute.startswith("get"):
        return "control As IRibbonControl, ByRef returnedVal"
    return "control As IRibbonControl"


def callback_stub(name: str, attribute: str, element: str, *, calls: str | None = None) -> str:
    inner = f"    {calls}" if calls else "    "
    return f"Public Sub {name}({callback_parameters(attribute, element)})\n{inner}\nEnd Sub\n"


def append_procedure(body: str, stub: str) -> tuple[str, int]:
    """Append stub to a module body. Returns (new_body, 1-based line of the Sub)."""
    trimmed = body.rstrip("\n")
    if not trimmed.strip():
        return stub, 1
    new_body = trimmed + "\n\n" + stub
    return new_body, trimmed.count("\n") + 3


def is_valid_procedure_name(name: str) -> bool:
    return bool(_PROC_NAME_RE.match(name)) and len(name) <= 255


# -- whole-project views ---------------------------------------------------------


@dataclass(frozen=True)
class RibbonEntry:
    callback: RibbonCallback
    resolution: Resolution


def ribbon_entries(draft: DocumentDraft) -> list[RibbonEntry]:
    callbacks = [
        cb for part in ribbon_parts(draft) for cb in find_callbacks(part.text or "", part.path)
    ]
    if not callbacks:
        return []  # skip parsing every module when there is nothing to resolve
    index = procedure_index(draft)
    return [RibbonEntry(cb, resolve(index, cb.name)) for cb in callbacks]


def callbacks_for_procedure(
    draft: DocumentDraft, module_id: str, proc_name: str
) -> list[RibbonCallback]:
    return [
        entry.callback
        for entry in ribbon_entries(draft)
        if any(
            m.module_id == module_id and m.name.casefold() == proc_name.casefold()
            for m in entry.resolution.matches
        )
    ]


def check_ribbon(draft: DocumentDraft) -> list[RibbonIssue]:
    issues: list[RibbonIssue] = []
    for entry in ribbon_entries(draft):
        if entry.resolution.status != STATUS_OK:
            cb = entry.callback
            what = f'{cb.element} "{cb.label or cb.control_id}"' if (cb.label or cb.control_id) else cb.element
            issues.append(
                RibbonIssue(
                    cb.part_path,
                    cb.line,
                    f"{what} {cb.attribute}: {entry.resolution.message}",
                    cb.offset,
                    cb.length,
                )
            )
    for part in ribbon_parts(draft):
        text = part.text or ""
        seen: dict[str, int] = {}
        for value, offset in control_ids(text):
            key = value.casefold()
            if key in seen:
                issues.append(
                    RibbonIssue(
                        part.path,
                        _line_col(text, offset)[0],
                        f'Duplicate control id "{value}" (ids must be unique or the ribbon will not load).',
                        offset,
                        len(value),
                    )
                )
            seen[key] = offset
    return sorted(issues, key=lambda i: (i.part_path, i.line))


def new_issues(draft: DocumentDraft) -> list[RibbonIssue]:
    """Issues present in the draft but not in the file as opened."""
    current = check_ribbon(draft)
    if not current:
        return []
    baseline = {issue.key for issue in check_ribbon(draft_from_snapshot(draft.baseline))}
    return [issue for issue in current if issue.key not in baseline]


# -- element structure: move, group into menu --------------------------------------


class RibbonEditError(ValueError):
    """A structural ribbon edit that cannot be made; the message is for the user."""


@dataclass(frozen=True)
class RibbonElement:
    index: int
    name: str
    start: int  # "<" of the opening tag
    open_end: int  # just past the opening tag
    end: int  # just past the closing tag; -1 while the element is unclosed
    parent: int  # index of the parent element, -1 at the top
    attrs: tuple[tuple[str, str, int], ...]  # (name, raw value, value offset)

    @property
    def complete(self) -> bool:
        return self.end >= 0

    def attr(self, name: str) -> str | None:
        return next((value for key, value, _o in self.attrs if key == name), None)

    def describe(self) -> str:
        label = self.attr("label") or self.attr("id") or self.attr("idMso") or ""
        return f'{self.name} "{_unescape(label)}"' if label else self.name


# Parents whose children follow a fixed schema order: never reorder inside them.
_FIXED_ORDER_PARENTS = {"customUI", "ribbon", "splitButton", "backstage"}
# Parents whose children are items followed by buttons: reorder only like with like.
_ITEMS_THEN_BUTTONS = {"gallery", "dropDown", "comboBox"}
# dialogBoxLauncher must stay the last child of its group.
_PINNED_ELEMENTS = {"dialogBoxLauncher"}
# Elements a <menu> may hold.
MENU_CHILD_ELEMENTS = {
    "button",
    "toggleButton",
    "checkBox",
    "gallery",
    "menu",
    "splitButton",
    "dynamicMenu",
    "menuSeparator",
    "control",
}
# Containers a new <menu> may go in.
MENU_PARENT_ELEMENTS = {"group", "menu", "box", "buttonGroup"}
# Elements that accept image="rId".
IMAGE_ELEMENTS = {
    "button",
    "toggleButton",
    "menu",
    "gallery",
    "dynamicMenu",
    "comboBox",
    "dropDown",
    "editBox",
    "labelControl",
    "group",
    "item",
}


def element_tree(text: str) -> list[RibbonElement]:
    """Every element of ``text`` in document order, with parent links."""
    raw: list[dict] = []
    stack: list[int] = []
    for tag in _scan_tags(text):
        if tag.closing:
            depth = next(
                (d for d in range(len(stack) - 1, -1, -1) if raw[stack[d]]["name"] == tag.name),
                None,
            )
            if depth is None:
                continue  # stray close tag
            raw[stack[depth]]["end"] = tag.end
            del stack[depth:]  # anything still open inside was never closed
            continue
        raw.append(
            {
                "name": tag.name,
                "start": tag.start,
                "open_end": tag.end,
                "end": tag.end if tag.self_closing else -1,
                "parent": stack[-1] if stack else -1,
                "attrs": tuple((k, v, o) for k, (v, o) in tag.attrs.items()),
            }
        )
        if not tag.self_closing:
            stack.append(len(raw) - 1)
    return [RibbonElement(index=i, **item) for i, item in enumerate(raw)]


def _children(elements: list[RibbonElement], parent: int) -> list[RibbonElement]:
    return [e for e in elements if e.parent == parent]


def _local_name(name: str) -> str:
    return name.rsplit(":", 1)[-1]


def element_at(text: str, offset: int) -> RibbonElement | None:
    """The innermost complete element whose markup contains ``offset``."""
    around = [
        e
        for e in element_tree(text)
        if e.complete and e.start <= offset < e.end
    ]
    return max(around, key=lambda e: e.start, default=None)


def _movable(elements: list[RibbonElement], element: RibbonElement) -> RibbonElement | None:
    """The element itself or, failing that, its nearest movable ancestor."""
    current: RibbonElement | None = element
    while current is not None:
        parent = elements[current.parent] if current.parent >= 0 else None
        if (
            parent is not None
            and current.complete
            and _local_name(parent.name) not in _FIXED_ORDER_PARENTS
        ):
            return current
        current = parent
    return None


def move_element(text: str, offset: int, direction: int) -> tuple[str, RibbonElement]:
    """Swap the element at ``offset`` with its previous (-1) or next (+1) sibling.

    Whitespace and comments between the two stay where they are. Returns the new
    text and the moved element as it is in the new text.
    """
    elements = element_tree(text)
    here = [e for e in elements if e.complete and e.start <= offset < e.end]
    if not here:
        raise RibbonEditError("Put the cursor on the ribbon control you want to move.")
    innermost = max(here, key=lambda e: e.start)
    if _local_name(innermost.name) in _PINNED_ELEMENTS:
        raise RibbonEditError("The dialog box launcher must stay last in its group.")
    element = _movable(elements, innermost)
    if element is None:
        raise RibbonEditError("This element has a fixed position in the ribbon XML.")
    parent = elements[element.parent]
    siblings = [e for e in _children(elements, parent.index) if e.complete]
    position = siblings.index(element)
    target = position + (1 if direction > 0 else -1)
    if not 0 <= target < len(siblings):
        where = "first" if direction < 0 else "last"
        raise RibbonEditError(f"{element.describe()} is already the {where} item in its {parent.name}.")
    other = siblings[target]
    if _local_name(other.name) in _PINNED_ELEMENTS:
        raise RibbonEditError(f"The dialog box launcher must stay last in its {parent.name}.")
    if _local_name(parent.name) in _ITEMS_THEN_BUTTONS and other.name != element.name:
        raise RibbonEditError(
            f"In a {parent.name}, items come before buttons; they cannot swap places."
        )
    first, second = (other, element) if direction < 0 else (element, other)
    new_text = (
        text[: first.start]
        + text[second.start : second.end]
        + text[first.end : second.start]
        + text[first.start : first.end]
        + text[second.end :]
    )
    if direction < 0:
        new_start = first.start
    else:
        new_start = first.start + (second.end - second.start) + (second.start - first.end)
    moved = element_at(new_text, new_start)
    assert moved is not None and moved.start == new_start
    return new_text, moved


def selected_siblings(text: str, start: int, end: int) -> tuple[RibbonElement, list[RibbonElement]]:
    """(parent, children) for the run of sibling elements a selection covers.

    The parent is the deepest element holding the whole selection; the run is
    its children the selection touches. A selection inside one control selects
    that control.
    """
    if end < start:
        start, end = end, start
    elements = element_tree(text)
    holders = [e for e in elements if e.complete and e.start <= start and end <= e.end]
    if not holders:
        raise RibbonEditError("Select the ribbon buttons you want to put in a menu.")
    holder = max(holders, key=lambda e: e.start)
    touched = [
        e
        for e in _children(elements, holder.index)
        if e.complete and e.start < max(end, start + 1) and e.end > start
    ]
    if touched:
        return holder, touched
    if holder.parent < 0:
        raise RibbonEditError("Select the ribbon buttons you want to put in a menu.")
    return elements[holder.parent], [holder]


def _remove_attrs(tag_text: str, names: set[str]) -> str:
    def drop(match: re.Match) -> str:
        return "" if match.group("name") in names else match.group(0)

    return re.sub(
        r"\s+(?P<name>[A-Za-z_][\w:.-]*)\s*=\s*(?:\"[^\"]*\"|'[^']*')", drop, tag_text
    )


def set_tag_attr(tag_text: str, name: str, value: str) -> str:
    """``tag_text`` (one opening tag) with attribute ``name`` set to ``value``."""
    pattern = re.compile(rf"(\s{re.escape(name)}\s*=\s*)(\"[^\"]*\"|'[^']*')")
    if pattern.search(tag_text):
        return pattern.sub(lambda m: f'{m.group(1)}"{_attr(value)}"', tag_text, count=1)
    close = re.search(r"\s*/?>$", tag_text)
    at = close.start() if close else len(tag_text)
    return f'{tag_text[:at]} {name}="{_attr(value)}"{tag_text[at:]}'


def group_into_menu(
    text: str,
    start: int,
    end: int,
    *,
    control_id: str,
    label: str,
    image_mso: str = "",
    size: str = "",
    item_size: str = "",
) -> tuple[str, int, int]:
    """Wrap the selected sibling controls in a new <menu>.

    Returns (new_text, offset, length) of the menu's opening tag. Controls in
    a menu take no size, so size/getSize are dropped from the moved controls.
    """
    parent, run = selected_siblings(text, start, end)
    if _local_name(parent.name) not in MENU_PARENT_ELEMENTS:
        raise RibbonEditError(
            f"A menu can only go in a group, menu, box or buttonGroup, not in a {parent.name}."
        )
    misfits = sorted({e.name for e in run if _local_name(e.name) not in MENU_CHILD_ELEMENTS})
    if misfits:
        raise RibbonEditError(f"A menu cannot hold {', '.join(misfits)} controls.")
    first, last = run[0], run[-1]
    indent, first_on_line = _line_indent(text, first.start)
    if not first_on_line:
        indent = _line_indent(text, parent.start)[0] + "    "
    parent_indent = _line_indent(text, parent.start)[0]
    unit = indent[len(parent_indent):] if indent.startswith(parent_indent) else ""
    unit = unit or "    "

    # Drop size from each moved control's opening tag, last first so offsets hold.
    block = text[first.start : last.end]
    for element in reversed(run):
        a, b = element.start - first.start, element.open_end - first.start
        block = block[:a] + _remove_attrs(block[a:b], {"size", "getSize"}) + block[b:]
    lines = block.split("\n")
    block = "\n".join([lines[0]] + [unit + line if line.strip() else line for line in lines[1:]])

    in_group = _local_name(parent.name) == "group"
    attrs = [f'<menu id="{_attr(control_id)}"', f'label="{_attr(label)}"']
    if size and in_group:
        attrs.append(f'size="{_attr(size)}"')
    if item_size:
        attrs.append(f'itemSize="{_attr(item_size)}"')
    if image_mso:
        attrs.append(f'imageMso="{_attr(image_mso)}"')
    opening = " ".join(attrs) + ">"
    menu = f"{opening}\n{indent}{unit}{block}\n{indent}</menu>"
    new_text = text[: first.start] + menu + text[last.end :]
    return new_text, first.start, len(opening)


# -- custom images ---------------------------------------------------------------------

IMAGE_REL_TYPE = "http://schemas.microsoft.com/office/2006/relationships/images"
_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def relationships_path(part_path: str) -> str:
    """customUI/customUI14.xml -> customUI/_rels/customUI14.xml.rels"""
    folder, _sep, name = part_path.rpartition("/")
    return f"{folder}/_rels/{name}.rels" if folder else f"_rels/{name}.rels"


def relationships(text: str) -> list[tuple[str, str, str]]:
    """(Id, Type, Target) of every Relationship in a .rels part."""
    return [
        (
            _unescape(tag.attrs.get("Id", ("", 0))[0]),
            _unescape(tag.attrs.get("Type", ("", 0))[0]),
            _unescape(tag.attrs.get("Target", ("", 0))[0]),
        )
        for tag in _scan_tags(text)
        if not tag.closing and _local_name(tag.name) == "Relationship"
    ]


def new_relationships_text() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<Relationships xmlns="{_RELS_NS}">\n</Relationships>\n'
    )


def add_relationship(text: str, rel_type: str, target: str) -> tuple[str, str]:
    """Append a Relationship; returns (new_text, its Id)."""
    taken = {rid.casefold() for rid, _t, _g in relationships(text)}
    n = 1
    while f"rIdImg{n}".casefold() in taken:
        n += 1
    rid = f"rIdImg{n}"
    element = f'<Relationship Id="{rid}" Type="{_attr(rel_type)}" Target="{_attr(target)}"/>'
    tags = _scan_tags(text)
    close = next(
        (t for t in reversed(tags) if t.closing and _local_name(t.name) == "Relationships"), None
    )
    if close is not None:
        indent, on_own_line = _line_indent(text, close.start)
        if on_own_line:
            at = text.rfind("\n", 0, close.start) + 1
            return text[:at] + indent + "  " + element + "\n" + text[at:], rid
        return text[: close.start] + element + text[close.start :], rid
    empty = next(
        (t for t in tags if t.self_closing and _local_name(t.name) == "Relationships"), None
    )
    if empty is None:
        raise RibbonEditError("The ribbon's relationships part has no <Relationships> element.")
    opening = re.sub(r"\s*/>$", ">", text[empty.start : empty.end])
    name = text[empty.start + 1 : empty.start + 1 + len(empty.name)]
    return (
        text[: empty.start] + opening + element + f"</{name}>" + text[empty.end :],
        rid,
    )


def ensure_default_content_type(text: str, extension: str, content_type: str) -> str:
    """[Content_Types].xml with a Default for ``extension`` (unchanged if it has one)."""
    tags = _scan_tags(text)
    for tag in tags:
        if (
            not tag.closing
            and _local_name(tag.name) == "Default"
            and tag.attrs.get("Extension", ("", 0))[0].casefold() == extension.casefold()
        ):
            return text
    root = next((t for t in tags if not t.closing and _local_name(t.name) == "Types"), None)
    if root is None or root.self_closing:
        raise RibbonEditError("[Content_Types].xml has no <Types> element to add to.")
    element = f'<Default Extension="{_attr(extension)}" ContentType="{_attr(content_type)}"/>'
    first_child = next((t for t in tags if t.start > root.end), None)
    if first_child is not None:
        indent, on_own_line = _line_indent(text, first_child.start)
        if on_own_line:
            at = text.rfind("\n", 0, first_child.start) + 1
            return text[:at] + indent + element + "\n" + text[at:]
    return text[: root.end] + element + text[root.end :]


def image_control_at(text: str, offset: int) -> RibbonElement | None:
    """The innermost control around ``offset`` that can show a custom image."""
    around = [
        e
        for e in element_tree(text)
        if e.complete and e.start <= offset < e.end and _local_name(e.name) in IMAGE_ELEMENTS
    ]
    return max(around, key=lambda e: e.start, default=None)


def set_control_image(text: str, element: RibbonElement, rid: str) -> str:
    """The control's opening tag with image=rid and no imageMso/getImage."""
    tag = text[element.start : element.open_end]
    tag = set_tag_attr(_remove_attrs(tag, {"imageMso", "getImage"}), "image", rid)
    return text[: element.start] + tag + text[element.open_end :]
