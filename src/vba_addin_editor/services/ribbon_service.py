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
    )


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


def unique_control_id(existing: set[str], base: str) -> str:
    stem = re.sub(r"[^A-Za-z0-9_]", "", base) or "Button"
    candidate = "btn" + stem[0].upper() + stem[1:]
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
