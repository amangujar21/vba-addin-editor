"""UserForms in a draft: create one, edit its layout, keep its code in step.

A form's layout lives in the draft as its designer storage (``FormStorage``).
Every edit here goes through the adapter, which applies it with pyOpenVBA to a
scratch copy of that storage and returns the new bytes; the draft only ever
holds finished storages, and a save copies them into the file unchanged.
Nothing here touches Tk.
"""

from __future__ import annotations

import re
from collections.abc import Callable

from vba_addin_editor.adapters import pyopenvba_adapter as pa
from vba_addin_editor.domain.document import (
    DocumentDraft,
    FormDesignSnapshot,
    FormStorage,
    ModuleDisplayKind,
    ModuleDraft,
    new_module_id,
)
from vba_addin_editor.services.validation_service import VBA_RESERVED

CONTROL_KINDS = pa.FORM_CONTROL_KINDS
FIXED_SIZE_KINDS = frozenset({"MultiPage", "TabStrip"})
CONTAINER_KINDS = frozenset({"Frame", "Form"})  # Form = a MultiPage's page

# The event the VBA editor opens when a control is double-clicked.
DEFAULT_EVENTS = {
    "UserForm": "Click",
    "CommandButton": "Click",
    "Label": "Click",
    "Image": "Click",
    "CheckBox": "Click",
    "OptionButton": "Click",
    "ToggleButton": "Click",
    "Frame": "Click",
    "TextBox": "Change",
    "ComboBox": "Change",
    "ListBox": "Click",
    "MultiPage": "Change",
    "TabStrip": "Change",
    "ScrollBar": "Change",
    "SpinButton": "Change",
}
# What each kind is called when it is added: the VBA editor's own names.
_NAME_STEMS = {kind: kind for kind in CONTROL_KINDS}
_CONTROL_NAME_RE = re.compile(r"^[A-Za-z][A-Za-z0-9_]{0,39}$")


class FormError(ValueError):
    """A form edit that cannot be made; the message is for the user."""


def _code_page(draft: DocumentDraft) -> int:
    return draft.baseline.code_page


_design_cache: dict[tuple[int, str, int], FormDesignSnapshot] = {}


def design_of(draft: DocumentDraft, module: ModuleDraft) -> FormDesignSnapshot | None:
    """The layout ``module`` has in this draft; None when it is not a form."""
    storage = draft.form_storage_of(module)
    if storage is None:
        if draft.is_form(module):
            return draft.baseline_form(module)
        return None
    if module.form_storage is None:
        base = draft.baseline_form(module)
        if base is not None and not base.problem:
            return base
    key = (id(storage), module.current_name, _code_page(draft))
    cached = _design_cache.get(key)
    if cached is None or cached.storage is not storage:
        if len(_design_cache) > 64:
            _design_cache.clear()
        cached = pa.read_form_design(storage, module.current_name, _code_page(draft))
        _design_cache[key] = cached
    return cached


def forms(draft: DocumentDraft) -> list[ModuleDraft]:
    """Every UserForm module the draft keeps, in project order."""
    return [m for m in draft.modules if not m.is_deleted and draft.is_form(m)]


def default_form_name(draft: DocumentDraft) -> str:
    taken = {m.current_name.casefold() for m in draft.modules if not m.is_deleted}
    taken |= draft.baseline.orphan_form_names()
    n = 1
    while f"userform{n}" in taken:
        n += 1
    return f"UserForm{n}"


def add_blank_form(draft: DocumentDraft, name: str, caption: str | None = None) -> ModuleDraft:
    """A new empty UserForm, as the VBA editor's Insert ▸ UserForm makes one."""
    storage, header = pa.new_form_storage(name, caption or name, _code_page(draft))
    module = ModuleDraft(
        id=new_module_id(),
        origin_name=None,
        current_name=name,
        body="Option Explicit\n",
        kind=ModuleDisplayKind.AMBIGUOUS,
        pyopenvba_kind="other",
        is_new=True,
        is_deleted=False,
        destructive_ops_safe=True,
        can_delete=True,
        can_rename=True,
        hidden_header=header.rstrip("\n"),
        project_item_kind="userform",
        form_storage=storage,
    )
    draft.modules.append(module)
    return module


def edited_layout(
    draft: DocumentDraft, module: ModuleDraft, edit: Callable[[pa.FormLayoutEditor], None]
) -> FormStorage:
    """The form's layout with ``edit`` applied; the draft is not changed."""
    storage = draft.form_storage_of(module)
    if storage is None:
        raise FormError(f"The layout of {module.current_name} could not be read, so it cannot be edited.")
    try:
        return pa.edit_form_storage(storage, module.current_name, _code_page(draft), edit)
    except pa.AdapterError as exc:
        raise FormError(str(exc)) from exc


def set_layout(draft: DocumentDraft, module: ModuleDraft, storage: FormStorage) -> None:
    base = draft.baseline_form(module)
    # Back to what the file holds (but for the TypeInfoVer an edit drops):
    # nothing to write for this layout.
    unchanged = base is not None and pa.same_layout(base.storage, storage)
    module.form_storage = None if unchanged else storage


def edit_layout(
    draft: DocumentDraft, module: ModuleDraft, edit: Callable[[pa.FormLayoutEditor], None]
) -> FormStorage:
    """Apply ``edit`` to the form's layout; the module keeps the result."""
    edited = edited_layout(draft, module, edit)
    set_layout(draft, module, edited)
    return edited


def next_control_name(design: FormDesignSnapshot | None, kind: str) -> str:
    taken = {c.name.casefold() for c in design.walk()} if design is not None else set()
    stem = _NAME_STEMS.get(kind, kind)
    n = 1
    while f"{stem}{n}".casefold() in taken:
        n += 1
    return f"{stem}{n}"


def control_name_problem(name: str, design: FormDesignSnapshot | None, *, current: str = "") -> str | None:
    if not _CONTROL_NAME_RE.match(name):
        return "A control name starts with a letter and uses letters, digits and underscores (40 at most)."
    if name.casefold() in VBA_RESERVED:
        return f"'{name}' is a reserved VBA word."
    clash = design is not None and any(c.name.casefold() == name.casefold() for c in design.walk())
    if clash and name.casefold() != current.casefold():
        return f"The form already has a control named {name}."
    return None


# -- the code behind the form -----------------------------------------------------------

_SUB_RE = r"(?im)^[ \t]*(?:(?:Private|Public|Friend)\s+)?(?:Static\s+)?Sub\s+{name}\s*\("


def find_handler(body: str, control: str, event: str) -> int | None:
    """1-based line of ``Sub <control>_<event>(`` in ``body``, if any."""
    pattern = re.compile(_SUB_RE.format(name=re.escape(f"{control}_{event}")))
    match = pattern.search(body)
    if match is None:
        return None
    return body.count("\n", 0, match.start()) + 1


def ensure_handler(body: str, control: str, event: str) -> tuple[str, int]:
    """``body`` with an event handler for ``control`` (added when missing), and its body line."""
    line = find_handler(body, control, event)
    if line is not None:
        return body, line + 1
    text = body if body.endswith("\n") or not body else body + "\n"
    if text.strip():
        text += "\n"
    start = text.count("\n") + 1
    text += f"Private Sub {control}_{event}()\n    \nEnd Sub\n"
    return text, start + 1


def rename_in_code(body: str, old: str, new: str) -> str:
    """Rename control ``old`` to ``new`` in its form's code.

    Whole-word references (``old.Caption``, ``Me.old``) and event handlers
    (``Sub old_Click``) follow; strings and comments do not change.
    """
    if not old or old == new:
        return body
    from vba_addin_editor.services.port_service import _lex

    toks = _lex(body)
    edits: list[tuple[int, int, str]] = []
    folded = old.casefold()
    prefix = folded + "_"
    for tok in toks:
        if tok.kind != "id":
            continue
        text = tok.text.casefold()
        if text == folded:
            edits.append((tok.start, tok.end, new))
        elif text.startswith(prefix) and text[len(prefix):] and _is_event_handler(body, tok.start):
            edits.append((tok.start, tok.start + len(old), new))
    for start, end, value in sorted(edits, reverse=True):
        body = body[:start] + value + body[end:]
    return body


def _is_event_handler(body: str, start: int) -> bool:
    """True when the identifier at ``start`` is the name in a ``Sub`` line."""
    line_start = body.rfind("\n", 0, start) + 1
    head = body[line_start:start]
    return re.fullmatch(r"\s*(?:(?:Private|Public|Friend)\s+)?(?:Static\s+)?Sub\s+", head, re.IGNORECASE) is not None


def event_for(kind: str) -> str:
    return DEFAULT_EVENTS.get(kind, "Click")
