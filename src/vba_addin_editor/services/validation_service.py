"""Draft validation: naming rules, uniqueness, encodability (plan 27.2, 42)."""

from __future__ import annotations

import re

from vba_addin_editor.domain.document import DocumentDraft

# Conservative VBA component identifier rules (plan 42): ASCII, letter or
# underscore first, then letters/digits/underscores. Max length 255 per
# MS-OVBA practical limits; keep conservative 64 until live VBE tests widen it.
_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]{0,63}$")

VBA_RESERVED = {
    "and", "as", "boolean", "byref", "byte", "byval", "call", "case", "class",
    "const", "currency", "date", "dim", "do", "double", "each", "else", "elseif",
    "empty", "end", "endif", "enum", "eqv", "error", "exit", "explicit", "false",
    "for", "friend", "function", "get", "global", "goto", "if", "imp", "implements",
    "in", "integer", "is", "let", "lib", "long", "loop", "lset", "me", "mod",
    "module", "new", "next", "not", "nothing", "null", "object", "on", "option",
    "optional", "or", "preserve", "private", "property", "public", "raiseevent",
    "redim", "rem", "resume", "return", "rset", "select", "set", "single",
    "static", "stop", "string", "sub", "then", "to", "true", "type", "typeof",
    "until", "variant", "wend", "while", "with", "xor",
}


def validate_module_name(name: str, draft: DocumentDraft, exclude_id: str | None = None) -> str | None:
    """Return a plain-language problem or None when acceptable."""
    if not name or not name.strip():
        return "Module name cannot be empty."
    if any(ord(c) > 127 for c in name):
        return "Module names must use ASCII letters, digits, and underscores in this version."
    if any(c in name for c in "\\/:*?\"<>|"):
        return "Module names cannot contain path or punctuation characters."
    if any(ord(c) < 32 for c in name):
        return "Module names cannot contain control characters."
    if not _NAME_RE.match(name):
        return (
            "Module names must start with a letter or underscore and use only "
            "letters, digits, and underscores (max 64 characters)."
        )
    if name.lower() in VBA_RESERVED:
        return f"'{name}' is a reserved VBA word and cannot be used as a module name."
    for m in draft.modules:
        if m.is_deleted or m.id == exclude_id:
            continue
        if m.current_name.casefold() == name.casefold():
            return f"A module named '{m.current_name}' already exists (names are case-insensitive)."
    return None


def validate_draft(draft: DocumentDraft) -> list[str]:
    problems: list[str] = []
    seen: dict[str, str] = {}
    for m in draft.final_module_state():
        is_new_or_renamed = m.is_new or (
            m.origin_name is not None
            and m.current_name.casefold() != m.origin_name.casefold()
        )
        if is_new_or_renamed:
            problem = validate_module_name(m.current_name, draft, exclude_id=m.id)
            if problem:
                problems.append(f"{m.current_name!r}: {problem}")
        key = m.current_name.casefold()
        if key in seen:
            problems.append(f"Duplicate module name: {m.current_name!r}.")
        seen[key] = m.id
    if draft.baseline.safety.password_protected:
        problems.append("The VBA project is password-protected.")
    return problems
