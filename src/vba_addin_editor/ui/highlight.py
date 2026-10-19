"""Syntax colouring for the code editors, limited to the lines on screen.

Colouring a whole 10,000-line module takes a quarter of a second, too slow to
redo on every keystroke; the visible lines plus a margin take a few
milliseconds. Lines outside that window may keep stale colours, but they are
recoloured before they can scroll into view.
"""

from __future__ import annotations

import bisect
import re
import tkinter as tk

TAGS = ("keyword", "string", "comment", "number", "xml_tag", "xml_attr")
MARGIN_LINES = 40
# An XML comment that opened further back than this is not looked for.
XML_COMMENT_LOOKBACK_LINES = 5000

_VBA_KEYWORDS = (
    "AddressOf", "And", "As", "Boolean", "ByRef", "Byte", "ByVal", "Call", "Case", "Const",
    "Currency", "Date", "Declare", "Dim", "Do", "Double", "Each", "Else", "ElseIf", "Empty",
    "End", "Enum", "Eqv", "Erase", "Error", "Event", "Exit", "Explicit", "False", "For",
    "Friend", "Function", "Get", "GoSub", "GoTo", "If", "Imp", "Implements", "In", "Integer",
    "Is", "Let", "Lib", "Like", "Long", "LongLong", "LongPtr", "Loop", "LSet", "Me", "Mod",
    "New", "Next", "Not", "Nothing", "Null", "Object", "On", "Option", "Optional", "Or",
    "ParamArray", "Preserve", "Private", "Property", "PtrSafe", "Public", "RaiseEvent", "ReDim",
    "Resume", "Return", "RSet", "Select", "Set", "Single", "Static", "Step", "Stop", "String",
    "Sub", "Then", "To", "True", "Type", "TypeOf", "Until", "Variant", "Wend", "While", "With",
    "WithEvents", "Xor",
)

VBA_PATTERN = re.compile(
    r"(?P<comment>'[^\n]*|(?m:^[ \t]*Rem\b[^\n]*))"
    r"|(?P<string>\"(?:[^\"\n]|\"\")*\"?)"
    r"|(?P<number>&[HO][0-9A-F]+&?|\b\d+(?:\.\d*)?(?:E[+-]?\d+)?\b)"
    # Not after a dot: rng.Select and obj.Type are members, not keywords.
    r"|(?P<keyword>(?<![\w.])#?(?:" + "|".join(_VBA_KEYWORDS) + r")\b)",
    re.IGNORECASE,
)

XML_PATTERN = re.compile(
    r"(?P<comment><!--.*?(?:-->|\Z))"
    r"|(?P<xml_tag></?[\w:.-]+|/?>|<\?[\w:.-]+|\?>)"
    r"|(?P<xml_attr>[\w:.-]+)(?=\s*=\s*[\"'])"
    r"|(?P<string>(?<==)\"[^\"]*\"|(?<==)'[^']*')",
    re.DOTALL,
)

LANGUAGES = {"vba": VBA_PATTERN, "xml": XML_PATTERN}


class Highlighter:
    """Recolours the visible part of one Text widget on request."""

    def __init__(self, text: tk.Text, language: str) -> None:
        self.text = text
        self.language = language
        self.pattern = LANGUAGES[language]

    def refresh(self) -> None:
        text = self.text
        first = int(text.index("@0,0").split(".")[0])
        last = int(text.index(f"@0,{max(text.winfo_height(), 1)}").split(".")[0])
        start = max(first - MARGIN_LINES, 1)
        if self.language == "xml":
            start = self._comment_start(start)
        end = f"{last + MARGIN_LINES}.end"
        begin = f"{start}.0"
        for tag in TAGS:
            text.tag_remove(tag, begin, end)
        chunk = text.get(begin, end)
        if not chunk:
            return
        line_starts = [0] + [m.end() for m in re.finditer("\n", chunk)]

        def index(offset: int) -> str:
            line = bisect.bisect_right(line_starts, offset) - 1
            return f"{start + line}.{offset - line_starts[line]}"

        for match in self.pattern.finditer(chunk):
            kind = match.lastgroup
            if kind is not None and match.end() > match.start():
                text.tag_add(kind, index(match.start()), index(match.end()))

    def _comment_start(self, line: int) -> int:
        """Back up to the line of an XML comment still open at ``line``."""
        text = self.text
        here = f"{line}.0"
        stop = f"{max(line - XML_COMMENT_LOOKBACK_LINES, 1)}.0"
        opened = text.search("<!--", here, backwards=True, stopindex=stop)
        if not opened:
            return line
        closed = text.search("-->", here, backwards=True, stopindex=opened)
        if closed:
            return line
        return int(opened.split(".")[0])
