"""Text fidelity helpers: newline policy, header/body split, strict code-page checks.

Pure functions — no pyOpenVBA imports here so domain/services can use them freely.
"""

from __future__ import annotations

from dataclasses import dataclass


def to_vba_crlf(text: str) -> str:
    """Normalize any newline style to CRLF (adapter boundary)."""
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")


def to_editor_text(text: str) -> str:
    """Normalize any newline style to LF (Tk editor boundary)."""
    return text.replace("\r\n", "\n").replace("\r", "\n")


_CLASS_PREAMBLE = (
    "VERSION 1.0 CLASS\n"
    "BEGIN\n"
    "  MultiUse = -1  'True\n"
    "END\n"
)


def replace_vb_name(header: str, current_name: str) -> str:
    """Rewrite Attribute VB_Name in a hidden header; keep other attributes."""
    lines = header.split("\n")
    found = False
    out: list[str] = []
    for line in lines:
        stripped = line.strip()
        if stripped.lower().startswith("attribute vb_name"):
            out.append(f'Attribute VB_Name = "{current_name}"')
            found = True
        else:
            out.append(line)
    if not found and header.strip():
        out.append(f'Attribute VB_Name = "{current_name}"')
    return "\n".join(out)


def synthesize_class_header(current_name: str) -> str:
    return (
        f"{_CLASS_PREAMBLE}"
        f'Attribute VB_Name = "{current_name}"\n'
        "Attribute VB_GlobalNameSpace = False\n"
        "Attribute VB_Creatable = False\n"
        "Attribute VB_PredeclaredId = False\n"
        "Attribute VB_Exposed = False"
    )


def synthesize_standard_header(current_name: str) -> str:
    return f'Attribute VB_Name = "{current_name}"'


def join_header_body(header: str, body: str) -> str:
    header = header.replace("\r\n", "\n").replace("\r", "\n").rstrip("\n")
    body = body.replace("\r\n", "\n").replace("\r", "\n")
    if not header:
        return body
    if not body:
        return header + "\n"
    return header + "\n" + body


def compose_module_source(
    *,
    body: str,
    hidden_header: str,
    current_name: str,
    kind: str,
    pyopenvba_kind: str = "standard",
) -> str:
    """Compose exportable source from current body and effective attributes.

    Hidden attributes stay separate from the visible body. Procedure-level
    Attribute lines that belong in the body are not moved into the header.
    """
    header = hidden_header or ""
    is_class = kind == "class" or pyopenvba_kind == "other" or "VERSION 1.0 CLASS" in header
    if is_class:
        header = replace_vb_name(header, current_name) if header.strip() else synthesize_class_header(current_name)
        if "VERSION 1.0 CLASS" not in header:
            header = synthesize_class_header(current_name)
            if hidden_header.strip():
                extra = "\n".join(
                    line
                    for line in hidden_header.split("\n")
                    if line.strip()
                    and not line.strip().lower().startswith("attribute vb_name")
                    and not line.strip().startswith("VERSION ")
                )
                if extra:
                    header = header + "\n" + extra
    elif header.strip():
        header = replace_vb_name(header, current_name)
    else:
        header = synthesize_standard_header(current_name)
    return join_header_body(header, body)


def split_attribute_header(source: str) -> tuple[str, str]:
    """Split leading ``Attribute VB_*`` lines (and a class ``VERSION`` preamble)
    from the visible body. Mirrors pyopenvba.vba.split_attribute_header without
    importing it, so this stays pure text logic.

    Returns (header, body); header is "" when no hidden attributes exist.
    """
    lines = source.split("\n")
    header_end = 0
    in_version_preamble = False
    for i, line in enumerate(lines):
        stripped = line.strip()
        if in_version_preamble:
            header_end = i + 1
            if stripped == "BEGIN":
                continue
            if stripped.startswith("VB_"):
                continue
            if in_version_preamble and stripped.startswith("END"):
                in_version_preamble = False
            continue
        if stripped.startswith("VERSION "):
            in_version_preamble = True
            header_end = i + 1
            continue
        if stripped.startswith("Attribute VB_"):
            header_end = i + 1
            continue
        break
    # Trim a single separating blank line between header and body.
    header_lines = lines[:header_end]
    while header_lines and header_lines[-1].strip() == "":
        header_lines.pop()
    header = "\n".join(header_lines)
    body = "\n".join(lines[header_end:])
    if body.startswith(("\r", "\n")):
        body = body[1:]
    return header, body


@dataclass(frozen=True)
class EncodingProblem:
    module_name: str
    char: str
    position: tuple[int, int]  # 1-based line, column within the body


class CodePageError(ValueError):
    def __init__(self, problems: list[EncodingProblem], code_page: int) -> None:
        self.problems = problems
        self.code_page = code_page
        first = problems[0]
        super().__init__(
            f"This VBA project uses code page {code_page}. The character "
            f"{first.char!r} in module {first.module_name!r} (line {first.position[0]}, "
            f"column {first.position[1]}) cannot be stored in that project. "
            "Replace it before saving."
        )


def validate_code_page(
    module_sources: dict[str, str], code_page: int, encoding: str
) -> None:
    """Strict-encode every changed module body; raise CodePageError on failure."""
    problems: list[EncodingProblem] = []
    for name, body in module_sources.items():
        try:
            body.encode(encoding, errors="strict")
            continue
        except UnicodeEncodeError:
            pass
        for line_no, line in enumerate(to_vba_crlf(body).split("\r\n"), start=1):
            for col_no, ch in enumerate(line, start=1):
                try:
                    ch.encode(encoding, errors="strict")
                except UnicodeEncodeError:
                    problems.append(EncodingProblem(name, ch, (line_no, col_no)))
                    if len(problems) >= 10:
                        raise CodePageError(problems, code_page)
    if problems:
        raise CodePageError(problems, code_page)
