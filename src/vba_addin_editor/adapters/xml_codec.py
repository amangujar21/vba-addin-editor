"""XML text codec: strict decode/encode, newline policy, well-formedness checks.

Parsing is used for VALIDATION ONLY (plan 5.6, 8): user text is written back as
exact bytes (BOM + newline policy + original encoding), never re-serialized by
ElementTree.
"""

from __future__ import annotations

import codecs
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass

_CONTENT_TYPES_PATH = "[content_types].xml"
_RELS_SUFFIX = ".rels"
_CONTENT_TYPES_ROOT = "{http://schemas.openxmlformats.org/package/2006/content-types}Types"
_RELS_ROOT = "{http://schemas.openxmlformats.org/package/2006/relationships}Relationships"

_DECL_RE = re.compile(rb'<\?xml[^>]*?encoding\s*=\s*["\']([A-Za-z0-9._-]+)["\']')
_DOCTYPE_RE = re.compile(r"<!\s*(DOCTYPE|ENTITY)", re.IGNORECASE)

_UTF8_BOM = b"\xef\xbb\xbf"
_UTF16LE_BOM = b"\xff\xfe"
_UTF16BE_BOM = b"\xfe\xff"


class XmlCodecError(ValueError):
    """Plain-language XML codec failure."""


@dataclass(frozen=True)
class DecodedXml:
    text: str  # LF-normalized, BOM stripped
    encoding: str  # canonical codec name for the payload (without BOM)
    bom: bytes
    newline: str  # dominant original newline: "\n" or "\r\n"


def _canonical_codec(name: str) -> str:
    try:
        codec = codecs.lookup(name)
    except LookupError as exc:
        raise XmlCodecError(
            f"This XML part declares an unsupported encoding ({name!r})."
        ) from exc
    canonical = codec.name  # e.g. "utf-8", "utf-16-le", "cp1252"
    if canonical == "ascii":
        canonical = "utf-8"  # ASCII content is valid UTF-8
    try:
        "".encode(canonical)
    except LookupError as exc:  # pragma: no cover - codec.name is always encodable
        raise XmlCodecError(f"Unsupported encoding {canonical!r}.") from exc
    return canonical


def _detect_encoding(raw: bytes) -> tuple[str, bytes]:
    if raw.startswith(_UTF8_BOM):
        return "utf-8-sig", _UTF8_BOM
    if raw.startswith(_UTF16LE_BOM):
        return "utf-16-le", _UTF16LE_BOM
    if raw.startswith(_UTF16BE_BOM):
        return "utf-16-be", _UTF16BE_BOM
    match = _DECL_RE.match(raw[:1024].lstrip()[:512])
    if match:
        declared = match.group(1).decode("ascii", errors="replace")
        return _canonical_codec(declared), b""
    return "utf-8", b""


def _dominant_newline(text: str) -> str:
    crlf = text.count("\r\n")
    lf = text.count("\n") - crlf
    return "\r\n" if crlf > lf else "\n"


def decode_xml(raw: bytes) -> DecodedXml:
    """Decode package XML bytes strictly; normalize editor text to LF."""
    encoding, bom = _detect_encoding(raw)
    payload = raw[len(bom):] if bom else raw
    try:
        if encoding == "utf-8-sig":
            text = payload.decode("utf-8")
        else:
            text = payload.decode(encoding)
    except (UnicodeDecodeError, LookupError) as exc:
        raise XmlCodecError(
            "This XML part is not decodable with its declared encoding; "
            "the file has not been changed."
        ) from exc
    newline = _dominant_newline(text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return DecodedXml(text=text, encoding=encoding, bom=bom, newline=newline)

def encode_xml(text: str, *, encoding: str, bom: bytes, newline: str) -> bytes:
    """Encode editor text to exact package bytes: BOM + newline policy + codec."""
    if bom and bom not in (_UTF8_BOM, _UTF16LE_BOM, _UTF16BE_BOM):
        raise XmlCodecError("Unknown byte-order mark for XML part.")
    payload = text.replace("\r\n", "\n").replace("\r", "\n")
    if newline == "\r\n":
        payload = payload.replace("\n", "\r\n")
    # Encode the payload with a BOM-less codec, then prepend the stored BOM so
    # BOM handling is fully explicit ("utf-8-sig" would otherwise add one).
    codec = "utf-8" if encoding == "utf-8-sig" else encoding
    try:
        data = payload.encode(codec)
    except (UnicodeEncodeError, LookupError) as exc:
        raise XmlCodecError(f"This XML part cannot be encoded as {encoding!r}.") from exc
    return bom + data


def declaration_conflict(text: str, encoding: str) -> str | None:
    """Return a problem string when the edited declaration contradicts the codec."""
    match = re.search(r'<\?xml[^>]*?encoding\s*=\s*["\']([A-Za-z0-9._-]+)["\']', text)
    if match is None:
        return None
    try:
        declared = codecs.lookup(match.group(1)).name
    except LookupError:
        return (
            f"This XML part declares unsupported encoding {match.group(1)!r}. "
            "Changing an XML part's character encoding is not supported in this version."
        )
    actual = codecs.lookup(encoding).name
    compatible = declared == actual or {declared, actual} <= {"ascii", "utf-8"}
    if not compatible:
        return (
            f"This XML part declares {declared}, but it was opened as {actual}. "
            "Changing an XML part's character encoding is not supported in this version."
        )
    return None


def validate_xml_bytes(raw: bytes, *, part_path: str) -> tuple[str, ...]:
    """Well-formedness + safe-shape checks. Empty tuple means valid."""
    problems: list[str] = []
    try:
        decoded = decode_xml(raw)
    except XmlCodecError as exc:
        return (str(exc),)
    if _DOCTYPE_RE.search(decoded.text):
        return (
            f"{part_path}: DOCTYPE/ENTITY declarations are not allowed in editable package XML.",
        )
    try:
        root = ET.fromstring(raw)
    except ET.ParseError as exc:
        line, col = exc.lineno, exc.offset
        where = f" (line {line}, column {col})" if line else ""
        problems.append(f"{part_path}: XML is not well-formed{where}: {exc.msg}")
        return tuple(problems)
    lowered = part_path.lower()
    if lowered == _CONTENT_TYPES_PATH and root.tag != _CONTENT_TYPES_ROOT:
        problems.append(
            f"{part_path}: unexpected root element {root.tag!r} "
            "(expected the OPC content-types Types element)."
        )
    elif lowered.endswith(_RELS_SUFFIX) and root.tag != _RELS_ROOT:
        problems.append(
            f"{part_path}: unexpected root element {root.tag!r} "
            "(expected the OPC Relationships element)."
        )
    return tuple(problems)


def validate_xml_draft(part) -> tuple[str, ...]:
    """Validate an XmlPartDraft the way it would be written to the package."""
    try:
        raw = encode_xml(part.text, encoding=part.encoding, bom=part.bom, newline=part.newline)
    except XmlCodecError as exc:
        return (f"{part.path}: {exc}",)
    conflict = declaration_conflict(part.text, part.encoding)
    if conflict:
        return (f"{part.path}: {conflict}",)
    return validate_xml_bytes(raw, part_path=part.path)
