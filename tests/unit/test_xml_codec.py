from __future__ import annotations

import pytest

from vba_addin_editor.adapters.xml_codec import (
    XmlCodecError,
    declaration_conflict,
    decode_xml,
    encode_xml,
    validate_xml_bytes,
    validate_xml_draft,
)


def _part(path, text, encoding="utf-8", bom=b"", newline="\n"):
    from vba_addin_editor.domain.document import XmlPartDraft

    return XmlPartDraft(
        path=path,
        text=text,
        original_text=text,
        encoding=encoding,
        bom=bom,
        newline=newline,
        is_relationships_part=path.lower().endswith(".rels"),
        is_content_types_part=path.lower() == "[content_types].xml",
    )


# -- decode/encode round trips ---------------------------------------------

def test_utf8_no_bom_roundtrip():
    d = decode_xml(b"<root/>")
    assert (d.text, d.encoding, d.bom, d.newline) == ("<root/>", "utf-8", b"", "\n")


def test_utf8_bom_preserved():
    raw = b"\xef\xbb\xbf<root/>"
    d = decode_xml(raw)
    assert d.encoding == "utf-8-sig" and d.bom == b"\xef\xbb\xbf"
    assert encode_xml(d.text, encoding=d.encoding, bom=d.bom, newline=d.newline) == raw


def test_utf16_le_bom_preserved():
    raw = b"\xff\xfe" + "<root/>".encode("utf-16-le")
    d = decode_xml(raw)
    assert d.encoding == "utf-16-le" and d.bom == b"\xff\xfe"
    assert encode_xml(d.text, encoding=d.encoding, bom=d.bom, newline=d.newline) == raw


def test_utf16_be_bom_preserved():
    raw = b"\xfe\xff" + "<root/>".encode("utf-16-be")
    d = decode_xml(raw)
    assert d.encoding == "utf-16-be" and d.bom == b"\xfe\xff"
    assert encode_xml(d.text, encoding=d.encoding, bom=d.bom, newline=d.newline) == raw


def test_declaration_encoding_detection():
    raw = b'<?xml version="1.0" encoding="ISO-8859-1"?><root a="\xe9"/>'
    d = decode_xml(raw)
    assert d.encoding == "iso8859-1"
    assert d.text == raw.decode("iso-8859-1")
    assert encode_xml(d.text, encoding=d.encoding, bom=d.bom, newline=d.newline) == raw


def test_invalid_declared_encoding_rejected():
    raw = b'<?xml version="1.0" encoding="NOT-A-CODEC"?><root/>'
    with pytest.raises(XmlCodecError):
        decode_xml(raw)


def test_crlf_normalized_and_restored():
    d = decode_xml(b"<root>\r\n<item/>\r\n</root>")
    assert d.text == "<root>\n<item/>\n</root>" and d.newline == "\r\n"
    out = encode_xml(d.text, encoding=d.encoding, bom=d.bom, newline=d.newline)
    assert out == b"<root>\r\n<item/>\r\n</root>"


def test_lf_preserved():
    d = decode_xml(b"<root>\n</root>")
    assert d.newline == "\n"


def test_mixed_newlines_use_dominant():
    d = decode_xml(b"<a>\r\n<b>\r\n<c>\n</c>\r\n</b>\r\n</a>")
    assert d.newline == "\r\n"


def test_undecodable_bytes_rejected():
    with pytest.raises(XmlCodecError):
        decode_xml(b"\xff\xfe\x00")  # truncated UTF-16 payload


# -- validation ---------------------------------------------------------------

def test_malformed_xml_reports_line_column():
    problems = validate_xml_bytes(b"<root>\n  <bad\n</root>", part_path="x.xml")
    assert problems and "line 3" in problems[0]


def test_doctype_rejected():
    problems = validate_xml_bytes(
        b'<!DOCTYPE root [<!ENTITY x "y">]><root>&x;</root>', part_path="x.xml"
    )
    assert problems and "DOCTYPE" in problems[0]


def test_content_types_root_accepted():
    raw = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        b"<Default Extension=\"rels\" ContentType=\"x\"/></Types>"
    )
    assert validate_xml_bytes(raw, part_path="[Content_Types].xml") == ()


def test_content_types_bad_root_rejected():
    assert validate_xml_bytes(b"<Wrong/>", part_path="[Content_Types].xml")


def test_rels_root_accepted_and_rejected():
    good = (
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        b"<Relationship Id=\"rId1\" Type=\"t\" Target=\"x\"/></Relationships>"
    )
    assert validate_xml_bytes(good, part_path="_rels/.rels") == ()
    assert validate_xml_bytes(b"<Other/>", part_path="_rels/.rels")


# -- draft validation ----------------------------------------------------------

def test_draft_encoding_conflict_rejected():
    part = _part(
        "a/customUI.xml",
        '<?xml version="1.0" encoding="UTF-16"?>\n<root/>',
        encoding="utf-8",
    )
    problems = validate_xml_draft(part)
    assert problems and "utf-16" in problems[0].lower() and "utf-8" in problems[0].lower()


def test_draft_undecodable_target_encoding_rejected():
    part = _part("a.xml", "<root>caf\u00e9</root>", encoding="ascii")
    assert validate_xml_draft(part)


def test_draft_valid_passes():
    part = _part("a.xml", '<?xml version="1.0" encoding="UTF-8"?>\n<root/>')
    assert validate_xml_draft(part) == ()


def test_declaration_conflict_helper():
    assert declaration_conflict("<root/>", "utf-8") is None
    assert declaration_conflict('<?xml version="1.0" encoding="utf-8"?><root/>', "utf-8") is None
    conflict = declaration_conflict('<?xml version="1.0" encoding="UTF-16"?><root/>', "utf-8")
    assert conflict
