from __future__ import annotations

import pytest

from vba_addin_editor.adapters.source_codec import (
    CodePageError,
    split_attribute_header,
    to_editor_text,
    to_vba_crlf,
    validate_code_page,
)


def test_crlf_normalization():
    assert to_vba_crlf("a\nb\rc\r\nd") == "a\r\nb\r\nc\r\nd"


def test_editor_normalization():
    assert to_editor_text("a\r\nb\rc") == "a\nb\nc"


def test_split_standard_header():
    src = 'Attribute VB_Name = "Module1"\nOption Explicit\nSub X()\nEnd Sub\n'
    header, body = split_attribute_header(src)
    assert header == 'Attribute VB_Name = "Module1"'
    assert body == "Option Explicit\nSub X()\nEnd Sub\n"


def test_split_class_header():
    src = (
        "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1\nEND\n"
        'Attribute VB_Name = "Class1"\nAttribute VB_GlobalNameSpace = False\n'
        "Public X As Long\n"
    )
    header, body = split_attribute_header(src)
    assert "VERSION 1.0 CLASS" in header and "VB_Name" in header
    assert body == "Public X As Long\n"


def test_split_no_header():
    header, body = split_attribute_header("Option Explicit\n")
    assert header == ""
    assert body == "Option Explicit\n"


def test_codepage_strict_failure_1252():
    with pytest.raises(CodePageError) as ei:
        validate_code_page({"Mod1": "x = \"\u2603\""}, 1252, "cp1252")
    assert "Mod1" in str(ei.value)


def test_codepage_pass():
    validate_code_page({"Mod1": "s = \"caf\xc3\xa9\"".replace("Ã©", "é")}, 1252, "cp1252")
    validate_code_page({"Mod1": "ok"}, 1252, "cp1252")
