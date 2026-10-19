"""Syntax colouring: what gets which tag, and that only the visible lines are paid for."""

from __future__ import annotations

import time
import tkinter as tk

import pytest

from vba_addin_editor.ui.highlight import MARGIN_LINES, Highlighter


@pytest.fixture()
def text():
    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    root.withdraw()
    widget = tk.Text(root, width=100, height=40)
    widget.pack()
    root.update()
    try:
        yield widget
    finally:
        root.destroy()


def _tagged(text: tk.Text, tag: str) -> list[str]:
    ranges = text.tag_ranges(tag)
    return [text.get(start, end) for start, end in zip(ranges[0::2], ranges[1::2])]


def _highlight(text: tk.Text, language: str, body: str) -> None:
    text.insert("1.0", body)
    text.update_idletasks()
    Highlighter(text, language).refresh()


def test_vba_tokens(text):
    _highlight(
        text,
        "vba",
        'Public Sub Go(ByVal n As Long)\n'
        '    msg = "it\'s ""quoted"""  \' a comment with "quotes"\n'
        '    rng.Select\n'
        '    x = &HFF + 12.5\n'
        'Rem old style comment\n'
        'End Sub\n',
    )
    assert _tagged(text, "keyword") == ["Public", "Sub", "ByVal", "As", "Long", "End", "Sub"]
    assert _tagged(text, "string") == ['"it\'s ""quoted"""']
    assert _tagged(text, "comment") == ["' a comment with \"quotes\"", "Rem old style comment"]
    assert _tagged(text, "number") == ["&HFF", "12.5"]


def test_xml_tokens_and_open_comment(text):
    body = "<!-- opened\n" + "still a comment\n" * (MARGIN_LINES + 60) + "-->\n"
    body += '<customUI xmlns="http://x"><button id="b1" onAction="Go"/></customUI>\n'
    text.insert("1.0", body)
    # Scroll so the comment's first line is far above the recoloured window.
    text.see("end")
    text.update_idletasks()
    Highlighter(text, "xml").refresh()
    assert _tagged(text, "xml_attr") == ["xmlns", "id", "onAction"]
    assert _tagged(text, "string") == ['"http://x"', '"b1"', '"Go"']
    assert "<customUI" in _tagged(text, "xml_tag")
    # The comment's closing line is still coloured as comment, not as text.
    closing = text.search("-->", "1.0")
    assert "comment" in text.tag_names(closing)


def test_refresh_cost_does_not_grow_with_module_size(text):
    line = "    total = total + Cells(i, 1).Value  ' running \"sum\"\n"
    text.insert("1.0", "Public Sub Big()\n" + line * 20000 + "End Sub\n")
    text.see("10000.0")
    text.update_idletasks()
    highlighter = Highlighter(text, "vba")
    started = time.perf_counter()
    for _ in range(5):
        highlighter.refresh()
    per_refresh_ms = (time.perf_counter() - started) * 1000 / 5
    # A few milliseconds in practice; the bound only catches whole-document colouring.
    assert per_refresh_ms < 60
    assert text.tag_nextrange("comment", "1.0", "9000.0") == ()
