"""Pure helpers of the Ribbon view, and the icon functions it uses."""

from __future__ import annotations

import io
from pathlib import Path

import pytest
from conftest import build_xlam_with_ribbon
from helpers import make_service
from PIL import Image

from vba_addin_editor.services import icon_service as ic
from vba_addin_editor.services import ribbon_service as rs
from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.ui.ribbon_designer import MAX_PEEK_LINES, drop_zone, procedure_source

PART = "customUI/customUI14.xml"
IMAGE = "customUI/images/btnHello.png"


def _png(color=(200, 30, 30, 255)) -> bytes:
    out = io.BytesIO()
    Image.new("RGBA", (8, 8), color).save(out, "PNG")
    return out.getvalue()


def _start(draft, control_id: str) -> int:
    text = draft.xml_part_by_path(PART).text
    return next(e.start for e in rs.element_tree(text) if e.attr("id") == control_id)


def _attrs(draft, control_id: str) -> dict[str, str | None]:
    element = next(e for e in rs.element_tree(draft.xml_part_by_path(PART).text) if e.attr("id") == control_id)
    return {"image": element.attr("image"), "imageMso": element.attr("imageMso")}


# -- procedure_source ----------------------------------------------------------------------


def test_procedure_source_stops_at_its_end():
    body = "Option Explicit\nSub A()\n    x = 1\nEnd Sub\nFunction B()\nEnd Function\n"
    assert procedure_source(body, 2) == ["Sub A()", "    x = 1", "End Sub"]
    assert procedure_source(body, 5) == ["Function B()", "End Function"]


def test_procedure_source_end_on_the_first_line_does_not_count():
    body = "Property Get P()\n  P = 1\n    End Property\nSub After()\nEnd Sub"
    assert procedure_source(body, 1) == ["Property Get P()", "  P = 1", "    End Property"]


def test_procedure_source_without_an_end_is_capped():
    body = "\n".join(["Sub Open()"] + ["    x = 1"] * (MAX_PEEK_LINES + 50))
    lines = procedure_source(body, 1)
    assert len(lines) == MAX_PEEK_LINES and lines[0] == "Sub Open()"
    assert procedure_source("Sub A()\nEnd Sub", 0)[0] == "Sub A()"  # line 0 reads as line 1


# -- drop_zone -------------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("fraction", "opened", "expected"),
    [(0.1, False, "before"), (0.5, False, "into"), (0.9, False, "after"), (0.9, True, "first")],
)
def test_drop_zone_on_a_row_that_holds_it_and_sits_beside_it(fraction, opened, expected):
    assert drop_zone(fraction, into=True, beside=True, opened=opened) == expected


def test_drop_zone_into_only_beside_only_and_nowhere():
    assert drop_zone(0.05, into=True, beside=False, opened=False) == "into"
    assert drop_zone(0.95, into=True, beside=False, opened=True) == "into"
    assert drop_zone(0.49, into=False, beside=True, opened=False) == "before"
    assert drop_zone(0.5, into=False, beside=True, opened=True) == "after"
    assert drop_zone(0.5, into=False, beside=False, opened=False) is None


# -- icon_service.image_bytes / clear_custom_icon --------------------------------------------


@pytest.fixture()
def ribbon_xlam(tmp_path: Path) -> Path:
    return build_xlam_with_ribbon(tmp_path / "work" / "Ribbon.xlam")


def test_image_bytes_from_the_draft_and_from_the_opened_package(ribbon_xlam: Path):
    draft = DocumentService().open(ribbon_xlam)
    assert ic.image_bytes(draft, PART, "rIdImg1") is None
    icon = ic.icon_from_bytes(_png())
    assert ic.set_custom_icon(draft, PART, _start(draft, "btnHello"), icon) == IMAGE
    assert ic.image_bytes(draft, PART, "rIdImg1") == icon.png  # pending, from the draft
    assert ic.image_bytes(draft, PART, "rIdNope") is None
    assert ic.image_bytes(draft, "customUI/customUI.xml", "rIdImg1") is None
    assert make_service().save_addin(draft).kind == "success"

    reopened = DocumentService().open(ribbon_xlam)
    assert not reopened.binary_parts
    assert ic.image_bytes(reopened, PART, "rIdImg1") == icon.png  # from the bytes kept at open


def test_clear_custom_icon_releases_a_pending_picture(ribbon_xlam: Path):
    draft = DocumentService().open(ribbon_xlam)
    ic.set_custom_icon(draft, PART, _start(draft, "btnHello"), ic.icon_from_bytes(_png()))
    ic.clear_custom_icon(draft, PART, _start(draft, "btnHello"), "HappyFace")
    assert _attrs(draft, "btnHello") == {"image": None, "imageMso": "HappyFace"}
    assert not draft.binary_parts
    assert ic.image_bytes(draft, PART, "rIdImg1") is None


def test_clear_custom_icon_removes_a_saved_picture_unless_shared(ribbon_xlam: Path):
    draft = DocumentService().open(ribbon_xlam)
    ic.set_custom_icon(draft, PART, _start(draft, "btnHello"), ic.icon_from_bytes(_png()))
    assert make_service().save_addin(draft).kind == "success"

    # Shared with another control: the picture stays in the file.
    shared = DocumentService().open(ribbon_xlam)
    part = shared.xml_part_by_path(PART)
    part.text = part.text.replace('onAction="NoSuchMacro"', 'onAction="NoSuchMacro" image="rIdImg1"')
    ic.clear_custom_icon(shared, PART, _start(shared, "btnHello"))
    assert _attrs(shared, "btnHello") == {"image": None, "imageMso": None}
    assert shared.removed_parts == []
    assert ic.image_bytes(shared, PART, "rIdImg1") is not None

    # Its only user: the picture leaves with the next save.
    alone = DocumentService().open(ribbon_xlam)
    ic.clear_custom_icon(alone, PART, _start(alone, "btnHello"))
    assert IMAGE in alone.removed_parts
    assert ic.image_bytes(alone, PART, "rIdImg1") is None
    assert make_service().save_addin(alone).kind == "success"


def test_clear_custom_icon_refuses_a_missing_control(ribbon_xlam: Path):
    draft = DocumentService().open(ribbon_xlam)
    with pytest.raises(ic.IconError, match="no longer"):
        ic.clear_custom_icon(draft, PART, _start(draft, "btnHello") + 1)
    with pytest.raises(ic.IconError, match="not editable"):
        ic.clear_custom_icon(draft, "customUI/missing.xml", 0)
