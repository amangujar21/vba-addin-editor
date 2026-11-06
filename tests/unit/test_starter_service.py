"""New blank add-in files, and a first ribbon for a file that has none."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from conftest import build_xlam, build_xlam_with_ribbon
from helpers import make_service

from vba_addin_editor.domain.document import draft_from_snapshot
from vba_addin_editor.services import ribbon_service as rs
from vba_addin_editor.services import starter_service as ss
from vba_addin_editor.services.document_service import DocumentService

SPEC = ss.StarterRibbonSpec("Tools & More", "Main")


def _content_types(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        return package.read("[Content_Types].xml").decode("utf-8")


@pytest.mark.parametrize("extension", [".xlam", ".ppam", ".pptm"])
def test_blank_file_opens_and_has_no_ribbon(tmp_path: Path, extension: str):
    path = tmp_path / f"New{extension}"
    data, fp = ss.create_blank_file(path)
    assert path.read_bytes() == data
    assert fp.written_identity is not None and fp.size == len(data)
    draft = DocumentService().open(path)
    assert any(m.current_name == "Module1" and m.pyopenvba_kind == "standard" for m in draft.modules)
    assert rs.ribbon_parts(draft) == []
    assert ss.ribbon_blocker(draft) is None


def test_ppam_is_typed_as_an_add_in(tmp_path: Path):
    ppam, pptm = tmp_path / "A.ppam", tmp_path / "A.pptm"
    ss.create_blank_file(ppam)
    ss.create_blank_file(pptm)
    main = "/ppt/presentation.xml"
    assert rs.override_content_type(_content_types(ppam), main) == (
        "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml"
    )
    assert rs.override_content_type(_content_types(pptm), main) == (
        "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
    )
    # Only the content types differ: every other member is the PowerPoint template's.
    with zipfile.ZipFile(ppam) as a, zipfile.ZipFile(pptm) as b:
        assert a.namelist() == b.namelist()
        for name in a.namelist():
            if name != "[Content_Types].xml":
                assert a.read(name) == b.read(name), name


def test_create_never_replaces_an_existing_file(tmp_path: Path):
    path = tmp_path / "Taken.xlam"
    path.write_bytes(b"keep me")
    with pytest.raises(ss.StarterError, match="already exists"):
        ss.create_blank_file(path)
    assert path.read_bytes() == b"keep me"
    with pytest.raises(ss.StarterError, match="Only"):
        ss.create_blank_file(tmp_path / "Book.xlsm")
    assert not (tmp_path / "Book.xlsm").exists()


@pytest.mark.parametrize("extension", [".xlam", ".ppam", ".pptm"])
def test_starter_ribbon_saves_and_links_from_the_package_root(tmp_path: Path, extension: str):
    path = tmp_path / f"New{extension}"
    ss.create_blank_file(path)
    draft = DocumentService().open(path)
    vba_entry = "xl/vbaProject.bin" if extension == ".xlam" else "ppt/vbaProject.bin"
    with zipfile.ZipFile(path) as package:
        project_before = package.read(vba_entry)
        types_before = package.read("[Content_Types].xml")

    part = ss.add_starter_ribbon(draft, SPEC)
    assert part.path == ss.RIBBON_PART_PATH and part.is_new
    assert draft.added_member_paths() == [ss.RIBBON_PART_PATH]
    containers = rs.find_containers(part.text, part.path)
    assert [(c.element, c.control_id, c.display) for c in containers] == [
        ("group", "grpMain", "Tools & More › Main")
    ]
    assert [t.control_id for t in rs.find_tabs(part.text, part.path)] == ["tabToolsMore"]

    result = make_service().save_addin(draft)
    assert result.kind == "success", result
    with zipfile.ZipFile(path) as package:
        rels = package.read("_rels/.rels").decode("utf-8")
        assert (
            "rIdUI1", ss.UI_EXTENSIBILITY_REL_TYPE, "customUI/customUI14.xml"
        ) in rs.relationships(rels)
        assert package.read(ss.RIBBON_PART_PATH).decode("utf-8").startswith("<?xml")
        # .xml is already application/xml in Office's templates: no content type edit.
        assert package.read("[Content_Types].xml") == types_before
        # An XML-only save leaves the VBA project byte-identical.
        assert package.read(vba_entry) == project_before
    reopened = DocumentService().open(path)
    assert [p.path for p in rs.ribbon_parts(reopened)] == [ss.RIBBON_PART_PATH]
    assert ss.ribbon_blocker(reopened) == "This file already has ribbon XML (customUI)."


def test_button_can_go_into_the_starter_group(tmp_path: Path):
    path = build_xlam(tmp_path / "Plain.xlam")
    draft = DocumentService().open(path)
    part = ss.add_starter_ribbon(draft, ss.StarterRibbonSpec())
    group = rs.find_containers(part.text, part.path)[0]
    element = rs.build_button_xml(control_id="btnHello", label="Hello", on_action="Hello")
    part.text, _offset, _length = rs.insert_into_container(part.text, group, element)
    assert (
        '        <group id="grpMacros" label="Macros">\n'
        '          <button id="btnHello" label="Hello" onAction="Hello" />\n'
        "        </group>\n"
    ) in part.text
    result = make_service().save_addin(draft)
    assert result.kind == "success", result
    with zipfile.ZipFile(path) as package:
        assert b'\r\n          <button id="btnHello"' in package.read(ss.RIBBON_PART_PATH)


def test_ribbon_blockers(tmp_path: Path):
    draft = DocumentService().open(build_xlam_with_ribbon(tmp_path / "R.xlam"))
    assert "already has ribbon XML" in ss.ribbon_blocker(draft)
    with pytest.raises(ss.StarterError):
        ss.add_starter_ribbon(draft, SPEC)

    plain = DocumentService().open(build_xlam(tmp_path / "P.xlam"))
    rels = plain.xml_part_by_path("_rels/.rels")
    rels.text = rels.text.replace(
        "</Relationships>",
        f'<Relationship Id="rIdX" Type="{ss.UI_EXTENSIBILITY_REL_TYPE}" '
        'Target="customUI/gone.xml"/></Relationships>',
    )
    assert "gone.xml" in ss.ribbon_blocker(plain)

    signed = draft_from_snapshot(plain.baseline)
    signed.baseline = signed.baseline.__class__(
        **{
            **signed.baseline.__dict__,
            "package_safety": signed.baseline.package_safety.__class__(True, ("x",)),
        }
    )
    assert "signature" in ss.ribbon_blocker(signed)


def test_existing_link_to_the_part_is_reused_and_content_type_added_when_needed(tmp_path: Path):
    draft = DocumentService().open(build_xlam(tmp_path / "P.xlam"))
    rels = draft.xml_part_by_path("_rels/.rels")
    rels.text = rels.text.replace(
        "</Relationships>",
        f'<Relationship Id="rIdOld" Type="{ss.UI_EXTENSIBILITY_REL_TYPE}" '
        'Target="customUI/customUI14.xml"/></Relationships>',
    )
    linked = rels.text
    types = draft.xml_part_by_path("[Content_Types].xml")
    types.text = types.text.replace(
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Default Extension="xml" ContentType="text/xml"/>',
    )
    ss.add_starter_ribbon(draft, SPEC)
    assert rels.text == linked
    assert rs.override_content_type(types.text, "/customUI/customUI14.xml") == "application/xml"


def test_spec_validation():
    assert ss.validate_spec(ss.StarterRibbonSpec("", "G")) == "Enter a label for the tab."
    assert ss.validate_spec(ss.StarterRibbonSpec("T", " ")) == "Enter a label for the group."
    assert ss.validate_spec(ss.StarterRibbonSpec("T", "G", tab_id="1bad")) is not None
    assert ss.validate_spec(ss.StarterRibbonSpec("T", "G", tab_id="x", group_id="X")) is not None
    assert ss.StarterRibbonSpec("Tools", "Tools").ids() == ("tabTools", "grpTools")
    xml = ss.starter_ribbon_xml(ss.StarterRibbonSpec('A "quoted" <tab>', "G"))
    assert 'label="A &quot;quoted&quot; &lt;tab&gt;"' in xml
