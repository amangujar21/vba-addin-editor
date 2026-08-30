from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from vba_addin_editor.adapters.ooxml_package_adapter import (
    CONTENT_TYPES_NAME,
    MAX_XML_PART_BYTES,
    OoxmlPackageAdapter,
    PackageError,
)
from vba_addin_editor.adapters.xml_codec import decode_xml
from vba_addin_editor.domain.document import PackageSafetyInfo, XmlPartDraft

CT_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    "</Types>"
)
RELS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="x" Target="a.xml"/></Relationships>'
)
SLIDE_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
)
SLIDE_XML_EDITED = SLIDE_XML.replace("/>", " edited='1'/>")

CT_SIG = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/_xmlsignatures/sig1.xml" '
    'ContentType="application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"/>'
    "</Types>"
)


def make_package(path: Path, entries: dict[str, str | bytes]) -> Path:
    with zipfile.ZipFile(path, "w") as zf:
        for name, data in entries.items():
            if isinstance(data, str):
                data = data.encode("utf-8")
            zf.writestr(name, data)
    return path


def base_entries() -> dict[str, str | bytes]:
    return {
        CONTENT_TYPES_NAME: CT_XML,
        "_rels/.rels": RELS_XML,
        "ppt/slides/slide1.xml": SLIDE_XML,
        "ppt/media/image1.png": b"\x89PNG\r\n\x1a\n-binary",
    }


@pytest.fixture()
def adapter():
    return OoxmlPackageAdapter()


# -- discovery -----------------------------------------------------------------

def test_discovers_xml_parts_and_ignores_binary(tmp_path, adapter):
    p = make_package(tmp_path / "a.pptm", base_entries())
    parts = adapter.snapshot_xml_parts(p)
    assert [x.path for x in parts] == [CONTENT_TYPES_NAME, "_rels/.rels", "ppt/slides/slide1.xml"]
    ct = parts[0]
    assert ct.is_content_types_part and not ct.is_relationships_part
    assert parts[1].is_relationships_part
    assert parts[0].original_size > 0 and parts[0].original_sha256


def test_part_metadata_roundtrip(tmp_path, adapter):
    entries = base_entries() | {"ppt/slides/slide2.xml": SLIDE_XML + "\r\n\r\n"}
    p = make_package(tmp_path / "a.pptm", entries)
    parts = {x.path: x for x in adapter.snapshot_xml_parts(p)}
    assert parts["ppt/slides/slide2.xml"].newline == "\r\n"
    assert parts["ppt/slides/slide1.xml"].newline == "\n"


def test_duplicate_member_names_rejected(tmp_path, adapter):
    path = tmp_path / "dup.pptm"
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("ppt/a.xml", SLIDE_XML)
        zf.writestr("ppt/a.xml", SLIDE_XML)
    with pytest.raises(PackageError):
        adapter.snapshot_xml_parts(path)


@pytest.mark.parametrize(
    ("payload", "expected_text"),
    [
        (b"", ""),
        (b"<not-well-formed", "<not-well-formed"),
        (b"\xff\xfe\x00", None),
    ],
)
def test_invalid_existing_xml_part_is_preserved_as_read_only_snapshot(
    tmp_path, adapter, payload, expected_text
):
    entries = base_entries() | {"ppt/presentation.xml": payload}
    p = make_package(tmp_path / "a.pptm", entries)

    parts = {part.path: part for part in adapter.snapshot_xml_parts(p)}

    malformed = parts["ppt/presentation.xml"]
    assert malformed.text == expected_text
    assert not malformed.editable
    assert not malformed.well_formed_on_open
    assert malformed.open_problem
    assert malformed.original_size == len(payload)
    assert parts["ppt/slides/slide1.xml"].editable
    assert parts["ppt/slides/slide1.xml"].well_formed_on_open
    assert parts["ppt/slides/slide1.xml"].open_problem is None


def test_oversized_part_blocked(tmp_path, adapter, monkeypatch):
    monkeypatch.setattr(
        "vba_addin_editor.adapters.ooxml_package_adapter.MAX_XML_PART_BYTES", 8
    )
    p = make_package(tmp_path / "a.pptm", base_entries())
    with pytest.raises(PackageError):
        adapter.snapshot_xml_parts(p)


# -- signature detection ----------------------------------------------------------

def test_signature_markers_detected(tmp_path, adapter):
    entries = base_entries() | {
        CONTENT_TYPES_NAME: CT_SIG,
        "_xmlsignatures/sig1.xml": SLIDE_XML,
    }
    p = make_package(tmp_path / "signed.pptm", entries)
    info = adapter.inspect_package_signature(p)
    assert info.opc_signature_present
    assert set(info.signature_part_names) == {"_xmlsignatures/sig1.xml", "[Content_Types].xml"}


def test_unsigned_package_has_no_signature(tmp_path, adapter):
    p = make_package(tmp_path / "clean.pptm", base_entries())
    assert adapter.inspect_package_signature(p) == PackageSafetyInfo(False, ())


# -- candidate rewrite --------------------------------------------------------------

def _draft_part(path: str, text: str, newline: str = "\n") -> XmlPartDraft:
    decoded = decode_xml(text.encode("utf-8"))
    return XmlPartDraft(
        path=path,
        text=text,
        original_text=decoded.text + "\n<!-- draft differs -->",  # ensures is_dirty()
        encoding="utf-8",
        bom=b"",
        newline=newline,
        is_relationships_part=False,
        is_content_types_part=False,
    )


def test_xml_only_rewrite_changes_exactly_one_payload(tmp_path, adapter):
    src = make_package(tmp_path / "src.pptm", base_entries())
    dest = tmp_path / "cand.pptm"
    part = _draft_part("ppt/slides/slide1.xml", SLIDE_XML_EDITED)
    adapter.write_xml_candidate(src, dest, [part])
    with zipfile.ZipFile(src) as a, zipfile.ZipFile(dest) as b:
        assert set(a.namelist()) == set(b.namelist())
        for name in a.namelist():
            if name == part.path:
                assert b.read(name) != a.read(name)
            else:
                assert a.read(name) == b.read(name), name
        assert b.testzip() is None


def test_rewrite_preserves_entry_metadata(tmp_path, adapter):
    src = tmp_path / "meta.pptm"
    with zipfile.ZipFile(src, "w") as zf:
        info = zipfile.ZipInfo("ppt/a.xml", date_time=(2020, 1, 2, 3, 4, 5))
        info.compress_type = zipfile.ZIP_DEFLATED
        info.comment = b"hello"
        zf.writestr(info, SLIDE_XML)
        zf.comment = b"archive-comment"
    dest = tmp_path / "meta2.pptm"
    adapter.write_xml_candidate(src, dest, [_draft_part("ppt/a.xml", SLIDE_XML)])
    with zipfile.ZipFile(dest) as zf:
        assert zf.comment == b"archive-comment"
        assert zf.getinfo("ppt/a.xml").date_time == (2020, 1, 2, 3, 4, 4)  # 2s DOS resolution


def test_rewrite_rejects_duplicate_source(tmp_path, adapter):
    src = tmp_path / "dup.pptm"
    with zipfile.ZipFile(src, "w") as zf:
        zf.writestr("ppt/a.xml", SLIDE_XML)
        zf.writestr("ppt/a.xml", SLIDE_XML)
    with pytest.raises(PackageError):
        adapter.write_xml_candidate(src, tmp_path / "out.pptm", [_draft_part("ppt/a.xml", SLIDE_XML)])


def test_rewrite_rejects_malformed_changed_xml(tmp_path, adapter):
    src = make_package(tmp_path / "src.pptm", base_entries())
    with pytest.raises(PackageError):
        adapter.write_xml_candidate(
            src, tmp_path / "out.pptm", [_draft_part("ppt/slides/slide1.xml", "<broken")]
        )


def test_rewrite_rejects_missing_part(tmp_path, adapter):
    src = make_package(tmp_path / "src.pptm", base_entries())
    with pytest.raises(PackageError):
        adapter.write_xml_candidate(src, tmp_path / "out.pptm", [_draft_part("nope.xml", SLIDE_XML)])


# -- verification ----------------------------------------------------------------

def _draft_with_parts(package: Path, changed: list[XmlPartDraft]):

    from vba_addin_editor.domain.document import DocumentDraft, DocumentSnapshot

    fp = type("FP", (), {"sha256": "", "size": 0, "mtime_ns": 0})()
    snap = DocumentSnapshot(
        path=package,
        host_kind="powerpoint",
        extension=".pptm",
        file_fingerprint=fp,  # type: ignore[arg-type]
        project_name="T",
        code_page=1252,
        safety=None,  # type: ignore[arg-type]
        modules=(),
        xml_parts=(),
    )
    return DocumentDraft(baseline=snap, modules=[], xml_parts=list(changed))


def test_verifier_accepts_exact_change(tmp_path, adapter):
    src = make_package(tmp_path / "src.pptm", base_entries())
    dest = tmp_path / "cand.pptm"
    part = _draft_part("ppt/slides/slide1.xml", SLIDE_XML.replace("<p:sld/>", "<p:sld edited='1'/>"))
    adapter.write_xml_candidate(src, dest, [part])
    draft = _draft_with_parts(src, [part])
    result = adapter.verify_candidate_xml(src, dest, draft)
    assert result.ok, result.problems


def test_verifier_rejects_malformed_xml_in_declared_change_set(tmp_path, adapter):
    src = make_package(tmp_path / "src.pptm", base_entries())
    malformed = b"<broken"
    candidate_entries = base_entries() | {"ppt/slides/slide1.xml": malformed}
    dest = make_package(tmp_path / "cand.pptm", candidate_entries)
    part = _draft_part("ppt/slides/slide1.xml", malformed.decode("utf-8"))
    draft = _draft_with_parts(src, [part])

    result = adapter.verify_candidate_xml(src, dest, draft)

    assert not result.ok
    assert any("not well-formed" in problem for problem in result.problems)


def test_verifier_rejects_unexpected_xml_change(tmp_path, adapter):
    src = make_package(tmp_path / "src.pptm", base_entries())
    dest = tmp_path / "cand.pptm"
    part = _draft_part("ppt/slides/slide1.xml", SLIDE_XML)
    adapter.write_xml_candidate(src, dest, [part])
    # Tamper with an untouched XML part after the rewrite.
    with zipfile.ZipFile(dest) as zf:
        payloads = {i.filename: zf.read(i.filename) for i in zf.infolist()}
    payloads["_rels/.rels"] = RELS_XML.replace("a.xml", "b.xml")
    with zipfile.ZipFile(dest, "w") as zf:
        for name, data in payloads.items():
            zf.writestr(name, data)
    draft = _draft_with_parts(src, [part])
    result = adapter.verify_candidate_xml(src, dest, draft)
    assert not result.ok
    assert any("_rels/.rels" in p for p in result.problems)


def test_verifier_rejects_mutation_of_untouched_malformed_xml(tmp_path, adapter):
    malformed_path = "ppt/presentation.xml"
    src_entries = base_entries() | {malformed_path: b""}
    candidate_entries = base_entries() | {malformed_path: b"x"}
    src = make_package(tmp_path / "src.pptm", src_entries)
    dest = make_package(tmp_path / "cand.pptm", candidate_entries)
    draft = _draft_with_parts(src, [])

    result = adapter.verify_candidate_xml(src, dest, draft)

    assert not result.ok
    assert any(malformed_path in problem for problem in result.problems)


def test_xml_verifier_scope_excludes_binary_parts(tmp_path, adapter):
    """Binary isolation is the pyOpenVBA verifier's job (plan 12.9); the XML
    verifier covers editable XML parts and vbaProject.bin only. The composite
    save pipeline runs both layers."""
    src = make_package(tmp_path / "src.pptm", base_entries())
    dest = tmp_path / "cand.pptm"
    part = _draft_part("ppt/slides/slide1.xml", SLIDE_XML)
    adapter.write_xml_candidate(src, dest, [part])
    with zipfile.ZipFile(dest) as zf:
        payloads = {i.filename: zf.read(i.filename) for i in zf.infolist()}
    payloads["ppt/media/image1.png"] = b"TAMPERED"
    with zipfile.ZipFile(dest, "w") as zf:
        for name, data in payloads.items():
            zf.writestr(name, data)
    draft = _draft_with_parts(src, [part])
    result = adapter.verify_candidate_xml(src, dest, draft)
    assert result.ok, result.problems


def test_verifier_rejects_vba_change_on_xml_only_save(tmp_path, adapter):
    src = make_package(tmp_path / "src.pptm", base_entries() | {"ppt/vbaProject.bin": b"VBA1"})
    dest = tmp_path / "cand.pptm"
    part = _draft_part("ppt/slides/slide1.xml", SLIDE_XML)
    adapter.write_xml_candidate(src, dest, [part])
    with zipfile.ZipFile(dest) as zf:
        payloads = {i.filename: zf.read(i.filename) for i in zf.infolist()}
    payloads["ppt/vbaProject.bin"] = b"VBA2"
    with zipfile.ZipFile(dest, "w") as zf:
        for name, data in payloads.items():
            zf.writestr(name, data)
    draft = _draft_with_parts(src, [part])
    result = adapter.verify_candidate_xml(src, dest, draft)
    assert not result.ok
    assert any("vbaProject.bin" in p for p in result.problems)


def test_verifier_flags_missing_changed_part(tmp_path, adapter):
    src = make_package(tmp_path / "src.pptm", base_entries())
    draft = _draft_with_parts(src, [_draft_part("ppt/missing.xml", SLIDE_XML)])
    result = adapter.verify_candidate_xml(src, src, draft)
    assert not result.ok


def test_max_total_constant_sane():
    assert MAX_XML_PART_BYTES > 0
