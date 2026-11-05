"""Custom ribbon icons: PNG making, draft wiring, save/verify of added members."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest
from conftest import build_xlam_with_ribbon
from helpers import make_service
from PIL import Image

from vba_addin_editor.adapters.ooxml_package_adapter import OoxmlPackageAdapter, PackageError
from vba_addin_editor.domain.changes import compute_changes, revert_xml_part
from vba_addin_editor.domain.document import BinaryPartDraft
from vba_addin_editor.services import icon_service as ic
from vba_addin_editor.services import ribbon_service as rs
from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.services.history_service import restore_package, snapshot_package
from vba_addin_editor.services.recovery_service import RecoveryOpenService, RecoveryService
from vba_addin_editor.services.session_service import SessionService

PART = "customUI/customUI14.xml"
RELS = "customUI/_rels/customUI14.xml.rels"


def _png(width: int, height: int, color=(200, 30, 30, 255)) -> bytes:
    out = io.BytesIO()
    Image.new("RGBA", (width, height), color).save(out, "PNG")
    return out.getvalue()


def _start(draft, control_id: str) -> int:
    text = draft.xml_part_by_path(PART).text
    element = rs.image_control_at(text, text.index(f'id="{control_id}"'))
    assert element is not None
    return element.start


@pytest.fixture()
def ribbon_xlam(tmp_path: Path) -> Path:
    return build_xlam_with_ribbon(tmp_path / "work" / "Ribbon.xlam")


def test_make_icon_is_square_trimmed_and_small():
    # A wide picture with a transparent border becomes a centred 32x32 PNG.
    image = Image.new("RGBA", (900, 300), (0, 0, 0, 0))
    image.paste(Image.new("RGBA", (600, 200), (0, 120, 255, 255)), (150, 50))
    icon = ic.make_icon(image)
    decoded = Image.open(io.BytesIO(icon.png))
    assert decoded.size == (32, 32) and decoded.mode == "RGBA"
    assert decoded.getpixel((16, 16))[3] == 255  # visible centre
    assert decoded.getpixel((16, 1))[3] == 0  # padding stays transparent
    assert len(icon.png) < 4096
    assert ic.icon_from_bytes(_png(10, 10), 16).size == 16
    with pytest.raises(ic.IconError, match="not an image"):
        ic.icon_from_bytes(b"not an image")


def test_icon_is_added_saved_and_verified(ribbon_xlam: Path):
    draft = DocumentService().open(ribbon_xlam)
    vba_before = zipfile.ZipFile(ribbon_xlam).read("xl/vbaProject.bin")
    icon = ic.icon_from_bytes(_png(300, 120))
    path = ic.set_custom_icon(draft, PART, _start(draft, "btnHello"), icon)
    assert path == "customUI/images/btnHello.png"

    ribbon = draft.xml_part_by_path(PART).text
    assert '<button id="btnHello" label="Say Hello" size="large" onAction="Hello" image="rIdImg1" />' in ribbon
    changes = compute_changes(draft)
    assert [c.path for c in changes.parts_added] == [RELS, path]
    assert {c.path for c in changes.xml_modified} == {"[Content_Types].xml", PART}

    result = make_service().save_addin(draft)
    assert result.kind == "success", result
    with zipfile.ZipFile(ribbon_xlam) as package:
        assert package.read(path) == icon.png
        assert 'Target="images/btnHello.png"' in package.read(RELS).decode("utf-8")
        types = package.read("[Content_Types].xml").decode("utf-8")
        assert 'Extension="png" ContentType="image/png"' in types
        # An XML-only save keeps the VBA project byte-identical.
        assert package.read("xl/vbaProject.bin") == vba_before
    # After the save the new members are ordinary baseline parts.
    assert not draft.is_dirty() and not draft.binary_parts
    assert not draft.xml_part_by_path(RELS).is_new

    # A second icon reuses the relationships part and the png default.
    other = ic.set_custom_icon(draft, PART, _start(draft, "btnMissing"), icon)
    assert other == "customUI/images/btnMissing.png"
    assert [c.path for c in compute_changes(draft).parts_added] == [other]
    assert make_service().save_addin(draft).kind == "success"


def test_replacing_a_pending_icon_overwrites_it(ribbon_xlam: Path):
    draft = DocumentService().open(ribbon_xlam)
    start = _start(draft, "btnHello")
    first = ic.set_custom_icon(draft, PART, start, ic.icon_from_bytes(_png(8, 8)))
    blue = ic.icon_from_bytes(_png(8, 8, (0, 0, 255, 255)))
    assert ic.set_custom_icon(draft, PART, start, blue) == first
    assert [b.data for b in draft.binary_parts] == [blue.png]
    assert len(rs.relationships(draft.xml_part_by_path(RELS).text)) == 1


def test_icon_with_vba_change_saves_in_one_pass(ribbon_xlam: Path):
    draft = DocumentService().open(ribbon_xlam)
    ic.set_custom_icon(draft, PART, _start(draft, "btnHello"), ic.icon_from_bytes(_png(4, 4)))
    m1 = draft.find_current("Module1")
    m1.body = m1.body.replace("ORIGINAL", "V2")
    assert make_service().save_addin(draft).kind == "success"
    reopened = DocumentService().open(ribbon_xlam)
    assert "V2" in reopened.find_current("Module1").body
    assert reopened.xml_part_by_path(RELS) is not None


def test_icon_refused_on_package_signed_file(ribbon_xlam: Path):
    from dataclasses import replace

    from vba_addin_editor.domain.document import PackageSafetyInfo

    draft = DocumentService().open(ribbon_xlam)
    draft.baseline = replace(draft.baseline, package_safety=PackageSafetyInfo(True, ("_xmlsignatures/sig1.xml",)))
    with pytest.raises(ic.IconError, match="signature"):
        ic.set_custom_icon(draft, PART, _start(draft, "btnHello"), ic.icon_from_bytes(_png(4, 4)))
    assert not draft.is_dirty()


def test_icon_on_control_that_cannot_show_one(ribbon_xlam: Path):
    draft = DocumentService().open(ribbon_xlam)
    text = draft.xml_part_by_path(PART).text
    tabs = text.index("<tabs>")
    with pytest.raises(ic.IconError, match="cursor"):
        ic.set_custom_icon(draft, PART, tabs, ic.icon_from_bytes(_png(4, 4)))
    assert not draft.is_dirty()


def test_verification_rejects_tampered_or_clashing_members(ribbon_xlam: Path):
    adapter = OoxmlPackageAdapter()
    draft = DocumentService().open(ribbon_xlam)
    ic.set_custom_icon(draft, PART, _start(draft, "btnHello"), ic.icon_from_bytes(_png(4, 4)))
    reference = draft.baseline.package_bytes
    replacements = adapter.xml_replacements(draft.changed_xml_parts())
    added = adapter.added_members(draft)
    good = adapter.build_package(reference, replacements, added)
    assert adapter.verify_candidate_xml(reference, good, draft).ok

    tampered = dict(added, **{"customUI/images/btnHello.png": b"other"})
    bad = adapter.build_package(reference, replacements, tampered)
    result = adapter.verify_candidate_xml(reference, bad, draft)
    assert not result.ok and "does not match" in " ".join(result.problems)

    missing = {k: v for k, v in added.items() if not k.endswith(".png")}
    result = adapter.verify_candidate_xml(reference, adapter.build_package(reference, replacements, missing), draft)
    assert not result.ok and "entry set" in " ".join(result.problems)

    with pytest.raises(PackageError, match="already exist"):
        adapter.build_package(reference, {}, {"customUI/customUI14.xml": b"x"})
    with pytest.raises(PackageError, match="not allowed"):
        adapter.build_package(reference, {}, {"../evil.png": b"x"})


def test_revert_undo_snapshot_and_recovery_keep_added_parts(ribbon_xlam: Path, tmp_path: Path):
    root = tmp_path / "sessions"
    sessions = SessionService(session_root=root)
    recovery = RecoveryService(session_root=root)
    session = sessions.open(ribbon_xlam)
    draft = session.draft
    before = snapshot_package(draft)
    icon = ic.icon_from_bytes(_png(6, 6))
    ic.set_custom_icon(draft, PART, _start(draft, "btnHello"), icon)
    after = snapshot_package(draft)

    restore_package(draft, before)
    assert draft.new_xml_parts() == [] and draft.binary_parts == []
    restore_package(draft, after)
    assert [p.path for p in draft.new_xml_parts()] == [RELS]
    assert [b.data for b in draft.binary_parts] == [icon.png]

    session.bump()
    assert recovery.checkpoint(session) is None
    directory = session.session_dir
    sessions.close(session)
    opened = RecoveryOpenService(recovery, sessions).open_recoverable(directory)
    assert opened.status == "clean" and opened.session is not None
    restored = opened.session.draft
    assert [b.data for b in restored.binary_parts] == [icon.png]
    assert restored.xml_part_by_path(RELS).is_new

    revert_xml_part(restored, RELS)
    assert restored.xml_part_by_path(RELS) is None


def test_recovery_rejects_unsafe_added_part_names(ribbon_xlam: Path, tmp_path: Path):
    root = tmp_path / "sessions"
    sessions = SessionService(session_root=root)
    recovery = RecoveryService(session_root=root)
    session = sessions.open(ribbon_xlam)
    ic.set_custom_icon(session.draft, PART, _start(session.draft, "btnHello"), ic.icon_from_bytes(_png(4, 4)))
    data = recovery._serialize(session, cursor=None)
    recovery._validate_checkpoint(data, session.session_dir)  # the real one is fine
    data["draft"]["added_parts"][0]["path"] = "../x.png"
    with pytest.raises(ValueError, match="path traversal"):
        recovery._validate_checkpoint(data, session.session_dir)
    data["draft"]["added_parts"][0].update(path="customUI/images/x.png", data="not base64!")
    with pytest.raises(ValueError, match="invalid encoding"):
        recovery._validate_checkpoint(data, session.session_dir)
    sessions.close(session)


def test_added_part_size_limit(ribbon_xlam: Path):
    draft = DocumentService().open(ribbon_xlam)
    draft.binary_parts.append(BinaryPartDraft("customUI/images/big.png", b"x" * (2 * 1024 * 1024), "image/png"))
    result = make_service().save_addin(draft)
    assert result.kind != "success"
    assert "at most" in (result.message or "")
