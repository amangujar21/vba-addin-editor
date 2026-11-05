"""Replacing or orphaning a custom ribbon icon removes the old image from the package."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path

import pytest
from conftest import build_xlam_with_ribbon
from helpers import make_service
from PIL import Image

from vba_addin_editor.adapters.ooxml_package_adapter import OoxmlPackageAdapter
from vba_addin_editor.domain.changes import compute_changes
from vba_addin_editor.services import icon_service as ic
from vba_addin_editor.services import ribbon_service as rs
from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.services.history_service import restore_package, snapshot_package
from vba_addin_editor.services.recovery_service import RecoveryOpenService, RecoveryService
from vba_addin_editor.services.review_service import ReviewService
from vba_addin_editor.services.session_service import SessionService

PART = "customUI/customUI14.xml"
RELS = "customUI/_rels/customUI14.xml.rels"
OLD = "customUI/images/btnHello.png"


def _icon(color) -> ic.IconImage:
    out = io.BytesIO()
    Image.new("RGBA", (40, 40), color).save(out, "PNG")
    return ic.icon_from_bytes(out.getvalue())


RED, BLUE, GREEN = _icon((220, 0, 0, 255)), _icon((0, 0, 220, 255)), _icon((0, 160, 0, 255))


def _start(draft, control_id: str) -> int:
    text = draft.xml_part_by_path(PART).text
    element = rs.image_control_at(text, text.index(f'id="{control_id}"'))
    assert element is not None
    return element.start


def _members(path: Path) -> set[str]:
    with zipfile.ZipFile(path) as package:
        return set(package.namelist())


def _read(path: Path, name: str) -> bytes:
    with zipfile.ZipFile(path) as package:
        return package.read(name)


def _save(draft) -> None:
    result = make_service().save_addin(draft)
    assert result.kind == "success", result


@pytest.fixture()
def saved_icon(tmp_path: Path) -> Path:
    """An add-in whose btnHello already shows a saved custom icon (RED)."""
    path = build_xlam_with_ribbon(tmp_path / "work" / "Ribbon.xlam")
    draft = DocumentService().open(path)
    ic.set_custom_icon(draft, PART, _start(draft, "btnHello"), RED)
    _save(draft)
    assert OLD in _members(path)
    return path


def test_replacing_a_saved_icon_removes_the_old_image(saved_icon: Path):
    draft = DocumentService().open(saved_icon)
    new = ic.set_custom_icon(draft, PART, _start(draft, "btnHello"), BLUE)
    assert new != OLD
    assert draft.removed_member_paths() == [OLD]
    changes = compute_changes(draft)
    assert [c.path for c in changes.parts_removed] == [OLD]
    assert "Package — Removed" in changes.summary_lines()

    _save(draft)
    members = _members(saved_icon)
    assert OLD not in members and new in members
    assert _read(saved_icon, new) == BLUE.png
    rels = _read(saved_icon, RELS).decode("utf-8")
    assert len(rs.relationships(rels)) == 1 and "btnHello.png" not in rels
    assert not draft.removed_parts and not draft.is_dirty()

    # Replacing again reuses the freed name, and again leaves one image.
    again = DocumentService().open(saved_icon)
    third = ic.set_custom_icon(again, PART, _start(again, "btnHello"), GREEN)
    assert third == OLD and again.removed_member_paths() == [new]
    _save(again)
    images = sorted(m for m in _members(saved_icon) if m.startswith("customUI/images/"))
    assert images == [OLD] and _read(saved_icon, OLD) == GREEN.png


def test_icon_shared_with_another_control_is_kept(saved_icon: Path):
    draft = DocumentService().open(saved_icon)
    ribbon = draft.xml_part_by_path(PART)
    rid = rs.element_at(ribbon.text, _start(draft, "btnHello")).attr("image")
    ribbon.text = ribbon.text.replace(
        '<button id="btnMissing" label="Broken"',
        f'<button id="btnMissing" label="Broken" image="{rid}"',
    )
    _save(draft)

    draft = DocumentService().open(saved_icon)
    ic.set_custom_icon(draft, PART, _start(draft, "btnHello"), BLUE)
    assert draft.removed_parts == []  # btnMissing still shows it
    _save(draft)
    assert OLD in _members(saved_icon)
    assert rid in {r for r, _t, _g in rs.relationships(_read(saved_icon, RELS).decode("utf-8"))}


def test_replacing_a_pending_icon_adds_nothing_extra(tmp_path: Path):
    path = build_xlam_with_ribbon(tmp_path / "Ribbon.xlam")
    draft = DocumentService().open(path)
    ic.set_custom_icon(draft, PART, _start(draft, "btnHello"), RED)
    ic.set_custom_icon(draft, PART, _start(draft, "btnHello"), BLUE)
    assert [b.path for b in draft.binary_parts] == [OLD] and not draft.removed_parts
    _save(draft)
    assert _read(path, OLD) == BLUE.png


def test_remove_unused_images_cleans_up_orphans(saved_icon: Path):
    draft = DocumentService().open(saved_icon)
    # The control stops using its picture (as if edited by hand): an orphan.
    ribbon = draft.xml_part_by_path(PART)
    element = rs.element_at(ribbon.text, _start(draft, "btnHello"))
    tag = ribbon.text[element.start : element.open_end]
    ribbon.text = ribbon.text.replace(tag, tag.replace(f' image="{element.attr("image")}"', ""))
    assert [link.used for link in ic.image_links(draft)] == [False]

    assert ic.remove_unused_images(draft) == [OLD]
    assert ic.image_links(draft) == []
    assert ic.remove_unused_images(draft) == []
    _save(draft)
    assert OLD not in _members(saved_icon)


def test_undo_snapshot_restores_removals(saved_icon: Path):
    draft = DocumentService().open(saved_icon)
    before = snapshot_package(draft)
    ic.set_custom_icon(draft, PART, _start(draft, "btnHello"), BLUE)
    after = snapshot_package(draft)
    restore_package(draft, before)
    assert draft.removed_parts == []
    restore_package(draft, after)
    assert draft.removed_parts == [OLD]


def test_review_lists_the_removed_image(saved_icon: Path, tmp_path: Path):
    sessions = SessionService(session_root=tmp_path / "sessions")
    session = sessions.open(saved_icon)
    ic.set_custom_icon(session.draft, PART, _start(session.draft, "btnHello"), BLUE)
    model = ReviewService().build(session, operation_type="save")
    removed = [item for item in model.items if item.operations == ("delete",)]
    assert [item.target_id for item in removed] == [OLD]
    sessions.close(session)


def test_recovery_keeps_removals_and_rejects_non_images(saved_icon: Path, tmp_path: Path):
    root = tmp_path / "sessions"
    sessions = SessionService(session_root=root)
    recovery = RecoveryService(session_root=root)
    session = sessions.open(saved_icon)
    ic.set_custom_icon(session.draft, PART, _start(session.draft, "btnHello"), BLUE)
    data = recovery._serialize(session, cursor=None)
    assert data["draft"]["removed_parts"] == [OLD]
    for bad in ("xl/vbaProject.bin", "customUI/customUI14.xml", "../x.png"):
        data["draft"]["removed_parts"] = [bad]
        with pytest.raises(ValueError, match="path traversal"):
            recovery._validate_checkpoint(data, session.session_dir)

    session.bump()
    assert recovery.checkpoint(session) is None
    directory = session.session_dir
    sessions.close(session)
    opened = RecoveryOpenService(recovery, sessions).open_recoverable(directory)
    assert opened.status == "clean" and opened.session is not None
    assert opened.session.draft.removed_parts == [OLD]


def test_only_images_may_be_removed(saved_icon: Path):
    draft = DocumentService().open(saved_icon)
    draft.removed_parts.append("xl/styles.xml")
    result = make_service().save_addin(draft)
    assert result.kind != "success"
    assert "Only unused images" in (result.message or "")
    assert "xl/styles.xml" in _members(saved_icon)


def test_verifier_rejects_a_candidate_that_drops_more_than_the_draft(saved_icon: Path):
    adapter = OoxmlPackageAdapter()
    draft = DocumentService().open(saved_icon)
    reference = saved_icon.read_bytes()
    candidate = adapter.build_package(reference, {OLD: None})
    result = adapter.verify_candidate_xml(reference, candidate, draft)
    assert not result.ok and "ZIP entry set changed" in result.problems[0]
    draft.removed_parts.append(OLD)
    assert adapter.verify_candidate_xml(reference, candidate, draft).ok


def test_remove_relationship_keeps_the_rest_of_the_part():
    text, first = rs.add_relationship(rs.new_relationships_text(), rs.IMAGE_REL_TYPE, "images/a.png")
    text, second = rs.add_relationship(text, rs.IMAGE_REL_TYPE, "images/b.png")
    removed = rs.remove_relationship(text, first)
    assert [r for r, _t, _g in rs.relationships(removed)] == [second]
    assert removed.count("<Relationship ") == 1 and "  \n" not in removed
    assert rs.remove_relationship(removed, "rIdNope") == removed
    assert rs.relationship_target_path("customUI/customUI14.xml", "images/a.png") == "customUI/images/a.png"
    assert rs.relationship_target_path("customUI/customUI14.xml", "/img/a.png") == "img/a.png"
    assert rs.relationship_target_path("customUI/customUI14.xml", "../a.png") == "a.png"
    assert rs.relationship_target_path("customUI/x.xml", "https://example.com/a.png") is None
