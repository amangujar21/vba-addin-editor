from __future__ import annotations

from pathlib import Path

from helpers import edit_module1, make_service

from vba_addin_editor.services.conflict_service import ConflictService
from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.services.recovery_service import RecoveryService
from vba_addin_editor.services.session_service import SessionService


def _disk_edit(path: Path, token: str = "DISK") -> None:
    draft = DocumentService().open(path)
    m1 = next(m for m in draft.modules if m.current_name == "Module1")
    m1.body = m1.body.replace("ORIGINAL", token)
    assert make_service().save_addin(draft).kind == "success"


def test_divergent_edits_require_choice_and_keep_draft(work_xlam: Path, tmp_path: Path):
    sessions = SessionService(session_root=tmp_path / "sessions")
    session = sessions.open(work_xlam)
    edit_module1(session.draft)
    session.bump()
    _disk_edit(work_xlam)
    external, digest, error = ConflictService().snapshot_external(work_xlam)
    assert error is None
    assert external is not None
    assert digest is not None
    service = ConflictService()
    proposal = service.compare(session, external)
    body_items = [item for item in proposal.items if item.property_name == "body" and item.requires_choice]
    assert body_items
    body_items[0].resolution = "keep_draft"
    for item in proposal.items:
        if item.requires_choice and item.resolution is None:
            item.resolution = "keep_draft"
    new_draft = service.apply(session, proposal, external, live_external_hash=digest)
    m1 = next(m for m in new_draft.modules if m.current_name == "Module1")
    assert "V2" in m1.body
    assert "DISK" not in m1.body


def test_draft_addition_is_not_dropped(work_xlam: Path, tmp_path: Path):
    from vba_addin_editor.domain.document import ModuleDisplayKind, ModuleDraft, new_module_id

    sessions = SessionService(session_root=tmp_path / "sessions")
    session = sessions.open(work_xlam)
    session.draft.modules.append(
        ModuleDraft(
            id=new_module_id(),
            origin_name=None,
            current_name="AddedLocal",
            body="Public Sub X()\r\nEnd Sub\r\n",
            kind=ModuleDisplayKind.STANDARD,
            pyopenvba_kind="standard",
            is_new=True,
            is_deleted=False,
            destructive_ops_safe=True,
            can_delete=True,
            can_rename=True,
            project_item_kind="standard",
        )
    )
    session.bump()
    _disk_edit(work_xlam)
    external, digest, err = ConflictService().snapshot_external(work_xlam)
    assert external is not None and digest is not None and err is None
    service = ConflictService()
    proposal = service.compare(session, external)
    assert any(item.structural == "draft_add" for item in proposal.items)
    for item in proposal.items:
        if item.requires_choice and not item.resolution:
            item.resolution = "keep_disk" if item.property_name == "body" else "keep_draft"
        if item.structural == "draft_add":
            item.resolution = "keep_draft"
    new_draft = service.apply(session, proposal, external, live_external_hash=digest)
    names = {m.current_name for m in new_draft.modules if not m.is_deleted}
    assert "AddedLocal" in names


def test_stale_revision_and_external_hash_are_rejected(work_xlam: Path, tmp_path: Path):
    sessions = SessionService(session_root=tmp_path / "sessions")
    session = sessions.open(work_xlam)
    edit_module1(session.draft)
    session.bump()
    _disk_edit(work_xlam)
    external, digest, err = ConflictService().snapshot_external(work_xlam)
    assert external is not None and digest is not None and err is None
    service = ConflictService()
    proposal = service.compare(session, external)
    for item in proposal.items:
        if item.requires_choice:
            item.resolution = "keep_draft"
    session.bump()
    try:
        service.apply(session, proposal, external, live_external_hash=digest)
        raise AssertionError("expected stale_proposal")
    except ValueError as exc:
        assert str(exc) == "stale_proposal"
    session.revision = proposal.binding.revision
    try:
        service.apply(session, proposal, external, live_external_hash="deadbeef")
        raise AssertionError("expected stale_external")
    except ValueError as exc:
        assert str(exc) == "stale_external"


def test_cancel_leaves_session_unchanged(work_xlam: Path, tmp_path: Path):
    sessions = SessionService(session_root=tmp_path / "sessions")
    session = sessions.open(work_xlam)
    edit_module1(session.draft)
    session.bump()
    before = session.draft.modules[0].body
    rev = session.revision
    # Cancel is a UI no-op: compare must not mutate.
    _disk_edit(work_xlam)
    external, _digest, err = ConflictService().snapshot_external(work_xlam)
    assert external is not None and err is None
    ConflictService().compare(session, external)
    assert session.revision == rev
    assert session.draft.modules[0].body == before


def test_commit_adopts_external_baseline(work_xlam: Path, tmp_path: Path):
    root = tmp_path / "sessions"
    sessions = SessionService(session_root=root)
    recovery = RecoveryService(session_root=root)
    session = sessions.open(work_xlam)
    edit_module1(session.draft)
    session.bump()
    recovery.checkpoint(session)
    _disk_edit(work_xlam)
    external, digest, err = ConflictService().snapshot_external(work_xlam)
    assert external is not None and digest is not None and err is None
    service = ConflictService()
    proposal = service.compare(session, external)
    for item in proposal.items:
        if item.requires_choice:
            item.resolution = "keep_draft"
    result = service.commit(
        session,
        proposal,
        external,
        live_external_hash=digest,
        external_path=work_xlam,
        session_service=sessions,
        recovery_service=recovery,
    )
    assert result.kind == "success", result
    assert session.baseline_sha256 == digest
    assert session.captured_path.read_bytes() == work_xlam.read_bytes()
    assert not session.conflicts_pending
    m1 = next(m for m in session.draft.modules if m.current_name == "Module1")
    assert "V2" in m1.body
