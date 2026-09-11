from __future__ import annotations

from pathlib import Path

from helpers import edit_module1, make_service

from vba_addin_editor.platform.paths import fingerprint
from vba_addin_editor.services.backup_service import BackupService
from vba_addin_editor.services.document_service import DocumentService


def test_successful_restore_replaces_current_and_keeps_safety(work_xlam: Path, tmp_path: Path):
    draft = DocumentService().open(work_xlam)
    original = work_xlam.read_bytes()
    edit_module1(draft)
    result = make_service().save_addin(draft)
    assert result.kind == "success"
    assert result.backup_path is not None
    edited = work_xlam.read_bytes()
    assert edited != original
    restore = BackupService().restore_transaction(work_xlam, result.backup_path)
    assert restore.kind == "success", restore
    assert work_xlam.read_bytes() == original
    assert restore.backup_path is not None
    assert restore.backup_path.read_bytes() == edited


def test_restore_blocks_when_target_changes_before_commit(work_xlam: Path, tmp_path: Path):
    draft = DocumentService().open(work_xlam)
    edit_module1(draft)
    saved = make_service().save_addin(draft)
    assert saved.kind == "success"
    backup = saved.backup_path
    target_hash = fingerprint(work_xlam).sha256
    backup_hash = fingerprint(backup).sha256

    def mutate():
        work_xlam.write_bytes(work_xlam.read_bytes() + b"x")

    result = BackupService().restore_transaction(
        work_xlam,
        backup,
        expected_backup_hash=backup_hash,
        expected_target_hash=target_hash,
        before_replace=mutate,
    )
    assert result.kind == "blocked"
    assert result.reason in {"external_change", "source_unstable"}


def test_restore_blocks_when_backup_changes_before_build(work_xlam: Path, tmp_path: Path):
    draft = DocumentService().open(work_xlam)
    edit_module1(draft)
    saved = make_service().save_addin(draft)
    backup = saved.backup_path
    from vba_addin_editor.platform.paths import fingerprint

    target_hash = fingerprint(work_xlam).sha256
    backup_hash = fingerprint(backup).sha256

    def mutate_backup():
        backup.write_bytes(backup.read_bytes() + b"nope")

    result = BackupService().restore_transaction(
        work_xlam,
        backup,
        expected_backup_hash=backup_hash,
        expected_target_hash=target_hash,
        before_build=mutate_backup,
    )
    assert result.kind == "blocked"


def test_invalid_backup_is_rejected(work_xlam: Path, tmp_path: Path):
    bogus = tmp_path / "not-a-backup.xlam"
    bogus.write_bytes(b"not zip")
    result = BackupService().restore_transaction(work_xlam, bogus)
    assert result.kind in {"blocked", "error"}
