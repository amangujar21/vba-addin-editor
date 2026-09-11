"""Backup catalog, browser metadata, and verified restore (IMP-06)."""

from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

from vba_addin_editor.adapters.pyopenvba_adapter import AdapterError, PyOpenVBAAdapter
from vba_addin_editor.domain.results import SaveResult
from vba_addin_editor.platform import paths
from vba_addin_editor.platform import windows_file_ops as wfo
from vba_addin_editor.platform import windows_processes as wp
from vba_addin_editor.platform.paths import backups_catalog_path, sha256_bytes
from vba_addin_editor.platform.session_store import write_json_atomic
from vba_addin_editor.version import build_identity


@dataclass(frozen=True)
class BackupRecord:
    schema_version: int
    original_path: str
    backup_path: str
    sha256: str
    size: int
    utc_time: str
    operation_type: str
    build_identity: dict
    extension: str


class BackupService:
    """Restore backup flow (plan 16.3, 78) plus catalog (IMP-06)."""

    def __init__(self, adapter: PyOpenVBAAdapter | None = None, catalog_path: Path | None = None) -> None:
        self.adapter = adapter or PyOpenVBAAdapter()
        self.catalog_path = catalog_path or backups_catalog_path()

    def record_success(
        self,
        *,
        original: Path,
        backup: Path | None,
        operation_type: str,
    ) -> str | None:
        if backup is None or not backup.exists():
            return None
        try:
            digest = sha256_bytes(backup.read_bytes())
            record = {
                "schema_version": 1,
                "original_path": str(original),
                "backup_path": str(backup),
                "sha256": digest,
                "size": backup.stat().st_size,
                "utc_time": datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z"),
                "operation_type": operation_type,
                "build_identity": build_identity(),
                "extension": original.suffix.lower(),
            }
            catalog = self._load()
            catalog.append(record)
            write_json_atomic(self.catalog_path, {"schema_version": 1, "entries": catalog})
        except OSError:
            return "The save succeeded, but the backup catalog could not be updated."
        return None

    def list_records(self, original: Path | None = None) -> list[BackupRecord]:
        records = []
        for item in self._load():
            if (
                original is not None
                and Path(item.get("original_path", "")).name != original.name
                and Path(item.get("backup_path", "")).parent != original.parent
            ):
                continue
            records.append(self._to_record(item))
        return records

    def discover_legacy(self, original: Path) -> list[Path]:
        pattern = f"{original.stem} - backup *{original.suffix}"
        return sorted(original.parent.glob(pattern))

    def validate_backup(self, backup_path: Path, *, expected_ext: str) -> str | None:
        if backup_path.suffix.lower() != expected_ext.lower():
            return "The backup file type does not match the open add-in."
        try:
            self.adapter.open_snapshot(backup_path, paths.fingerprint(backup_path))
        except AdapterError as exc:
            return str(exc)
        except OSError as exc:
            return f"The backup could not be read: {exc.__class__.__name__}"
        return None

    def restore(
        self,
        current_path: Path,
        backup_path: Path,
        *,
        process_probe=None,
    ) -> Path:
        if process_probe is None:
            process_probe = wp.host_process_running
        if process_probe(current_path):
            raise AdapterError("Close Excel/PowerPoint before restoring a backup.")
        if current_path.suffix.lower() != backup_path.suffix.lower():
            raise AdapterError("The backup file type does not match the open add-in.")
        if paths.paths_are_same_file(current_path, backup_path):
            raise AdapterError("The selected backup is the current file.")
        self.adapter.open_snapshot(backup_path, paths.fingerprint(backup_path))
        backup_bytes = backup_path.read_bytes()
        backup_hash = sha256_bytes(backup_bytes)
        candidate = paths.sibling_temp_path(current_path, kind="restore")
        candidate.write_bytes(backup_bytes)
        if sha256_bytes(candidate.read_bytes()) != backup_hash:
            candidate.unlink(missing_ok=True)
            raise AdapterError("The restore candidate did not match the selected backup.")
        recovery_backup = paths.backup_path_for(current_path, label="pre-restore-safety")
        wfo.replace_file(current_path, candidate, recovery_backup)
        if sha256_bytes(current_path.read_bytes()) != backup_hash:
            raise AdapterError("Restored file did not verify; use the safety backup.")
        if sha256_bytes(backup_path.read_bytes()) != backup_hash:
            raise AdapterError("The selected backup changed during restore.")
        verified = self.adapter.open_snapshot(current_path, paths.fingerprint(current_path))
        if not verified.modules:
            raise AdapterError("Restored file did not verify; use the safety backup.")
        return recovery_backup

    def _load(self) -> list[dict]:
        if not self.catalog_path.exists():
            return []
        try:
            data = json.loads(self.catalog_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return []
        return list(data.get("entries", []))

    def _to_record(self, item: dict) -> BackupRecord:
        return BackupRecord(
            schema_version=int(item.get("schema_version", 1)),
            original_path=item.get("original_path", ""),
            backup_path=item.get("backup_path", ""),
            sha256=item.get("sha256", ""),
            size=int(item.get("size", 0)),
            utc_time=item.get("utc_time", ""),
            operation_type=item.get("operation_type", "save"),
            build_identity=item.get("build_identity") or {},
            extension=item.get("extension", ""),
        )


def restore_result_from_exception(exc: Exception) -> SaveResult:
    return SaveResult.error(str(exc), reason="commit_failed", operation_type="restore")
