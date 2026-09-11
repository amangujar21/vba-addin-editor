"""Backup catalog, browser metadata, and verified restore (IMP-06)."""

from __future__ import annotations

import json
from collections.abc import Callable
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path

from vba_addin_editor.adapters.pyopenvba_adapter import AdapterError, PyOpenVBAAdapter
from vba_addin_editor.domain.changes import compute_changes
from vba_addin_editor.domain.document import draft_from_snapshot
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


@dataclass
class BackupInspect:
    path: Path
    sha256: str
    size: int
    extension: str
    problems: list[str] = field(default_factory=list)
    summary_lines: list[str] = field(default_factory=list)
    catalog_match: BackupRecord | None = None
    build_identity: dict = field(default_factory=dict)


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

    def inspect(self, backup_path: Path, *, current_path: Path | None = None) -> BackupInspect:
        backup_path = Path(backup_path)
        problems: list[str] = []
        summary: list[str] = []
        digest = ""
        size = 0
        try:
            fp = paths.fingerprint(backup_path)
            digest = fp.sha256
            size = fp.size
        except OSError as exc:
            problems.append(f"The backup could not be read: {type(exc).__name__}")
            return BackupInspect(backup_path, "", 0, backup_path.suffix.lower(), problems)
        if current_path is not None and backup_path.suffix.lower() != current_path.suffix.lower():
            problems.append("The backup file type does not match the open add-in.")
        if paths.paths_are_same_file(backup_path, current_path) if current_path else False:
            problems.append("The selected backup is the current file.")
        catalog_match = next((rec for rec in self.list_records() if rec.sha256 == digest), None)
        try:
            snap = self.adapter.open_snapshot(backup_path, paths.fingerprint(backup_path))
            summary.append(f"Backup: {backup_path.name}")
            summary.append(f"SHA-256: {digest}")
            summary.append(f"Size: {size} bytes")
            summary.append(f"Modules: {len(snap.modules)}")
            if catalog_match:
                summary.append(f"Catalog: {catalog_match.operation_type} at {catalog_match.utc_time}")
                summary.append(f"Build: {catalog_match.build_identity}")
            if current_path is not None and current_path.exists():
                try:
                    current = self.adapter.open_snapshot(current_path, paths.fingerprint(current_path))
                    backup_draft = draft_from_snapshot(snap)
                    # Compare backup bodies against current snapshot as a semantic diff.
                    current_by_name = {m.original_name.casefold(): m for m in current.modules}
                    for module in backup_draft.modules:
                        cur = current_by_name.get(module.current_name.casefold())
                        if cur is None:
                            summary.append(f"Only in backup: {module.current_name}")
                        elif cur.body != module.body:
                            summary.append(f"Body differs: {module.current_name}")
                    backup_names = {m.current_name.casefold() for m in backup_draft.modules}
                    for name, module in current_by_name.items():
                        if name not in backup_names:
                            summary.append(f"Only in current: {module.original_name}")
                    _ = compute_changes(backup_draft)
                except (AdapterError, OSError):
                    summary.append("Current file could not be compared.")
        except AdapterError as exc:
            problems.append(str(exc))
        except OSError as exc:
            problems.append(f"The backup could not be parsed: {type(exc).__name__}")
        return BackupInspect(
            path=backup_path,
            sha256=digest,
            size=size,
            extension=backup_path.suffix.lower(),
            problems=problems,
            summary_lines=summary,
            catalog_match=catalog_match,
            build_identity=(catalog_match.build_identity if catalog_match else {}),
        )

    def restore(
        self,
        current_path: Path,
        backup_path: Path,
        *,
        process_probe=None,
    ) -> Path:
        result = self.restore_transaction(
            current_path,
            backup_path,
            process_probe=process_probe,
        )
        if result.kind != "success":
            raise AdapterError(result.message or "Restore failed.", result.details)
        if result.backup_path is None:
            raise AdapterError("Restore succeeded but no safety backup path was returned.")
        return result.backup_path

    def restore_transaction(
        self,
        current_path: Path,
        backup_path: Path,
        *,
        expected_backup_hash: str | None = None,
        expected_target_hash: str | None = None,
        process_probe=None,
        exclusive_probe: Callable[[Path], object] | None = None,
        before_build=None,
        before_replace=None,
    ) -> SaveResult:
        current_path = Path(current_path)
        backup_path = Path(backup_path)
        details: dict = {
            "target": str(current_path),
            "selected_backup": str(backup_path),
        }
        try:
            backup_fp = paths.fingerprint(backup_path)
            target_fp = paths.fingerprint(current_path)
        except FileNotFoundError:
            return SaveResult.blocked(
                "source_missing",
                "The current file or selected backup is missing.",
                operation_type="restore",
                details=details,
            )
        except paths.UnstableSourceError:
            return SaveResult.blocked(
                "source_unstable",
                "The current file or selected backup changed while it was being read.",
                operation_type="restore",
                details=details,
            )
        except OSError as exc:
            return SaveResult.error(
                f"The restore files could not be read ({type(exc).__name__}).",
                reason="access_denied",
                operation_type="restore",
                details=details,
            )
        if expected_backup_hash is None:
            expected_backup_hash = backup_fp.sha256
        if expected_target_hash is None:
            expected_target_hash = target_fp.sha256

        blocked = self._restore_preflight(
            current_path,
            backup_path,
            expected_backup_hash=expected_backup_hash,
            expected_target_hash=expected_target_hash,
            process_probe=process_probe,
            exclusive_probe=exclusive_probe,
            details=details,
            stage="preflight",
        )
        if blocked is not None:
            return blocked

        if before_build is not None:
            before_build()

        blocked = self._restore_preflight(
            current_path,
            backup_path,
            expected_backup_hash=expected_backup_hash,
            expected_target_hash=expected_target_hash,
            process_probe=process_probe,
            exclusive_probe=exclusive_probe,
            details=details,
            stage="pre_build",
        )
        if blocked is not None:
            return blocked

        try:
            backup_bytes = backup_path.read_bytes()
        except OSError as exc:
            return SaveResult.error(
                "The selected backup could not be read.",
                reason="access_denied",
                operation_type="restore",
                details={**details, "exception_type": type(exc).__name__},
            )
        if sha256_bytes(backup_bytes) != expected_backup_hash:
            return SaveResult.blocked(
                "destination_changed",
                "The selected backup changed after it was reviewed.",
                operation_type="restore",
                details=details,
            )
        candidate = paths.sibling_temp_path(current_path, kind="restore")
        details["candidate"] = str(candidate)
        try:
            candidate.write_bytes(backup_bytes)
            if sha256_bytes(candidate.read_bytes()) != expected_backup_hash:
                candidate.unlink(missing_ok=True)
                return SaveResult.error(
                    "The restore candidate did not match the selected backup.",
                    reason="candidate_failed",
                    operation_type="restore",
                    candidate_path=candidate,
                    details=details,
                )
        except OSError as exc:
            return SaveResult.error(
                "The restore candidate could not be written.",
                reason="candidate_failed",
                operation_type="restore",
                details={**details, "exception_type": type(exc).__name__},
            )

        if before_replace is not None:
            before_replace()

        blocked = self._restore_preflight(
            current_path,
            backup_path,
            expected_backup_hash=expected_backup_hash,
            expected_target_hash=expected_target_hash,
            process_probe=process_probe,
            exclusive_probe=exclusive_probe,
            details=details,
            stage="pre_commit",
        )
        if blocked is not None:
            candidate.unlink(missing_ok=True)
            return blocked

        safety = paths.backup_path_for(current_path, label="pre-restore-safety")
        details["safety_backup"] = str(safety)
        try:
            wfo.replace_file(current_path, candidate, safety)
        except wfo.FileOpsError as exc:
            return SaveResult.error(
                "Windows could not safely restore the backup. Recovery artifacts were kept.",
                reason="commit_failed",
                stage="commit",
                win32_error=exc.win32_error,
                win32_name=(
                    wfo.describe_win32_error(exc.win32_error)
                    if exc.win32_error is not None
                    else None
                ),
                details={**details, **dict(exc.details)},
                candidate_path=candidate if candidate.exists() else None,
                backup_path=safety if safety.exists() else None,
                operation_type="restore",
            )

        try:
            restored_hash = sha256_bytes(current_path.read_bytes())
            safety_hash = sha256_bytes(safety.read_bytes()) if safety.exists() else ""
            backup_after = sha256_bytes(backup_path.read_bytes())
        except OSError:
            return SaveResult.recovery_required(
                safety if safety.exists() else None,
                ("Post-restore verification could not read the restored files.",),
            )
        problems: list[str] = []
        if restored_hash != expected_backup_hash:
            problems.append("Restored file did not equal the selected backup.")
        if safety_hash != expected_target_hash:
            problems.append("The safety backup does not match the pre-restore file.")
        if backup_after != expected_backup_hash:
            problems.append("The selected backup changed during restore.")
        if problems:
            result = SaveResult.recovery_required(safety if safety.exists() else None, tuple(problems))
            result.details.update(details)
            result.candidate_path = candidate if candidate.exists() else None
            result.operation_type = "restore"
            return result
        try:
            verified = self.adapter.open_snapshot(current_path, paths.fingerprint(current_path))
            if not verified.modules:
                result = SaveResult.recovery_required(
                    safety, ("Restored file did not verify; use the safety backup.",)
                )
                result.details.update(details)
                result.operation_type = "restore"
                return result
        except AdapterError as exc:
            result = SaveResult.recovery_required(safety, (str(exc),))
            result.details.update(details)
            result.operation_type = "restore"
            return result
        warning = self.record_success(
            original=current_path, backup=safety, operation_type="restore"
        )
        result = SaveResult.success(safety, current_path, operation_type="restore")
        result.details.update(details)
        result.details["hash_backup"] = expected_backup_hash
        result.details["hash_safety"] = safety_hash
        if warning:
            result.details["catalog_warning"] = warning
        return result

    def _restore_preflight(
        self,
        current_path: Path,
        backup_path: Path,
        *,
        expected_backup_hash: str,
        expected_target_hash: str,
        process_probe,
        exclusive_probe,
        details: dict,
        stage: str,
    ) -> SaveResult | None:
        if process_probe is None:
            process_probe = wp.host_process_running
        if process_probe(current_path):
            return SaveResult.blocked(
                "office_running",
                "Close Excel/PowerPoint before restoring a backup.",
                stage=stage,
                operation_type="restore",
                details=details,
            )
        if current_path.suffix.lower() != backup_path.suffix.lower():
            return SaveResult.blocked(
                "unsupported_extension",
                "The backup file type does not match the open add-in.",
                stage=stage,
                operation_type="restore",
                details=details,
            )
        if paths.paths_are_same_file(current_path, backup_path):
            return SaveResult.blocked(
                "same_path",
                "The selected backup is the current file.",
                stage=stage,
                operation_type="restore",
                details=details,
            )
        probe = exclusive_probe(current_path) if exclusive_probe is not None else wfo.probe_exclusive_access(current_path)
        if probe is not None:
            ok = getattr(probe, "ok", probe is None)
            if probe is False or ok is False:
                reason = getattr(probe, "reason", "file_in_use")
                message = getattr(probe, "message", None) or "The current file is in use."
                return SaveResult.blocked(
                    reason if isinstance(reason, str) else "file_in_use",
                    message,
                    stage=stage,
                    operation_type="restore",
                    details=details,
                )
        try:
            backup_now = paths.fingerprint(backup_path)
            target_now = paths.fingerprint(current_path)
        except FileNotFoundError:
            return SaveResult.blocked("source_missing", stage=stage, operation_type="restore", details=details)
        except paths.UnstableSourceError:
            return SaveResult.blocked("source_unstable", stage=stage, operation_type="restore", details=details)
        except OSError:
            return SaveResult.blocked("access_denied", stage=stage, operation_type="restore", details=details)
        if backup_now.sha256 != expected_backup_hash:
            return SaveResult.blocked(
                "destination_changed",
                "The selected backup changed after it was reviewed.",
                stage=stage,
                operation_type="restore",
                details=details,
            )
        if target_now.sha256 != expected_target_hash:
            return SaveResult.blocked(
                "external_change",
                "The current file changed after restore was reviewed.",
                stage=stage,
                operation_type="restore",
                details=details,
            )
        problem = self.validate_backup(backup_path, expected_ext=current_path.suffix)
        if problem:
            return SaveResult.blocked(
                "invalid_draft",
                problem,
                stage=stage,
                operation_type="restore",
                details=details,
            )
        return None

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
