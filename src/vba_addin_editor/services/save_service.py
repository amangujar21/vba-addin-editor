"""The critical safe-save pipeline (original plan sections 14, 76; XML plan 12, 31).

Never calls host.save() on the original path. Commit primitive is injectable
for failure-injection tests. One save => one verified candidate => one backup
=> one ReplaceFileW. XML changes are a second draft mutation type flowing
through the same transaction; only explicitly changed XML part paths are
allowed to differ from the reference package.
"""

from __future__ import annotations

import zipfile
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path

from vba_addin_editor.adapters.ooxml_package_adapter import (
    XML_EDITABLE_EXTENSIONS,
    OoxmlPackageAdapter,
    PackageError,
)
from vba_addin_editor.adapters.pyopenvba_adapter import (
    AdapterError,
    PyOpenVBAAdapter,
)
from vba_addin_editor.adapters.source_codec import CodePageError, validate_code_page
from vba_addin_editor.domain.changes import ChangeSet, compute_changes
from vba_addin_editor.domain.document import DocumentDraft, draft_from_snapshot
from vba_addin_editor.domain.results import SaveResult
from vba_addin_editor.platform import paths
from vba_addin_editor.platform import windows_file_ops as wfo
from vba_addin_editor.platform import windows_processes as wp
from vba_addin_editor.platform.session_store import sessions_root

CommitFn = Callable[[Path, Path, Path | None], None]


@dataclass
class SaveService:
    adapter: PyOpenVBAAdapter
    package_adapter: OoxmlPackageAdapter = field(default_factory=OoxmlPackageAdapter)
    commit: CommitFn | None = None  # default: real ReplaceFileW
    process_probe: Callable[[Path], bool] | None = None
    exclusive_probe: Callable[[Path], str | None] | None = None
    progress: Callable[[str], None] | None = None  # stage label callback

    def _report(self, stage: str) -> None:
        if self.progress is not None:
            self.progress(stage)

    def _probe_process(self, path: Path) -> bool:
        if self.process_probe is not None:
            return self.process_probe(path)
        probe = wp.probe_host_process(path)
        return probe.corresponding_host_running or probe.enumeration_failed

    def _host_probe(self, path: Path) -> wp.HostProcessProbe:
        if self.process_probe is not None:
            running = bool(self.process_probe(path))
            return wp.HostProcessProbe(
                corresponding_host_running=running,
                enumeration_failed=False,
                label=wp.corresponding_host_label(path),
            )
        return wp.probe_host_process(path)

    def _probe_exclusive(self, path: Path) -> str | None:
        result = self._exclusive_probe(path)
        if result is None or result.ok:
            return None
        return result.message

    def _exclusive_probe(self, path: Path) -> wfo.AccessProbe | None:
        if self.exclusive_probe is not None:
            message = self.exclusive_probe(path)
            if message:
                return wfo.AccessProbe(ok=False, reason="file_in_use", message=message)
            return wfo.AccessProbe(ok=True)
        return wfo.probe_exclusive_access(path)

    def _fingerprint_block(self, original: Path, draft: DocumentDraft, *, stage: str) -> SaveResult | None:
        status = paths.fingerprint_status(original, draft.baseline.file_fingerprint)
        if status == "match":
            return None
        messages = {
            "external_change": (
                "The file changed on disk after you opened it. The draft was not overwritten."
            ),
            "source_missing": "The original add-in file could not be found.",
            "source_unstable": "The original add-in file changed while it was being read.",
            "access_denied": "Windows denied access to the original add-in file.",
        }
        return SaveResult.blocked(status, messages.get(status), stage=stage)

    def _commit(self, replaced: Path, replacement: Path, backup: Path | None) -> None:
        if self.commit is not None:
            self.commit(replaced, replacement, backup)
        else:
            wfo.replace_file(replaced, replacement, backup)

    def _open_full_snapshot(self, path: Path):
        """Open VBA + XML state for baseline refresh (XML plan 12.13)."""
        from dataclasses import replace

        fp = paths.fingerprint(path)
        snapshot = self.adapter.open_snapshot(path, fp, host_process_running=False)
        if path.suffix.lower() in XML_EDITABLE_EXTENSIONS:
            snapshot = replace(
                snapshot,
                xml_parts=self.package_adapter.snapshot_xml_parts(path),
                package_safety=self.package_adapter.inspect_package_signature(path),
            )
        return snapshot

    # ------------------------------------------------------------------

    def _composite_preflight(self, draft: DocumentDraft, changes: ChangeSet) -> SaveResult | None:
        """Mutation-specific protection/signature policy (XML plan 12.4)."""
        has_vba = changes.has_vba_changes
        has_xml = changes.has_xml_changes
        # Password-protected VBA: block only VBA mutations; XML-only edits do
        # not touch vbaProject.bin, so the protected project stays intact.
        if has_vba and draft.baseline.safety.password_protected:
            return SaveResult.blocked(
                "password_protected",
                "This VBA project is password-protected and cannot be modified safely.",
            )
        # VBA signature: confirmation required only when VBA will be rewritten.
        if has_vba and draft.baseline.safety.signature_present and not draft.signed_save_confirmed:
            return SaveResult.needs_signature_confirmation()
        # OPC/package signature: any XML edit invalidates it — blocked in v1.
        if has_xml and draft.baseline.package_safety.opc_signature_present:
            return SaveResult.blocked(
                "package_signed",
                "This Office package contains an OPC digital signature.\n"
                "Editing package XML would invalidate that signature.\n"
                "XML editing of package-signed files is blocked in this version.",
            )
        return None

    def _validate_draft_parts(self, draft: DocumentDraft, has_vba: bool) -> SaveResult | None:
        # Stage E0/E1 — VBA validation gates VBA mutations only (XML plan 12.5/12.6).
        if has_vba:
            from vba_addin_editor.services.validation_service import validate_draft

            draft_problems = validate_draft(draft)
            if draft_problems:
                reason = (
                    "unsupported_component_operation"
                    if any(
                        item.startswith(("Cannot delete", "Cannot rename"))
                        for item in draft_problems
                    )
                    else "invalid_draft"
                )
                return SaveResult.error("; ".join(draft_problems), reason=reason)
            self._report("Validating text encoding")
            encoding = self.adapter.codepage_encoding(draft.baseline.code_page)
            changed = {m.current_name: m.body for m in draft.changed_existing_modules()}
            changed.update({m.current_name: m.body for m in draft.new_modules()})
            try:
                validate_code_page(changed, draft.baseline.code_page, encoding)
            except CodePageError as exc:
                return SaveResult.error(str(exc), reason="encoding")
        # Stage E2 — XML validation of every changed part (XML plan 12.7).
        if draft.changed_xml_parts():
            self._report("Validating XML")
            problems: list[str] = []
            for part in draft.changed_xml_parts():
                problems.extend(
                    f"{p}" for p in self.package_adapter.validate_draft_part(part)
                )
            if problems:
                return SaveResult.error("\n".join(problems), reason="invalid_xml")
        return None

    def _build_candidate(
        self,
        original: Path,
        draft: DocumentDraft,
        changes: ChangeSet,
        candidate: Path,
    ) -> SaveResult | None:
        """Stage F/G/H/I — build matrix (XML plan 12.8). Returns error result or None."""
        has_vba = changes.has_vba_changes
        has_xml = changes.has_xml_changes
        try:
            if has_vba:
                self._report("Building edited file")
                self.adapter.build_candidate(
                    original,
                    draft,
                    candidate,
                    allow_signature_removal=draft.signed_save_confirmed,
                )
                if has_xml:
                    # Combined: patch the disposable VBA candidate, never the original.
                    xml_candidate = paths.xml_patch_path_for(candidate)
                    try:
                        self.package_adapter.write_xml_candidate(
                            source_path=candidate,
                            dest_path=xml_candidate,
                            changed_parts=draft.changed_xml_parts(),
                        )
                        xml_candidate.replace(candidate)
                    finally:
                        _discard(xml_candidate)
            else:
                self._report("Building edited XML package")
                self.package_adapter.write_xml_candidate(
                    source_path=original,
                    dest_path=candidate,
                    changed_parts=draft.changed_xml_parts(),
                )
        except AdapterError as exc:
            _discard(candidate)
            reason = (
                "unsupported_component_operation"
                if "is not permitted" in str(exc)
                else "candidate_failed"
            )
            return SaveResult.error(str(exc), reason=reason)
        except PackageError as exc:
            _discard(candidate)
            return SaveResult.error(str(exc), reason="invalid_xml")
        except (OSError, zipfile.BadZipFile) as exc:
            _discard(candidate)
            return SaveResult.error(
                f"The edited package could not be built: {exc!r}", reason="build"
            )
        return None

    def _verify_candidate(
        self,
        reference: Path,
        candidate: Path,
        draft: DocumentDraft,
        changes: ChangeSet,
    ):
        allowed_xml = frozenset(c.path for c in changes.xml_modified)
        problems: list[str] = []
        vba_verification = self.adapter.verify_candidate(
            reference,
            candidate,
            draft,
            allowed_non_vba_changes=allowed_xml,
        )
        problems.extend(vba_verification.problems)
        # XML layer only applies to editable packages; skip for pure .xlam saves.
        if draft.baseline.extension in XML_EDITABLE_EXTENSIONS or changes.has_xml_changes:
            xml_verification = self.package_adapter.verify_candidate_xml(
                reference_path=reference,
                candidate_path=candidate,
                draft=draft,
            )
            problems.extend(xml_verification.problems)
        return not problems, tuple(problems)

    def _office_block(self, original: Path, *, stage: str) -> SaveResult | None:
        probe = self._host_probe(original)
        if probe.enumeration_failed:
            return SaveResult.blocked(
                "process_probe_failed",
                "Windows could not list running programs, so the editor cannot "
                f"confirm {probe.label or 'Office'} is closed.",
                stage=stage,
                win32_error=probe.win32_error,
            )
        if probe.corresponding_host_running:
            label = probe.label or wp.corresponding_host_label(original)
            return SaveResult.blocked(
                "office_running",
                f"Close {label} before saving this add-in.",
                stage=stage,
            )
        return None

    def _lock_block(self, original: Path, *, stage: str) -> SaveResult | None:
        probe = self._exclusive_probe(original)
        if probe is None or probe.ok:
            return None
        return SaveResult.blocked(
            probe.reason or "file_in_use",
            probe.message,
            stage=stage,
            win32_error=probe.win32_error,
            win32_name=probe.win32_name,
        )

    def _race_narrowing(self, original: Path, draft: DocumentDraft, candidate: Path) -> SaveResult | None:
        self._report("Final pre-commit checks")
        blocked = self._office_block(original, stage="pre_commit")
        if blocked is not None:
            _discard(candidate)
            return blocked
        blocked = self._fingerprint_block(original, draft, stage="pre_commit")
        if blocked is not None:
            _discard(candidate)
            return blocked
        blocked = self._lock_block(original, stage="pre_commit")
        if blocked is not None:
            _discard(candidate)
            return blocked
        return None

    def _initial_race_narrowing(self, draft: DocumentDraft) -> SaveResult | None:
        """Stage B/C preflight before any candidate work."""
        original = draft.baseline.path
        self._report("Checking Office is closed")
        blocked = self._office_block(original, stage="preflight")
        if blocked is not None:
            return blocked
        self._report("Checking the file is not in use")
        blocked = self._lock_block(original, stage="preflight")
        if blocked is not None:
            return blocked
        self._report("Checking the original file")
        return self._fingerprint_block(original, draft, stage="preflight")

    def save_addin(self, draft: DocumentDraft) -> SaveResult:
        original = draft.baseline.path

        # Stage A — composite change detection (no write when clean).
        self._report("Checking for changes")
        changes = compute_changes(draft)
        if changes.is_empty:
            return SaveResult.no_changes()

        # Stage B/C — host/process, lock, external-change preflight.
        preflight = self._initial_race_narrowing(draft)
        if preflight is not None:
            return preflight

        # Stage D — mutation-specific protection/signature policy.
        policy = self._composite_preflight(draft, changes)
        if policy is not None:
            return policy

        # Stage E0/E1/E2 — validation.
        validation = self._validate_draft_parts(draft, changes.has_vba_changes)
        if validation is not None:
            return validation

        # Stage F/G/H/I — candidate build matrix.
        candidate = paths.candidate_path_for(original)
        build_error = self._build_candidate(original, draft, changes, candidate)
        if build_error is not None:
            return build_error

        # Stage J — dual candidate verification.
        self._report("Verifying the edited file")
        ok, problems = self._verify_candidate(original, candidate, draft, changes)
        if not ok:
            keep = candidate.with_suffix(candidate.suffix + ".failed")
            _rename_or_discard(candidate, keep)
            return SaveResult.candidate_failed(problems, keep)

        # Stage L — narrow the race window.
        race = self._race_narrowing(original, draft, candidate)
        if race is not None:
            return race

        # Stage K/M — backup + ReplaceFileW commit.
        backup = paths.backup_path_for(original)
        self._report("Creating backup and replacing original")
        try:
            self._commit(original, candidate, backup)
        except wfo.FileOpsError as exc:
            return SaveResult.error(
                "Windows could not safely replace the file. Your files have been "
                "preserved for recovery.",
                reason="commit_failed",
                stage="commit",
                win32_error=exc.win32_error,
                win32_name=(
                    wfo.describe_win32_error(exc.win32_error)
                    if exc.win32_error is not None
                    else None
                ),
                details=dict(exc.details),
                candidate_path=candidate if candidate.exists() else None,
                backup_path=backup if backup.exists() else None,
            )

        # Stage N — post-commit verification: the saved ORIGINAL path compared
        # against the pre-save BACKUP (XML plan 12.12).
        self._report("Final verification")
        ok, problems = self._verify_candidate(backup, original, draft, changes)
        if not ok:
            return SaveResult.recovery_required(backup, problems)

        self._refresh_baseline(draft, original)
        return SaveResult.success(backup, original)

    def _refresh_baseline(self, draft: DocumentDraft, original: Path) -> None:
        new_snapshot = self._open_full_snapshot(original)
        refreshed = draft_from_snapshot(new_snapshot)
        draft.baseline = new_snapshot
        draft.modules[:] = refreshed.modules
        draft.xml_parts[:] = refreshed.xml_parts
        draft.signed_save_confirmed = False

    # -- Save a Copy (original plan 47; XML plan 12.14) -------------------

    def save_copy(
        self,
        draft: DocumentDraft,
        dest: Path,
        *,
        source_path: Path | None = None,
        allow_overwrite: bool = False,
        recovered_copy: bool = False,
        reviewed_dest_hash: str | None = None,
        session_dir: Path | None = None,
    ) -> SaveResult:
        dest = Path(dest)
        source_ext = draft.baseline.extension
        original = draft.baseline.path
        build_from = Path(source_path) if source_path is not None else original
        if dest.suffix.lower() != source_ext.lower():
            return SaveResult.error(
                f"The copy must keep the same file type ({source_ext}); format "
                "conversion is not supported in this version.",
                reason="unsupported_extension",
                operation_type="save_copy",
            )
        if paths.paths_are_same_file(original, dest):
            return SaveResult.blocked(
                "same_path",
                "The copy destination is the same file as the original.",
                operation_type="save_copy",
            )
        if session_dir is not None and paths.is_inside(dest, session_dir):
            return SaveResult.blocked(
                "same_path",
                "The copy destination cannot be inside the editor's session folder.",
                operation_type="save_copy",
            )
        try:
            if paths.is_inside(dest, sessions_root()):
                return SaveResult.blocked(
                    "same_path",
                    "The copy destination cannot be inside the editor's session folder.",
                    operation_type="save_copy",
                )
        except OSError:
            pass
        if paths.path_or_ancestor_is_reparse_point(dest):
            return SaveResult.blocked(
                "reparse_point",
                "Choose a regular folder. This version does not write through links.",
                operation_type="save_copy",
            )
        dest_exists = dest.exists()
        if dest_exists and not allow_overwrite:
            return SaveResult.blocked(
                "destination_exists",
                "The destination already exists. Confirm overwrite to replace it.",
                operation_type="save_copy",
            )
        source_status = paths.fingerprint_status(original, draft.baseline.file_fingerprint)
        if source_status != "match" and not recovered_copy:
            return SaveResult.blocked(
                source_status if source_status != "match" else "external_change",
                "The original file is missing or has changed. Use Save recovered "
                "draft as a separate copy; newer disk changes will not be included.",
                operation_type="save_copy",
            )
        if recovered_copy and paths.paths_are_same_file(original, dest):
            return SaveResult.blocked(
                "same_path",
                "A recovered copy must be saved to a different path.",
                operation_type="save_copy",
            )
        changes = compute_changes(draft)
        if dest_exists:
            blocked = self._office_block(dest, stage="preflight")
            if blocked is not None:
                blocked.operation_type = "save_copy"
                return blocked
            blocked = self._lock_block(dest, stage="preflight")
            if blocked is not None:
                blocked.operation_type = "save_copy"
                return blocked
            if reviewed_dest_hash is not None:
                current = paths.fingerprint(dest)
                if current.sha256 != reviewed_dest_hash:
                    return SaveResult.blocked(
                        "destination_changed",
                        "The destination file changed after it was reviewed.",
                        operation_type="save_copy",
                    )
        elif dest.parent.exists() and paths.path_or_ancestor_is_reparse_point(dest.parent):
            return SaveResult.blocked(
                "reparse_point",
                "Choose a regular folder. This version does not write through links.",
                operation_type="save_copy",
            )
        if not changes.is_empty:
            policy = self._composite_preflight(draft, changes)
            if policy is not None:
                policy.operation_type = "save_copy"
                return policy
            validation = self._validate_draft_parts(draft, changes.has_vba_changes)
            if validation is not None:
                validation.operation_type = "save_copy"
                return validation
        temp = paths.sibling_temp_path(dest, kind="copy")
        dest_backup = None
        try:
            if changes.is_empty:
                temp.write_bytes(build_from.read_bytes())
                if paths.sha256_bytes(temp.read_bytes()) != draft.baseline.file_fingerprint.sha256:
                    _discard(temp)
                    return SaveResult.error(
                        "The copy source no longer matches the captured baseline.",
                        reason="source_unstable",
                        operation_type="save_copy",
                    )
            else:
                build_error = self._build_candidate(build_from, draft, changes, temp)
                if build_error is not None:
                    build_error.operation_type = "save_copy"
                    return build_error
                ok, problems = self._verify_candidate(build_from, temp, draft, changes)
                if not ok:
                    keep = temp.with_suffix(temp.suffix + ".failed")
                    _rename_or_discard(temp, keep)
                    result = SaveResult.candidate_failed(problems, keep)
                    result.operation_type = "save_copy"
                    return result
            committed_hash = paths.sha256_bytes(temp.read_bytes())
            if dest_exists:
                dest_backup = paths.backup_path_for(dest, label="copy-backup")
                try:
                    self._commit(dest, temp, dest_backup)
                except wfo.FileOpsError as exc:
                    return SaveResult.error(
                        "Windows could not safely replace the destination file.",
                        reason="commit_failed",
                        stage="commit",
                        win32_error=exc.win32_error,
                        win32_name=(
                            wfo.describe_win32_error(exc.win32_error)
                            if exc.win32_error is not None
                            else None
                        ),
                        details=dict(exc.details),
                        candidate_path=temp if temp.exists() else None,
                        backup_path=dest_backup if dest_backup.exists() else None,
                        operation_type="save_copy",
                    )
            else:
                if dest.exists():
                    _discard(temp)
                    return SaveResult.blocked(
                        "destination_exists",
                        "Another file appeared at the destination before the copy was committed.",
                        operation_type="save_copy",
                        candidate_path=None,
                    )
                try:
                    wfo.move_new_file(temp, dest)
                except wfo.FileOpsError as exc:
                    return SaveResult.error(
                        "Windows could not move the copy into place.",
                        reason="commit_failed",
                        stage="commit",
                        win32_error=exc.win32_error,
                        details=dict(exc.details),
                        candidate_path=temp if temp.exists() else None,
                        operation_type="save_copy",
                    )
            if paths.sha256_bytes(dest.read_bytes()) != committed_hash:
                return SaveResult.recovery_required(
                    dest_backup,
                    ("Committed copy does not match the verified candidate hash.",),
                )
        finally:
            _discard(temp)
        result = SaveResult.success(dest_backup, dest, operation_type="save_copy")
        result.details["hash_candidate"] = committed_hash
        return result

def _discard(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except OSError:
        pass


def _rename_or_discard(path: Path, dest: Path) -> None:
    try:
        path.replace(dest)
    except OSError:
        _discard(path)
