"""The critical safe-save pipeline (original plan sections 14, 76; XML plan 12, 31).

Never calls host.save() on the original path. Commit primitive is injectable
for failure-injection tests. One save => one verified candidate => one atomic
commit, with the old file kept as a backup unless auto_backup is off. XML
changes are a second draft mutation type flowing through the same transaction;
only explicitly changed XML part paths are allowed to differ from the
reference package.

The candidate is built and verified in memory from the baseline bytes captured
at open, written once, and never read back: antivirus scans a freshly written
macro package on its first open for read, which costs seconds per save.
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
    auto_backup: bool = True  # keep the replaced file as a same-folder backup

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

    def _commit(self, replaced: Path, replacement: Path, backup: Path | None) -> str | None:
        """Commit; returns how ("rename" / "replace_file"), None for an injected primitive."""
        if self.commit is not None:
            self.commit(replaced, replacement, backup)
            return None
        return wfo.commit_replace(replaced, replacement, backup)

    def _snapshot_of_written(self, path: Path, data: bytes):
        """VBA + XML state of bytes this save just committed to ``path`` (XML plan 12.13)."""
        from dataclasses import replace

        fp = paths.written_fingerprint(data, path)
        snapshot = self.adapter.open_snapshot(path, fp, host_process_running=False, data=data)
        if path.suffix.lower() in XML_EDITABLE_EXTENSIONS:
            snapshot = replace(
                snapshot,
                xml_parts=self.package_adapter.snapshot_xml_parts(data),
                package_safety=self.package_adapter.inspect_package_signature(data),
            )
        return snapshot

    def _reference_bytes(self, draft: DocumentDraft, source: Path) -> bytes | None:
        """The baseline package bytes, or None when they are no longer available.

        Normally the bytes read when the file was opened. A draft without them
        (an adopted external baseline, a hand-built snapshot) reads ``source``.
        """
        data = draft.baseline.package_bytes
        if data is None:
            data = source.read_bytes()
        if paths.sha256_bytes(data) != draft.baseline.file_fingerprint.sha256:
            return None
        return data

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
        reference: bytes,
        draft: DocumentDraft,
        changes: ChangeSet,
    ) -> bytes | SaveResult:
        """Stage F/G/H/I — build matrix (XML plan 12.8). Returns the package or an error."""
        try:
            xml = (
                self.package_adapter.xml_replacements(draft.changed_xml_parts())
                if changes.has_xml_changes
                else {}
            )
            if changes.has_vba_changes:
                self._report("Building edited file")
                # Combined saves write the XML parts in the same pass.
                return self.adapter.build_candidate_bytes(
                    reference,
                    draft,
                    allow_signature_removal=draft.signed_save_confirmed,
                    xml_replacements=xml,
                )
            self._report("Building edited XML package")
            return self.package_adapter.build_package(reference, dict(xml))
        except AdapterError as exc:
            reason = (
                "unsupported_component_operation"
                if "is not permitted" in str(exc)
                else "candidate_failed"
            )
            return SaveResult.error(str(exc), reason=reason)
        except PackageError as exc:
            return SaveResult.error(str(exc), reason="invalid_xml")
        except (OSError, zipfile.BadZipFile) as exc:
            return SaveResult.error(
                f"The edited package could not be built: {exc!r}", reason="build"
            )

    def _verify_candidate(
        self,
        reference: Path | bytes,
        candidate: Path | bytes,
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
                signature_edits=self.adapter.signature_removal_edits(reference, draft),
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

        # Stage F/G/H/I — candidate build matrix, in memory.
        try:
            reference = self._reference_bytes(draft, original)
        except OSError as exc:
            return SaveResult.error(
                f"The original add-in file could not be read: {exc!r}", reason="build"
            )
        if reference is None:
            return SaveResult.blocked(
                "external_change",
                "The file changed on disk after you opened it. The draft was not overwritten.",
            )
        built = self._build_candidate(reference, draft, changes)
        if isinstance(built, SaveResult):
            return built

        # Stage J — dual candidate verification.
        self._report("Verifying the edited file")
        candidate = paths.candidate_path_for(original)
        ok, problems = self._verify_candidate(reference, built, draft, changes)
        if not ok:
            return SaveResult.candidate_failed(problems, _keep_failed(candidate, built))

        self._report("Writing the edited file")
        try:
            wfo.write_verified(candidate, built)
        except OSError as exc:
            _discard(candidate)
            return SaveResult.error(
                f"The edited package could not be written: {exc!r}", reason="build"
            )

        # Stage L — narrow the race window.
        race = self._race_narrowing(original, draft, candidate)
        if race is not None:
            return race

        # Stage K/M — atomic commit; the replaced file becomes the backup.
        backup = paths.backup_path_for(original) if self.auto_backup else None
        self._report("Creating backup and replacing original" if backup else "Replacing original")
        try:
            candidate_identity = paths.file_identity(candidate)
            original_identity = paths.file_identity(original)
        except OSError as exc:
            _discard(candidate)
            return SaveResult.error(
                f"The original add-in file could not be inspected: {exc!r}", reason="build"
            )
        try:
            commit_method = self._commit(original, candidate, backup)
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
                backup_path=backup if backup is not None and backup.exists() else None,
            )

        # Stage N — post-commit verification (XML plan 12.12).
        self._report("Final verification")
        problems = self._verify_commit(
            original,
            built,
            candidate_identity,
            backup,
            reference,
            original_identity,
        )
        if problems:
            return SaveResult.recovery_required(backup, problems)

        replaced_hash = draft.baseline.file_fingerprint.sha256
        self._refresh_baseline(draft, original, built)
        result = SaveResult.success(backup, original)
        result.details["hash_candidate"] = draft.baseline.file_fingerprint.sha256
        if backup is not None:
            result.details["hash_backup"] = replaced_hash
        if commit_method:
            result.details["commit_method"] = commit_method
        return result

    def _verify_commit(
        self,
        committed: Path,
        committed_bytes: bytes,
        committed_identity: tuple[int, int, int, int] | None,
        backup: Path | None,
        backup_bytes: bytes,
        backup_identity: tuple[int, int, int, int] | None,
    ) -> tuple[str, ...]:
        """The committed path holds the verified candidate; the backup holds the old file.

        Both were moved, not copied, so each is proven by being the same file
        (volume, file id, size, mtime) as before the commit. A file that is not
        provably the same one is compared byte for byte instead.
        """
        problems: list[str] = []
        try:
            if not _holds(committed, committed_bytes, committed_identity):
                problems.append("The saved file does not match the verified candidate.")
            if backup is not None and not _holds(backup, backup_bytes, backup_identity):
                problems.append("The backup does not match the file as it was before the save.")
        except OSError as exc:
            problems.append(f"The saved files could not be checked: {exc!r}")
        return tuple(problems)

    def _refresh_baseline(self, draft: DocumentDraft, original: Path, written: bytes) -> None:
        new_snapshot = self._snapshot_of_written(original, written)
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
            try:
                reference = self._reference_bytes(draft, build_from)
            except OSError as exc:
                return SaveResult.error(
                    f"The copy source could not be read: {exc!r}",
                    reason="build",
                    operation_type="save_copy",
                )
            if reference is None:
                return SaveResult.error(
                    "The copy source no longer matches the captured baseline.",
                    reason="source_unstable",
                    operation_type="save_copy",
                )
            if changes.is_empty:
                built = reference
            else:
                built = self._build_candidate(reference, draft, changes)
                if isinstance(built, SaveResult):
                    built.operation_type = "save_copy"
                    return built
                ok, problems = self._verify_candidate(reference, built, draft, changes)
                if not ok:
                    result = SaveResult.candidate_failed(problems, _keep_failed(temp, built))
                    result.operation_type = "save_copy"
                    return result
            try:
                wfo.write_verified(temp, built)
            except OSError as exc:
                return SaveResult.error(
                    f"The copy could not be written: {exc!r}",
                    reason="build",
                    operation_type="save_copy",
                )
            committed_hash = paths.sha256_bytes(built)
            temp_identity = paths.file_identity(temp)
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
            if not _holds(dest, built, temp_identity):
                return SaveResult.recovery_required(
                    dest_backup,
                    ("Committed copy does not match the verified candidate hash.",),
                )
        finally:
            _discard(temp)
        result = SaveResult.success(dest_backup, dest, operation_type="save_copy")
        result.details["hash_candidate"] = committed_hash
        return result


def _holds(path: Path, data: bytes, identity: tuple[int, int, int, int] | None) -> bool:
    """True when ``path`` is the file known to hold ``data``, or reads back as it."""
    if identity is not None and paths.file_identity(path) == identity:
        return True
    return path.read_bytes() == data


def _keep_failed(candidate: Path, data: bytes) -> Path | None:
    """Keep a candidate that failed verification beside the original for diagnosis."""
    keep = candidate.with_suffix(candidate.suffix + ".failed")
    try:
        keep.write_bytes(data)
    except OSError:
        return None
    return keep


def _discard(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except OSError:
        pass
