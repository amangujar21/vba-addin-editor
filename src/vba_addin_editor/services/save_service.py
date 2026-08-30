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
    host_process_for,
)
from vba_addin_editor.adapters.source_codec import CodePageError, validate_code_page
from vba_addin_editor.domain.changes import ChangeSet, compute_changes
from vba_addin_editor.domain.document import DocumentDraft, draft_from_snapshot
from vba_addin_editor.domain.results import SaveResult
from vba_addin_editor.platform import paths
from vba_addin_editor.platform import windows_file_ops as wfo
from vba_addin_editor.platform import windows_processes as wp

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
        return wp.host_process_running(path)

    def _probe_exclusive(self, path: Path) -> str | None:
        if self.exclusive_probe is not None:
            return self.exclusive_probe(path)
        return wfo.exclusive_access_error(path)

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
                return SaveResult.error("; ".join(draft_problems), reason="invalid_draft")
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
            return SaveResult.error(str(exc), reason="build")
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

    def _race_narrowing(self, original: Path, draft: DocumentDraft, candidate: Path) -> SaveResult | None:
        self._report("Final pre-commit checks")
        if self._probe_process(original):
            _discard(candidate)
            return SaveResult.blocked(
                "office_running", "Excel/PowerPoint started again before the save."
            )
        if not paths.fingerprint_matches(original, draft.baseline.file_fingerprint):
            _discard(candidate)
            return SaveResult.blocked(
                "external_change",
                "The file changed on disk while the edit was being prepared. "
                "Nothing was overwritten.",
            )
        lock_err = self._probe_exclusive(original)
        if lock_err:
            _discard(candidate)
            return SaveResult.blocked("locked", lock_err)
        return None

    def _initial_race_narrowing(self, draft: DocumentDraft) -> SaveResult | None:
        """Stage B/C preflight before any candidate work."""
        original = draft.baseline.path
        self._report("Checking Office is closed")
        if self._probe_process(original):
            return SaveResult.blocked(
                "office_running",
                f"Close {host_process_for(original)}'s application (Excel/PowerPoint) "
                "before saving this file.",
            )
        self._report("Checking the file is not in use")
        lock_err = self._probe_exclusive(original)
        if lock_err:
            return SaveResult.blocked("locked", lock_err)
        self._report("Checking the original file")
        if not paths.fingerprint_matches(original, draft.baseline.file_fingerprint):
            return SaveResult.blocked(
                "external_change",
                "The file changed on disk after you opened it. Reload it before saving.",
            )
        return None

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
        except wfo.FileOpsError:
            return SaveResult.error(
                "Windows could not safely replace the file. Your files have been "
                "preserved for recovery.",
                reason="commit_failed",
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

    def save_copy(self, draft: DocumentDraft, dest: Path) -> SaveResult:
        changes = compute_changes(draft)
        if changes.is_empty:
            return SaveResult.no_changes()
        source_ext = draft.baseline.extension
        if dest.suffix.lower() != source_ext.lower():
            return SaveResult.error(
                f"The copy must keep the same file type ({source_ext}); format "
                "conversion is not supported in this version.",
                reason="unsupported_extension",
            )
        policy = self._composite_preflight(draft, changes)
        if policy is not None:
            return policy
        validation = self._validate_draft_parts(draft, changes.has_vba_changes)
        if validation is not None:
            return validation
        build_error = self._build_candidate(
            draft.baseline.path, draft, changes, dest
        )
        if build_error is not None:
            return build_error
        ok, problems = self._verify_candidate(draft.baseline.path, dest, draft, changes)
        if not ok:
            _discard(dest)
            return SaveResult.candidate_failed(problems, None)
        return SaveResult.success(None, dest)

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
