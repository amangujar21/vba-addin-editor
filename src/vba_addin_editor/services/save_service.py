"""The critical safe-save pipeline (plan sections 14, 76).

Never calls host.save() on the original path. Commit primitive is injectable
for failure-injection tests.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path

from vba_addin_editor.adapters.pyopenvba_adapter import (
    AdapterError,
    PyOpenVBAAdapter,
    host_process_for,
)
from vba_addin_editor.adapters.source_codec import CodePageError, validate_code_page
from vba_addin_editor.domain.changes import compute_changes
from vba_addin_editor.domain.document import DocumentDraft, draft_from_snapshot
from vba_addin_editor.domain.results import SaveResult
from vba_addin_editor.platform import paths
from vba_addin_editor.platform import windows_file_ops as wfo
from vba_addin_editor.platform import windows_processes as wp

CommitFn = Callable[[Path, Path, Path | None], None]


@dataclass
class SaveService:
    adapter: PyOpenVBAAdapter
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

    # ------------------------------------------------------------------

    def save_addin(self, draft: DocumentDraft) -> SaveResult:
        adapter = self.adapter
        original = draft.baseline.path

        # Stage A — change detection (no write when clean; plan 59).
        self._report("Checking for changes")
        changes = compute_changes(draft)
        if changes.is_empty:
            return SaveResult.no_changes()

        # Stage B — host/process preflight.
        self._report("Checking Office is closed")
        if self._probe_process(original):
            return SaveResult.blocked(
                "office_running",
                f"Close {host_process_for(original)}'s application (Excel/PowerPoint) "
                "before saving this add-in.",
            )
        self._report("Checking the file is not in use")
        lock_err = self._probe_exclusive(original)
        if lock_err:
            return SaveResult.blocked("locked", lock_err)

        # Stage C — external-change preflight.
        self._report("Checking the original file")
        if not paths.fingerprint_matches(original, draft.baseline.file_fingerprint):
            return SaveResult.blocked(
                "external_change",
                "The add-in changed on disk after you opened it. "
                "Reload it before saving.",
            )

        # Stage D — protection/signature preflight.
        if draft.baseline.safety.password_protected:
            return SaveResult.blocked(
                "password_protected",
                "This VBA project is password-protected and cannot be modified safely.",
            )
        if draft.baseline.safety.signature_present and not draft.signed_save_confirmed:
            return SaveResult.needs_signature_confirmation()

        # Stage E0 — draft structural validation (names, uniqueness).
        from vba_addin_editor.services.validation_service import validate_draft

        draft_problems = validate_draft(draft)
        if draft_problems:
            return SaveResult.error(
                "; ".join(draft_problems), reason="invalid_draft"
            )

        # Stage E — strict code-page validation of changed/new bodies.
        self._report("Validating text encoding")
        encoding = adapter.codepage_encoding(draft.baseline.code_page)
        changed = {
            m.current_name: m.body for m in draft.changed_existing_modules()
        }
        changed.update({m.current_name: m.body for m in draft.new_modules()})
        try:
            validate_code_page(changed, draft.baseline.code_page, encoding)
        except CodePageError as exc:
            return SaveResult.error(str(exc), reason="encoding")

        # Stage F/G/H/I — candidate build from a fresh open.
        candidate = paths.candidate_path_for(original)
        self._report("Building edited add-in")
        try:
            adapter.build_candidate(
                original,
                draft,
                candidate,
                allow_signature_removal=draft.signed_save_confirmed,
            )
        except AdapterError as exc:
            _discard(candidate)
            return SaveResult.error(str(exc), reason="build")

        # Stage J — candidate verification.
        self._report("Verifying the edited add-in")
        verification = adapter.verify_candidate(original, candidate, draft)
        if not verification.ok:
            keep = candidate.with_suffix(candidate.suffix + ".failed")
            _rename_or_discard(candidate, keep)
            return SaveResult.candidate_failed(verification.problems, keep)

        # Stage L — narrow the race window.
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
                "The add-in changed on disk while the edit was being prepared. "
                "Nothing was overwritten.",
            )
        lock_err = self._probe_exclusive(original)
        if lock_err:
            _discard(candidate)
            return SaveResult.blocked("locked", lock_err)

        # Stage K/M — backup + ReplaceFileW commit.
        backup = paths.backup_path_for(original)
        self._report("Creating backup and replacing original")
        try:
            self._commit(original, candidate, backup)
        except wfo.FileOpsError:
            return SaveResult.error(
                "Windows could not safely replace the add-in. Your files have been "
                "preserved for recovery.",
                reason="commit_failed",
            )

        # Stage N — post-commit verification (candidate was moved onto the
        # original by ReplaceFileW, so verify the ORIGINAL path).
        self._report("Final verification")
        final = adapter.verify_candidate(original, original, draft)
        if not final.ok:
            return SaveResult.recovery_required(backup, final.problems)

        new_snapshot = adapter.open_snapshot(
            original, paths.fingerprint(original), host_process_running=False
        )
        draft.baseline = new_snapshot
        refreshed = draft_from_snapshot(new_snapshot)
        draft.modules[:] = refreshed.modules
        draft.signed_save_confirmed = False
        return SaveResult.success(backup, original)

    # -- Save a Copy (plan 47) ------------------------------------------

    def save_copy(self, draft: DocumentDraft, dest: Path) -> SaveResult:
        changes = compute_changes(draft)
        if changes.is_empty:
            return SaveResult.no_changes()
        if draft.baseline.safety.password_protected:
            return SaveResult.blocked(
                "password_protected",
                "This VBA project is password-protected and cannot be modified safely.",
            )
        if draft.baseline.safety.signature_present and not draft.signed_save_confirmed:
            return SaveResult.needs_signature_confirmation()
        encoding = self.adapter.codepage_encoding(draft.baseline.code_page)
        changed = {m.current_name: m.body for m in draft.changed_existing_modules()}
        changed.update({m.current_name: m.body for m in draft.new_modules()})
        try:
            validate_code_page(changed, draft.baseline.code_page, encoding)
        except CodePageError as exc:
            return SaveResult.error(str(exc), reason="encoding")
        try:
            self.adapter.build_candidate(
                draft.baseline.path,
                draft,
                dest,
                allow_signature_removal=draft.signed_save_confirmed,
            )
        except AdapterError as exc:
            return SaveResult.error(str(exc), reason="build")
        verification = self.adapter.verify_candidate(draft.baseline.path, dest, draft)
        if not verification.ok:
            _discard(dest)
            return SaveResult.candidate_failed(verification.problems, None)
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
