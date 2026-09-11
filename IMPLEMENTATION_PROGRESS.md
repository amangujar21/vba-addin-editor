# Implementation progress

Updated: 2026-09-11

Post-implementation review repairs from
`POST_IMPLEMENTATION_REVIEW_FIX_HANDOFF.md` are in source with automated
coverage. Live Office and packaged GUI qualification remain incomplete.

| Work ID | Status | Commits | Evidence | Blockers |
|---|---|---|---|---|
| CORE-SAVE | automated_pass | pending push | `tests/integration/test_save_retry_ui.py`, `tests/unit/test_fingerprint_and_access.py`, `tests/unit/test_save_pipeline.py` | Live Excel/PowerPoint retry on authentic fixtures |
| CORE-CLASS | automated_pass | pending push | `tests/unit/test_classification_and_delete.py`, `tests/unit/test_adapter_seams.py` | Authentic Office-authored class fixtures; VBE compile/load |
| IMP-01 | automated_pass | pending push | `tests/unit/test_recovery_session_copy.py`, `tests/unit/test_recovery_open.py`, `tests/unit/test_session_baseline_publication.py` | Packaged kill/restart recovery |
| IMP-02 | automated_pass | pending push | post-save copy identity tests in `test_session_baseline_publication.py` | Live overwrite/copy GUI |
| IMP-03 | automated_pass | pending push | `tests/unit/test_conflict_resolution.py`; comparison dialog in `ui/conflict_dialog.py` | Full rebase GUI on authentic files |
| IMP-04 | automated_pass | pending push | `test_undo_text_and_rename`; replace/folder batch undo | Packaged GUI undo chain |
| IMP-05 | automated_pass | pending push | `test_review_is_revision_bound` | Large-diff background loader |
| IMP-06 | automated_pass | pending push | `tests/unit/test_backup_restore.py` | Packaged backup browser restore |
| IMP-07 | automated_pass | pending push | `test_diagnostics_and_review.py`; rotating logs | Copy-report GUI on packaged exe |
| IMP-08 | automated_pass | pending push | `tests/integration/test_search_replace_ui.py` | Packaged project search GUI |
| IMP-09 | automated_pass | pending push | `tests/unit/test_folder_sync_safety.py` | Path-traversal/reserved-name live folder |
| IMP-10 | live_pending | pending push | `scripts/build.ps1 -Release` is a strict gate; `tests/unit/test_release_build_script.py` | Authentic fixtures, packaged rebuild, Office qualification |

Review-fix sections 1–8 have automated regressions. They are **not** `complete`
while authentic XLAM/PPAM/PPTM fixtures, packaged GUI workflows, and live
Office open/VBE/compile/macro/repair-dialog checks remain unexecuted.

Live gates cannot be completed with mocks. Authentic XLAM/PPAM fixtures were
not present under `tests/fixtures/` at implementation time.
