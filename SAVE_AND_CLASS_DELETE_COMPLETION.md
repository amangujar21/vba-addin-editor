# Completion note: save retry, class delete, IMP-01–IMP-10 source work

Date: 2026-09-11
Repository: https://github.com/amangujar21/vba-addin-editor

## Automated evidence

- `python -m pytest tests -q` → 179 passed, 5 skipped (live Office fixtures absent)
- `python -m ruff check .` → clean
- Pinned dependency remains `pyopenvba==3.4.0`

## CORE-SAVE

Save is enabled whenever a document is open and no save is in progress. Open-time host detection is advisory only. Live host status refreshes on focus; the service still blocks in-place save while the corresponding host runs. A blocked or failed save leaves the draft and keeps Save retryable. Close uses Save / Discard / Cancel; Cancel and failed Save keep the window. SHA-256 plus size are authoritative; timestamp-only changes are not conflicts. Win32 exclusive-access probes distinguish in-use, access denied, read-only, and missing files. Process enumeration failure is `process_probe_failed`, not “no host”.

## CORE-CLASS

Ordinary classes are classified from the PROJECT stream plus dir records. Document and designer components stay locked. Draft flags cannot authorize delete/rename; `validate_draft` and `build_candidate` use trusted snapshot capabilities. After delete, PROJECT declarations are scrubbed by logical name to cover pyOpenVBA 3.4.0’s stream-name delete map. Synthetic XLAM/PPTM/PPAM class deletion and stream-name mismatch cases pass automated tests.

## IMP-01–IMP-10 (source)

Session capture, draft recovery checkpoints, hardened Save a Copy, unified undo (Tk text undo disabled), revision-bound review, external-change comparison service, backup catalog/restore, diagnostic envelope and rotating logs, project search/procedure outline, folder export/import, `--version-json` and build-identity plumbing are in source with unit coverage.

## Not completed here (live gates)

- Authentic Office-authored XLAM/PPAM fixtures and VBE compile/load
- Packaged onefile/onedir rebuild and packaged GUI workflows
- Human Excel/PowerPoint repair-dialog qualification

Those remain `live_pending` in `IMPLEMENTATION_PROGRESS.md`. No user add-in was modified. `HANDOFF.md` local edits and `latest_draft_vba.md` were left in place.
