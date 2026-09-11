# Handoff: repair post-implementation review findings

Date: 2026-09-11

Repository: vba-addin-editor

Implementation commit reviewed: 4b37e68

Review fixed point: 206aaabca2d4eb390f278c28408d22ad2a3bc982

Status: implementation instructions only. None of the defects below has been fixed by this handoff.

## Objective

Repair the safety and completeness defects found while reviewing commit `4b37e68`, then re-run the full automated, packaged, and live Office qualification required by the original plan.

The authoritative feature specification remains [SAVE_AND_CLASS_DELETE_IMPLEMENTATION_HANDOFF.md](SAVE_AND_CLASS_DELETE_IMPLEMENTATION_HANDOFF.md). Read it completely, especially the shared contracts, IMP-01 through IMP-10, the acceptance matrices, and the release gates. This document supplements that plan with concrete repairs for review findings; it does not replace or narrow the original requirements.

Also read:

- [README.md](README.md), especially the safety model and release gating.
- [HANDOFF.md](HANDOFF.md), especially key decisions and maintenance notes.
- [SAVE_AND_CLASS_DELETE_COMPLETION.md](SAVE_AND_CLASS_DELETE_COMPLETION.md) and [IMPLEMENTATION_PROGRESS.md](IMPLEMENTATION_PROGRESS.md) as claims to re-verify, not as proof.
- The three-dot diff: `git diff 206aaabca2d4eb390f278c28408d22ad2a3bc982...HEAD`.

Workspace precautions:

- `HANDOFF.md` has user-owned modifications and `latest_draft_vba.md` is untracked. Preserve both.
- Do not edit installed `pyopenvba` files. Keep `pyopenvba==3.4.0` pinned unless separately justified and fully qualified.
- Do not modify real user add-ins. Use disposable copies and dedicated fixtures.
- Do not claim the expanded plan complete while any required packaged or live Office gate remains unexecuted.

## Required implementation order

Implement in the following order. Add a failing regression before each production fix and keep earlier tests green.

1. Repair post-save session/baseline publication.
2. Repair missing-source and changed-source recovery.
3. Complete conflict comparison, resolution, and baseline adoption.
4. Make backup restore a verified transaction.
5. Make folder export/import safe, validated, and atomic.
6. Restore the documented adapter seams and fail-closed verification.
7. Complete the project search/replace UI.
8. Enforce release qualification and update completion evidence.

The first three items are one dependency chain. Do not patch recovery by rebasing the recovered draft directly onto current external bytes.

## 1. Publish a new session baseline after every verified in-place save

### Defect

`SaveService.save_addin` refreshes `DocumentDraft.baseline`, but `MainWindow._handle_save_result` only clears history and creates `complete.marker`. `DocumentSession.captured_path`, `captured_fingerprint`, `baseline_sha256`, and revision remain tied to the pre-save package.

This breaks proposal invalidation, later recovery checkpoints, folder manifest identity, and Save a Copy. A permanent `complete.marker` also causes recovery discovery to ignore edits made after the first successful save.

### Required design

Add one session-level baseline publication operation, owned by `SessionService` or a dedicated service. On a verified successful in-place save it must:

1. Use the committed and post-commit-verified original as the new source of truth.
2. Copy it to a sibling temporary file inside the session directory, flush it, hash it consistently, and atomically replace the session's captured baseline package. Do not mutate the only valid captured package in place.
3. Verify the captured copy is content-equal to the committed original.
4. Update `captured_path`, `captured_fingerprint`, `session.json`, `DocumentDraft.baseline`, and the baseline generation as one logical transition.
5. Increment the session revision so every pre-save review, conflict proposal, folder preview, or search result becomes stale.
6. Clear undo/redo and signature approval at the new baseline boundary.
7. Publish the clean completion/checkpoint state only after the new captured baseline and metadata are durable.
8. Return a structured failure if baseline publication fails after the add-in was committed. Report the actual committed file and recovery artifacts; never describe the save as if the original were unchanged.

Do not merely call the existing `SessionService.replace_baseline`: it updates the fingerprint but does not refresh the captured package or manage marker/checkpoint durability.

Any mutation after a clean save must transition the session back to active/dirty recovery state. Remove or supersede `complete.marker` before publishing the next dirty checkpoint. A later crash must expose those edits in `list_recoverable`.

Save a Copy must always build from the captured package belonging to the current baseline generation. A clean copy after an in-place save must be byte-identical to that new baseline, not rejected as `source_unstable` and not built from the original session package.

### Regression tests

- Save successfully, edit again, checkpoint, simulate process death, and recover the second edit.
- Save successfully and assert session revision and baseline hash change.
- Assert a proposal created before Save is rejected afterward.
- Save successfully, then perform a clean Save a Copy; assert byte identity with the new baseline.
- Save module A, then edit module B and Save a Copy; assert both edits survive.
- Export a source folder after Save and assert its baseline hash matches the committed file.
- Inject failure during captured-baseline publication and verify the committed original and recovery instructions are accurately reported.

## 2. Recover drafts safely when the source is missing or changed

### Defect

`MainWindow._offer_recovery` returns when the original is missing. When the original exists but changed, it opens the current file and applies the old checkpoint directly, bypassing three-way conflict resolution. A later save can overwrite newer external changes.

### Required behavior

Recovery must start from the checkpoint's immutable captured baseline, not from whatever bytes currently occupy the original path.

Create a recovery-opening service that:

1. Validates `session.json`, the published checkpoint generation, schema, IDs, lengths, hashes, paths, baseline package size/hash, and all referenced payloads before constructing a draft.
2. Falls back to the previous complete generation when the newest generation or pointer is corrupt or interrupted.
3. Reconstructs the draft over the captured baseline and revalidates component capabilities from that trusted package. Serialized capability flags must never grant permission.
4. Compares the current original with the checkpoint baseline:
   - Matching original: open the reconstructed dirty draft normally.
   - Changed original: retain the reconstructed draft and open the IMP-03 comparison/resolution workflow using `B = captured baseline`, `D = recovered draft`, and `E = stable current external snapshot`.
   - Missing original: open a recovery-only session with `original_exists=False`; disable in-place Save and enable an explicitly labelled recovered Save a Copy flow.
5. Keeps invalid recovery data available for inspection/export. Do not silently delete or mark it complete.

For recovered Save a Copy, explain that newer/missing source bytes are excluded, require a different destination, and build from the captured baseline. Apply the normal protection, signature, encoding, candidate verification, destination-race, and reparse-point rules.

### Regression tests

- Fresh-process recovery with the original unchanged, changed, and missing.
- Changed source cannot reach in-place Save until conflicts are resolved.
- Missing source can save a verified separate copy without recreating or overwriting the original.
- Corrupt newest generation falls back to the prior valid generation.
- Reject wrong hashes, wrong lengths, duplicate IDs, path traversal, unknown schema, invalid BOM/encoding metadata, oversized metadata, and a mismatched captured package.
- Simulate failure before and after checkpoint flush and pointer publication.

## 3. Complete external-change resolution

### Defect

`MainWindow.compare_external_changes` only lists unresolved items. It never collects choices or applies a proposal. `ConflictService.compare/apply` omits parts of the three-way model and can discard draft-added/deleted modules or merged VBA text.

### Required engine behavior

Build comparisons over the union of stable component identities and XML paths from baseline, draft, and external snapshots. Cover:

- Body edits and renames.
- Draft/external additions and deletions.
- Delete-versus-edit and rename-versus-rename conflicts.
- Module kind/capability changes.
- Editable XML changes, additions, and removals allowed by existing package rules.
- Unparseable or unstable external files as explicit blocked outcomes.

Use stable stream/project identity rather than transient IDs created by independently opening the external file. Do not infer identity solely from a mutable logical name.

For each true conflict require exactly one resolution: Keep Disk, Keep Draft, or validated manual merged text where text merging is supported. Structural choices must be explicit. Do not silently drop local additions, resurrect deletions, or ignore merged VBA text.

Apply must:

1. Verify the proposal's session ID, revision, baseline hash, and external hash immediately before applying.
2. Apply into a cloned draft or immutable candidate state.
3. Validate all names, encodings, XML, protection/signature policy, component capabilities, and cross-component constraints before mutating the live session.
4. Commit the whole resolution atomically as one history/checkpointed operation.
5. Adopt the stable external snapshot as the new baseline, publish a new captured baseline generation, clear signature approval, reset history as specified, and retain the named pre-resolution checkpoint until the next verified save or discard.

### Required UI

Provide a real comparison surface displaying labelled `B-to-E` and `B-to-D` differences. Allow per-conflict resolution, manual text editing where supported, Cancel, and Apply. Cancel must not mutate the session or filesystem. Disable Apply until every required choice is valid.

### Regression tests

Exercise local/external edits, additions, deletions, renames, XML changes, all structural conflicts, manual VBA/XML merge text, stale session revision, stale external hash, validation failure, Cancel, and successful atomic Apply.

## 4. Replace direct backup restore with a verified restore transaction

### Defect

The current UI calls `BackupService.restore` directly. The service reads the backup and replaces the current file without the complete dirty-state, comparison, host/lock, race, safety-backup, and post-commit verification protocol required by IMP-06.

### Required transaction

1. If the draft is dirty, require Save / Discard / Cancel. A failed or blocked Save leaves restore cancelled and the draft intact.
2. Inspect and parse the selected backup read-only. Display identity, timestamp, size/hash, build/catalog provenance when available, and a semantic diff against the current baseline.
3. Capture stable hashes of both the selected backup and current target before confirmation.
4. Before building and immediately before replacement, recheck the selected backup hash, target hash, corresponding Office host, exclusive access, permissions, and path identity.
5. Create and verify a safety backup of the pre-restore current file.
6. Restore through the existing safe replacement primitive. Never use a direct overwrite.
7. Verify the target equals the selected backup and the safety backup equals the pre-restore target.
8. On success, publish the restored file as the new session baseline, refresh every view, reset history/signature state, checkpoint, and update the backup catalog.
9. On partial or post-commit failure, return `recovery_required` with accurate target, candidate, selected backup, and safety-backup paths. Do not start an automatic restore loop.

Test mutation of the target and selected backup at both race boundaries, host/lock failures, dirty Cancel, invalid backup, commit failure, verification failure, catalog-warning behavior, and successful restore/reopen.

## 5. Make folder synchronization safe and atomic

### Export repairs

`FolderSyncService.export_folder` currently replaces managed source and ancestor files without comparing them to the previous manifest. Before any replacement:

- Load and fully validate the previous manifest when present.
- Verify every existing managed file against its recorded hash.
- If a managed file changed externally, produce a conflict preview and do not overwrite it.
- Never modify or delete unmanaged files.
- Stage a complete export generation and publish the manifest last.
- Include and verify separate source, effective-attribute, and ancestor hashes as required by IMP-09.
- Check every resolved source, ancestor, staging, and manifest path for containment and reparse points.

### Preview validation repairs

Treat every manifest value as untrusted. Validate types and required fields before indexing. Reject duplicate IDs/names/paths, type-extension mismatches, malformed UTF-8, invalid headers, corrupt ancestor hashes, oversized individual and aggregate source plus ancestor inputs, absolute/escaped/reparse paths, reserved-name/case collisions, and unsupported attributes. Return structured problems rather than leaking `KeyError`, `UnicodeDecodeError`, or filesystem exceptions.

### Atomic Apply repairs

Do not mutate the live draft while iterating selected changes. Clone the complete relevant draft state, apply every selected addition/edit/rename/delete to the clone, and validate the entire result. Only then replace the live state in one session mutation, record one `folder_sync` history command, increment revision once, and checkpoint once. Any failure must leave draft, history, revision, selection, and recovery state unchanged.

The UI must preserve unchecked defaults, require explicit conflict choices, and never silently select every edit/add after preview. After successful Apply, open Review Changes without saving automatically.

Tests must cover partial-failure rollback, undo/redo of the complete batch, modified export destinations, malicious manifests, invalid encodings/hashes, reparse paths, size limits, header/attribute preservation, host-component restrictions, unchanged round trips, and stale previews.

## 6. Restore architectural seams and fail closed

### OOXML mutation seam

Move ZIP/package-member rewriting currently performed in `PyOpenVBAAdapter` behind `OoxmlPackageAdapter`. The pyOpenVBA adapter remains responsible for VBA/CFB/PROJECT semantics; it may return the replacement `vbaProject.bin` bytes or a mutation result, but it must not own ZIP entry rewriting.

Add an architecture regression test or dependency-direction test that makes accidental package manipulation in `pyopenvba_adapter.py` visible.

### Candidate verification

Do not suppress exceptions while enumerating VBA streams or reading `PROJECT`, `PROJECTwm`, or related structural metadata. Missing, malformed, undecodable, duplicated, or unreadable metadata must append a candidate-verification problem and fail the save before commit.

For class deletion, prove all of the following:

- The deleted module stream is absent.
- Its `dir` record is absent.
- Its `PROJECT` declaration is absent.
- Its `PROJECTwm` mappings are absent.
- Surviving component kinds, bindings, attributes, sources, and unrelated package members remain correct.

Keep semantic verification tolerant only where the serializer is documented to normalize equivalent metadata. Never translate inability to verify into success.

## 7. Complete IMP-08 project search and replace

The service layer alone is insufficient. Replace the current truncated message-box UI with a non-modal project search surface providing:

- Query, Match Case, Whole Word, and Include XML controls.
- Incrementally loaded, paged, clickable results with module/path, line, column, and escaped snippet.
- Procedure outline/navigation for VBA modules.
- Revision-bound results that visibly become stale after mutation.
- Replace preview with individually selected occurrences and exact before/after text.
- One atomic, validated, undoable replacement batch using descending offsets and non-overlapping literal matches.

Cancel or stale results must make no changes. Invalid encoding, XML, capability, or naming results must roll back the complete batch. Add GUI tests for navigation, paging, options, stale results, preview cancellation, atomic replacement, rollback, and undo/redo.

## 8. Make release mode enforce qualification

`scripts/build.ps1 -Release` must not print `Build OK` or publish a release manifest when required fixtures, packaged smoke tests, or qualification evidence are absent. Either:

- Make `-Release` a strict gate that fails on every missing required input/check; or
- Rename any provenance-only mode so it cannot be mistaken for a qualified release, and reserve `-Release` for the strict gate.

Check the exit code of every dependency, test, lint, PyInstaller, self-test, round-trip, and manifest step. Build onedir and onefile from the same clean commit. Record actual embedded identity and executable hashes. A development/dirty build must be explicitly labelled and cannot satisfy release qualification.

Do not change `IMPLEMENTATION_PROGRESS.md` or the completion note to `complete` until authentic XLAM, PPAM, and PPTM fixtures; packaged GUI workflows; Office open/VBE/compile/macro checks; repair-dialog checks; backup restore; recovery restart; folder round trip; and external-conflict resolution have all been executed and recorded.

## Validation commands and evidence

At minimum, run and record:

```powershell
python -m pytest tests -q
python -m ruff check .
python -m pyright
```

At review time, pytest reported `179 passed, 5 skipped`, Ruff was clean, and Pyright was not installed. Install project development dependencies in the approved environment before treating type checking as passed.

Then run the focused packaged and live gates from the original plan. Record:

- Exact source commit and dirty-tree state.
- Python architecture/version and installed dependency versions.
- Fixture paths and SHA-256 values.
- Office application build/bitness.
- Onedir and onefile executable hashes.
- Expected and observed outcome for every workflow.
- Every skip or unavailable gate as incomplete, never passed.

## Definition of done

The repair is complete only when:

- All eight work sections above have regression coverage and pass.
- Recovery cannot overwrite changed external content without explicit resolution.
- Post-save edits remain crash-recoverable and use the newly published baseline.
- Conflict resolution, restore, and folder import are atomic and race-checked.
- Candidate verification fails closed.
- Architectural adapter boundaries match `HANDOFF.md`.
- Search/replace meets the full IMP-08 UI contract.
- Automated, packaged, and authentic Office gates from the original plan are complete.
- README, progress, and completion documentation accurately distinguish automated evidence from live qualification.

## Suggested skills

- `diagnose`: reproduce and minimize each safety defect before changing production code.
- `tdd`: add failing regressions for every repair, especially recovery and transaction races.
- `implement`: execute the ordered work packages while preserving the existing safety transaction.
- `vba-delivery`: qualify authentic Office fixtures, packaged executables, VBE compilation, and exact artifacts.
- `code-review` or `review`: repeat independent Standards and Spec reviews against commit `4b37e68` after implementation.
- `handoff`: record verified results and remaining external gates if the work spans another session.
