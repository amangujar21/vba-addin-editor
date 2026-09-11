# Implementation handoff: save/class fixes and editor reliability improvements

Date: 2026-09-11
Repository: https://github.com/amangujar21/vba-addin-editor
Inspected local HEAD: 206aaabca2d4eb390f278c28408d22ad2a3bc982
Status: investigation and implementation plan only; no production code changed. Expanded scope requires the original fixes plus IMP-01 through IMP-10; see the milestone table and explicit decisions below.

## Task for the implementing model

Fix these user-reported workflows for existing XLAM and PPAM files:

1. Open an add-in in the editor while its Office host is running, edit it, close Office, and save the same draft successfully without reopening the editor, reloading the file, or losing edits.
2. Delete an existing ordinary VBA class module, save in place, reopen, and confirm the class remains deleted.

Include the related retry and close-window data-loss defects described below. Keep PPTM and XML editing working. Preserve the safe candidate/backup/replacement transaction.

This request does not require hot-patching a file that Office still has loaded. Default to retaining the current corresponding-host-must-be-closed policy, but make it current, retryable, and accurately explained. Running Excel blocks XLAM under this policy; running PowerPoint blocks PPAM/PPTM. The unrelated host must not block saving. Allowing saves while the corresponding host runs with unrelated documents is a separate policy expansion: do not silently equate a successful momentary file probe with proof that Office has not loaded or cached the target.

## Existing context and workspace precautions

Read README.md and HANDOFF.md for the transaction and release model. Read XML_MALFORMED_PART_FIX_HANDOFF.md and relevant parts of VBA_Addin_Editor_XML_and_Context_Menu_Implementation_Plan.md before changing candidate verification. Reference those documents rather than duplicating their specifications.

At inspection, HANDOFF.md had pre-existing modifications and latest_draft_vba.md was untracked. Preserve both. Do not overwrite existing handoffs. No AGENTS.md was found by the repository file search; check applicable parent instructions at implementation start. The historical core-plan filename mentioned in README is absent from the current root listing; do not assume it exists.

Pinned dependency: pyopenvba==3.4.0. Keep dependency-specific access inside adapters. Do not edit installed site-packages as the delivered fix. Existing dist artifacts may differ from source; the prior handoff explicitly says the onedir build is stale.

## Evidence and confidence

The following are code findings and isolated reproductions, not a completed live-Office diagnosis of the user's exact files.

### A. Open-time process state permanently gates the UI

- services/document_service.py, DocumentService.open: stores wp.host_process_running(path) in immutable baseline.safety.
- ui/main_window.py, _refresh_state (around lines 181-208): Save button and Office banner use baseline.safety.host_process_running. No live refresh occurs there.
- services/save_service.py, _initial_race_narrowing and _race_narrowing: already probe current process state for each save. Therefore toolbar state and actual service eligibility disagree; Ctrl+S/menu can reach save() even while the toolbar is disabled.
- platform/windows_processes.py: detection is executable-name based and host-specific, not target-file ownership based. Blocking with an unrelated document in the corresponding host is intentional current policy, not evidence of a lock leak.

Executed a temporary synthetic XLAM reproduction with a withdrawn real Tk window, patching the host probe true for load and false afterward. Calling _refresh_state after the transition produced:

    save_button_at_open= disabled
    save_button_after_host_closes= disabled

### B. Failure results also leave Save disabled

save() sets _saving=True and refreshes the UI, disabling Save. Its finally resets the flag but does not refresh. _handle_save_result refreshes on success, but not most blocked/error outcomes. This can strand even an initially eligible draft after a failed attempt until another action happens to refresh state. Add a test invoking the actual button twice, not only the service.

### C. Closing the editor can discard the draft

on_close (around lines 740-760) checks the stale baseline flag, calls showwarning with text asking “Exit anyway?”, then unconditionally destroys the window. showwarning is not a Yes/No/Cancel decision. This is a direct data-loss defect in the reported workflow and belongs in this fix.

### D. Existing ordinary classes are deliberately denied deletion

- adapters/pyopenvba_adapter.py, make_display_kind: all existing non-standard modules become AMBIGUOUS. open_snapshot never supplies is_new_app_class=True.
- open_snapshot grants destructive_ops_safe only to non-read-only standard modules.
- ui/main_window.py, delete_module: blocks modules with destructive_ops_safe=False. rename_module shares the same flag.
- Newly created classes receive permission in the current draft but lose it after save/reopen, because snapshot classification repeats the ambiguous mapping.

The temporary fixture created OrdinaryClass using the pinned library, saved it, and opened it in the editor:

    reopened_class_kind= ambiguous delete_safe= False

This proves the application classification restriction; it does not prove all real-world class serialization cases are correct.

### E. Timestamp-only changes are treated as content conflicts

platform/paths.py says SHA-256 is authoritative, but fingerprint_matches compares the whole FileFingerprint dataclass, including mtime_ns. An isolated os.utime change with identical bytes produced:

    same_bytes_changed_mtime_matches= False

This is a potential additional blocker after Office closes; actual Office timestamp/content changes were not measured. A true byte change must continue to block overwrite.

### F. Windows diagnostics lose useful distinctions

exclusive_access_error translates every failed CreateFileW probe into “Another program is using this add-in.” This conflates sharing violations with permission errors, missing paths, etc. save_addin catches FileOpsError without retaining its Win32 code or file-state details. Fix diagnostic propagation so remaining failures can be identified.

The adapter uses host context managers, and pinned VBAHostFile.close closes its ZIP handle. A persistent handle leak was not demonstrated; investigate only if a real lock remains after host exit.

## Implementation sequence

### 1. Establish deterministic regressions before production changes

Use tests/integration/test_gui_end_to_end.py, tests/unit/test_save_pipeline.py, and tests/unit/test_adapter_and_platform.py. Add focused test files if that keeps responsibilities clear.

Use temporary fixtures and injected process probes. Drive a real Tk Text widget and Save button with messageboxes stubbed; existing GUI tests commonly invoke save_service directly, which misses the disabled-button defect. Preserve baseline and draft identity across a blocked attempt. Capture the current failures before fixing them.

Record actual build identity and exact error reason when reproducing on customer fixture copies. Do not alter installed add-ins or terminate the user's Office sessions for diagnosis.

### 2. Make Save retryable and host state transient

Files: ui/main_window.py, services/document_service.py, domain/document.py, platform/windows_processes.py as needed.

- Do not use open-time host state as a permanent action gate. Prefer enabling Save whenever a document exists and no save is in progress; the service performs authoritative fresh preflight on each attempt. Clean drafts can retain the existing no-change behavior.
- Treat process availability as transient state, separate from immutable file/project facts. Remove or clearly demote the snapshot flag so banners, status reports, and close handling cannot reuse it as current truth.
- Refresh current advisory host status on suitable UI events, such as focus return and completion of every save attempt. Optional polling must use cancellable Tk after callbacks and must not reload the draft or block the event loop.
- Keep current service probes before build and immediately before replacement. Never auto-reload to “fix” stale process status; that loses edits.
- In finally, reset _saving and restore UI state for every outcome, including exceptions. Translate unexpected boundary failures into a useful result/dialog without swallowing diagnostic context.
- Make toolbar, File menu, keyboard shortcut, and save-on-exit share the same command semantics.
- Replace the stale-host close branch with a real Save/Discard/Cancel decision. Cancel keeps the window and draft. A failed/blocked Save keeps the window and draft. Only an explicit Discard or a completed save permits closing a dirty window.
- Phrase the host policy accurately: “Close Excel before saving this add-in” rather than claiming the exact file is necessarily locked.

### 3. Correct content-conflict handling and diagnostics

Files: platform/paths.py, platform/windows_file_ops.py, domain/results.py, services/save_service.py, affected UI paths.

- Make content equality authoritative for timestamp-only differences: compare SHA-256 and content size, with metadata used only as an optimization/diagnostic. Do not skip hashing merely because size and mtime match; equal-length external edits can preserve both.
- Capture a consistent fingerprint: detect changes during reading, using a bounded retry or explicit unstable-source failure. Keep final conflict checks immediately before commit; document the remaining check-to-replace race instead of claiming the probe eliminates it.
- If actual content changed, retain the draft and baseline and block overwrite. Do not update the baseline hash merely to allow Save. Reload must explicitly disclose loss of edits and offer cancellation.
- Preserve native error codes immediately after failed Win32 calls. Distinguish in-use/sharing, access denied/read-only, missing source, process-enumeration failure, and commit/recovery failures. Unknown process enumeration must not silently mean no host exists.
- Carry stage, reason, Win32 code, and candidate/backup paths through SaveResult.details where appropriate. Avoid logging VBA source or sensitive contents.
- Do not turn every replacement failure into an automatic retry: ReplaceFileW can have partial recovery states. Check actual paths and preserve recovery artifacts.

Save a Copy caution: it currently builds from the current source path with no original fingerprint preflight and writes directly to the chosen destination. Do not present it as a universally safe external-conflict escape. If this work adds it to recovery dialogs, first add source consistency checks, source/destination identity protection, and safe handling of existing destinations. Otherwise leave that larger change separately tracked.

### 4. Classify existing components from actual project metadata

Files: adapters/pyopenvba_adapter.py, domain/document.py, services/validation_service.py, ui/main_window.py.

Use the PROJECT stream to distinguish ordinary classes from host/designer components, cross-checking the dir module records. Microsoft identifies PROJECT as the source of project-item types and requires correspondence with dir records:

- https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/1883ad6a-628b-47e0-aac0-c292a87655bc
- https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/dadaf851-0a23-465b-9f64-952ca6d88031

Pinned pyopenvba.vba already exposes parse_project_stream(raw, code_page=...), with standard_modules, class_modules, document_modules, and base_classes. CFB.from_bytes(host.vba_project_bytes()).get_stream('PROJECT') is an available adapter-local read route. Inspect the installed version before choosing the exact seam.

Classification rules:

- Consistent ordinary class declaration plus matching non-standard dir record, no document/designer conflict, no read-only restriction: display as class and permit deletion.
- Document declarations: preserve host binding and deny deletion.
- Designer/BaseClass declarations and designer storage evidence: preserve designer components and deny deletion.
- Missing, malformed, duplicated/conflicting declarations, decoding ambiguity, or mismatched identity: fail closed for destructive operations with a reason the UI can explain.
- Match logical names case-insensitively and map them to stream names explicitly. Never infer subtype from a name such as Class1 or ThisWorkbook, from VB_PredeclaredId alone, or from the broad other enum.
- The library parser decodes with errors='replace'; destructive classification must detect decoding loss rather than treating permissive parsing as proof.

Keep project protection and signature rules unchanged. Do not simply set every other module's destructive_ops_safe=True.

The shared flag also enables rename. Decide explicitly whether classification qualifies both operations. If only deletion is qualified, introduce separate capabilities instead of accidentally expanding rename support. Add a reason field if necessary; the existing single boolean cannot explain unknown versus host-bound versus read-only restrictions.

Enforce deletion permissions below the GUI as well, resolving original modules against trusted snapshot/fresh metadata. validate_draft currently checks surviving names and protection, not deletion capability. A crafted ModuleDraft or headless caller must not bypass the UI safety rule by toggling is_deleted/destructive_ops_safe.

### 5. Qualify deletion serialization and verification

Files: adapters/pyopenvba_adapter.py and adapter/save tests; dependency compatibility shim only if evidence requires it.

build_candidate already calls project.delete_module for deleted originals. Pinned delete_module records stream names; _host.py removes streams and rewrites dir, PROJECT, and PROJECTwm on structural changes. Verify these operations before adding any new binary mutation code.

Pay particular attention to logical name versus stream-name differences: pending deletions use stream names, while PROJECT declarations identify logical names. Prove that declaration cleanup works when these differ; this is an investigation target, not a confirmed dependency defect.

After deletion, verify the candidate and saved original contain exactly the expected modules, no deleted module stream/declaration/mapping, and coherent PROJECT/dir/PROJECTwm structures. Verify surviving module types and semantic attributes, host/document bindings, designer storage, references, and source content are preserved. Compare relevant structure semantically where serializer changes are expected; require untouched non-VBA package payloads to stay byte-identical under the existing XML allowed-change rules.

An independent Office/VBE check is required because reopening with the same parser can repeat the writer's assumptions. Keep compilation and execution qualification outside the normal editor runtime.

## Acceptance matrix

Run XLAM and PPAM for all applicable cases; include PPTM regression coverage.

| Scenario | Required result |
|---|---|
| Host running at open, edit, host closes, click Save | Same draft saves without reload; backup matches pre-save original |
| Host stays running through attempt | Current-policy block; draft retained; Save remains retryable |
| Host absent at open, starts before save | Fresh preflight blocks |
| Host starts between candidate verification and commit | Final check blocks; original unchanged |
| Blocked/error attempt followed by resolution | Actual toolbar can retry and succeed |
| Only unrelated Office host running | No process-policy block |
| Close window with stale/running host state | Save/Discard/Cancel; no implicit discard |
| Sharing lock held then released | Specific blocker then successful retry without reload |
| Access denied or missing source | Accurate reason, no misleading in-use assertion |
| Timestamp-only change | No false content conflict |
| Actual content changed, including same-size/mtime case | Overwrite blocked; dirty draft preserved |
| Existing ordinary class deleted | Absent after candidate verification, commit, reopen, Office load |
| New class saved then reopened then deleted | Same supported behavior as existing class |
| Logical name differs from stream name | Correct deletion across all relevant structures |
| Document/UserForm/unknown/conflicting/read-only component | Deletion denied by GUI and service/adapter |
| Delete class plus edit VBA/XML in one save | One verified candidate and backup; allowed differences only |
| No changes / protection / signatures / malformed untouched XML | Existing guarantees retained |
| Commit or post-commit verification failure | Correct recoverable state and artifact paths; no silent success |

For every pre-commit failure assert original bytes unchanged, no inappropriate backup/commit, and draft edits/deletion flags retained. On success assert draft clean, fresh baseline, exactly one backup, and a second save performs no write.

## Validation and delivery

1. Run focused failing regressions, implement, then rerun them.
2. Run python -m pytest tests -q and python -m ruff check . using the project's Windows Python environment. Record actual results; prior HANDOFF counts are historical, not current evidence.
3. Add authentic Office-authored XLAM/PPAM fixtures with an ordinary class plus relevant host/designer modules. Synthetic PPAM tests rename a PPTM package and are insufficient for release qualification.
4. On disposable copies, reproduce the user's sequence in Excel and PowerPoint, confirm no repair dialog, inspect class removal in VBE, compile the project, and run a benign known entry point. Use a class with no remaining references or update fixture references intentionally so compilation failures test the tool, not a deliberately broken VBA program.
5. Rebuild both desired packaging variants with scripts/build.ps1; inspect its fixture prerequisites first. Run packaged --self-test and --self-roundtrip, then the actual packaged GUI retry/delete workflows. Headless modes alone do not exercise the Save button. Record source commit, dependency versions, executable SHA-256, fixture identities, Office versions, and skipped/live gates.
6. Update README to describe the retry workflow and verified ordinary-class support accurately. Record final results in a new completion note; preserve pre-existing HANDOFF.md edits.

Suggested small commits: (1) failing workflow regressions; (2) transient host/UI retry and close behavior; (3) content comparison and Win32 diagnostics; (4) metadata classification plus below-UI capability enforcement; (5) deletion integrity and live-fixture qualification; (6) documentation and packaged release evidence. Keep tests passing at review checkpoints after paired fix commits.

## Suggested skills

- diagnose: reproduce each user symptom before selecting additional causes.
- vba-delivery: when building or validating Office/VBA fixtures and exact delivered artifacts; read its requirements before claiming Office qualification.
- implement: carry out the bounded work above using repository conventions.
- handoff: record completed changes, evidence, and remaining live gates for the next session.

## Investigation limits

Only temporary synthetic files were created during this planning session. The current full test suite, real Office sessions, and packaged executable were not exercised. No user add-in was modified. The temporary reproduction confirmed the stale-button state, existing-class restriction, and timestamp-only mismatch. Other findings are grounded in inspected code but require regression/live qualification. Windows sandbox process startup failed locally; approved read/reproduction commands ran outside that broken sandbox. No implementation changes or GitHub issues/PRs were created.

## Expanded implementation scope: ten required improvements

Added 2026-09-11 at the user's request. All ten work packages below are required scope, not optional suggestions. Priorities determine order, not permission to omit work. Complete the original save/class fixes first. Deliver reviewable increments; do not claim the whole plan complete while required work or release gates remain unverified.

This section supersedes earlier optional wording about Save a Copy and narrows earlier implementation choices where explicitly stated. Existing safety and compatibility requirements still apply. Saving into an add-in that Office has loaded remains out of scope. No cloud service, telemetry, COM-dependent editor runtime, automatic Office termination/unloading, or automatic Git commit/push is required.

### Shared contracts and implementation boundaries

These are proposed interfaces/files, not claims that they already exist. Equivalent existing abstractions may be extended if they meet the contracts.

1. Add a DocumentSession separate from DocumentDraft: unique session_id, monotonically increasing revision, immutable baseline package copy/fingerprint, mutable draft, undo history, and last checkpoint revision. Every accepted mutation, undo/redo, and baseline replacement increments revision. Dirty state remains a comparison against baseline, not revision inequality.
2. At Open, capture the source into an application-owned session directory, hash the captured bytes, and parse the snapshot from those exact bytes. Retain the original user-facing path in the snapshot. Recheck original content after capture; retry twice on instability, then return source_unstable without adopting a mismatched snapshot. Open never writes to the original.
3. Use %LOCALAPPDATA%/VBAAddinEditor/sessions/<session_id>/ with inherited per-user permissions. Keep package bytes separate from metadata. Do not store drafts beside installed add-ins or in PyInstaller extraction directories. Recovery files contain user code; exclude them from diagnostics and releases.
4. Route edits through a session mutation boundary: validate capability/identity, apply one change or atomic batch, increment revision, emit one event, schedule recovery. Programmatic loading/refresh must not generate edits. Stable module IDs are valid within a baseline generation; list indexes are not identities.
5. Save review, conflict resolution, replacement, and import proposals bind to session revision and baseline identity. Applying a stale proposal returns stale_proposal and changes nothing. Recompute its preview instead of using stale offsets/names.
6. Add a shared source composer/parser in adapters/source_codec.py. Keep hidden attributes separate from visible body; compose from current draft body and effective attributes, including approved name changes. Never append imported headers blindly to body. Test module-level and procedure-level attributes so source splitting cannot drop attributes following procedure declarations.
7. Candidate builders read the immutable baseline selected for the operation. In-place save still compares the actual original before building and immediately before commit. Post-commit verification uses the actual backup. A baseline copy does not authorize overwriting an externally changed original.
8. Tk access stays on the main thread. Background I/O/hash/diff jobs consume immutable inputs; deliver results via after/queue and discard stale revisions. Permit one save/restore commit per session and freeze editing during commit. Other instances are handled through ownership and current-file checks.
9. Keep logic in services/domain/platform; MainWindow coordinates views. Proposed modules: recovery_service.py, session_service.py, conflict_service.py, review_service.py, search_service.py, folder_sync_service.py, diagnostics_service.py, domain/session.py. Extract BackupService into backup_service.py if useful while migrating callers safely.
10. Revalidate recovered/imported capabilities from trusted package metadata. JSON flags never authorize destructive operations. Clear signed_save_confirmed after recovery, rebase, or baseline replacement; confirmation belongs to the current transaction.

### IMP-01 — P0: draft autosave and crash recovery

**Behavior**

Autosave recovery data by default without saving the add-in. Show “Draft recovery saved” with its time separately from “File saved.” Checkpoint after 1 second without edits and at least every 10 seconds during continuous typing. Inject the scheduler/clock for tests. A crash may lose changes since the last completed checkpoint; do not promise per-keystroke durability.

Checkpoint immediately after structural/batch operations, before save/restore/reload/rebase, and when focus leaves the app. Flush active editor text first. Recovery write failure shows a persistent nonmodal warning, leaves editing enabled, and offers “Save Recovery Now.” Checkpoint success is never add-in save success.

**Storage/lifecycle**

- Versioned UTF-8 JSON: schema_version=1, session_id, revision, UTC times, source locator, baseline hash, and full reconstructable draft state: originals, hidden attributes, tombstones, names, XML encoding/newline/BOM, selection/cursor. Binary payloads are separate hash-verified files; no pickle/executable serialization.
- Keep two complete checkpoint generations. Write/flush the new generation and atomically publish a manifest pointer only after its payloads are durable. Never overwrite a committed generation in place. Previous generation survives an interrupted write.
- Validate schema, IDs, hashes, lengths, and paths. Reject payload references outside the session directory, duplicate IDs, unknown schemas, and corrupt hashes. Retain invalid data for explicit inspection/export, not silent deletion. Set a 64 MiB draft-metadata limit, reject larger metadata with recovery_write_failed, and retain the previous valid checkpoint; baseline package size is separately recorded and checked against actual file size.
- Hold an exclusive per-session ownership handle released on process exit. Discovery skips actively owned sessions; PID alone is insufficient evidence.
- Startup lists unclosed dirty sessions with filename/path/time/revision. Actions: Recover, Keep for Later, Delete Recovery. Only explicit deletion removes selected unsaved recovery.
- If disk matches baseline, reconstruct draft without writing disk. If disk differs, recover against the saved baseline and enter IMP-03. If source is missing, retain draft and allow IMP-02 recovered Save a Copy; block in-place Save.
- Successful in-place save publishes a new baseline and clean completion marker before cleanup. Explicit Discard persists a discard marker before closing. Clean/discarded sessions may be removed next startup; dirty/invalid sessions never expire automatically. Warn at 500 MiB total storage and offer recovery management; never evict unsaved work automatically.
- Save a Copy does not clear original dirty state/history/recovery. Record candidate hash and operation state before commit. If a crash occurs after replacement but before completion publication, compare disk with that hash and reconcile instead of replaying edits blindly.

**Acceptance**

Fresh-process recovery of VBA/XML edits, add/rename/delete, non-ASCII, and missing-source sessions. Inject failures before/after flush and pointer publication; recover latest intact generation. Test concurrent ownership, disk full/write denial, truncated JSON, unknown schema, path traversal, size limits, corrupted baseline, explicit Discard, and crash after commit. Autosave must never change add-in bytes.

### IMP-02 — P0: safe Save a Copy

Refactor save_copy to build a unique temporary sibling of the destination, never directly over it. Replay the draft against the verified session baseline.

1. Flush editors and bind to revision. Require the original extension. Reject source/destination identity through normalized Windows paths plus file identity for existing files: case aliases, relative paths, hard links, and resolved links. Reject internal session destinations. For v1 reject reparse-point destinations/ancestors with an explanation to choose a regular folder; recheck this before commit.
2. Missing/changed original requires explicit “Save recovered draft as a separate copy,” explaining newer disk changes are excluded. Never silently build from current external bytes. Destination must differ from original.
3. Existing destination: explicit overwrite review, current hash capture, host/lock policy, and final hash/availability check. Cancel changes nothing. New destination: a running host alone does not block creating a new separate file; do not overwrite a file another process creates before commit.
4. Build/verify with existing VBA/XML protection/signature rules. A clean draft creates a verified byte-identical baseline copy; it must not return no_changes. Clean in-place Save still performs no write.
5. Existing destination commits through ReplaceFileW with mandatory backup of that destination. New destination uses a same-volume atomic move that fails if destination exists. Never delete-then-rename. Verify committed bytes against the candidate hash and semantic expectations.
6. Report destination and recovery/backup paths. Source bytes, path, baseline, dirty state, and history remain unchanged. Opening the copy is separate and first resolves unsaved original-session state.
7. Partial commit returns recovery_required with actual path states. Delete only temporary files proven not to be the sole recoverable copy.

**Acceptance**

Same path/case alias/hard link; new/existing destination; overwrite Cancel; destination content/identity change; destination appearance race; locks; missing/changed source with explicit recovered-copy choice; clean copy; build/commit/post-commit failures. Destination backup equals destination-before bytes; source remains unchanged and dirty. Include XML-only protected-VBA and approved signature removal.

### IMP-03 — P1: external-change comparison and resolution

Inputs: B = immutable baseline, D = draft, E = stable current external snapshot. Display labeled B-to-E and B-to-D diffs. Comparison never reloads or changes the draft/source.

Match modules by baseline logical identity plus stream mapping. Unprovable correspondence becomes add/delete conflicts; no fuzzy matching. XML identity is exact package path. Detect body/attribute/capability changes, additions, deletions, renames, and noneditable package differences.

**Resolution rules**

- Per component/property: D=B takes E; E=B takes D; D=E takes the common value. All other combinations require a choice. Treat a whole module body or XML part as one property in v1; automatic line merging is not required.
- Group structural operations with dependent body/attribute changes. Delete-versus-edit, competing rename, name collisions, type changes, removed locally edited XML, and protection/codepage/signature changes require explicit review.
- Text conflicts offer Keep Disk, Keep Draft, or manual merged text. Validate naming/encoding/XML/capability before Apply. Never insert conflict markers automatically.
- Uneditable external package payloads come from E; no old-payload override. If E cannot be parsed safely, offer Cancel, recovered-copy/export, or explicit discard-and-Reload; do not fabricate a merge baseline.
- Enable Apply only after all conflicts resolve and validation passes. Recheck disk hash against E. Another change invalidates the proposal and leaves D intact. Otherwise adopt E as baseline and apply resolved changes as the new dirty draft; do not save automatically.
- Recompute capabilities, clear signature approval, checkpoint, and reset undo history at this baseline boundary. Retain a named pre-resolution checkpoint until the next verified save or explicit discard.

**Acceptance**

Local-only/disk-only/equal/divergent edits, add/delete-versus-edit, rename collisions, stream identity mismatch, codepage/protection changes, XML add/remove, malformed E, second external mutation during review. Cancel/stale proposal preserve D. Accepted rebase preserves unrelated E bytes and leaves local edits dirty until Save.

### IMP-04 — P1: unified undo/redo

Implement one session history for VBA/XML typing, add/delete/rename, import, revert, and batch replacement. Route Ctrl+Z, Ctrl+Y, Edit menu, and context-menu undo/redo through it. Do not leave a separate Tk undo stack able to diverge from draft state.

- Commands have stable target identity and before/after state. Use reversible text deltas and affected-object snapshots, not entire package copies per keystroke.
- Coalesce adjacent typing/deletion in one target for up to 750 ms. Cursor/selection movement, target switch, paste, explicit command, or timeout ends a group. Paste and Replace All are one command each; import/sync is one atomic batch.
- Programmatic loads/refreshes do not record history. Switching modules preserves history and text. Undo restores IDs, names, body, attributes, ordering, selection, and deletion state. Redo reproduces the operation.
- Editing after Undo clears Redo. Recompute dirty state against baseline; undo back to baseline becomes clean.
- Clear history after verified in-place save, accepted rebase, reload, or restore. Explain that boundary in transition dialogs. Save a Copy retains history. Recovery restores edits with empty history; cross-crash undo is not required.
- Limit history to 1,000 commands or 64 MiB of payload, whichever comes first. Evict whole oldest commands, retaining at least the newest. If one command exceeds the limit, keep it and warn older history was cleared. Never discard current contents or recovery due to history limits.

**Acceptance**

Edit A, edit XML, rename B, delete B, import two modules, then undo/redo the chain across views. Verify identity/text/attributes. Test fake-clock typing groups, paste/context menu, undo-to-clean, redo invalidation, limits, save boundaries, and no phantom commands from <<Modified>> refresh loops.

### IMP-05 — P1: pre-save change review

Extend review_changes/_preflight_dialog with a revision-bound ReviewModel derived from compute_changes and the shared composer. Show operation badges and read-only unified diffs with line numbers. Group VBA/XML; show old/new names for renames and removed/added source for delete/add. Rename plus body edit appears as one item with both operations.

Display destination/type, backup policy, operation counts, encoding errors, and applicable signature consequences. Show VBA signature removal only when signed VBA changes. Preserve the separate XML-only/protection/package-signature rules.

Every dirty in-place Save and Save a Copy opens this review, offering Save/Save Copy and Cancel. Replace the existing duplicate preflight summary with this window. Require an initially unchecked signature-removal checkbox when relevant; approval is bound to revision/baseline. Clean Save a Copy needs only a destination summary.

Review is read-only in v1: no hunk editing or partial staging. Navigate to the editor after closing/cancelling review. Save verifies the reviewed revision then runs fresh preflight. A stale review must be rebuilt. Load large diffs per item and cancel obsolete background work.

**Acceptance**

Rename+edit, add/delete class, unsaved class body, XML BOM/newlines, XML-only protected VBA, signature accept/cancel, stale revision, external change after review. Review/cancel must mutate neither draft nor filesystem and create no backup/candidate.

### IMP-06 — P1: backup browser and verified restore

Catalog successful save/copy/restore backups: schema version, original identity/path, backup path, SHA-256, size, UTC time, operation type, and build identity. Update atomically after commit. Catalog failure after a verified commit is a warning, not a falsely reported failed save; offer discovery/reindexing.

Backup Browser shows date/size/source/integrity/change summary. Compare selected backup with current file through the review engine. Validate on selection; filename alone proves neither ownership nor validity. Discover matching legacy same-folder names as unindexed candidates and allow a user-selected backup through identical validation. No automatic backup pruning in this release.

Restore sequence:

1. Resolve dirty state through Save/Discard/Cancel. Cancel or failed Save aborts. Checkpoint before explicit discard/restore and retain that generation through completion.
2. Require matching extension and readable full VBA/XML package under existing malformed-baseline tolerance. Preserve signatures/protection as bytes; restoration is not source mutation. Reject aliases and corrupt selection.
3. Show backup-versus-current diff and require Restore. Capture both hashes. Check host/exclusive access initially and before commit. Block if either file changes while preparing.
4. Copy selected backup to a unique sibling candidate; verify hash/byte equality and readability. Replace current with mandatory pre-restore safety backup. Never consume/change selected backup.
5. Verify current equals selected backup and safety backup equals pre-restore current. On success adopt fresh baseline, reset history/signature approval, update recovery/catalog and all views. Partial/post-commit failure returns recovery_required retaining artifacts; no automatic restore loop.

**Acceptance**

Corrupt/changed/missing catalog entries, legacy backup, extension mismatch, aliases, dirty Cancel, locks, backup/current mutation during preparation, catalog failure after success, partial commit, and final failure. Selected backup remains identical; safety backup contains actual previous current bytes. Preserve untouched malformed XML.

### IMP-07 — P1: actionable diagnostics

Use a common result envelope across services: operation_id UUID, operation type/stage, stable reason, UTC time, retryable flag, Win32 code/name, revision, relevant hashes, artifact paths, and original/candidate/backup existence. Capture native error immediately after the failed call.

Minimum reasons: office_running, process_probe_failed, file_in_use, access_denied, read_only, source_missing, source_unstable, external_change, destination_changed, destination_exists, invalid_draft, encoding, invalid_xml, unsupported_component_operation, candidate_failed, commit_failed, recovery_required, recovery_write_failed, stale_proposal. Unknown errors remain unknown; do not relabel them as locks.

Map each reason to plain explanation and a useful action: Retry availability, Compare Changes, Choose Destination, fix validation, or open recovery artifacts. Retry runs full preflight; it never bypasses policy or repeats an uncertain commit without state inspection.

Add Copy Diagnostic Report with a preview. Include build/dependency/Python/Windows versions, extension, operation/stage/reason, numeric errors, counts, and artifact states. Exclude VBA/XML, source imports, recovery payloads, environment dumps, and exception strings containing user text. Redact full paths by default; an unchecked Include Full Paths option changes only the previewed report. No upload.

Replace app.py's unbounded FileHandler with rotation: five files total, at most 2 MiB each. Log the same sanitized fields. Keep full recovery paths usable in the local recovery UI even when shared reports redact them.

**Acceptance**

Inject all minimum reasons at service boundaries; assert correct stage/reason/action. Distinguish sharing/access denial and preserve last-error codes. Test partial-commit artifact reporting, rotation, report copying, and sanitization using sentinel secrets in VBA/XML/paths/exceptions. No report/log may contain source sentinels.

### IMP-08 — P2: project-wide search and procedure navigation

Ctrl+Shift+F opens project search against the flushed current draft, not disk/baseline. Default scope is surviving VBA bodies, including new modules; explicit Include XML toggle adds editable/readable XML text. Exclude deleted modules and hidden attributes. Support literal Match Case/Whole Word; regex and semantic refactoring are out of scope.

Whole Word boundaries mean adjacent characters are not Unicode letters/digits or underscore. Show target/path, 1-based line/column, bounded snippet. Clicking opens and selects the current match; revision changes invalidate offsets and require refresh. Count all matches but render pages of 200. Background search uses immutable snapshots and discards cancelled/stale jobs.

Replace preview lists exact selected occurrences with before/after text and read-only exclusions. Apply only at the reviewed revision, in one undoable batch, descending offsets per target. Use left-to-right non-overlapping matches and literal replacement text. Validate naming-relevant constraints where applicable, encoding, XML, and capabilities for the entire batch; any invalid result changes nothing. No matches means no mutation.

Procedure outline lists Sub, Function, Property Get/Let/Set with visibility/static modifiers and continued signatures. Use a tokenizer that skips comments/strings and joins continuation lines, not substring matching. List conditional-compilation branches without claiming which is active. Navigation is lexical, not compilation or symbol resolution; no reference-renaming feature.

**Acceptance**

Unsaved cross-module/XML matches, case/word/non-ASCII boundaries, deleted exclusions, stale results, cancellation, pagination counts, exact selection. Test batch undo, read-only exclusion, non-overlap behavior, literal replacement, and all-or-nothing validation. Outline ignores declarations in comments/strings and recognizes continued signatures/properties.

### IMP-09 — P2: folder source export and synchronization

First fix ImportExportService.export_module: existing classes currently export baseline.full_source, omitting current body edits. Use the shared composer for single-module export and folder export. Preserve attributes and current names/body; add a regression before fixing it.

**Managed export**

- Add Export Source Folder. Write vbaae-project.json with schema_version=1, project identity, extension, baseline hash, revision, UTF-8 encoding, LF newline convention, and entries containing opaque export ID, logical name, verified kind, safe relative path, original stream identity, source/attribute hashes, and explanatory capabilities.
- Include a .vbaae-baseline directory containing the exact exported source texts keyed by opaque export ID, with hashes in the manifest. These are the common ancestor for later folder-versus-draft comparison; hashes alone cannot reconstruct that ancestor. Both baseline texts and manifest are versionable, contain user source, and must never enter diagnostics automatically. Original add-in package bytes are not exported.
- Ordinary modules use .bas; ordinary classes use .cls with appropriate attributes/preamble. Host/document/designer/unknown components use .vba.txt reference files, supporting only permitted existing-body updates. They must not be represented as freely importable classes. UserForm layouts/binaries and XML folder synchronization are out of scope.
- Use Windows-safe filenames with opaque ID suffixes. Never derive a path solely from arbitrary VBA names. Reject/escape reserved names, separators, trailing spaces/dots, and case collisions. Every resolved path must remain inside the export folder; reject reparse-point paths.
- UTF-8/LF is the exchange encoding, not a claim of byte identity to the project. Store original codepage and effective attributes and validate strictly on return to the add-in.
- Stage export and publish manifest last. Hashes must detect interrupted mixed generations. Before replacing managed files, compare them to the previous manifest. Externally edited managed files trigger a conflict preview; never overwrite silently. Unmanaged files are never changed/deleted. Do not run Git commands.

**Import preview/apply**

- Add Preview Folder Changes. Treat manifest as untrusted. Reject unknown schema, escaped/absolute/reparse paths, duplicate IDs/names, type/extension mismatch, corrupt baseline text hashes, lossy decoding, inputs above 8 MiB each, or aggregate source/baseline-text inputs above 64 MiB. Fail explicitly; do not truncate. Manifest itself is limited to 4 MiB.
- Require manifest baseline identity to match the current session baseline. A mismatch requires fresh export or explicit per-module import, not automatic application to an unrelated/rebased project.
- Compare the exported ancestor text with both folder and current draft: folder-only changes are proposed; identical changes are no-ops; divergent folder/draft changes require Keep Draft or Use Folder. Retain local-only edits.
- Export ID establishes identity. Changing filename alone never renames a module. An ordinary-component rename needs manifest logical-name change, consistent VB_Name attribute, and an explicit rename preview row. Missing files become deletion proposals unchecked by default. Apply only explicitly selected deletions that pass fresh capability checks.
- Unlisted .bas/.cls files are potential additions unchecked by default. Propose filename stem as name and explicitly reconcile a differing VB_Name. Ignore unrelated files. Reject .frm/.frx and new host/designer components with a clear result.
- Existing ordinary components may edit body/rename only under current capabilities. Reject attribute changes other than consistent approved VB_Name rename in v1; preserve existing attributes. For new classes preserve validated attributes through adapter/save verification. Unsupported preambles/attributes are errors, never silently stripped. Header text cannot create host/designer identity.
- Reference exports may update existing body only if current protection/read-only policy permits. Preserve identity, attributes, and designer payloads. Manifest capability flags never authorize rename/delete.
- Validate the entire selection against revision, names, encoding, protection, metadata, and header rules. Apply as one undoable batch and checkpoint. Open Review Changes afterward. Do not save the add-in automatically.

**Acceptance**

Unsaved class export includes current body/attributes. Export-then-import unchanged is a no-op. Cover additions/renames/edits/deletions, unchecked defaults, stale revision/baseline, folder-versus-draft conflicts, filename-only changes, import headers, class attributes, codepages, reference-body updates, denied host deletion, path traversal, reserved names/case collisions, partial exports, modified destination files, invalid baseline hashes, and size limits. Failure rolls back the whole draft batch and preserves unrelated files.

### IMP-10 — P2: release provenance and qualification

Generate and embed metadata through packaging/VBAAddinEditor.spec: semantic version, full source commit, dirty-tree flag, UTC build time, Python version/architecture, actual exact dependency versions, and source/onedir/onefile mode. Generate at build time; runtime does not need Git or rewrite source metadata. Source runs can say development/unbuilt instead of inventing a commit.

Display metadata in About and diagnostics. Add --version-json with machine-readable metadata, including packaged mode. Generate adjacent release-manifest.json with final executable SHA-256, embedded identity, exact dependency inventory, and validation artifact references. Do not embed an executable's own final hash inside itself. Release mode rejects dirty trees; development builds must be explicitly labeled.

Update scripts/build.ps1 to check failure of every dependency/build/smoke subprocess, resolve fixture paths explicitly, and label variants separately. Release onedir/onefile builds must use the same clean commit. Retain pyopenvba==3.4.0 unless documented compatibility evidence and full gates justify a change. Record installed version, not just the constant.

Required qualification:

- Full automated tests/lint with commands/results. Missing-fixture live skips mean live qualification incomplete, never a qualified release.
- Authentic XLAM/PPAM and PPTM regression fixtures with ordinary classes, relevant host/designer components, benign macros, and fixture hashes. Synthetic packages remain unit-test inputs.
- Packaged GUI: open with host running, edit, close host, retry/save; existing class deletion/save/reopen; close-window Cancel; kill/restart recovery on disposable drafts; copy overwrite/conflict; backup restore; folder roundtrip; cross-project replacement/undo.
- Office opens without repair dialog, VBE inspection and compilation, benign macro execution, and backup restore. Record Office build/bitness, Windows version, executable hash, fixture hash, expected/observed outcomes. Confidential fixtures must not be bundled with releases.
- Restart the exact output executable to validate recovery and embedded identity. Existing --self-test/--self-roundtrip remain required but do not substitute for GUI/Office gates.

Deliver source/tests, README workflows, build metadata/manifests, and completion report mapping all work IDs to evidence. Publishing/uploading releases is outside this plan unless separately requested; produce local reviewable artifacts.

### Required order and completion tracking

| Milestone | Work | Dependencies | Exit evidence |
|---|---|---|---|
| M0 | Original save retry/close/class fixes | Existing baseline | Original automated matrix; live outcomes recorded |
| M1 | Session baseline/revision contracts; IMP-07 envelope foundation | M0 | Stable capture, stale proposal, sanitized results |
| M2 | IMP-01 recovery; IMP-02 safe copy | M1 | Fresh-process recovery/copy failure injection |
| M3 | IMP-04 history; shared source composer | M2 | Cross-target undo; current-class export regression |
| M4 | IMP-05 review; IMP-03 external resolution | M3 | Revision-bound review/rebase and retained recovery |
| M5 | IMP-06 backup browser/restore; IMP-07 UI completion | M4 | Restore/partial failure and diagnostics tests |
| M6 | IMP-08 search/navigation; IMP-09 folder sync | M5 | Batch replacement/import and strict source roundtrips |
| M7 | IMP-10 packaging/qualification; final docs | M0-M6 | Exact-artifact automated and live evidence |

Build-metadata plumbing may start earlier, but final qualification depends on every feature. Add tests with each feature. Avoid unrelated broad refactors.

During implementation create IMPLEMENTATION_PROGRESS.md: one row each for CORE-SAVE, CORE-CLASS, IMP-01 through IMP-10; status (not_started/in_progress/automated_pass/live_pending/complete), commits, evidence paths, and blockers. Live gates cannot be completed with mocks. On context handoff update the ledger and resume at the first incomplete milestone.

### Cross-feature release acceptance

Run an integrated scenario on each authentic host format: open while host runs; edit ordinary class/XML; checkpoint; terminate/restart editor; recover; close host; preview/apply project replacement and undo it; export/edit/import an ordinary class; encounter and resolve a controlled external change; review/save; browse and restore backup; reopen in Office. Keep fixture code compilable and record transitions/hashes. Separate scenarios exercise class deletion and explicit discard so they do not invalidate macro references in the integrated fixture.

No milestone may regress untouched-payload verification, malformed baseline XML tolerance, strict codepage validation, protection/signature rules, same-path replacement, mandatory backups, or clean in-place no-write behavior. Keep add-in registration and Office settings unchanged.

### Explicit decisions replacing earlier open choices

- Save toolbar eligibility is based on a document being present and no operation being active. Live host state is advisory in UI and authoritative in service preflight; never a permanently cached button gate.
- Separate can_delete and can_rename capabilities with reasons. The original class fix qualifies deletion first; before IMP-09, also qualify ordinary-class renaming with PROJECT/dir/PROJECTwm/stream and live tests. Unknown/document/designer components retain destructive restrictions. Existing standard-module renaming remains supported. A new class is removable/renameable in draft, and a saved ordinary class must retain verified capabilities on reopen.
- Review and conflict windows do not write the add-in. Import/replace/rebase apply only to the draft. Only explicit Save/Save Copy/Restore commands commit package bytes.
- The expanded plan requires Save a Copy hardening; the earlier suggestion to leave it separately tracked no longer applies.
- No automatic three-way line merge, partial save staging, cross-crash undo, UserForm layout editing, regex search, semantic reference renaming, XML folder synchronization, or loaded-target hot patching is required.

### Expanded-scope inspection notes

This amendment inspected import_export_service.py, source_codec.py, app.py, code_editor.py, domain/changes.py, packaging/VBAAddinEditor.spec, and version.py. The baseline-source class export defect is a code finding, not yet a fresh runtime reproduction. The additions specify future implementation behavior; none of these features was implemented or tested in this planning amendment. Preserve the investigation limits above when transferring the task.
