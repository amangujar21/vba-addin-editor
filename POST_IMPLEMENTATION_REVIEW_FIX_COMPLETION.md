# Completion note: post-implementation review repairs

Date: 2026-09-11
Repository: https://github.com/amangujar21/vba-addin-editor
Review fixed point: `206aaab`
Implementation reviewed: `4b37e68`

## Automated evidence

Recorded after the repairs in this change. Live Office and packaged GUI
qualification were not executed and are **not** passed.

Commands:

```powershell
python -m pytest tests -q
python -m ruff check .
python -m pyright
```

Recorded on this change:

- `python -m pytest tests -q` → 212 passed, 5 skipped (live Office fixtures absent)
- `python -m ruff check .` → clean
- `python -m pyright` → installed from `.[dev]`; reports remaining type errors,
  mostly in pre-existing adapter/test files. Type checking is **not** treated
  as passed.

Prior handoff counts are historical. Pinned dependency remains `pyopenvba==3.4.0`.

## Repairs implemented (source + automated tests)

1. **Session baseline publication** after verified in-place save: captured
   package is replaced atomically, session metadata/revision advance, proposals
   go stale, Save a Copy uses the new captured generation, and
   `complete.marker` is cleared before the next dirty checkpoint.
2. **Recovery opening** reconstructs from the captured baseline. Matching /
   changed / missing originals are distinct; changed sources cannot in-place
   save until conflicts are resolved; missing sources use recovered Save a
   Copy. Corrupt newest generation falls back; invalid data is retained.
3. **Conflict engine** covers the three-way union, structural cases, and
   atomic apply with external-hash / revision checks. Comparison UI can Apply
   or Cancel without mutating on Cancel.
4. **Backup restore** is a verified transaction with race rechecks, safety
   backup, and `recovery_required` artifacts. The UI no longer auto-restores
   after a post-commit verification failure.
5. **Folder export/import** checks previous managed hashes, stages then
   publishes the manifest last, validates untrusted manifests into structured
   problems, and applies to a cloned draft as one undoable batch.
6. **Adapter seam**: ZIP member rewriting lives in `OoxmlPackageAdapter`.
   Candidate verification fails closed on unreadable VBA stream / PROJECT /
   PROJECTwm metadata. Class deletion checks stream, dir, PROJECT, and
   PROJECTwm.
7. **Project search/replace UI**: non-modal find surface with options, paging,
   outline, stale results, preview, and atomic undoable replacement.
8. **`scripts/build.ps1 -Release`** is a strict gate (clean tree, pytest, ruff,
   pyright, required fixtures, onedir+onefile, packaged smoke). Missing
   evidence fails the release; development builds are labelled and do not
   write a qualified release manifest.

## Not completed here (live gates)

- Authentic Office-authored XLAM/PPAM/PPTM fixtures and hashes
- Packaged onefile/onedir rebuild against this commit
- Human Excel/PowerPoint repair-dialog, VBE compile, and macro checks
- Packaged GUI recovery/restore/folder/conflict/search workflows

Those remain `live_pending`. No user add-in was modified. `HANDOFF.md` local
edits and `latest_draft_vba.md` were left in place.
