# HANDOFF — VBA Add-in Editor project state

_Date: 2026-08-30. Source: working session that implemented
`..\VBA_Addin_Editor_Implementation_Plan.md` (3977 lines) end-to-end._

## What exists

Single app in this directory (`vba-addin-editor/`), Windows-only, offline,
Tkinter/ttk GUI. Edits VBA source inside installed Excel `.xlam` /
PowerPoint `.ppam` add-ins **in place** (same path) with mandatory backup,
candidate verification, and `ReplaceFileW` commit. No COM, no Python
install needed at runtime.

| Area | Location |
|---|---|
| Entry / headless self-tests | `src/vba_addin_editor/__main__.py` (`--self-test`, `--self-roundtrip`), `app.py` |
| Domain (snapshot/draft/changes) | `src/vba_addin_editor/domain/` |
| pyOpenVBA adapter (the only upgrade seam) | `src/vba_addin_editor/adapters/pyopenvba_adapter.py` |
| Codepage + header/body codec | `src/vba_addin_editor/adapters/source_codec.py` |
| Save pipeline (plan stages A–N) | `src/vba_addin_editor/services/save_service.py` |
| Import/export (.bas/.cls) + backup restore | `src/vba_addin_editor/services/import_export_service.py` |
| ReplaceFileW / lock probe / process detect | `src/vba_addin_editor/platform/` |
| GUI | `src/vba_addin_editor/ui/main_window.py`, `ui/code_editor.py` |
| Tests (33, all green) | `tests/` (unit + integration; GUI drives real Text widget) |
| Packaging | `packaging/VBAAddinEditor.spec`, `scripts/build.ps1`, `scripts/test.ps1` |

Dependency pin: `pyopenvba==3.4.0` exact. Python 3.10.10 used for dev/build.

## Verified state

- `python -m pytest tests -q` → 33 passed. Includes: full save→backup→
  in-place-replace round trip on an Excel-authored template; rename chains;
  combined add+rename+delete; blocked saves (Office running / locked /
  external change / commit failure) leave the original byte-identical;
  no-change save writes nothing; scripted GUI end-to-end save.
- Onedir build (`dist\VBAAddinEditor\VBAAddinEditor.exe`) packaged and
  smoke-tested: `--self-test` and `--self-roundtrip` exit 0.
- Not yet run: `scripts\build.ps1 -Onefile` (single-file release exe + SHA-256).
- Repo is **not** under git yet — no commits made.

## Key decisions (why the code looks like this)

- **Post-commit verification runs on the original path**, not the candidate:
  `ReplaceFileW` moves the candidate onto the original, so Stage N calls
  `adapter.verify_candidate(original, original, draft)`.
- **CRLF canonical comparison in the verifier**: editor text is LF; compare
  `to_vba_crlf(split_attribute_header(...))` against CRLF draft bodies, or you
  get false "Body mismatch".
- **Draft name validation gates save** (Stage E0) — pyopenvba would happily
  accept `Bad Name!`; only new/renamed modules are name-checked so
  pre-existing names like `ThisWorkbook` never block saves.
- **Rename replay uses two-stage unique temp names** (`VBAAE_TMP_XXXXXXXX`);
  relies on pyopenvba pending_renames chain-collapsing (proven by
  `test_rename_chain_a_b_c`).
- **Import bodies are LF in drafts**; CRLF applied only at the adapter
  boundary (`to_vba_crlf` in build_candidate / replay).
- Never call `host.save()` on the original path; never
  `allow_protected=True`; no signature removal without explicit confirmation;
  password-protected projects are read-only.
- Ambiguous `other` modules (ThisWorkbook/Sheet1/UserForm code-behind):
  body editing allowed, rename/delete disabled in v1.

## Safety model (enforced, tested)

No-change → no write. Backup mandatory on every in-place save. Office
running → save blocked. External change (fingerprint) → save blocked.
Candidate fully verified before anything is replaced. Recovery path
restores the backup if final verification fails.

## Remaining work

1. **Run the onefile release build** (the plan's release artifact):
   ```powershell
   powershell -File scripts\build.ps1 -Onefile
   ```
   Confirms `dist\VBAAddinEditor.exe` (single file) + `dist\VBAAddinEditor.exe.sha256`.
2. **Live Office release qualification** (plan §4, §29) — **requires a human
   with real Office**; the agent cannot do this:
   - Create authentic `.xlam` and `.ppam` via Office "Save As" (not synthetic
     templates) and drop them in `tests/fixtures/xlam/` and `tests/fixtures/ppam/`.
   - Run the edit / add / rename / delete cycle against them.
   - Reopen from the installed path in Excel/PowerPoint; verify **no repair
     dialog** and that the updated macro runs.
   - Everything is staged for this: add adapter round-trip tests marked
     `@pytest.mark.live` that skip when the fixtures are absent.
3. **Initialize git** in this directory and commit (nothing is version
   controlled yet). Suggested sequence mirrors plan §66.
4. Optional: fix the 7 remaining ruff findings (intentional local-time backup
   stamp DTZ005, C408, PIE810, RUF059 unused unpacks) or add `# noqa` where
   intentional.
5. After live PPAM/XLAM proof passes: update README "Release gating" section;
   consider upstreaming the PPAM patch to pyOpenVBA (plan §49).

## Maintenance notes

- If pyopenvba ever bumps off 3.4.0: re-read installed `_host.py` save path
  and `vba.py` mutation methods first (plan §75 rules 2–3); the internals
  import in `pyopenvba_adapter.py` (`_encoding_for_codepage`, `detect_signature`,
  `split_attribute_header`, `CFB`) is the only upgrade seam.
- After adding new modules to `src/`, re-run `pip install -e .` — a fresh
  module was missed once until reinstall.
- Windows quirk: some shells can't execute `dist\...exe` directly; run via
  `python -c "subprocess.run([...])"` or `cmd /c`.
- Synthetic test fixtures (`tests/conftest.py` → `build_xlam` via
  `ExcelFile.create_new`) are acceptable for parser/save-pipeline tests but
  are NOT a substitute for the live-Office gate.
