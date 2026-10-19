# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

Windows-only, offline Tkinter/ttk app that edits VBA source (and existing OOXML parts) inside Excel `.xlam`, PowerPoint `.ppam` and `.pptm` files **in place**: verified candidate → atomic commit, keeping the old file as a same-folder backup unless the user turned automatic backup off. No COM, no Office automation. VBA parsing/writing goes through `pyopenvba==6.2.0` (exact pin; the adapter imports its internals).

## Commands

```powershell
pip install -e .[dev]                      # re-run after adding new modules under src/
python -m pytest tests -q                  # full suite (5 "live" tests skip without real fixtures)
python -m pytest tests/unit/test_ribbon_service.py -q
python -m pytest tests -k "class_delete" -q
python -m pytest tests -m live -v          # live Office gate; needs tests/fixtures/* (gitignored)
python -m ruff check .
python -m pyright src                      # baseline has ~26 pre-existing errors; don't add new ones
python -m vba_addin_editor                 # run the GUI
python -m vba_addin_editor --self-test path\to\file.xlam
python -m vba_addin_editor --self-roundtrip path\to\file.ppam
scripts/build.ps1 -Onefile                 # dist\VBAAddinEditor.exe (dev build)
scripts/build.ps1 -Release                 # strict gate: clean tree, tests, lint, pyright, real fixtures
```

- Build smoke tests need `tests/fixtures/xlam/Demo.xlam`; `-Release` also needs `RealAddin.xlam`, `RealAddin.ppam`, `RealPresentation.pptm`. `tests/fixtures/` is gitignored. A synthetic Demo can be made with `conftest.build_xlam`. Never copy the user's own add-ins into the repo.
- Releases so far: onefile exe + `.sha256` zipped as `VBAAddinEditor-vX.Y.Z-Windows.zip` on GitHub. Version lives in three places: `pyproject.toml`, `src/vba_addin_editor/version.py`, `packaging/VBAAddinEditor.spec`.
- Some shells can't launch `dist\...exe` directly; use PowerShell `& ".\dist\VBAAddinEditor.exe"` or `cmd /c`.

## Architecture

Layers under `src/vba_addin_editor/`:

- **domain/** — `DocumentSnapshot` (immutable open-time state, the *trusted* source) vs `DocumentDraft` (editable `ModuleDraft`s + `XmlPartDraft`s). `capabilities.py` classifies each component from the VBA `PROJECT` stream + dir records (standard / class / document / designer / userform); only snapshot capabilities authorize delete/rename — draft flags never do.
- **adapters/** — the only places bytes are touched. Both adapters take a package as a path *or* as bytes (`PackageSource`); open and save pass bytes. `pyopenvba_adapter.py` is the single VBA seam (in-memory hosts via `_open_host`, a decoded-source cache keyed by compressed stream, build/verify candidate, delete/rename replay, PROJECT scrub shim written for pyOpenVBA 3.4.0; read-only UserForm layouts via `pyopenvba.forms`; signature-part removal edits via `pyopenvba._package_signature`). `ooxml_package_adapter.py` is the only code that reads/rewrites ZIP entries; the GUI never touches package members. `xml_codec.py` preserves each part's codec/BOM/newline; editor text is LF-normalized.
- **services/** — `save_service.py` runs the composite transaction in memory: VBA replay and changed XML parts go into one candidate built from `baseline.package_bytes`; `verify_candidate(..., allowed_non_vba_changes=...)` permits differences only on the exact changed XML paths; the candidate is written once and committed; post-commit verification proves by file identity (volume, file id, size, mtime) that the path holds the verified candidate and the backup holds the old file, reading a file only when identity cannot prove it. `session_service`/`recovery_service`/`history_service` implement draft sessions, recovery checkpoints and undo (`HistoryCommand` ops: `text`, `xml_text`, `rename`, `add`, `delete`, `replace_all` with `{"modules","xml"}` snapshots for multi-target undo). `ribbon_service.py` is pure-text customUI ↔ VBA callback linking (lexical XML scan so it works on mid-edit XML).
- **platform/** — `commit_replace` (rename when it preserves the file's security/attributes/streams, else `ReplaceFileW`), exclusive-access probes, Office process detection, `settings.py` (user preferences).
- **ui/** — `main_window.py` owns everything; VBA and XML tabs each use a `CodeEditor` (`xml_editor.py` subclasses it). `text_context_menu.py` provides right-click menus with `add_extra()` hooks. Find/replace is one inline bar per `CodeEditor` (`show_find`); the text widget's Ctrl+F/Ctrl+H bindings return `"break"` so the window-level bindings (for focus outside an editor) never fire a second time. `ribbon_dialogs.py` holds the Add Button / Add Menu forms; `MainWindow.apply_ribbon_button` / `apply_ribbon_menu` do the edit and are what tests call.
  - `theme.py` holds the only colours (`DARK` / `LIGHT` palettes); `apply_theme` (called in `MainWindow.__init__` before any widget exists) sets ttk `clam` styles, the Tk option database for classic widgets, and named fonts (`VbaaeCode` is shared by every editor, so zoom is one font change). Use a named style or `theme.palette()`, never a hex colour in a widget call. Dark title bars come from `DwmSetWindowAttribute`; `run_gui` keeps the root withdrawn until built because Windows 10 only paints it before first show.
  - There is no Win32 menu bar (Tk cannot colour it): the header holds `ttk.Menubutton`s (`_header_menu`), Alt+letter/F10 post them, and `_set_menubar_state` disables them while a worker runs.
  - `highlight.py` colours only the visible lines ± `MARGIN_LINES` on idle (whole-module colouring is ~250 ms on 10k lines). `CodeEditor._schedule_view_update` batches gutter + colouring; scrolling of any kind reaches it through `yscrollcommand`.
  - Pack bars (status bar, banner, wait bar, dialog button rows) *before* the expanding widget in packing order, or a small window pushes them off-screen.

## Invariants that are easy to break

- Never read back a file the editor wrote (candidate, committed add-in, backup, captured session copy). Defender scans a freshly written macro package on its first open **for read**, seconds for a large project; stat, rename and write-only opens do not trigger it. Verify writes through the writing handle (`write_verified`, `write_bytes_atomic`), and treat `FileFingerprint.written_identity` as proof of content while the path still has that identity. The one deliberate exception is `platform/prescan.py`: after a successful save `MainWindow._start_prescan` opens the saved file on a daemon thread so the scan is spent before Office loads it (status bar: "Antivirus scanning …" → "… ready for Office"). Its handle fails the exclusive-access probe while the scan runs, so `_run_in_background` and `restore_backup` call `Prescan.settle()` first; the pending open cannot be cancelled. It holds the GC pause (`_hold_gc`, nestable) until the thread has ended.
- The fast rename commit is only allowed when `rename_commit_preserves_metadata` holds; otherwise `ReplaceFileW` must do it. Lock probes ask for write+delete access, never read.
- Automatic backup is a user setting (`Settings.auto_backup` → `SaveService.auto_backup`, default on). UI text that promises a backup must check it.
- Open and save both run through `MainWindow._run_in_background` with `_saving` set; it disables cyclic GC while the worker runs so Tk variables are never collected off the Tk thread. Don't touch Tk from the worker.
- No change → no write. Office host running / file locked / fingerprint changed → save blocked, draft kept.
- Untouched package members (including malformed baseline XML) must stay byte-identical; only XML in the change set is parsed/validated. Don't reintroduce parse-all verification.
- Password-protected VBA blocks only VBA mutations; XML-only saves must keep `vbaProject.bin` byte-identical. OPC signatures block XML edits.
- Document modules (`ThisWorkbook`, sheets) are never deletable/renamable. Verified UserForms (`project_item_kind == "userform"`) are delete-only; other designers stay locked.
- The editor never writes UserForm layouts. Never call `host.forms()` on a host that will `save()` — pyOpenVBA writes back any form it has read. Verification requires every surviving form storage to be byte-identical.
- In `MainWindow`, editor widgets hold unflushed text. Call `_flush_all_editors()` before reading/mutating the draft. `_checkpoint_now()` flushes editors, so after mutating the draft programmatically reload editors **before** checkpointing (see `_batch_edit`). `on_module_selected`/`on_xml_part_selected` ignore re-selection of the already-shown item; navigate with `_open_module_at` / `_open_xml_at`.

## Tests

- `tests/conftest.py` builds synthetic packages (`build_xlam`, `build_xlam_with_class`, `build_pptm`, `build_xlam_with_ribbon`); these are not a substitute for the live-Office gate.
- GUI tests create a real withdrawn `tk.Tk()`, set `VBAAE_SESSION_ROOT` to a temp dir (settings live there too, and the on-disk source cache stays off), and monkeypatch `messagebox` / `ribbon_dialogs.choose` instead of opening modal dialogs. Never call `tk_popup()` in tests (use `TextContextMenu.prepare_menu()`).
- `pyopenvba` upgrades: re-read its `_host.py` save path and `vba.py` mutation methods first; `pyopenvba_adapter.py` is the only upgrade seam.

## Project state

`README.md` holds user-facing behavior, the safety model and the release gate. Code comments cite "plan N.N" sections from implementation plans that were removed from the tree; they are in git history before the commit that added this file. Outstanding: the live Office qualification (`pytest -m live` with authentic fixtures plus a human no-repair-dialog check) has never been run, so no build is a qualified `-Release`. Known bug: a module body that starts with a blank line fails candidate verification ("Body mismatch"), because `split_attribute_header` drops the blank line on re-read; the save is refused and the original is untouched.
