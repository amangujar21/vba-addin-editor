# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

Windows-only, offline Tkinter/ttk app that edits VBA source (and existing OOXML parts) inside Excel `.xlam`, PowerPoint `.ppam` and `.pptm` files **in place**: verified candidate → mandatory same-folder backup → atomic `ReplaceFileW`. No COM, no Office automation. VBA parsing/writing goes through `pyopenvba==3.4.0` (exact pin; the adapter imports its internals).

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

- **domain/** — `DocumentSnapshot` (immutable open-time state, the *trusted* source) vs `DocumentDraft` (editable `ModuleDraft`s + `XmlPartDraft`s). `capabilities.py` classifies each component from the VBA `PROJECT` stream + dir records (standard / class / document / designer); only snapshot capabilities authorize delete/rename — draft flags never do.
- **adapters/** — the only places bytes are touched. `pyopenvba_adapter.py` is the single VBA seam (build/verify candidate, delete/rename replay, PROJECT scrub shim for pyOpenVBA 3.4.0). `ooxml_package_adapter.py` is the only code that reads/rewrites ZIP entries; the GUI never touches package members. `xml_codec.py` preserves each part's codec/BOM/newline; editor text is LF-normalized.
- **services/** — `save_service.py` runs the composite transaction: VBA candidate first, then XML patch applied to the disposable candidate; `verify_candidate(..., allowed_non_vba_changes=...)` permits differences only on the exact changed XML paths; post-commit verification compares against the backup. `session_service`/`recovery_service`/`history_service` implement draft sessions, recovery checkpoints and undo (`HistoryCommand` ops: `text`, `xml_text`, `rename`, `add`, `delete`, `replace_all` with `{"modules","xml"}` snapshots for multi-target undo). `ribbon_service.py` is pure-text customUI ↔ VBA callback linking (lexical XML scan so it works on mid-edit XML).
- **platform/** — `ReplaceFileW`, exclusive-access probes, Office process detection.
- **ui/** — `main_window.py` owns everything; VBA and XML tabs each use a `CodeEditor` (`xml_editor.py` subclasses it). `text_context_menu.py` provides right-click menus with `add_extra()` hooks.

## Invariants that are easy to break

- No change → no write. Office host running / file locked / fingerprint changed → save blocked, draft kept.
- Untouched package members (including malformed baseline XML) must stay byte-identical; only XML in the change set is parsed/validated. Don't reintroduce parse-all verification.
- Password-protected VBA blocks only VBA mutations; XML-only saves must keep `vbaProject.bin` byte-identical. OPC signatures block XML edits.
- Document modules (`ThisWorkbook`, sheets) and designers (UserForms) are never deletable/renamable.
- In `MainWindow`, editor widgets hold unflushed text. Call `_flush_all_editors()` before reading/mutating the draft. `_checkpoint_now()` flushes editors, so after mutating the draft programmatically reload editors **before** checkpointing (see `_batch_edit`). `on_module_selected`/`on_xml_part_selected` ignore re-selection of the already-shown item; navigate with `_open_module_at` / `_open_xml_at`.

## Tests

- `tests/conftest.py` builds synthetic packages (`build_xlam`, `build_xlam_with_class`, `build_pptm`, `build_xlam_with_ribbon`); these are not a substitute for the live-Office gate.
- GUI tests create a real withdrawn `tk.Tk()`, set `VBAAE_SESSION_ROOT` to a temp dir, and monkeypatch `messagebox` / `ribbon_dialogs.choose` instead of opening modal dialogs. Never call `tk_popup()` in tests (use `TextContextMenu.prepare_menu()`).
- `pyopenvba` upgrades: re-read its `_host.py` save path and `vba.py` mutation methods first; `pyopenvba_adapter.py` is the only upgrade seam.

## Project state

`README.md` holds user-facing behavior, the safety model and the release gate. Code comments cite "plan N.N" sections from implementation plans that were removed from the tree; they are in git history before the commit that added this file. Outstanding: the live Office qualification (`pytest -m live` with authentic fixtures plus a human no-repair-dialog check) has never been run, so no build is a qualified `-Release`.
