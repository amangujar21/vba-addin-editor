# VBA Add-in Editor

Edit VBA source inside installed Excel `.xlam` and PowerPoint `.ppam` add-ins —
directly, in place, without converting through `.xlsm`/`.pptm`, without the VBA
IDE, and without reinstalling the add-in.

VBA Add-in Editor safely patches the VBA project inside the actual installed
add-in file while Office is closed, verifies the edited add-in before replacing
anything, creates a backup, keeps the same path, and lets Office load the
updated code the next time it starts.

## What it does

1. Close Excel/PowerPoint.
2. Open the installed `.xlam` / `.ppam` in VBA Add-in Editor.
3. Edit code; add/rename/delete standard and class modules.
4. **Save Add-in** — the tool verifies the original has not changed, builds a
   verified candidate file, creates a backup, and atomically replaces the
   original using Windows `ReplaceFileW` (preserving ACLs and metadata).
5. Reopen Office — the same installed add-in path now contains the new code.

## Safety model

- **Backups are mandatory.** Every in-place save creates a same-folder backup
  before the original is replaced.
- **Candidate verification.** The edited copy is fully parsed and validated
  (VBA structure, ZIP entry set, non-VBA payload byte-equality, module set and
  bodies) *before* the original is touched. Any failure leaves the original
  untouched.
- **No change means no write.** Saving without edits never rewrites the file.
- **Office running → save blocked.** Excel/PowerPoint must be closed.
- **Password-protected projects are read-only.** No bypass exists.
- **Digitally signed projects** require explicit confirmation; saving removes
  the (now invalid) VBA signature.
- **Strict encoding.** Characters that cannot be stored in the project's code
  page block the save instead of being silently replaced.
- The tool never executes VBA, never changes Trust Center settings, never
  installs/uninstalls add-ins, and makes no network calls.

## Limitations (honest scope)

- The tool validates structure, not VBA compilation. A syntax error you type
  will surface when Office runs the macro. Use the automatic backup to roll back.
- Existing document/designer modules (`ThisWorkbook`, `Sheet1`, UserForm
  code-behind) can be viewed and their body edited, but renaming/deleting them
  is disabled because their subtype cannot be verified safely yet.
- UserForm layout cannot be created or edited; `.frm/.frx` import is not
  supported (`.bas`/`.cls` only).
- The add-in must be closed in Office. The tool does not manage add-in
  registration or file locations — the edited file keeps the same path, so
  existing registration stays valid.

## Development

```powershell
pip install -e .[dev]
scripts/test.ps1          # pytest
scripts/build.ps1         # onedir build + packaged smoke test
scripts/build.ps1 -Onefile   # release exe + SHA-256
```

Headless verification (also runs inside the packaged exe):

```text
VBAAddinEditor.exe --self-test path\to\addin.xlam
VBAAddinEditor.exe --self-roundtrip path\to\addin.ppam
```

### Release gating

Parser tests use a synthetic add-in built from pyOpenVBA's Excel-authored
template. The release promise (Office opens the edited add-in with no repair
dialog and runs the updated macro from the same installed path) requires the
live Office qualification matrix from the implementation plan
(`VBA_Addin_Editor_Implementation_Plan.md`, sections 4 and 29) run against
authentic Office-authored `.xlam`/`.ppam` fixtures.

### Dependency pin

- `pyopenvba==3.4.0` (MIT) — see `THIRD_PARTY_NOTICES.md`.
- Python 3.10+; build with one pinned Python minor version.
