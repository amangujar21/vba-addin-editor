# VBA Add-in Editor

Edit VBA source inside installed Excel `.xlam` and PowerPoint `.ppam` add-ins,
and PowerPoint `.pptm` presentations — directly, in place, without the VBA
IDE, and without reinstalling the add-in. For `.xlam`/`.ppam`/`.pptm`, existing XML
package parts (`.xml`, `.rels`, `[Content_Types].xml`) can be edited too.

VBA Add-in Editor safely patches the VBA project inside the actual installed
add-in file while Office is closed, verifies the edited add-in before replacing
anything, creates a backup, keeps the same path, and lets Office load the
updated code the next time it starts.

## What it does

1. You can open the add-in while Excel or PowerPoint is running. Edits are kept
   in the editor; the corresponding host must still be closed before a successful
   in-place save. Close Excel for `.xlam`, or PowerPoint for `.ppam`/`.pptm`.
   The unrelated host does not block saving. If Save is blocked, keep the window
   open, close the host, and click **Save File** again — do not reload.
2. Edit code; add/rename/delete standard modules and ordinary class modules.
   UserForms can be deleted (code and layout together) but not renamed; host
   document modules and other designer components stay locked for
   rename/delete. The **Forms** tab shows each UserForm's controls and stored
   properties (read-only). On `.xlam`/`.ppam`/`.pptm`, switch to the **XML**
   tab to edit existing package parts.
3. **Save File** reviews the current draft, verifies the original has not
   changed, builds a verified candidate, creates a backup, and atomically
   replaces the original using Windows `ReplaceFileW`. **Save a Copy** writes a
   separate file and does not clear the original dirty state.
4. Closing the window with unsaved changes asks Save / Discard / Cancel.
   Cancel keeps the draft. A blocked save keeps the window open.
5. Reopen Office — the same installed file path now contains the new content.
   Draft recovery is written separately from add-in saves; a crash may lose
   only changes since the last recovery checkpoint. If the original file is
   missing, in-place Save stays disabled and **Save a Copy** writes the
   recovered draft to a new path. If the original changed, compare and resolve
   external changes before saving in place.
6. **Find in Project** (Ctrl+Shift+F) searches the current draft, not disk.
   Replace applies one undoable batch. **Restore Backup** is a verified
   replace with a pre-restore safety copy; it does not run automatically after
   a failed final verification.

## Ribbon buttons ↔ macros

For files with ribbon XML (`customUI/customUI.xml` or `customUI14.xml`):

- **Go to Macro** — in the XML tab, Ctrl+click (or right-click, or F12) any
  `onAction="…"` or other callback attribute to jump to that Sub. If it does
  not exist, you are offered a stub with the right callback signature,
  placed in the module that already holds most of your ribbon callbacks.
- **Add Ribbon Button for This Macro** — right-click inside a Sub in the VBA
  tab. Pick the label, group or menu, icon and size. The id is generated
  and unique, the XML is inserted with matching indentation, and a small
  `…Callback(control As IRibbonControl)` wrapper is added when the Sub has
  no ribbon argument. One Undo reverts both.
- **Show Ribbon Buttons Using This Macro** — F12 or right-click in the VBA tab.
- **Ribbon → Ribbon Buttons…** lists every callback with its module and
  status; **Check Ribbon Callbacks** and **Validate XML** report callbacks
  that point to missing, ambiguous or class-module procedures, and duplicate
  control ids. Save warns (never blocks) only about problems your edits
  introduced.

## Safety model

- **Backups are mandatory.** Every in-place save creates a same-folder backup
  before the original is replaced.
- **Candidate verification.** The edited copy is fully parsed and validated
  (VBA structure, ZIP entry set, non-VBA payload byte-equality, module set and
  bodies) *before* the original is touched. Any failure leaves the original
  untouched.
- **No change means no write.** Saving without edits never rewrites the file.
- **Password-protected projects are read-only.** No bypass exists. (XML-only
  package edits on `.xlam`/`.ppam`/`.pptm` remain possible; they never modify the
  protected VBA project.)
- **Strict encoding.** Characters that cannot be stored in the project's code
  page block the save instead of being silently replaced.
- **XML edits stay inside their package.** Only the exact XML parts you edited
  may differ; every other payload is verified byte-identical. An XML-only save
  never touches `vbaProject.bin`.
- **Digitally signed projects** require explicit confirmation; saving removes
  the (now invalid) VBA signature. Files with an **OPC package digital
  signature** block XML editing entirely.
- The tool never executes VBA, never changes Trust Center settings, never
  installs/uninstalls add-ins, and makes no network calls.

## Limitations (honest scope)

- The tool validates structure, not VBA compilation. A syntax error you type
  will surface when Office runs the macro. Use the automatic backup to roll back.
- **XML editing** covers existing parts only: no add/delete/rename of package
  parts, no schema validation against the full OOXML schemas (syntax and
  package integrity are validated; PowerPoint remains the final format
  validator), no re-signing, and no editing of binary parts such as
  `vbaProject.bin`, images, or ActiveX `.bin`.
- Existing document/designer modules (`ThisWorkbook`, `Sheet1`, UserForm
  code-behind) can be viewed and their body edited, but renaming/deleting them
  is disabled. Ordinary class modules are classified from the VBA PROJECT
  stream plus dir records and can be deleted or renamed when that metadata
  agrees. Unknown or conflicting components stay locked.
- A UserForm can be deleted when the PROJECT stream declares it (`BaseClass`),
  the dir record agrees, and its designer storage holds a form. Deleting
  removes the code, the layout storage and every declaration; verification
  confirms nothing is left behind and that every other form's layout is
  byte-identical. Renaming a UserForm stays disabled (the layout storage would
  be orphaned). A deleted form's name cannot be reused until the deletion is
  saved. Folder import never deletes a UserForm.
- UserForm layout is read-only: it can be viewed in the **Forms** tab but not
  created or edited; `.frm/.frx` import/export is not supported
  (`.bas`/`.cls` only).
- A VBA signature is detected both inside `vbaProject.bin` and in the package
  parts beside it (where Office stores it). After you confirm, a save that
  changes VBA removes the stale signature parts, their relationships and
  content-type overrides as Office does; verification allows exactly those
  package edits.
- The corresponding Office host must be closed to save in place. The tool does
  not hot-patch a loaded add-in, manage add-in registration, or change file
  locations — the edited file keeps the same path, so existing registration
  stays valid.

## Development

```powershell
pip install -e .[dev]
scripts/test.ps1          # pytest
scripts/build.ps1         # development onedir package (not a qualified release)
scripts/build.ps1 -Onefile   # development onefile package
scripts/build.ps1 -Release   # strict gate: tests, lint, fixtures, onedir+onefile
```

Headless verification (also runs inside the packaged exe):

```text
VBAAddinEditor.exe --self-test path\to\addin.xlam
VBAAddinEditor.exe --self-roundtrip path\to\addin.ppam
VBAAddinEditor.exe --self-roundtrip path\to\RealPresentation.pptm
VBAAddinEditor.exe --version-json
```

### Release gating

Parser tests use a synthetic add-in built from pyOpenVBA's Excel-authored
template. The release promise (Office opens the edited add-in with no repair
dialog and runs the updated macro from the same installed path) requires the
live Office qualification run (`pytest tests -m live`) against authentic
Office-authored fixtures placed in `tests/fixtures/` (`xlam/RealAddin.xlam`,
`ppam/RealAddin.ppam`, `pptm/RealPresentation.pptm`; the PPAM/PPTM need a
`customUI` part containing `VBAAE_XML_ORIGINAL`). The automated live cycle
edits that XML marker and VBA code in one save; a human then verifies Excel or
PowerPoint opens the saved file with **no repair dialog** and runs the macro.

### Dependency pin

- `pyopenvba==6.2.0` (MIT) — see `THIRD_PARTY_NOTICES.md`.
- Python 3.10+; build with one pinned Python minor version.
