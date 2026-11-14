# VBA Add-in Editor

Edit VBA source inside installed Excel `.xlam` and PowerPoint `.ppam` add-ins,
and PowerPoint `.pptm` presentations — directly, in place, without the VBA
IDE, and without reinstalling the add-in. For `.xlam`/`.ppam`/`.pptm`, existing XML
package parts (`.xml`, `.rels`, `[Content_Types].xml`) can be edited too.

VBA Add-in Editor safely patches the VBA project inside the actual installed
add-in file while Office is closed, verifies the edited add-in before replacing
anything, keeps a backup (unless you turn that off), keeps the same path, and
lets Office load the updated code the next time it starts.

## What it does

1. You can open the add-in while Excel or PowerPoint is running. Edits are kept
   in the editor; the corresponding host must still be closed before a successful
   in-place save. Close Excel for `.xlam`, or PowerPoint for `.ppam`/`.pptm`.
   The unrelated host does not block saving. If you click **Save** while
   the host is running (or still shutting down, or another program has the file
   open), the editor first gives a briefly held file (antivirus, cloud sync,
   indexing) a few seconds, and names the program holding it when Windows can
   tell. Otherwise it offers to wait: a bar shows it is waiting, and it saves as
   soon as the file is free (**Cancel Auto-Save** stops waiting). Edits made
   while it waits are reviewed again before saving. Keep the window open — do
   not reload.
   If Office saved the add-in itself in the meantime (an add-in that calls
   `ThisWorkbook.Save`, Ctrl+S in the VBE), Save rebases your edits onto that
   version: changes that do not overlap merge automatically and the save goes
   ahead; overlapping ones open **Compare External Changes** so you choose,
   then Save again. Nothing is overwritten until the merge is resolved.
2. Edit code; add/rename/delete standard modules and ordinary class modules.
   UserForms can be deleted (code and layout together) but not renamed; host
   document modules and other designer components stay locked for
   rename/delete. The bar on the left switches between four views (Ctrl+1–4):
   **Code** (modules and the procedures of the open module), **Ribbon** (the
   ribbon designer, below), **UserForms** (each UserForm's controls and stored
   properties, read-only) and **Package XML** (the existing package parts of a
   `.xlam`/`.ppam`/`.pptm`, edited as text).
3. **Save** (Ctrl+S) reviews the current draft, verifies the original has not
   changed, builds a verified candidate, and atomically replaces the original,
   keeping the old file as a same-folder backup. **File → Back Up
   Automatically on Save** turns that backup off (the choice is remembered);
   **File → Back Up Now** makes one whenever you want. **Save a Copy** writes a
   separate file and does not clear the original dirty state.
4. Closing the window with unsaved changes asks Save / Discard / Cancel.
   Cancel keeps the draft. A blocked save keeps the window open.
5. Reopen Office — the same installed file path now contains the new content.
   Draft recovery is written separately from add-in saves; a crash may lose
   only changes since the last recovery checkpoint. If the original file is
   missing, in-place Save stays disabled and **Save a Copy** writes the
   recovered draft to a new path. If the original changed, compare and resolve
   external changes before saving in place.
6. **Find** (Ctrl+F) and **Replace** (Ctrl+H) open a bar under the editor you
   are in: Enter / F3 for the next match, Shift+Enter / Shift+F3 for the
   previous, Esc to close. **Find in Project** (Ctrl+Shift+F) searches the
   current draft, not disk.
   Replace applies one undoable batch. **Restore Backup** is a verified
   replace with a pre-restore safety copy; it does not run automatically after
   a failed final verification.
7. The window is dark by default; **View → Light Theme** switches it, and
   **View → Zoom In / Zoom Out** (Ctrl+= / Ctrl+-, Ctrl+0 to reset) sizes the
   code font. Both are remembered. With no file open the window offers
   **Open file…** and your recently opened files (also under **File → Open
   Recent**). In the module list, ● marks an edited module, + a new one, → a
   renamed one, and 🔒 one that cannot be renamed or deleted (its code can
   still be edited; selecting it says why). Under it, **PROCEDURES** lists the
   open module's Subs, Functions (`fn`) and properties (`get`/`let`/`set`),
   follows the cursor, and marks the macros a ribbon control runs; **+** adds a
   new Sub, right-click adds or shows its ribbon button. VBA and XML are
   syntax-coloured.
8. **File → New Add-in…** (Ctrl+N) makes a new, empty `.xlam`, `.ppam` or
   `.pptm` (pick the type in the file dialog) and opens it. It starts from the
   blank file Excel or PowerPoint itself writes (shipped inside pyOpenVBA):
   `ThisWorkbook`, `Sheet1` and an empty `Module1` for Excel, an empty
   `Module1` for PowerPoint. A `.ppam` is that PowerPoint presentation with its
   main part typed as an add-in. The form that follows offers a ribbon tab with
   one group for your buttons (**No Ribbon** skips it); the ribbon is an
   unsaved edit until **Save File**. An existing file is never replaced.

## Ribbon buttons ↔ macros

**Ribbon → Create Ribbon…** gives a file that has no ribbon XML a first one:
a custom tab with one empty group, in a new `customUI/customUI14.xml` linked
from `_rels/.rels` (one Undo removes it again; nothing is written until Save
File). **Add Ribbon Button**, **Add Ribbon Menu** and **Import Button** offer
the same when the file has no ribbon yet. Files with an OPC package signature,
or whose `_rels/.rels` already links a ribbon part that is missing, are
refused with the reason.

For files with ribbon XML (`customUI/customUI.xml` or `customUI14.xml`):

- **Ribbon view** (Ctrl+2, or **Ribbon → Design Ribbon**) edits the ribbon as
  a form instead of XML. On the left is the ribbon as a tree (tab › group ›
  menu › button) with a filter and a health line ("2 callbacks need a macro";
  click it to show only the problems). On the right: a drawing of the tab, a
  form for the selected control (Label, Macro, Icon, Size, Screentip,
  Supertip, Id and its other callbacks), and the code of the macro it runs
  (F12 or double-click opens it in the Code view).
  - **+** or **Ins** adds a button at the selection and puts the cursor in
    Label; its id follows the label until you give it one. The + menu also
    adds a menu, group or separator.
  - Typing a macro name that exists links it; a plain `Sub X()` gets its
    `XCallback(control As IRibbonControl)` wrapper. A name that does not
    exist shows **Create Sub**.
  - In the tree: **Alt+Up/Down** moves, **Del** deletes (asks first when a
    group or menu still holds controls), **F2** renames, drag and drop moves
    before, after or into another control (only where the ribbon schema
    allows it). Select several controls to delete them or group them into a
    menu.
  - Every change is an ordinary text edit of the ribbon XML: one Undo step
    each, untouched XML stays byte-identical, and **Edit XML** shows the
    control in the Package XML view. A ribbon part shown there has a
    **Design view** button that goes back to the same control.
  - Results of ribbon commands (add button or menu, icons, import, F12 from
    the code) appear in whichever ribbon view, designer or XML, you used last.
  - The Ribbon menu commands below that work on "the control under the
    cursor" (move, group into menu, icons, Go to Macro) use the designer's
    selection while the Ribbon view is showing.
  - Built-in (`imageMso`) icons cannot be drawn offline; the drawing shows a
    placeholder for them and the real picture for custom icons.
- **Go to Macro** — in the Package XML view, Ctrl+click (or right-click, or F12) any
  `onAction="…"` or other callback attribute to jump to that Sub. If it does
  not exist, you are offered a stub with the right callback signature,
  placed in the module that already holds most of your ribbon callbacks.
- **Add Ribbon Button for This Macro** — right-click inside a Sub in the VBA
  code tab. Pick the label, group or menu, icon and size. The id is generated
  and unique, the XML is inserted with matching indentation, and a small
  `…Callback(control As IRibbonControl)` wrapper is added when the Sub has
  no ribbon argument. One Undo reverts both.
- **Add Ribbon Menu** — Ribbon menu, or right-click in the ribbon XML. Pick
  the label, the group or menu to put it in (the one under the cursor is
  preselected), icon and sizes, and optionally tick the macros to add as its
  buttons. The `<menu>` block is inserted with matching indentation and
  unique ids; more buttons can be added to it later with Add Ribbon Button.
- **Move Up / Move Down** — put the cursor on a button, menu, group or tab in
  the ribbon XML and press **Alt+Up** / **Alt+Down**, click ▲ / ▼ above the
  part list, or right-click. The control swaps places with its neighbour;
  comments between them stay put. Elements whose order Office fixes (the
  dialog box launcher, the parts of a split button, items before buttons in a
  gallery) are not moved.
- **Group Selected into Menu** — select one or more buttons in the ribbon XML,
  right-click and pick a label and icon. They are wrapped in a new `<menu>`
  with matching indentation; their `size` attribute is removed, since buttons
  in a menu have none.
- **Custom icons** — in the browser, right-click a picture and choose **Copy
  image**; then put the cursor on a button (or menu, group…) in the ribbon XML,
  right-click and choose **Set Icon from Clipboard** (or **Set Icon from
  File…**). The **Add Ribbon Button** and **Group into Menu** forms have the
  same **Paste Image** / **Browse…** buttons with a preview. The picture is
  trimmed, centred on a transparent square and resized to a 32×32 PNG (a few
  KB, so the add-in loads as fast as with a built-in icon), stored as
  `customUI/images/<id>.png`, linked from the ribbon's relationships part
  (created if missing) and given the `image="rId…"` attribute in place of
  `imageMso`. `[Content_Types].xml` gets a `png` entry when it lacks one. One
  Undo reverts all of it. Replacing a saved icon removes the old image unless
  another control still shows it; **Ribbon → Remove Unused Icons** removes
  images no control shows. Icons saved by 0.6.0 show blank in Office (wrong
  relationship type); opening such a file offers to repair them, as does
  **Ribbon → Repair Custom Icons**. Mind the picture's licence when you share
  the add-in.
- **Import Button from Another Add-in** — **Ribbon → Import Button from
  Another Add-in…**, then pick another `.xlam`/`.ppam`/`.pptm`. It is only
  read (it may stay loaded in Office). Choose one of its buttons (or toggles,
  menus, whole groups…), and choose where it goes in this file: a group,
  menu or tab of yours, or a new group on one of your tabs. The preview lists
  everything that will be added and renamed. Imported along with the control:
  - every procedure its callbacks call, followed through the whole call
    chain. That includes constants, module variables, `Declare` statements
    (with their `#If VBA7` blocks), Types and Enums, and macros run by name
    (`Application.Run`, `OnTime`, `OnKey`, `OnAction`). Each source module
    becomes one new module here, keeping its `Option` lines, and Private
    names keep their meaning. Class modules it uses are copied whole, with
    their attributes;
  - its icon, if it has a custom one.

  So that it works with both add-ins loaded at once, the button's macros (and
  macros run by name) always get a new name: the **name suffix**, `_<this
  file's name>` by default. Any other module, procedure, variable or control
  id that this file already uses is renamed with the same suffix. Every
  reference follows the renames: calls, `Module.Proc`, control ids in strings
  (`InvalidateControl "btnX"`), and `"'Other.xlam'!Proc"` strings, which then
  name this file. Code this file already has, character for character, is
  reused rather than copied again. If the code keeps the ribbon in an
  `IRibbonUI` variable, a small `RibbonLoaded…` Sub that sets it is added and
  called from this ribbon's `onLoad` (or becomes the `onLoad`). Attributes
  that are not valid in the new place are dropped: `size` outside a group,
  `description` outside a menu, and `insertAfter…`/`insertBefore…`. One Undo
  reverts it all; nothing is written until Save File.

  The import is refused if the code uses a UserForm this file does not have
  (import the form in the VBA editor first), or if either project is
  password-protected. The summary lists what to check afterwards:
  references this file lacks (Tools → References), sheets of the other
  add-in the code reads, and a cross-host import (Excel ↔ PowerPoint).
- **Show Ribbon Buttons Using This Macro** — F12 or right-click in the VBA code tab.
- **Ribbon → Ribbon Buttons…** lists every callback with its module and
  status; **Check Ribbon Callbacks** and **Validate XML** report callbacks
  that point to missing, ambiguous or class-module procedures, and duplicate
  control ids. Save warns (never blocks) only about problems your edits
  introduced.

## Safety model

- **Backups are automatic unless you switch them off.** By default every
  in-place save keeps the replaced file as a same-folder backup. With
  **Back Up Automatically on Save** off, a save leaves no backup file; use
  **Back Up Now** before risky edits.
- **Candidate verification.** The edited copy is fully parsed and validated
  (VBA structure, ZIP entry set, non-VBA payload byte-equality, module set and
  bodies) *before* the original is touched. Any failure leaves the original
  untouched.
- **Atomic commit.** The verified candidate is written beside the original and
  moved over it. When the original carries security settings, attributes or
  alternate data streams the new file would not have, Windows `ReplaceFileW`
  does the move so they are kept.
- **No change means no write.** Saving without edits never rewrites the file.
- **Password-protected projects are read-only.** No bypass exists. (XML-only
  package edits on `.xlam`/`.ppam`/`.pptm` remain possible; they never modify the
  protected VBA project.)
- **Strict encoding.** Characters that cannot be stored in the project's code
  page block the save instead of being silently replaced.
- **XML edits stay inside their package.** Only the exact XML parts you edited
  may differ; every other payload is verified byte-identical. An XML-only save
  never touches `vbaProject.bin`. The only members a save may add are the ones
  the draft adds (a first ribbon part, ribbon icons, made or imported with a
  button, and their relationships part); each must hold
  exactly the draft's bytes, and none may already exist. The only members a
  save may remove are images no ribbon control shows any more; nothing else is
  ever removed.
- **Digitally signed projects** require explicit confirmation; saving removes
  the (now invalid) VBA signature. Files with an **OPC package digital
  signature** block XML editing entirely.
- The tool never executes VBA, never changes Trust Center settings, never
  installs/uninstalls add-ins, and makes no network calls.

## Speed and antivirus

Windows Defender scans a macro package the first time it is opened for reading
after it was written, and for a large VBA project that takes seconds. The
editor therefore keeps the package in memory, builds and verifies the edited
file there, writes it once and never reads it back while you wait.

The scan itself cannot be skipped, so the editor starts it as soon as a save
(or Save a Copy) returns: it opens the saved file on a background thread and
the status bar shows **Antivirus scanning *name*…**, then ***name* ready for
Office** when the scan is done. Start Excel or PowerPoint after that and it
loads the add-in without waiting. If you start Office sooner, it waits for the
scan. If you save again while a scan is still running, that save waits for it
to finish first.

The same scan runs if you open an add-in that nothing has read since it was
last written.

Decoded module source is cached under
`%LOCALAPPDATA%\VBAAddinEditor\source-cache` so reopening an add-in does not
decompress modules that have not changed. The folder can be deleted at any time.

## Limitations (honest scope)

- The tool validates structure, not VBA compilation. A syntax error you type
  will surface when Office runs the macro. Use a backup to roll back.
- **XML editing** covers existing parts. The only parts the editor adds are
  a first ribbon part (`customUI/customUI14.xml`, for a file without one),
  ribbon icons (made, or imported with a button) and the relationships part
  that links them, and the only parts
  it removes are icon images nothing shows; it never renames package parts.
  No schema validation against the full OOXML schemas (syntax and
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
- **Import Button** follows calls lexically (no VBA compiler): a macro whose
  name is only built at run time (`Application.Run "Do" & n`) is not found,
  and library references are reported but not added. Sheets, names and other
  workbook content of the other add-in are never copied.
- UserForm layout is read-only: it can be viewed in the **UserForms** tab but not
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
- `pillow>=10.1,<13` (HPND) for custom ribbon icons; loaded only when an icon
  is used. The packaged build leaves out its font, colour-management, WebP
  and Tk modules.
- Python 3.10+; build with one pinned Python minor version.
