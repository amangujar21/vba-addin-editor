# Implementation Plan — PowerPoint OOXML/XML Editing + Right-Click Cut/Copy/Paste

**Repository:** `amangujar21/vba-addin-editor`  
**Target branch:** `main`  
**Plan date:** 2026-08-30  
**Purpose:** Handoff-ready implementation plan for adding XML editing for PowerPoint `.ppam` and `.pptm` packages, adding `.pptm` as a supported document type, and adding Windows-style right-click Cut/Copy/Paste to both the VBA and XML text editors.

---

## 1. Executive summary

The current application is intentionally conservative: it opens `.xlam` and `.ppam`, models VBA modules in an immutable snapshot plus editable draft, builds a candidate file through the pinned `pyopenvba==3.4.0` adapter, verifies that only `vbaProject.bin` changed, creates a mandatory backup, and atomically replaces the original with `ReplaceFileW`.

This work should preserve that architecture and safety model.

The implementation should **not** add a second, direct-to-file XML save path. XML edits must become a second type of draft change that flows through the same transaction as VBA changes:

1. Open the Office file and take a full fingerprint.
2. Snapshot VBA state.
3. For `.ppam` / `.pptm`, snapshot editable XML package parts.
4. Let the user edit VBA and XML independently in the UI.
5. Compute one composite change set.
6. Validate VBA changes and changed XML parts.
7. Build one candidate file.
8. Verify that the candidate contains exactly the intended VBA/XML changes and no unintended package payload changes.
9. Re-check process/lock/fingerprint.
10. Create one mandatory backup.
11. Commit once with `ReplaceFileW`.
12. Re-open and verify the saved original path.
13. Refresh both the VBA and XML baselines so the draft is clean.

Right-click Cut/Copy/Paste should be implemented once as shared text-widget behavior and installed into both editors. Do not duplicate clipboard/menu code in the VBA and XML editors.

### Recommended scope for this phase

- Keep existing `.xlam` VBA editing behavior unchanged.
- Keep existing `.ppam` VBA editing behavior unchanged.
- Add `.pptm` as a first-class supported PowerPoint file for VBA editing.
- Enable XML editing for `.ppam` and `.pptm`.
- XML editing applies to **existing XML package parts only**:
  - `*.xml`
  - `*.rels`
  - `[Content_Types].xml`
- Do **not** add/delete/rename package parts in this phase.
- Do **not** edit binary package entries.
- Do **not** build a schema-aware Ribbon designer or PowerPoint object model editor.
- Validate XML well-formedness and package isolation, but do not claim full OOXML schema validation.
- Preserve the existing “Office must be closed before in-place save” rule.
- Preserve mandatory backup and `ReplaceFileW`.
- Preserve the existing explicit VBA-signature-removal confirmation for VBA mutations.
- Detect OPC/package digital signatures and **block XML mutation of package-signed files in v1** rather than silently invalidating the signature.

---

## 2. Current repository baseline that this plan assumes

The implementation agent should first re-run the current tests and read the following files before changing anything:

| Area | Current file |
|---|---|
| Entry / self-test | `src/vba_addin_editor/__main__.py` |
| App startup | `src/vba_addin_editor/app.py` |
| Domain snapshot/draft | `src/vba_addin_editor/domain/document.py` |
| Change calculation | `src/vba_addin_editor/domain/changes.py` |
| Structured results | `src/vba_addin_editor/domain/results.py` |
| pyOpenVBA seam | `src/vba_addin_editor/adapters/pyopenvba_adapter.py` |
| VBA text codec | `src/vba_addin_editor/adapters/source_codec.py` |
| Open/reload workflow | `src/vba_addin_editor/services/document_service.py` |
| Safe save pipeline | `src/vba_addin_editor/services/save_service.py` |
| Import/export + backup restore | `src/vba_addin_editor/services/import_export_service.py` |
| Validation | `src/vba_addin_editor/services/validation_service.py` |
| Candidate/backup paths | `src/vba_addin_editor/platform/paths.py` |
| Windows process probe | `src/vba_addin_editor/platform/windows_processes.py` |
| ReplaceFileW / lock probe | `src/vba_addin_editor/platform/windows_file_ops.py` |
| Main UI | `src/vba_addin_editor/ui/main_window.py` |
| Current text editor | `src/vba_addin_editor/ui/code_editor.py` |
| Unit tests | `tests/unit/` |
| GUI integration | `tests/integration/test_gui_end_to_end.py` |
| Authentic Office release gate | `tests/integration/test_live_office_fixtures.py` |

Important current behaviors to preserve:

- `DocumentDraft.is_dirty()` currently derives from VBA module changes.
- `compute_changes()` is the UI/save/review source of truth.
- `SaveService.save_addin()` has the staged safe-save flow.
- `PyOpenVBAAdapter.build_candidate()` always opens the original fresh and writes a candidate, never the original.
- Candidate verification currently requires the ZIP entry set to remain identical and requires every non-VBA package payload to remain byte-identical.
- The in-place commit primitive is `ReplaceFileW`.
- Every in-place save gets a backup.
- The current GUI uses `self.tree` / `self.editor` for VBA.
- `CodeEditor` is a Tk `Text` wrapper with line numbers, undo, find and replace.
- `Ctrl+F` / `Ctrl+H` currently always target the VBA editor.
- The app is Windows-only, so `<Button-3>` is the primary context-menu trigger.
- The current file extension literals and UI file filters include only `.xlam` and `.ppam`.
- The current post-commit verifier opens the saved original path. Preserve that fact; XML support only changes what is used as the *reference* for package-payload comparison.

Before implementation:

```powershell
python -m pytest tests -q
python -m ruff check .
```

Record the baseline result in the implementation PR/commit notes.

---

## 3. User-visible target behavior

### 3.1 Supported file types

After this work:

- `.xlam`
  - VBA editor: yes, unchanged.
  - XML editor: not exposed in this phase.
- `.ppam`
  - VBA editor: yes.
  - XML editor: yes.
- `.pptm`
  - VBA editor: yes.
  - XML editor: yes.

The application can keep the product name `VBA Add-in Editor` for now, but user-facing strings that say only “add-in” should be generalized where they are now incorrect for `.pptm`.

Recommended UI wording:

- `Open Office VBA File…`
- `Save File`
- `Save a Copy…`
- Status: `Opened Presentation.pptm`
- Error: `The selected file is not a supported .xlam, .ppam, or .pptm file.`

Do not rename the package/project or executable in this task unless explicitly desired.

### 3.2 Editing model

For `.ppam` and `.pptm`, the main window should expose two editor tabs:

- **VBA**
  - Existing module tree.
  - Existing VBA code editor.
  - Existing module add/rename/delete/import/export behavior.
- **XML**
  - A tree/list of editable existing XML parts in the OOXML ZIP package.
  - A text editor showing the selected XML part.
  - Part path and encoding metadata.
  - Dirty marker for modified XML parts.
  - Optional `Validate XML` action.
  - No add/delete/rename part commands in v1.

For `.xlam`, the XML tab should be disabled or hidden. Do not silently expand XML editing to Excel in this phase.

### 3.3 Save behavior

A single Save may contain:

- VBA-only changes.
- XML-only changes.
- Both VBA and XML changes.

All three paths must create one final verified candidate and use the same backup/commit/recovery path.

### 3.4 Right-click editor behavior

Right-click in both the VBA and XML text editors should show:

- Cut
- Copy
- Paste

Behavior should feel like a normal Windows text editor:

- If the user right-clicks inside the current selection, preserve the selection.
- If the user right-clicks outside the selection, move the insertion cursor to the click location and clear the old selection.
- Cut is enabled only when text is selected and the widget is editable.
- Copy is enabled when text is selected.
- Paste is enabled when the widget is editable and text is available from the clipboard.
- Commands should use Tk virtual events (`<<Cut>>`, `<<Copy>>`, `<<Paste>>`) so undo/selection behavior remains native.
- Clipboard lookup failures must not crash the app.
- The menu must not mutate read-only/disabled editors.

---

## 4. Explicit non-goals

The implementation agent must not expand scope accidentally.

Not part of this phase:

- Editing `.xlam` OOXML parts.
- `.pptx`, `.potm`, `.xlsm`, `.docm`, `.dotm`, etc.
- Adding/removing/renaming arbitrary ZIP entries.
- Adding new Ribbon XML parts from scratch.
- Automatic relationship creation.
- Automatic `[Content_Types].xml` updates for new parts.
- Binary XML-adjacent parts such as ActiveX `.bin`, images, embedded OLE objects, media, or `vbaProject.bin`.
- XML schema validation against all ECMA/ISO Office schemas.
- VBA compilation.
- Package re-signing.
- VBA re-signing.
- Editing password-protected VBA.
- Bypassing package-level digital signatures.
- Direct write to the original ZIP.
- Extracting the Office package to a temporary directory.
- COM automation.
- Opening PowerPoint to perform saves.
- A visual Ribbon designer.

If future work needs package part creation/deletion, add it as a separate feature after this save/verification foundation is proven.

---

## 5. Architectural decisions

### 5.1 Keep pyOpenVBA isolated

Do not put ZIP/XML manipulation into `pyopenvba_adapter.py` beyond the small verification allowance needed to distinguish intended XML changes.

Create a separate OOXML/package adapter:

```text
src/vba_addin_editor/adapters/ooxml_package_adapter.py
```

All standard-library `zipfile` and XML package inspection/write logic should live there.

The existing pyOpenVBA rule remains:

> GUI and services do not import pyOpenVBA directly.

The new rule should be:

> GUI code does not manipulate ZIP entries or XML bytes directly.

### 5.2 XML is a draft concern, not a UI-only buffer

Do not leave XML text only inside Tk widgets.

The domain snapshot/draft must contain XML part state so:

- dirty state works,
- Review Changes works,
- Revert All works,
- Save can be driven headlessly,
- tests can mutate XML without GUI,
- the same model supports both GUI and CLI verification.

### 5.3 One composite transaction

Never implement:

```text
save VBA -> replace original
then
save XML -> replace original again
```

That would create two race windows, two backups, and a partial-success state.

The intended flow is:

```text
original
  |
  +-- VBA-only ----------> candidate
  |
  +-- XML-only ----------> candidate
  |
  +-- VBA + XML -> VBA candidate -> XML-patched candidate
                                      |
                                      v
                                 verification
                                      |
                                   backup
                                      |
                                 ReplaceFileW
```

### 5.4 Existing package entry set is immutable in v1

Candidate verification should continue to require:

```python
set(candidate.namelist()) == set(reference.namelist())
```

Additionally detect duplicate member names, because a set comparison hides duplicate ZIP entries.

No XML edit should add or remove a package part.

### 5.5 Preserve exact untouched uncompressed payloads

For every entry except:

- `ppt/vbaProject.bin` when there are VBA changes, and
- the exact XML part names in `changed_xml_parts()`,

the uncompressed bytes must be identical between reference and candidate.

Recompression of unchanged entries may change ZIP-level compressed bytes/offsets. That is acceptable. The verifier should compare `ZipFile.read(name)` payload bytes, not raw compressed stream bytes.

### 5.6 Do not serialize user XML with ElementTree

ElementTree may normalize namespace prefixes, whitespace, declaration formatting, or attribute ordering.

Use XML parsing for **validation only**.

When writing a changed XML part:

1. Take the user's editor text.
2. Apply the controlled original newline/encoding policy.
3. Encode it directly.
4. Write those exact bytes into the package.

---

## 6. Domain model changes

Modify:

```text
src/vba_addin_editor/domain/document.py
```

### 6.1 Extension types

Change:

```python
Extension = Literal[".xlam", ".ppam"]
```

to:

```python
Extension = Literal[".xlam", ".ppam", ".pptm"]
```

`HostKind` can remain:

```python
Literal["excel", "powerpoint"]
```

### 6.2 Add XML part snapshot metadata

Recommended model:

```python
@dataclass(frozen=True)
class XmlPartSnapshot:
    path: str
    text: str
    encoding: str
    bom: bytes
    newline: str
    original_sha256: str
    original_size: int
    is_relationships_part: bool
    is_content_types_part: bool
```

Notes:

- `path` is the exact ZIP member name, without a leading `/`.
- Use exact case from the package.
- `text` is normalized to LF for Tk editing.
- `encoding` is the actual codec used to decode the source.
- `bom` is retained separately to reproduce BOM behavior.
- `newline` should be `"\n"` or `"\r\n"` based on the original bytes; if mixed, choose a deterministic policy and record that the original was mixed if useful.
- `original_sha256` is for diagnostics and targeted verification.
- Do not use a random ID; the exact package path is already stable and unique after duplicate-entry validation.

### 6.3 Add XML part draft

Recommended model:

```python
@dataclass
class XmlPartDraft:
    path: str
    text: str
    original_text: str
    encoding: str
    bom: bytes
    newline: str
    is_relationships_part: bool
    is_content_types_part: bool

    def is_dirty(self) -> bool:
        return self.text != self.original_text
```

No `is_new` / `is_deleted` flags in v1.

### 6.4 Add package safety metadata

Recommended:

```python
@dataclass(frozen=True)
class PackageSafetyInfo:
    opc_signature_present: bool
    signature_part_names: tuple[str, ...] = ()
```

Then add to `DocumentSnapshot`:

```python
package_safety: PackageSafetyInfo
xml_parts: tuple[XmlPartSnapshot, ...]
```

For `.xlam` in this phase:

- `xml_parts=()`
- package safety can still be default/empty.

### 6.5 Extend `DocumentDraft`

Add:

```python
xml_parts: list[XmlPartDraft] = field(default_factory=list)
```

Add queries:

```python
def xml_part_by_path(self, path: str) -> XmlPartDraft | None: ...
def changed_xml_parts(self) -> list[XmlPartDraft]: ...
```

Update:

```python
def is_dirty(self) -> bool:
```

so it returns true if VBA or XML is dirty.

### 6.6 Draft creation/revert

Update `draft_from_snapshot()` to copy both module and XML state.

Update `revert_all()` indirectly through `draft_from_snapshot()`.

Add a single-part XML revert helper in `domain/changes.py`:

```python
def revert_xml_part(draft: DocumentDraft, path: str) -> None:
    ...
```

This is useful even if the first UI exposes only Revert All; it makes tests and future context menus straightforward.

---

## 7. Composite change-set changes

Modify:

```text
src/vba_addin_editor/domain/changes.py
```

Preserve the public `ChangeSet` name to minimize churn.

Add:

```python
@dataclass(frozen=True)
class XmlPartChange:
    path: str
```

Extend:

```python
@dataclass(frozen=True)
class ChangeSet:
    modified: tuple[ModuleChange, ...] = ()
    added: tuple[ModuleChange, ...] = ()
    renamed: tuple[ModuleRename, ...] = ()
    deleted: tuple[ModuleChange, ...] = ()
    xml_modified: tuple[XmlPartChange, ...] = ()
```

Add convenience properties:

```python
@property
def has_vba_changes(self) -> bool:
    return bool(self.modified or self.added or self.renamed or self.deleted)

@property
def has_xml_changes(self) -> bool:
    return bool(self.xml_modified)
```

Update `is_empty`.

Update `summary_lines()` to make sections unambiguous:

```text
VBA — Modified
  • Module1
VBA — Renamed
  • Old → New
XML — Modified
  • customUI/customUI14.xml
  • ppt/presentation.xml
```

Update `compute_changes(draft)` to append every `XmlPartDraft` where:

```python
part.text != part.original_text
```

Update `dirty_count()` to include XML part changes.

Do not count every changed XML line as a change. Count one changed package part as one change.

---

## 8. New XML text codec

Create:

```text
src/vba_addin_editor/adapters/xml_codec.py
```

Keep XML encoding/newline rules out of the UI and OOXML ZIP adapter.

### 8.1 Required functions

Suggested API:

```python
@dataclass(frozen=True)
class DecodedXml:
    text: str
    encoding: str
    bom: bytes
    newline: str

class XmlCodecError(ValueError):
    pass

def decode_xml(raw: bytes) -> DecodedXml:
    ...

def encode_xml(
    text: str,
    *,
    encoding: str,
    bom: bytes,
    newline: str,
) -> bytes:
    ...

def validate_xml_bytes(raw: bytes, *, part_path: str) -> tuple[str, ...]:
    ...

def validate_xml_draft(part: XmlPartDraft) -> tuple[str, ...]:
    ...
```

### 8.2 Encoding detection order

Implement deterministic detection:

1. UTF-8 BOM.
2. UTF-16 LE BOM.
3. UTF-16 BE BOM.
4. XML declaration encoding from the initial bytes.
5. XML default encoding: UTF-8.

Normalize aliases with `codecs.lookup()`.

Reject unknown codecs with a plain-language `XmlCodecError`.

### 8.3 Encoding policy on save

For v1, preserve the original encoding.

If the user changes:

```xml
<?xml version="1.0" encoding="UTF-8"?>
```

to a declaration that conflicts with the original codec, reject the save with a specific message rather than emitting bytes that contradict the declaration.

Example:

```text
customUI/customUI14.xml declares UTF-16, but this part was opened as UTF-8.
Changing an XML part's character encoding is not supported in this version.
```

Removing the declaration may be allowed if the resulting bytes are valid under the original encoding.

### 8.4 Newline policy

Normalize decoded editor text to LF.

On save:

- convert LF back to the part's original dominant newline style,
- preserve whether the original used CRLF or LF,
- do not alter unchanged parts at all.

Mixed-newline XML should either:

- retain a recorded `mixed` mode and save changed text with LF, or
- normalize changed text to the dominant original style.

Choose one deterministic behavior and cover it with tests. Recommended: dominant original style, defaulting to LF.

### 8.5 Well-formedness validation

Use the standard library parser for validation only.

Recommended:

```python
xml.etree.ElementTree.fromstring(raw)
```

Catch `ParseError` and surface line/column information.

Also reject obviously unsafe/non-OOXML DTD/entity declarations in editable package XML:

- `<!DOCTYPE`
- `<!ENTITY`

This avoids entity-related surprises and is appropriate for Office OOXML parts.

### 8.6 Lightweight shape checks

Do not attempt schema validation, but add two cheap structural checks:

For `[Content_Types].xml`, require root QName:

```text
{http://schemas.openxmlformats.org/package/2006/content-types}Types
```

For any `.rels` part, require root QName:

```text
{http://schemas.openxmlformats.org/package/2006/relationships}Relationships
```

Other XML parts only need to be well-formed.

---

## 9. New OOXML package adapter

Create:

```text
src/vba_addin_editor/adapters/ooxml_package_adapter.py
```

Use only Python standard library dependencies unless a later requirement proves insufficient.

### 9.1 Responsibilities

This adapter owns:

- Detecting editable XML package parts.
- Rejecting duplicate ZIP member names.
- Decoding XML using `xml_codec.py`.
- Detecting OPC/package digital signatures.
- Rewriting selected XML parts into a new package file.
- Verifying expected XML bytes after candidate build.
- Checking that XML changes are isolated to the intended paths.

It must not understand VBA module semantics.

### 9.2 Extension scope

Constants:

```python
XML_EDITABLE_EXTENSIONS = {".ppam", ".pptm"}
XML_PART_SUFFIXES = {".xml", ".rels"}
```

Special-case:

```text
[Content_Types].xml
```

Even though it already ends in `.xml`, call it out explicitly in tests and metadata.

### 9.3 XML part discovery

Suggested:

```python
def snapshot_xml_parts(self, path: Path) -> tuple[XmlPartSnapshot, ...]:
    ...
```

Algorithm:

1. Open with `zipfile.ZipFile`.
2. Run duplicate member validation.
3. Iterate `ZipInfo` entries.
4. Skip directories.
5. Include if name is `[Content_Types].xml`, ends with `.xml`, or ends with `.rels`, case-insensitive.
6. Read raw bytes.
7. Enforce decompression limits.
8. Decode strictly with `decode_xml()`.
9. Validate that existing XML is parseable.
10. Build snapshots sorted by case-insensitive path.

If an existing Office XML part cannot be decoded/parsed, opening should fail conservatively with an `AdapterError`-style user message. Do not show a corrupted part as editable and later rewrite the package.

### 9.4 ZIP safety limits

Because the app may open untrusted Office files, add conservative limits.

Recommended starting constants:

```python
MAX_XML_PART_BYTES = 32 * 1024 * 1024
MAX_TOTAL_XML_BYTES = 256 * 1024 * 1024
MAX_XML_PART_COUNT = 10_000
```

These are decompressed byte limits.

Do not extract ZIP members to disk, so path traversal from member names cannot write outside the package.

Reject duplicate member names before any modification because:

- `set(namelist())` hides duplicates,
- `ZipFile.read(name)` can produce ambiguous behavior,
- creating a second same-name XML entry is unsafe.

### 9.5 Package signature detection

Add:

```python
def inspect_package_signature(self, path: Path) -> PackageSafetyInfo:
    ...
```

Detect at least:

- `_xmlsignatures/` parts,
- the OPC digital-signature origin relationship type,
- digital-signature content types.

Return the exact signature-related part names for diagnostics.

V1 policy:

- XML edits + OPC/package signature => save blocked.
- Do not silently remove package signatures.
- VBA-only saves should continue to follow the current VBA-signature behavior; package-signature handling for a VBA-only mutation can remain conservative if detected, but do not regress existing signed-VBA semantics without tests.

### 9.6 XML candidate rewrite

Suggested API:

```python
def write_xml_candidate(
    self,
    source_path: Path,
    dest_path: Path,
    changed_parts: list[XmlPartDraft],
) -> None:
    ...
```

Important: `source_path` may be:

- the original file for XML-only changes, or
- the already-built VBA candidate for combined changes.

Never append changed XML entries to an existing ZIP because append mode would create duplicate member names.

Rewrite the ZIP to `dest_path`:

1. Validate all changed parts first.
2. Build `replacement_bytes: dict[str, bytes]`.
3. Open source `ZipFile` read-only.
4. Reject duplicates.
5. Open destination `ZipFile` in write mode.
6. Iterate `source.infolist()` in original order.
7. For each member:
   - if changed XML path: write encoded replacement bytes;
   - otherwise: write original uncompressed payload.
8. Preserve `ZipInfo` metadata as far as `zipfile` supports:
   - filename
   - date/time
   - compression type
   - comment
   - extra
   - internal/external attributes
   - create system/version fields where possible
9. Preserve archive comment.
10. Close both archives.
11. `testzip()` destination before returning.

Do not use `shutil.make_archive` or filesystem extraction/repack.

### 9.7 Atomic handling inside the candidate workspace

For combined changes, do not rewrite the candidate “in place” with `ZipFile`.

Use:

```text
candidate.pptm
candidate.xmlpatch.pptm
```

Then after XML patch validation succeeds:

```python
patched.replace(candidate)
```

This is safe because both are disposable candidate files, not the installed original.

Add a helper in `platform/paths.py` if useful:

```python
xml_patch_path_for(candidate: Path) -> Path
```

or generate a sibling temporary name inside the save service.

---

## 10. Add `.pptm` support to the existing VBA adapter

Modify:

```text
src/vba_addin_editor/adapters/pyopenvba_adapter.py
```

### 10.1 Supported extensions

Change:

```python
SUPPORTED_EXTENSIONS = (".xlam", ".ppam")
```

to:

```python
SUPPORTED_EXTENSIONS = (".xlam", ".ppam", ".pptm")
```

### 10.2 VBA entry mapping

Add:

```python
_VBA_ENTRY = {
    ".xlam": "xl/vbaProject.bin",
    ".ppam": "ppt/vbaProject.bin",
    ".pptm": "ppt/vbaProject.bin",
}
```

### 10.3 Host process mapping

Add:

```python
_HOST_PROCESS = {
    ".xlam": "EXCEL.EXE",
    ".ppam": "POWERPNT.EXE",
    ".pptm": "POWERPNT.EXE",
}
```

### 10.4 Host class dispatch

Use normal `PowerPointFile` for `.pptm`.

Recommended:

```python
def _host_class(path: Path) -> type[VBAHostFile]:
    suffix = path.suffix.lower()
    if suffix == ".xlam":
        return ExcelFile
    if suffix == ".ppam":
        return PPAMPowerPointFile
    if suffix == ".pptm":
        return PowerPointFile
    raise AdapterError(...)
```

Do not unnecessarily route `.pptm` through the PPAM compatibility subclass.

### 10.5 Host kind

Make host mapping explicit rather than relying on “anything non-xlam is PowerPoint”:

```python
def host_kind_for(path: Path) -> HostKind:
    suffix = path.suffix.lower()
    if suffix == ".xlam":
        return "excel"
    if suffix in {".ppam", ".pptm"}:
        return "powerpoint"
    raise AdapterError(...)
```

### 10.6 Candidate verifier allowance

Current verification treats every non-VBA difference as an error.

Change the method signature from conceptually:

```python
verify_candidate(original_path, candidate_path, expected)
```

to:

```python
verify_candidate(
    reference_path: Path,
    candidate_path: Path,
    expected: DocumentDraft,
    *,
    allowed_non_vba_changes: frozenset[str] = frozenset(),
) -> CandidateVerificationResult:
```

In the non-VBA payload comparison:

```python
if name == vba_entry:
    continue
if name in allowed_non_vba_changes:
    continue
```

Every other package entry still must be payload-identical.

Rename the local concept from `original_path` to `reference_path` where appropriate because post-commit verification may use the backup as the reference.

Do not remove existing VBA structural/module verification.

### 10.7 Optional stronger invariant

When there are no VBA changes, have either the save service or verifier explicitly assert:

```python
sha256(reference.read(vba_entry)) == sha256(candidate.read(vba_entry))
```

That proves an XML-only save did not touch the VBA project, including a VBA signature.

---

## 11. Document open/reload changes

Modify:

```text
src/vba_addin_editor/services/document_service.py
```

### 11.1 Constructor

Add package adapter injection:

```python
class DocumentService:
    def __init__(
        self,
        adapter: PyOpenVBAAdapter | None = None,
        package_adapter: OoxmlPackageAdapter | None = None,
    ) -> None:
        self.adapter = adapter or PyOpenVBAAdapter()
        self.package_adapter = package_adapter or OoxmlPackageAdapter()
```

### 11.2 Open

After the existing VBA snapshot:

```python
snapshot = self.adapter.open_snapshot(...)
```

If extension is `.ppam` or `.pptm`:

1. Snapshot XML parts.
2. Inspect package signature.
3. Use `dataclasses.replace()` to attach:
   - `xml_parts`
   - `package_safety`

For `.xlam`:

- leave XML parts empty.

Then call `draft_from_snapshot(snapshot)`.

### 11.3 Error messages

Change the extension error from only `.xlam/.ppam` to:

```text
Only .xlam, .ppam, and .pptm files are supported.
```

### 11.4 Snapshot report

Optionally append:

```text
package digitally signed
```

when package safety reports an OPC signature.

Do not confuse that with the VBA project digital signature.

---

## 12. Revised safe-save pipeline

Modify:

```text
src/vba_addin_editor/services/save_service.py
```

This is the most important part of the implementation.

### 12.1 Constructor

Inject:

```python
package_adapter: OoxmlPackageAdapter | None = None
```

Default to a real adapter in `__post_init__` or use a default factory.

Keep commit/process/lock/progress injection for tests.

### 12.2 Stage A — composite change detection

Existing:

```python
changes = compute_changes(draft)
if changes.is_empty:
    return SaveResult.no_changes()
```

remains, but `is_empty` now includes XML.

Define:

```python
has_vba = changes.has_vba_changes
has_xml = changes.has_xml_changes
allowed_xml_paths = frozenset(c.path for c in changes.xml_modified)
```

### 12.3 Stage B/C — process, lock, external fingerprint

No conceptual change.

PowerPoint process blocking applies to both `.ppam` and `.pptm`.

### 12.4 Stage D — protection and signatures

Split the current all-or-nothing protection logic by mutation type.

#### Password-protected VBA

If `has_vba` and VBA project is password-protected:

```text
block: password_protected
```

For XML-only changes, the preferred policy is:

- allow the XML save,
- require `vbaProject.bin` to remain byte-identical,
- do not call any VBA mutation APIs.

This lets the package XML be edited without bypassing the protected VBA project.

If testing shows pyOpenVBA cannot safely reopen/verify a protected project after an XML-only package rewrite, fall back to conservative blocking and document that limitation.

#### VBA signature

Require signature-removal confirmation only if:

```python
has_vba and draft.baseline.safety.signature_present
```

An XML-only save must not set `allow_invalidate_signature=True` or rewrite `vbaProject.bin`.

#### OPC/package signature

If:

```python
has_xml and draft.baseline.package_safety.opc_signature_present
```

block:

```text
This PowerPoint package contains an OPC digital signature.
Editing package XML would invalidate that signature.
XML editing of package-signed files is blocked in this version.
```

Do not offer to remove it in this task.

### 12.5 Stage E0 — VBA structural validation

Run `validate_draft(draft)` when `has_vba`.

It is acceptable to run it for all saves, but do not let irrelevant pre-existing VBA naming oddities block an XML-only save if the current validator would not have blocked the clean original.

Preferred rule:

- VBA validation gates only VBA mutations.

### 12.6 Stage E1 — VBA code-page validation

Run existing changed-module encoding validation only when `has_vba`.

### 12.7 Stage E2 — XML validation

Before creating any candidate:

```python
for part in draft.changed_xml_parts():
    problems.extend(package_adapter.validate_draft_part(part))
```

Block with reason:

```text
invalid_xml
```

Include part path and parser line/column.

### 12.8 Stage F/G/H/I — candidate build matrix

Implement explicit branches.

#### Case 1: VBA-only

Use the current build path unchanged:

```python
adapter.build_candidate(
    original,
    draft,
    candidate,
    allow_signature_removal=draft.signed_save_confirmed,
)
```

This is the regression-sensitive path; minimize changes.

#### Case 2: XML-only

Do **not** call `host.save()`.

Instead:

```python
package_adapter.write_xml_candidate(
    source_path=original,
    dest_path=candidate,
    changed_parts=draft.changed_xml_parts(),
)
```

This guarantees `vbaProject.bin` is copied from the original payload unchanged.

#### Case 3: VBA + XML

1. Build the VBA candidate through pyOpenVBA.
2. Produce a second XML-patched candidate from the VBA candidate.
3. Replace the first disposable candidate with the second.

Pseudo:

```python
adapter.build_candidate(original, draft, candidate, ...)

xml_candidate = paths.xml_patch_path_for(candidate)
try:
    package_adapter.write_xml_candidate(
        source_path=candidate,
        dest_path=xml_candidate,
        changed_parts=draft.changed_xml_parts(),
    )
    xml_candidate.replace(candidate)
finally:
    _discard(xml_candidate)
```

Do not patch the original and then ask pyOpenVBA to save over it.

### 12.9 Stage J — candidate verification

Run both layers.

#### Package/VBA verifier

```python
vba_verification = adapter.verify_candidate(
    original,
    candidate,
    draft,
    allowed_non_vba_changes=allowed_xml_paths,
)
```

This continues to prove:

- ZIP opens,
- entry set identical,
- untouched non-VBA payloads identical,
- VBA structure valid,
- expected module state/bodies present,
- code page unchanged,
- signature behavior correct.

#### XML verifier

Add:

```python
xml_verification = package_adapter.verify_candidate_xml(
    reference_path=original,
    candidate_path=candidate,
    draft=draft,
)
```

It must prove:

- every changed XML part exists,
- each changed XML part's bytes exactly equal `encode_xml(...)` for the draft,
- changed XML parses,
- no XML part outside the expected change set differs,
- no duplicate names exist,
- if no VBA change, `ppt/vbaProject.bin` is payload-identical.

Merge verification problems into one `CandidateVerificationResult` or into the existing `SaveResult.candidate_failed()` path.

Keep the `.failed` candidate behavior for diagnostics.

### 12.10 Stage L — race narrowing

No conceptual change:

- Office process check.
- fingerprint check.
- lock check.

The original file is still untouched at this point.

### 12.11 Stage K/M — backup + ReplaceFileW

No conceptual change.

One save => one backup => one ReplaceFileW.

### 12.12 Stage N — post-commit verification

This requires a deliberate adjustment for XML support.

The file that must be verified is still:

```text
the saved original path
```

Do not try to verify the candidate path after `ReplaceFileW`, because the candidate has been moved.

However, for package payload comparison, the saved file now needs a pre-save reference. The backup is exactly that reference.

Therefore use:

```python
final = adapter.verify_candidate(
    backup,
    original,
    draft,
    allowed_non_vba_changes=allowed_xml_paths,
)
```

and:

```python
final_xml = package_adapter.verify_candidate_xml(
    reference_path=backup,
    candidate_path=original,
    draft=draft,
)
```

This is stronger than comparing `original` to itself.

Important implementation note:

- `backup` has the same extension as the original.
- Existing backup tests already assert it contains the pre-save bytes.
- The verifier should use the reference path only for package comparison and extension/VBA entry lookup; it should reopen the candidate/saved original for semantic VBA verification.

If either final verifier fails:

- return `recovery_required`,
- offer Restore Backup exactly as today.

### 12.13 Refresh baseline after success

The current save service calls `adapter.open_snapshot()` directly, which only knows VBA.

After XML support, refresh both VBA and XML state.

Recommended private helper in `SaveService`:

```python
def _open_full_snapshot(self, path: Path) -> DocumentSnapshot:
    fp = paths.fingerprint(path)
    snapshot = self.adapter.open_snapshot(path, fp, host_process_running=False)
    if path.suffix.lower() in XML_EDITABLE_EXTENSIONS:
        snapshot = replace(
            snapshot,
            xml_parts=self.package_adapter.snapshot_xml_parts(path),
            package_safety=self.package_adapter.inspect_package_signature(path),
        )
    return snapshot
```

Then:

```python
new_snapshot = self._open_full_snapshot(original)
refreshed = draft_from_snapshot(new_snapshot)

draft.baseline = new_snapshot
draft.modules[:] = refreshed.modules
draft.xml_parts[:] = refreshed.xml_parts
draft.signed_save_confirmed = False
```

Ensure the result is completely clean.

### 12.14 Save a Copy

Apply the same build/verify matrix to `save_copy()`.

Differences:

- no backup,
- no ReplaceFileW,
- original fingerprint can remain the reference for verification,
- destination extension should normally match the source extension.

Do not let the user save a `.pptm` draft to `.ppam` merely by changing the suffix. In v1 require exact extension match unless an explicit format-conversion feature is implemented later.

---

## 13. OOXML candidate verification details

The XML/package verifier should be strict enough that an implementation agent cannot accidentally weaken the current guarantees.

Suggested function:

```python
def verify_candidate_xml(
    self,
    reference_path: Path,
    candidate_path: Path,
    draft: DocumentDraft,
) -> CandidateVerificationResult:
    ...
```

Algorithm:

1. Validate both reference and candidate are ZIP packages.
2. Reject duplicate member names in either.
3. Run `candidate.testzip()`.
4. Assert exact member-name list/set equality.
5. Build:
   ```python
   changed_paths = {p.path for p in draft.changed_xml_parts()}
   ```
6. For every changed path:
   - read candidate bytes,
   - encode expected draft bytes,
   - exact byte equality required,
   - parse candidate bytes,
   - run special root checks for `.rels` / `[Content_Types].xml`.
7. For every editable XML path not in `changed_paths`:
   - reference bytes == candidate bytes.
8. If no VBA change:
   - `ppt/vbaProject.bin` reference bytes == candidate bytes.
9. Optionally record SHA-256 detail maps for diagnostics.
10. Return all problems, not only the first.

The existing pyOpenVBA verifier continues checking all *other* non-VBA payloads too, so binary images/media/ActiveX entries remain protected against accidental mutation.

### 13.1 Intended-difference exactness

Do not merely say “this XML part was allowed to differ.”

Require:

```text
candidate bytes == exact bytes generated from current draft
```

This catches:

- wrong selected part,
- stale editor flush,
- encoding bugs,
- writing an earlier version of the text,
- wrong newline handling.

---

## 14. UI architecture

Modify:

```text
src/vba_addin_editor/ui/main_window.py
```

Add:

```text
src/vba_addin_editor/ui/xml_editor.py
src/vba_addin_editor/ui/text_context_menu.py
```

Optionally refactor common text-editor logic into:

```text
src/vba_addin_editor/ui/text_editor.py
```

### 14.1 Minimize regression in existing VBA UI

Keep:

```python
self.tree
self.editor
```

as the VBA tree/editor so current module methods need minimal change.

Add:

```python
self.editor_notebook
self.xml_tree
self.xml_editor
self.current_xml_part_path
```

### 14.2 Recommended layout

Use a top-level `ttk.Notebook` in the main content area:

```text
+---------------------------------------------------------+
| toolbar                                                 |
+---------------------------------------------------------+
| [ VBA ] [ XML ]                                         |
|                                                         |
| VBA tab:                                                |
| +----------------+------------------------------------+ |
| | module tree    | VBA CodeEditor                     | |
| +----------------+------------------------------------+ |
|                                                         |
| XML tab:                                                |
| +----------------+------------------------------------+ |
| | XML part tree  | XML editor                        | |
| |                | part path / encoding header       | |
| +----------------+------------------------------------+ |
+---------------------------------------------------------+
| status                                                  |
+---------------------------------------------------------+
```

Alternative nested notebook/pane layouts are fine, but do not merge VBA modules and ZIP parts into one overloaded tree. Their operations and selection semantics are different.

### 14.3 XML tree

Sort by path case-insensitively.

At minimum a flat tree is acceptable:

```text
[Content_Types].xml
_rels/.rels
customUI/customUI.xml
customUI/customUI14.xml
ppt/_rels/presentation.xml.rels
ppt/presentation.xml
ppt/slides/slide1.xml
...
```

A hierarchical folder tree is a nice improvement but not required for correctness.

Dirty marker:

```text
customUI/customUI14.xml *
```

Use a stable item ID that cannot collide with module IDs, e.g.:

```python
iid = "xml::" + part.path
```

### 14.4 XML editor header

Show:

```text
customUI/customUI14.xml
UTF-8 · CRLF
```

Optional warning near `[Content_Types].xml` and relationship parts:

```text
Package-critical XML. Invalid relationships/content types can make PowerPoint repair or reject the file.
```

Do not make those parts read-only; the user explicitly asked to modify XML.

### 14.5 XML tab on `.xlam`

Either:

- disable the XML notebook tab, or
- show a small message:
  `XML package editing is enabled for .ppam and .pptm in this version.`

Do not populate `.xlam` XML parts.

---

## 15. Editor flush and selection plumbing

Current VBA logic uses `_flush_active_editor()`.

Refactor to:

```python
def _flush_active_vba_editor(self) -> None:
    ...

def _flush_active_xml_editor(self) -> None:
    ...

def _flush_all_editors(self) -> None:
    self._flush_active_vba_editor()
    self._flush_active_xml_editor()
```

Replace all save/review/revert/close call sites that must capture current text with `_flush_all_editors()`.

Critical call sites:

- before opening a new file,
- before Save,
- before Save a Copy,
- before Review Changes,
- before Revert All,
- on app close,
- before switching away from a selected VBA module,
- before switching away from a selected XML part.

### 15.1 XML selection

Implement:

```python
def on_xml_part_selected(self, _event=None) -> None:
    self._flush_active_xml_editor()
    ...
    self.current_xml_part_path = selected_path
    self.xml_editor.set_text(part.text)
```

Programmatic `set_text()` must not mark the draft dirty.

### 15.2 Save success

After successful save:

- refresh both trees,
- preserve selected VBA module if it still exists,
- preserve selected XML path if it still exists,
- reload editor text from the refreshed clean draft,
- refresh dirty markers/count/status.

---

## 16. Shared text editor refactor

The existing `CodeEditor` is already mostly generic.

There are two acceptable approaches.

### Preferred approach: generic base

Create:

```text
src/vba_addin_editor/ui/text_editor.py
```

Move common behavior:

- `Text` creation.
- line number canvas.
- scrollbars.
- undo.
- find.
- replace.
- cursor position.
- set/get text.
- context-menu installation.

Then:

```python
class CodeEditor(TextEditor):
    pass
```

and:

```python
class XmlEditor(TextEditor):
    pass
```

`XmlEditor` may add:

- validation command,
- path/encoding display outside the base widget,
- optional XML-specific highlighting later.

### Lower-risk approach

Keep `CodeEditor` unchanged except for installing shared context-menu behavior, and subclass/reuse it for XML.

If doing this, update the class/module naming comments so future maintainers know the widget is intentionally generic.

Do not copy/paste the entire `CodeEditor` implementation into `xml_editor.py`.

---

## 17. Right-click Cut/Copy/Paste implementation

Create:

```text
src/vba_addin_editor/ui/text_context_menu.py
```

Suggested public helper:

```python
class TextContextMenu:
    def __init__(self, text: tk.Text) -> None:
        ...

def install_text_context_menu(text: tk.Text) -> TextContextMenu:
    ...
```

Keep a reference on the editor object so Tk does not garbage-collect relevant objects unexpectedly.

### 17.1 Menu creation

One `tk.Menu` per text widget:

```python
menu = tk.Menu(text, tearoff=0)
menu.add_command(label="Cut", command=...)
menu.add_command(label="Copy", command=...)
menu.add_command(label="Paste", command=...)
```

### 17.2 Right-click handler

Bind:

```python
text.bind("<Button-3>", self._show_menu, add=True)
```

Windows-only app means no cross-platform button abstraction is required for this task.

### 17.3 Selection behavior

Helper:

```python
def _has_selection(text: tk.Text) -> bool:
    try:
        text.index("sel.first")
        text.index("sel.last")
        return True
    except tk.TclError:
        return False
```

To determine whether the click is inside selection:

```python
clicked = text.index(f"@{event.x},{event.y}")
```

Compare with `sel.first` / `sel.last`.

If click is outside selection:

```python
text.tag_remove("sel", "1.0", "end")
text.mark_set("insert", clicked)
```

Then:

```python
text.focus_set()
```

### 17.4 Enabled state

Before `tk_popup`:

- Copy = normal if selection exists.
- Cut = normal if selection exists and widget state is not disabled.
- Paste:
  - widget editable,
  - `clipboard_get()` succeeds.

Catch:

```python
tk.TclError
```

for empty/non-text clipboard.

### 17.5 Commands

Use virtual events:

```python
text.event_generate("<<Cut>>")
text.event_generate("<<Copy>>")
text.event_generate("<<Paste>>")
```

Do not manually manipulate clipboard text unless tests show a Tk virtual-event issue.

This preserves:

- platform clipboard conventions,
- selection replacement,
- undo stack behavior.

### 17.6 Keyboard shortcuts

Existing native Ctrl+X/C/V behavior should remain.

Do not bind Ctrl+X/C/V globally unless needed; Tk `Text` already handles them.

### 17.7 Find/replace shortcuts

Current `MainWindow._bind_shortcuts()` sends Ctrl+F/Ctrl+H to `self.editor`.

Change it to resolve the active text editor:

```python
def _active_text_editor(self):
    if self.editor_notebook.select() == str(self.xml_tab):
        return self.xml_editor
    return self.editor
```

Then:

```python
Ctrl+F -> active.find_dialog()
Ctrl+H -> active.replace_dialog()
```

This is required for the XML editor to behave as a peer editor.

---

## 18. Dirty-state UI and Review Changes

### 18.1 Global dirty count

`dirty_count(draft)` includes VBA and XML.

Toolbar:

```text
Presentation.pptm — unsaved changes: 3
```

still works.

### 18.2 Per-tree markers

VBA tree keeps current markers.

XML tree:

- unchanged: no marker,
- changed: ` *`.

### 18.3 Review Changes

Update `review_changes()` and `_preflight_dialog()` to use the composite summary.

Example:

```text
Review changes to MyAddin.ppam

VBA — Modified
  • Module1

XML — Modified
  • customUI/customUI14.xml

Safety
  ✓ File has not changed on disk
  ✓ PowerPoint is closed
  ✓ VBA project is not password-protected
  ✓ No OPC package signature detected
  ! VBA digital signature will be removed

A backup will be created before the original is replaced.
```

Only show relevant warnings.

For XML-only changes with a signed VBA project:

- do **not** say the VBA signature will be removed,
- because `vbaProject.bin` must remain identical.

### 18.4 Save button enablement

Current UI disables Save for any password-protected project.

That must be revised if XML-only edits are allowed on password-protected VBA.

Recommended logic:

- if no draft: disabled,
- if Office running: disabled,
- if saving: disabled,
- otherwise normal.

Let `SaveService` make change-specific safety decisions.

If desired, disable VBA module mutation commands when password-protected while keeping XML editor writable.

---

## 19. XML editor validation UX

Save-time validation is mandatory.

Also add a lightweight manual action:

- Button or XML menu item: `Validate XML`.

Behavior:

1. Flush selected XML editor.
2. Validate only selected part.
3. On success:
   `XML is well-formed.`
4. On failure:
   show path + line/column.

This action does not write anything.

It is optional for Definition of Done if save-time errors are clear, but it is strongly recommended because users editing raw XML need fast feedback.

---

## 20. File menu, filters, CLI, and wording

### 20.1 Main UI file types

Update `_FILETYPES`:

```python
_FILETYPES = [
    ("Supported Office VBA files", "*.xlam;*.ppam;*.pptm"),
    ("Excel Add-ins", "*.xlam"),
    ("PowerPoint Add-ins", "*.ppam"),
    ("PowerPoint Macro-Enabled Presentations", "*.pptm"),
]
```

### 20.2 Save a Copy

Use the source extension as default.

Reject a destination with a different supported extension unless an explicit conversion path exists.

### 20.3 CLI help

Update `__main__.py`:

```text
file: Office VBA file (.xlam/.ppam/.pptm)
```

`--self-test` should also report XML part count for `.ppam/.pptm`.

Recommended output:

```text
OK Demo.pptm: project='VBAProject' code_page=1252 modules=3 xml_parts=47
```

### 20.4 Self-roundtrip

Extend self-roundtrip so `.pptm` can exercise a combined change when possible:

- append VBA marker,
- choose a safe XML part and append/change an XML comment or known synthetic marker in a test fixture,
- build candidate,
- verify.

Do not mutate authentic user files; current behavior already works on a temporary copy.

### 20.5 About/README wording

Update:

- `.pptm` support.
- XML editing scope.
- existing-parts-only limitation.
- package signature blocking.
- no full OOXML schema validation.

---

## 21. Backup restore behavior

`BackupService.restore()` should continue to restore the entire file, so XML changes require no special restore logic.

Ensure `.pptm` is accepted by `PyOpenVBAAdapter.open_snapshot()` so backup restore works for `.pptm`.

Recommended enhancement:

After restore verification, use the full snapshot logic so a restored `.ppam/.pptm` also has parseable XML parts.

At minimum:

- VBA snapshot must open,
- ZIP must pass,
- XML parts must parse.

Do not restore only `vbaProject.bin` or only XML. Backup restore remains whole-file recovery.

---

## 22. Error/result changes

Modify `domain/results.py` only as needed.

Possible new reasons:

```text
invalid_xml
xml_encoding
package_signed
unsupported_extension
package_corrupt
```

`SaveResult.kind` can remain unchanged.

Use `blocked` for policy conditions:

- Office running.
- locked.
- external change.
- password-protected VBA when VBA changes.
- package-signed when XML changes.

Use `error` for invalid user edit/build issues:

- malformed XML,
- conflicting XML encoding declaration,
- ZIP rewrite failure.

Use `candidate_failed` when candidate was built but verification found unintended/missing changes.

### 22.1 Generalize messages

Current messages repeatedly say `add-in`.

For code touched by this feature, prefer:

- `Office file`
- `PowerPoint file`
- `edited copy`
- `original file`

Keep “add-in” where it specifically means `.xlam/.ppam`.

---

## 23. Tests — required automated coverage

Do not consider this feature complete based only on manual UI testing.

### 23.1 New unit test file: XML codec

Create:

```text
tests/unit/test_xml_codec.py
```

Required cases:

1. UTF-8 no BOM decode/encode.
2. UTF-8 BOM preserved.
3. UTF-16 LE BOM preserved.
4. UTF-16 BE if supported by decoder.
5. XML declaration encoding detection.
6. Invalid declared encoding rejected.
7. Encoding declaration changed to conflicting encoding rejected.
8. CRLF normalized to LF in editor and restored on encode.
9. LF preserved.
10. Malformed XML reports line/column.
11. DTD/ENTITY policy rejected.
12. `[Content_Types].xml` correct root accepted.
13. bad `[Content_Types].xml` root rejected.
14. `.rels` correct root accepted.
15. bad `.rels` root rejected.

### 23.2 New unit test file: OOXML package adapter

Create:

```text
tests/unit/test_ooxml_package_adapter.py
```

Build tiny synthetic ZIP packages with standard `zipfile`; these tests do not need valid PowerPoint/VBA.

Required cases:

1. Discovers `.xml`, `.rels`, `[Content_Types].xml`.
2. Ignores binary entries.
3. Exact ZIP member names retained.
4. Duplicate member names rejected.
5. XML-only rewrite changes exactly one payload.
6. Unchanged payload bytes identical after rewrite.
7. Entry set unchanged.
8. Archive passes `testzip()`.
9. Changed XML expected bytes exact.
10. Unexpected second XML payload change fails verifier.
11. Binary payload change fails verifier/integration isolation check.
12. OPC signature markers detected.
13. Compression/archive metadata preservation tested for representative fields.
14. Oversized XML part blocked by configured limit.
15. Total XML size/count limits enforced.

### 23.3 Domain tests

Extend:

```text
tests/unit/test_domain.py
```

Required:

- `draft_from_snapshot()` creates XML drafts.
- XML text change makes draft dirty.
- returning text to original makes it clean.
- `compute_changes()` contains `xml_modified`.
- `dirty_count()` includes XML.
- `revert_all()` reverts XML.
- `revert_xml_part()` reverts only selected part.

### 23.4 `.pptm` adapter tests

Extend:

```text
tests/unit/test_adapter_and_platform.py
```

Use `PowerPointFile.create_new()` from pinned pyOpenVBA to create a synthetic `.pptm`.

Required:

- `.pptm` dispatch uses PowerPoint.
- `host_kind == "powerpoint"`.
- `vba_entry_for(.pptm) == "ppt/vbaProject.bin"`.
- process mapping is `POWERPNT.EXE`.
- open snapshot succeeds.
- VBA candidate roundtrip succeeds.

### 23.5 Save pipeline tests

Extend:

```text
tests/unit/test_save_pipeline.py
```

Add synthetic `.pptm` fixture support in `tests/conftest.py`.

Create a helper to inject a benign XML marker into an existing XML part if necessary.

Required matrix:

#### Regression

- existing XLAM no-change.
- existing XLAM VBA edit.
- existing XLAM add/rename/delete.
- existing signature behavior.
- existing commit failure.
- existing external-change block.

#### PPTM VBA-only

- edit VBA only.
- XML parts remain payload-identical.
- backup equals original pre-save bytes.

#### PPTM XML-only

- edit one XML part.
- VBA project bytes unchanged.
- all other payloads unchanged.
- backup created.
- reopen shows XML text.
- draft clean after save.

#### PPTM combined

- edit VBA module.
- edit XML part.
- one save succeeds.
- both changes present on reopen.
- no unintended payload changes.
- one backup.

#### XML failure paths

- malformed XML => original byte-identical.
- encoding conflict => original byte-identical.
- candidate verifier detects unintended XML change => original untouched.
- package rewrite exception => original untouched.
- external change => no write.
- Office running => no write.
- lock => no write.
- ReplaceFileW failure => original preserved.
- post-commit XML verification failure => recovery required.

#### Signature/protection policy

- signed VBA + XML-only => VBA signature not removed and no confirmation required, if supported by test fixture/mocking.
- signed VBA + VBA change => existing confirmation still required.
- package-signed + XML change => blocked.
- password-protected VBA + XML-only => allowed only if chosen policy is proven safe; otherwise explicit block test.
- password-protected VBA + VBA change => blocked.

### 23.6 PPAM parser-level tests

There is no need to pretend a renamed PPTM is an authentic PPAM release fixture.

For parser/unit coverage, it is acceptable to use a synthetic PPTM-derived package renamed to `.ppam` only if the test is explicitly labeled synthetic and proves only adapter mechanics.

Authentic PPAM release qualification remains mandatory.

### 23.7 GUI integration tests

Extend:

```text
tests/integration/test_gui_end_to_end.py
```

Add:

#### XML edit roundtrip

1. Create/open synthetic `.pptm`.
2. Select XML tab.
3. Select a known XML part.
4. Insert/change text in the real Tk `Text`.
5. Flush.
6. call save pipeline.
7. reopen.
8. assert XML change present.
9. assert draft clean.

#### Combined GUI roundtrip

- edit VBA Text widget,
- edit XML Text widget,
- save once,
- assert both.

### 23.8 Context-menu GUI tests

Create:

```text
tests/integration/test_text_context_menu.py
```

Use a real Tk root when display is available; skip otherwise as existing GUI tests do.

Run the same tests against both `CodeEditor` and `XmlEditor`.

Required:

1. Right-click menu can be created.
2. Copy enabled with selection.
3. Cut enabled with selection/editable.
4. Paste enabled with clipboard text/editable.
5. Cut removes selected text.
6. Copy preserves text.
7. Paste inserts clipboard text.
8. Paste replaces selection.
9. Right-click inside selection preserves selection.
10. Right-click outside selection moves cursor/clears selection.
11. Disabled editor: Cut/Paste disabled.
12. Disabled editor: Copy still possible if selection is available.
13. Empty clipboard does not raise.
14. Context-menu operations participate in undo.

If directly testing `tk_popup()` is flaky, expose small testable helpers for menu-state calculation and command invocation, while still keeping at least one real widget integration test.

---

## 24. Test fixtures

### 24.1 Synthetic PPTM fixture

Update:

```text
tests/conftest.py
```

Add:

```python
from pyopenvba import PowerPointFile

def build_pptm(path: Path) -> Path:
    with PowerPointFile.create_new(path) as host:
        host.set_module(...)
        host.save()
    return path
```

Then, if needed, use the new OOXML adapter to replace an existing benign XML part with a known marker or to make a deterministic XML change.

Do not add arbitrary new package parts merely for the test unless the package is still valid.

### 24.2 Authentic live fixtures

Extend live fixture contract:

```text
tests/fixtures/ppam/RealAddin.ppam
tests/fixtures/pptm/RealPresentation.pptm
```

Continue to gitignore authentic binary fixtures.

Recommended contents:

#### PPAM

- `ModuleKeep`
- `ModuleEdit`
- `ModuleDelete`
- custom Ribbon XML (`customUI/customUI.xml` or `customUI/customUI14.xml`) with a known label:
  `VBAAE_XML_ORIGINAL`

#### PPTM

- same VBA module contract,
- a known XML marker in a safe existing part or customUI:
  `VBAAE_XML_ORIGINAL`

The automated live test should copy fixtures before mutation.

---

## 25. Live Office release qualification

Current parser tests are not enough to promise that PowerPoint will accept arbitrary edited XML without repair.

Add live test coverage in:

```text
tests/integration/test_live_office_fixtures.py
```

### 25.1 Automated live mutation cycle

For `.ppam` and `.pptm`:

1. Copy authentic fixture to temp.
2. Open via `DocumentService`.
3. Assert XML parts are available.
4. Edit VBA marker.
5. Edit known XML marker:
   `VBAAE_XML_ORIGINAL` -> `VBAAE_XML_EDITED`.
6. Add VBA module.
7. Rename VBA module.
8. Delete VBA module.
9. Save through real `SaveService`:
   - real process probe,
   - real lock probe,
   - real candidate,
   - real backup,
   - real `ReplaceFileW`.
10. Reopen via `DocumentService`.
11. Assert all VBA changes.
12. Assert XML marker changed.
13. Assert untouched XML marker/control part unchanged.
14. Assert draft clean.
15. Assert backup exists.

### 25.2 Human PowerPoint gate

For both authentic `.ppam` and `.pptm`:

1. Ensure PowerPoint is closed.
2. Edit VBA + XML through the packaged executable.
3. Save in place.
4. Start PowerPoint.
5. Verify **no repair dialog**.
6. Verify presentation/add-in loads.
7. Verify edited VBA macro runs.
8. If customUI XML was changed, verify the Ribbon control/label reflects the XML edit.
9. Close PowerPoint.
10. Restore backup with the tool.
11. Reopen and verify original behavior.

This is a release-blocking gate for claiming XML editing is safe.

---

## 26. Packaging and dependency policy

No new third-party dependency should be required.

Use:

- `zipfile`
- `hashlib`
- `codecs`
- `re`
- `xml.etree.ElementTree`

If an implementation agent proposes `lxml`, require a concrete reason because it:

- enlarges the onefile executable,
- adds native wheel/build considerations,
- changes third-party notices,
- is unnecessary for well-formedness-only validation.

If no new dependency is added:

- `pyproject.toml` runtime dependencies stay pinned to `pyopenvba==3.4.0`.
- PyInstaller spec should pick up new Python modules automatically.
- Re-run onefile and onedir smoke tests.

---

## 27. Detailed implementation sequence

Implement in small commits so regressions are easy to isolate.

### Commit 0 — Baseline checkpoint

- Run tests/lint.
- Record results.
- No code changes.

### Commit 1 — `.pptm` extension support only

Files:

- `domain/document.py`
- `adapters/pyopenvba_adapter.py`
- `services/document_service.py`
- `platform/windows_processes.py`
- `ui/main_window.py`
- `__main__.py`
- tests

Acceptance:

- `.pptm` opens.
- VBA edit/save works on synthetic PPTM.
- existing `.xlam/.ppam` tests remain green.
- no XML editor yet.

### Commit 2 — XML codec + XML domain models

Files:

- new `adapters/xml_codec.py`
- `domain/document.py`
- `domain/changes.py`
- unit tests

Acceptance:

- no save/UI changes yet.
- codec and dirty/revert semantics fully unit-tested.

### Commit 3 — OOXML package read/verification adapter

Files:

- new `adapters/ooxml_package_adapter.py`
- `services/document_service.py`
- tests

Acceptance:

- `.ppam/.pptm` drafts contain XML snapshots.
- duplicate/invalid package XML is rejected safely.
- `.xlam` XML list remains empty.

### Commit 4 — XML UI tab

Files:

- new `ui/xml_editor.py`
- optional `ui/text_editor.py`
- `ui/main_window.py`
- GUI tests

Acceptance:

- XML parts browse and edit in memory.
- dirty count/review/revert include XML.
- no actual XML save yet if needed to keep commit small.

### Commit 5 — Shared right-click context menu

Files:

- new `ui/text_context_menu.py`
- `ui/code_editor.py` / shared base
- `ui/xml_editor.py`
- GUI tests

Acceptance:

- Cut/Copy/Paste works in both editors.
- no regression in undo/find/replace.

This commit is independent enough that it can happen before Commit 4 if preferred.

### Commit 6 — XML candidate writer

Files:

- `ooxml_package_adapter.py`
- `platform/paths.py` if temp helper added
- package adapter tests

Acceptance:

- targeted XML rewrite into a new file.
- entry set unchanged.
- untouched payloads unchanged.
- exact expected bytes.

### Commit 7 — Composite safe-save pipeline

Files:

- `services/save_service.py`
- `adapters/pyopenvba_adapter.py`
- `domain/results.py` if needed
- tests

Acceptance:

- VBA-only regression path green.
- XML-only save green.
- combined save green.
- all failure paths leave original safe.

### Commit 8 — Post-commit verification + full snapshot refresh

Files:

- `save_service.py`
- package adapter
- tests

Acceptance:

- post-commit reference uses backup.
- saved original reopens.
- draft clean for modules + XML.
- recovery path triggers on injected final XML verification failure.

### Commit 9 — Live fixtures/test matrix + docs

Files:

- `test_live_office_fixtures.py`
- `README.md`
- `HANDOFF.md` or a new handoff note
- optional release checklist

Acceptance:

- normal automated suite green.
- live tests skip clearly without fixtures.
- authentic PPAM/PPTM gate documented.

### Commit 10 — Packaging smoke

- `ruff`
- `pytest`
- `pyright` if part of project workflow
- onedir build
- onefile build
- `--self-test`
- `--self-roundtrip`
- manual GUI smoke

---

## 28. Specific file-by-file change checklist

### `src/vba_addin_editor/domain/document.py`

- Add `.pptm`.
- Add `PackageSafetyInfo`.
- Add `XmlPartSnapshot`.
- Add `XmlPartDraft`.
- Add XML fields to snapshot/draft.
- Add XML query/dirty methods.
- Update `draft_from_snapshot()`.

### `src/vba_addin_editor/domain/changes.py`

- Add `XmlPartChange`.
- Extend `ChangeSet`.
- Add VBA/XML convenience properties.
- Extend `summary_lines()`.
- Extend `compute_changes()`.
- Extend `dirty_count()`.
- Add `revert_xml_part()`.

### `src/vba_addin_editor/domain/results.py`

- Add new reasons/messages only as necessary.
- Generalize “add-in” wording in candidate/recovery messages.

### `src/vba_addin_editor/adapters/pyopenvba_adapter.py`

- Add `.pptm`.
- Add mappings.
- Explicit host dispatch.
- Add `allowed_non_vba_changes`.
- Rename verification reference parameter semantics.
- Preserve existing module verification.

### `src/vba_addin_editor/adapters/xml_codec.py`

- New.
- Strict decode/encode.
- newline handling.
- well-formedness.
- special root checks.
- encoding declaration consistency.

### `src/vba_addin_editor/adapters/ooxml_package_adapter.py`

- New.
- XML discovery.
- duplicate detection.
- signature detection.
- limits.
- candidate rewrite.
- XML candidate verification.

### `src/vba_addin_editor/services/document_service.py`

- Inject package adapter.
- attach XML/package safety to snapshots.
- support `.pptm`.
- wording.

### `src/vba_addin_editor/services/save_service.py`

- Inject package adapter.
- composite change branches.
- change-specific signature/protection policy.
- XML validation.
- combined candidate build.
- dual verification.
- post-commit comparison against backup.
- full snapshot refresh.
- same logic in Save a Copy.

### `src/vba_addin_editor/services/import_export_service.py`

- Backup restore `.pptm` support via adapter.
- Consider full package/XML validation after restore.
- Generalize wording.

### `src/vba_addin_editor/platform/windows_processes.py`

Current fallback of “non-xlam = PowerPoint” happens to handle `.pptm`, but make supported mapping explicit or keep it only if tests make the assumption intentional.

Recommended:

```python
_HOST_PROCESS = {
    ".xlam": "EXCEL.EXE",
    ".ppam": "POWERPNT.EXE",
    ".pptm": "POWERPNT.EXE",
}
```

### `src/vba_addin_editor/platform/paths.py`

- optional XML-patch candidate helper.
- no change to backup naming semantics.

### `src/vba_addin_editor/ui/text_context_menu.py`

- New shared context menu.

### `src/vba_addin_editor/ui/code_editor.py`

- Install shared context menu.
- or become subclass of shared text editor.

### `src/vba_addin_editor/ui/xml_editor.py`

- New.
- Reuse common text editor behavior.

### `src/vba_addin_editor/ui/main_window.py`

- file filters.
- notebook/tabs.
- XML tree/editor.
- flush-all logic.
- active find/replace.
- composite review/preflight.
- selection restore.
- Save button policy.
- wording/about.

### `src/vba_addin_editor/__main__.py`

- help strings.
- `.pptm`.
- XML count in self-test.
- optional combined self-roundtrip.

### `README.md`

- `.pptm`.
- XML scope and limitations.
- package signature limitation.
- existing-parts-only limitation.
- live PowerPoint release gate.

---

## 29. Implementation pitfalls to avoid

### 29.1 Do not append into ZIPs

This is wrong:

```python
with ZipFile(candidate, "a") as z:
    z.writestr("ppt/presentation.xml", new_bytes)
```

It can create duplicate entries.

Always reconstruct a new candidate ZIP.

### 29.2 Do not extract/repack through a directory

Avoid:

```text
extractall -> modify files -> zip directory
```

It creates:

- path traversal concerns,
- metadata drift,
- accidental entry additions/removals,
- ordering differences,
- more filesystem failure modes.

### 29.3 Do not parse-and-reserialize XML

Do not do:

```python
root = ET.fromstring(text)
ET.ElementTree(root).write(...)
```

for the final user content.

Only parse to validate.

### 29.4 Do not weaken untouched-payload verification

The current safety promise is strong because non-VBA payload changes are rejected.

XML support should change:

```text
all non-VBA differences forbidden
```

to:

```text
only exact user-selected XML differences are allowed;
everything else remains forbidden
```

Do not replace this with a broad “candidate ZIP opens” check.

### 29.5 Flush both editors before computing changes

A stale XML Tk buffer is an easy bug.

Every save/review/close path must call `_flush_all_editors()` before `compute_changes()`.

### 29.6 Avoid programmatic dirty events

`set_text()` should:

- replace text,
- reset undo,
- clear modified flag,
- not call the domain `on_change` callback as if the user typed.

If refactoring the editor, explicitly test this.

### 29.7 Keep XML-only saves away from pyOpenVBA mutation

If XML-only save calls `host.save()`, it may unnecessarily rewrite `vbaProject.bin`, defeating the guarantee that VBA signatures/protected content are untouched.

### 29.8 Do not confuse VBA signatures with OPC package signatures

They are separate safety concerns.

- VBA signature: current detector inside VBA CFB.
- OPC/package signature: package-level XML signature parts/relationships.

### 29.9 Do not save XML editor text using platform default encoding

Never:

```python
text.encode()
```

without the captured XML codec.

### 29.10 Do not allow extension conversion by Save a Copy accidentally

`.pptm` -> `.ppam` is not a valid conversion just because both are ZIP packages.

Require matching suffix.

### 29.11 Do not declare schema correctness

A well-formed XML file can still violate PowerPoint's OOXML schema or relationship rules.

UI/docs should state:

> The editor validates XML syntax and package integrity. PowerPoint remains the final format validator.

The live no-repair gate is therefore required.

---

## 30. Suggested helper APIs

These signatures are not mandatory, but keeping close to them will make the implementation coherent.

### XML codec

```python
def decode_xml(raw: bytes) -> DecodedXml: ...
def encode_xml(text: str, *, encoding: str, bom: bytes, newline: str) -> bytes: ...
def validate_xml_draft(part: XmlPartDraft) -> tuple[str, ...]: ...
```

### Package adapter

```python
class OoxmlPackageAdapter:
    def snapshot_xml_parts(self, path: Path) -> tuple[XmlPartSnapshot, ...]: ...
    def inspect_package_signature(self, path: Path) -> PackageSafetyInfo: ...
    def write_xml_candidate(
        self,
        source_path: Path,
        dest_path: Path,
        changed_parts: list[XmlPartDraft],
    ) -> None: ...
    def verify_candidate_xml(
        self,
        reference_path: Path,
        candidate_path: Path,
        draft: DocumentDraft,
    ) -> CandidateVerificationResult: ...
```

### Domain

```python
class DocumentDraft:
    def changed_xml_parts(self) -> list[XmlPartDraft]: ...
    def xml_part_by_path(self, path: str) -> XmlPartDraft | None: ...
```

### UI

```python
def _flush_active_vba_editor(self) -> None: ...
def _flush_active_xml_editor(self) -> None: ...
def _flush_all_editors(self) -> None: ...
def on_xml_part_selected(self, _event=None) -> None: ...
def _refresh_xml_tree(self) -> None: ...
def _active_text_editor(self): ...
```

---

## 31. Save pipeline pseudocode

Implementation agent should be able to map the final code to this structure.

```python
def save_addin(self, draft: DocumentDraft) -> SaveResult:
    original = draft.baseline.path

    changes = compute_changes(draft)
    if changes.is_empty:
        return SaveResult.no_changes()

    has_vba = changes.has_vba_changes
    has_xml = changes.has_xml_changes
    allowed_xml = frozenset(c.path for c in changes.xml_modified)

    # Process / lock / external file fingerprint
    preflight_original(original, draft)

    # Mutation-specific safety
    if has_vba and draft.baseline.safety.password_protected:
        return blocked_password()

    if has_vba and draft.baseline.safety.signature_present:
        if not draft.signed_save_confirmed:
            return needs_signature_confirmation()

    if has_xml and draft.baseline.package_safety.opc_signature_present:
        return blocked_package_signature()

    # Validation
    if has_vba:
        validate_vba_draft(draft)
        validate_vba_encoding(draft)

    if has_xml:
        validate_changed_xml(draft)

    candidate = paths.candidate_path_for(original)

    try:
        if has_vba:
            adapter.build_candidate(
                original,
                draft,
                candidate,
                allow_signature_removal=draft.signed_save_confirmed,
            )
        else:
            package_adapter.write_xml_candidate(
                original,
                candidate,
                draft.changed_xml_parts(),
            )

        if has_vba and has_xml:
            patched = temp_xml_candidate(candidate)
            package_adapter.write_xml_candidate(
                candidate,
                patched,
                draft.changed_xml_parts(),
            )
            patched.replace(candidate)

        # Candidate verification
        v = adapter.verify_candidate(
            original,
            candidate,
            draft,
            allowed_non_vba_changes=allowed_xml,
        )
        x = package_adapter.verify_candidate_xml(
            original,
            candidate,
            draft,
        )
        fail_if_any(v, x)

        # Race narrowing
        preflight_original_again(original, draft)

        backup = paths.backup_path_for(original)
        self._commit(original, candidate, backup)

        # Verify SAVED ORIGINAL, compare against PRE-SAVE BACKUP
        v2 = adapter.verify_candidate(
            backup,
            original,
            draft,
            allowed_non_vba_changes=allowed_xml,
        )
        x2 = package_adapter.verify_candidate_xml(
            backup,
            original,
            draft,
        )
        if not v2.ok or not x2.ok:
            return SaveResult.recovery_required(...)

        refresh_full_draft_from_disk(draft, original)
        return SaveResult.success(backup, original)

    except expected_errors as exc:
        cleanup_candidate_files()
        return mapped_result(exc)
```

For VBA-only saves, the XML verifier can be skipped or run with an empty changed set. Running it provides extra package coverage but must not create new failure modes for `.xlam`, where XML editing is out of scope.

---

## 32. Acceptance criteria / Definition of Done

The feature is complete only when all applicable items are true.

### Functional

- [ ] Existing `.xlam` VBA editing still works.
- [ ] Existing `.ppam` VBA editing still works.
- [ ] `.pptm` opens and VBA editing works.
- [ ] `.ppam` shows XML parts.
- [ ] `.pptm` shows XML parts.
- [ ] User can edit existing `.xml` parts.
- [ ] User can edit existing `.rels` parts.
- [ ] User can edit `[Content_Types].xml`.
- [ ] User cannot add/delete/rename package parts in v1.
- [ ] VBA-only save works.
- [ ] XML-only save works.
- [ ] combined VBA+XML save works.
- [ ] Revert All reverts both.
- [ ] Review Changes shows both.
- [ ] close-with-unsaved-changes detects both.
- [ ] Save a Copy handles both.
- [ ] backup restore restores both.

### Context menu

- [ ] Right-click Cut works in VBA editor.
- [ ] Right-click Copy works in VBA editor.
- [ ] Right-click Paste works in VBA editor.
- [ ] Same three actions work in XML editor.
- [ ] Menu enabled/disabled state is correct.
- [ ] selection behavior is Windows-like.
- [ ] undo still works after Cut/Paste.
- [ ] empty clipboard cannot crash.

### Safety

- [ ] No-change still performs no write.
- [ ] Office running blocks in-place save.
- [ ] lock blocks save.
- [ ] external fingerprint change blocks save.
- [ ] original is untouched until candidate verifies.
- [ ] one mandatory backup per in-place save.
- [ ] one `ReplaceFileW` commit.
- [ ] candidate ZIP entry set cannot change.
- [ ] duplicate ZIP members are rejected.
- [ ] untouched package payloads are byte-identical.
- [ ] intended XML bytes match draft exactly.
- [ ] malformed XML cannot reach commit.
- [ ] XML-only save preserves `vbaProject.bin`.
- [ ] VBA signature confirmation happens only for VBA mutation.
- [ ] OPC/package-signed XML save is blocked.
- [ ] final verification reads saved original path.
- [ ] final payload comparison uses backup as the pre-save reference.
- [ ] recovery-required flow still works.

### Tests/release

- [ ] `python -m pytest tests -q` green, excluding intentional live skips.
- [ ] `python -m ruff check .` clean.
- [ ] pyright passes if used as release gate.
- [ ] synthetic PPTM tests green.
- [ ] GUI integration tests green on a display-capable Windows machine.
- [ ] authentic PPAM live automated test passes.
- [ ] authentic PPTM live automated test passes.
- [ ] human PowerPoint PPAM no-repair + VBA + XML check passes.
- [ ] human PowerPoint PPTM no-repair + VBA + XML check passes.
- [ ] onefile build smoke-tested.
- [ ] onedir build smoke-tested.
- [ ] packaged `--self-test` passes on `.xlam`, `.ppam`, `.pptm`.
- [ ] packaged roundtrip passes on qualified fixtures.

---

## 33. Handoff notes for the implementing agent

1. **Do not start in the UI.** First implement `.pptm`, XML codec/domain, and package adapter tests. The save guarantees are more important than rendering the XML tab.
2. **Keep the existing VBA-only candidate path as unchanged as possible.** It is already well-tested.
3. **The central new invariant is “allowed XML part paths”, not “non-VBA changes are okay now.”** Every difference must remain explicit.
4. **XML-only save should not call pyOpenVBA mutation/save.** Copy/rewrite the package while preserving `vbaProject.bin`.
5. **Combined save order should be VBA first, XML second.** pyOpenVBA then receives the known original file, and the package adapter patches only the disposable candidate.
6. **Never patch ZIP members in append mode.**
7. **Never serialize the user's XML via ElementTree.**
8. **Keep package signatures separate from VBA signatures.**
9. **Use the backup as the Stage N comparison reference.** The saved original path is still the verification target.
10. **Do not claim PowerPoint schema validity from XML parsing.** The authentic Office no-repair test remains a release gate.
11. **Do not broaden XML editing to `.xlam` in the same implementation unless explicitly requested.**
12. **Do not implement package part add/delete/rename under the guise of XML editing.** That requires new relationship/content-type semantics and a separate plan.
13. **Update the final `HANDOFF.md` after implementation** with:
    - exact test counts,
    - new files,
    - signature/protection policy,
    - live fixture requirements,
    - onefile hash/build status,
    - any deviations from this plan and why.

---

## 34. Recommended final verification commands

```powershell
# Full automated suite
python -m pytest tests -q

# Focused XML/package tests
python -m pytest tests/unit/test_xml_codec.py -q
python -m pytest tests/unit/test_ooxml_package_adapter.py -q

# Save pipeline
python -m pytest tests/unit/test_save_pipeline.py -q

# GUI tests
python -m pytest tests/integration/test_gui_end_to_end.py -q
python -m pytest tests/integration/test_text_context_menu.py -q

# Authentic Office fixtures
python -m pytest tests -m live -v

# Static checks
python -m ruff check .
python -m pyright

# Build
powershell -File scripts\build.ps1
powershell -File scripts\build.ps1 -Onefile

# Packaged verification examples
dist\VBAAddinEditor.exe --self-test path\to\RealAddin.ppam
dist\VBAAddinEditor.exe --self-test path\to\RealPresentation.pptm
dist\VBAAddinEditor.exe --self-roundtrip path\to\RealPresentation.pptm
```

On shells that cannot directly execute the built `.exe`, use the existing documented subprocess or `cmd /c` workaround.

---

## 35. Final implementation outcome

When this plan is complete, the application will still have one safety model:

> Open a known file, keep an immutable baseline, edit a draft, build a disposable candidate, permit only explicitly intended VBA/XML package changes, verify the candidate, re-check the original, create a mandatory backup, atomically replace the original, and verify the saved file again.

XML editing becomes another verified payload mutation inside that model rather than a separate ZIP editor bolted onto the application.

The right-click clipboard work is intentionally orthogonal: it should be a shared `Text` behavior used by both editors, with no domain/save coupling.
