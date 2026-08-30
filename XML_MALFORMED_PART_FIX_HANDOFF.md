# HANDOFF — Fix PPAM open failure caused by eager OOXML XML validation

**Project:** `amangujar21/vba-addin-editor`
**Target issue:** A PowerPoint add-in (`.ppam`) that opens in Office RibbonX Editor is rejected by VBA Add-in Editor with:

```text
Package XML part 'ppt/presentation.xml' is not valid:
ppt/presentation.xml XML is not well-formed: no element found: line 1, column 0
```

**Observed executable SHA-256:**

```text
892bccff558843e5eef8d1e83d6f76e11b87db572b313d8e665b789ee3352239
```

This executable contains XML/package functionality not currently visible on the public `main` branch, including modules/classes with names such as:

```text
vba_addin_editor.adapters.ooxml_package_adapter
vba_addin_editor.adapters.xml_codec
vba_addin_editor.ui.xml_editor
vba_addin_editor.ui.text_context_menu
```

The implementation agent must work from the source tree that produced the executable above, or from the newer local source containing those modules. **Do not implement against the older public `main` UI and assume the XML feature exists there.** If those files are missing, first locate/recover the source corresponding to this executable before modifying behavior.

---

## 1. Required outcome

Opening an otherwise usable `.xlam` or `.ppam` **must not fail merely because an unrelated existing OOXML XML part is empty, malformed, undecodable, or otherwise not parseable by our XML editor**.

The core policy after this change must be:

> **Existing package oddities are tolerated and preserved byte-for-byte. New invalid XML introduced by VBA Add-in Editor is blocked.**

Specifically:

1. A `.ppam` containing a malformed/empty `ppt/presentation.xml` must still open if the ZIP container and VBA project can be read.
2. VBA modules must remain editable.
3. A VBA-only save must preserve the malformed/empty XML part **byte-for-byte**.
4. The XML UI may show the malformed part, but it must not silently normalize, repair, re-encode, or rewrite it.
5. Prefer making an already-malformed part read-only in v1 of this fix. The UI should explain why.
6. XML parts that were valid on open remain editable.
7. User-edited XML must be validated before save; invalid edited XML must block save.
8. Candidate verification must validate changed XML and verify untouched parts are byte-identical; it must **not** require every pre-existing XML part in the package to parse successfully.
9. The same tolerance must apply during post-save reload/verification. Fixing only the first Open path is insufficient.
10. The exact PPAM that currently reproduces the error must pass a live regression test before this issue is considered closed.

---

## 2. Root cause

The XML/package feature currently performs eager validation of package XML while creating the document/package snapshot.

The failing flow is approximately:

```text
DocumentService.open()
  -> package-state enrichment / _with_package_state()
     -> OoxmlPackageAdapter.snapshot_xml_parts()
        -> enumerate *.xml / *.rels parts
        -> read each part
        -> decode XML
        -> validate XML
        -> raise PackageError if any part is invalid
```

For the failing PPAM, validation of:

```text
ppt/presentation.xml
```

returns an XML parser failure equivalent to:

```text
no element found: line 1, column 0
```

That message commonly means the parser received empty XML content, although the implementation must inspect the actual failing file and record the exact byte length/content before assuming this is literally a zero-byte part.

The architectural error is not simply “ElementTree rejects this XML.” The architectural error is:

> **The app treats successful parsing of every XML part as a prerequisite for opening the VBA project.**

That is too strict for this product. VBA editing only requires the Office package and `vbaProject.bin` to be readable. Unrelated outer-package XML should not prevent VBA work if it can be left untouched.

---

## 3. Mandatory reproduction before changing code

Before editing implementation code, reproduce the issue against the exact source build and exact PPAM.

### 3.1 Record the failing fixture

Keep the real customer/reproduction file outside version control if licensing/privacy prevents committing it, for example:

```text
tests/fixtures/live/ppam/ExistingMalformedPart.ppam
```

Add that location to `.gitignore` if necessary.

### 3.2 Inspect the relevant ZIP member without modifying the file

Use a small diagnostic such as:

```python
from pathlib import Path
from zipfile import ZipFile

path = Path(r"tests/fixtures/live/ppam/ExistingMalformedPart.ppam")
with ZipFile(path, "r") as zf:
    raw = zf.read("ppt/presentation.xml")
    print("length:", len(raw))
    print("head:", repr(raw[:200]))
    print("contains VBA:", "ppt/vbaProject.bin" in zf.namelist())
```

Record in the test/issue notes whether the member is:

- zero bytes,
- whitespace only,
- BOM only,
- valid bytes that our decoder mishandles,
- truncated XML,
- or another form of malformed XML.

Do **not** “repair” the fixture. The purpose of the regression is to prove that the editor tolerates and preserves the pre-existing state.

### 3.3 Verify external behavior

Record that the same original PPAM can be opened by Office RibbonX Editor. If practical, also verify PowerPoint itself can load it. This is a manual qualification datapoint, not a substitute for automated tests.

---

## 4. Implementation strategy

The fix should be made at the **package adapter / XML snapshot boundary**, not as a one-off exception for `ppt/presentation.xml` and not as a PPAM-specific filename hack.

Do **not** add code such as:

```python
if part_name == "ppt/presentation.xml":
    ignore_error = True
```

That would merely move the bug to the next unusual package part.

Instead, change the model from:

```text
part is invalid -> abort opening document
```

to:

```text
part is invalid -> record per-part problem -> preserve it -> continue opening document
```

Fatal errors should remain fatal only when the editor cannot safely work with the package at all, for example:

- file is not a readable ZIP/OOXML package where one is required,
- VBA project entry is missing when opening as a VBA add-in,
- VBA project cannot be parsed,
- package cannot be read due to I/O/permission corruption,
- a save candidate cannot be created or verified safely.

A malformed unrelated XML member is **not** a fatal Open error.

---

## 5. `OoxmlPackageAdapter.snapshot_xml_parts()` changes

Locate the implementation corresponding to:

```text
vba_addin_editor.adapters.ooxml_package_adapter
```

and specifically the method/function responsible for enumerating and snapshotting XML parts, expected to be named approximately:

```python
snapshot_xml_parts(...)
```

### 5.1 Current behavior to remove

The current open-time logic is believed to be structurally equivalent to:

```python
raw = zf.read(part_name)
decoded = decode_xml(raw)
problems = validate_xml_bytes(raw, part_name=part_name)
if problems:
    raise PackageError(
        f"Package XML part {part_name!r} is not valid: " + "; ".join(problems)
    )
```

Remove the **fatal escalation** for a per-part validation/decode problem during snapshot creation.

### 5.2 New behavior

Each XML part must produce a snapshot even when it cannot be parsed.

Conceptually:

```python
raw = zf.read(part_name)

try:
    decoded = decode_xml(raw)
except Exception as exc:
    return XmlPartSnapshot(
        ...,
        text=None,
        editable=False,
        well_formed=False,
        open_problem=f"Could not decode existing package part: {exc}",
        # retain enough original-byte identity to guarantee preservation
    )

problems = validate_xml_bytes(raw, part_name=part_name)
if problems:
    return XmlPartSnapshot(
        ...,
        text=decoded.text,
        editable=False,
        well_formed=False,
        open_problem="; ".join(problems),
        # retain original-byte identity
    )

return XmlPartSnapshot(
    ...,
    text=decoded.text,
    editable=True,
    well_formed=True,
    open_problem=None,
)
```

Adapt this to the actual model/API rather than copying it literally.

### 5.3 Preserve raw identity

The snapshot/draft model must retain enough information to prove that an untouched part is unchanged. One of these patterns is acceptable:

**Preferred if already consistent with the architecture:**

- original raw bytes are retained in the XML part baseline/snapshot, or
- an SHA-256 hash + byte length are retained while save rebuilds from the original ZIP.

The critical invariant is:

> **An untouched part is copied from the original package bytes. It is never reconstructed from parsed XML text.**

Do not parse and reserialize untouched XML, even when it is valid. Reserialization can change:

- XML declaration,
- BOM,
- encoding,
- whitespace,
- attribute ordering,
- namespace prefix spelling,
- line endings,
- processing instructions,
- comments,
- empty-element formatting.

The save pipeline should only replace raw bytes for parts the user actually changed.

### 5.4 Catch errors per part, not around the entire enumeration

The loop should be shaped so one bad part cannot prevent other parts from being listed:

```python
for part_name in xml_parts:
    try:
        snapshot = snapshot_one_part(...)
    except ExpectedPerPartXmlError as exc:
        snapshot = make_uneditable_problem_snapshot(...)
    parts.append(snapshot)
```

Do not swallow broad package-level failures that make safe editing impossible. Distinguish:

- **per-part XML decode/parse failure** -> nonfatal,
- **ZIP/I/O/package access failure** -> fatal.

If the current code wraps everything in `except Exception`, narrow it.

---

## 6. XML snapshot/draft domain model

Locate the XML part snapshot/draft dataclasses or equivalent domain objects. Expected concepts include `XmlPartSnapshot` and a mutable XML draft type.

The model needs, either explicitly or equivalently:

```python
part_name: str
text: str | None
editable: bool
well_formed_on_open: bool
open_problem: str | None
original_encoding: str | None
original_bom: ...
original_byte_length: int
original_sha256: str
```

Not every field above is mandatory if the existing model stores equivalent state, but the following behaviors are mandatory:

1. The UI can distinguish a healthy editable XML part from a preserved malformed one.
2. A malformed baseline part cannot accidentally become dirty merely because it was selected/displayed.
3. Revert restores exact baseline state.
4. `changed_xml_parts()` must not report untouched malformed parts as changes.
5. Save verification can prove untouched bytes were preserved.

### 6.1 Recommended v1 policy for malformed baseline parts

Make them **read-only**.

This avoids ambiguous behavior such as letting the editor implicitly “repair” a package part that was invalid before the user touched the file.

Do not implement automatic repair in this bug fix.

A future feature may provide an explicit “Edit/repair malformed part” workflow, but that must be a deliberate user action with a separate safety design.

---

## 7. `xml_codec` behavior

Locate:

```text
vba_addin_editor.adapters.xml_codec
```

Expected functions may include concepts such as:

```python
decode_xml(...)
validate_xml_bytes(...)
encode_xml(...)
```

### 7.1 Validation must return problems, not crash the open path

`validate_xml_bytes()` should safely handle:

- `b""`,
- whitespace-only bytes,
- BOM-only input,
- malformed XML,
- invalid encoding declaration,
- valid UTF-8 XML,
- valid UTF-16 XML if the app claims to support it.

For invalid bytes, it should return a clear problem object/string (or a typed result) rather than leaking a raw parser exception through the document-open path.

Example expected result:

```text
ppt/presentation.xml XML is not well-formed: no element found: line 1, column 0
```

The message itself can remain. What changes is **where it is used**: as per-part status, not as a reason to reject the whole add-in.

### 7.2 Do not decode before deciding whether the app needs text

For an existing malformed/undecodable part, the application only needs to preserve the bytes. Therefore, decoding failure must not prevent VBA editing.

### 7.3 User-edited XML remains strict

When the user modifies a healthy XML part:

1. encode using the supported/original encoding policy,
2. validate the exact encoded bytes that would be written,
3. block save if not well-formed.

Do not validate one representation and write another representation.

---

## 8. `DocumentService` open/reload changes

Locate the newer `DocumentService` code that enriches the VBA draft with package/XML state, expected to include a helper similar to:

```python
_with_package_state(...)
```

The Open contract must become:

```text
read/fingerprint file
-> open VBA snapshot
-> snapshot package XML parts
   -> record per-part XML problems nonfatally
-> return complete draft
```

Do not catch the XML adapter error only in the GUI and suppress it there. The domain/service API itself should return a valid document draft with per-part diagnostics.

That makes headless tests, CLI, reload, and GUI all behave consistently.

### 8.1 Reload must use the same tolerant policy

Any reload after save, revert, or external refresh must not resurrect the original bug.

---

## 9. `SaveService` changes

Inspect the newer save service carefully. Previous inspection indicates it has concepts similar to:

```python
_validate_draft_parts(...)
_open_full_snapshot(...)
changed_xml_parts()
package_adapter.validate_draft_part(...)
```

### 9.1 Keep strict validation for changed XML

This is the correct place for strict XML validation.

Before building a candidate:

```python
for part in draft.changed_xml_parts():
    problems.extend(package_adapter.validate_draft_part(part))
```

If there are problems, return a blocked/validation result and do not write anything.

### 9.2 Do not validate untouched baseline XML as a save prerequisite

Remove/avoid any call equivalent to:

```python
validate_every_xml_part(candidate_or_original)
```

when that call treats a pre-existing malformed untouched part as a save failure.

The save gate should distinguish:

- malformed **changed** XML -> block,
- malformed **unchanged baseline** XML -> preserve and permit VBA-only save.

### 9.3 `_open_full_snapshot()` must be tolerant too

This is essential.

If save verification/reload calls `snapshot_xml_parts()` again, the tolerant behavior must be centralized in the adapter so both initial Open and post-save verification behave identically.

Do not patch only `DocumentService.open()`.

---

## 10. Candidate package writer

Inspect how XML changes are written into the candidate ZIP.

Required behavior:

```text
original package
  + replace ppt/vbaProject.bin if VBA changed
  + replace only XML members explicitly changed by the user
  + copy every other member from the original package unchanged
= candidate
```

For an untouched malformed `ppt/presentation.xml`, the candidate member bytes must be exactly equal to original bytes.

### 10.1 Never rewrite all XML parts

Do not use a strategy that serializes every XML snapshot back into the candidate.

### 10.2 Preserve ZIP member metadata where the project already promises this

Continue preserving relevant `ZipInfo` metadata/compression behavior according to the existing safety model. The important regression assertion is byte equality of the uncompressed member payload; preserve container metadata as well where existing code does so.

---

## 11. Candidate verification changes

This is the most important safety section.

The verifier must classify package parts into:

1. VBA part(s) intentionally changed,
2. XML parts intentionally changed,
3. untouched parts.

### 11.1 Untouched parts

For every untouched ZIP member:

```python
sha256(candidate.read(name)) == sha256(original.read(name))
```

must hold.

This includes malformed XML.

If `ppt/presentation.xml` was malformed/empty on open and the user only changed VBA, candidate verification must require exact byte equality for that part.

### 11.2 Changed XML parts

For each intentionally changed XML part:

- candidate contains the expected part,
- candidate bytes equal the exact encoded draft bytes that passed pre-save validation,
- candidate bytes parse as well-formed XML,
- no unrelated package member changed.

### 11.3 Do not require malformed untouched XML to parse

The verifier must not do this:

```python
for xml_part in all_xml_parts:
    assert parse(candidate.read(xml_part))
```

That recreates the original bug at a later stage.

### 11.4 VBA verification remains unchanged

Continue existing VBA project structural/module/body verification. This fix must not weaken VBA safety checks.

---

## 12. UI behavior

Locate the newer XML UI, expected under:

```text
vba_addin_editor.ui.xml_editor
```

### 12.1 XML part list

A malformed baseline part should still be visible, e.g.:

```text
⚠ ppt/presentation.xml
```

or with an equivalent status indicator.

### 12.2 Selecting malformed XML

If it can be decoded for display:

- show its text,
- editor is read-only,
- show a warning banner/status.

Suggested wording:

```text
This XML part was already malformed when the add-in was opened.
VBA Add-in Editor will preserve it exactly and will not edit it.
VBA editing and saving remain available.
```

If it cannot be decoded:

- do not insert fake replacement text into the draft,
- show a read-only diagnostic pane/message,
- preserve the bytes invisibly in the package model.

### 12.3 “Validate XML” button

For a malformed baseline/read-only part, the button may display the recorded baseline problem, but must not throw an exception or block access to the VBA tab.

For editable XML, it should validate the current editor contents.

### 12.4 Save state

A malformed untouched XML part must **not** disable Save Add-in / Save File when the only changes are VBA changes.

Invalid XML that the user has actually edited **must** disable/block save via the normal validation result.

### 12.5 Dirty tracking

Selecting a malformed part, viewing it, switching tabs, or clicking Validate must not mark it dirty.

---

## 13. Automated tests — required

Do not close this issue with only a manual smoke test. Add regression coverage at multiple levels.

### 13.1 Unit tests for XML codec

Add tests for at least:

```python
b""
b"   \r\n"
UTF-8 BOM only
b"<root>"                  # truncated
b"<root/>"                 # valid
valid XML with declaration
valid supported non-UTF8 encoding if supported by product
invalid declared encoding
```

Assertions:

- invalid inputs produce validation problems without crashing,
- valid inputs remain valid,
- encoding/decoding metadata is retained correctly.

### 13.2 Unit tests for package snapshot

Construct a ZIP containing:

```text
ppt/presentation.xml        -> empty or malformed bytes
ppt/other.xml               -> valid XML
ppt/vbaProject.bin          -> valid test VBA project
```

Assert:

```python
parts = adapter.snapshot_xml_parts(path)
```

returns successfully.

Assert the malformed part has equivalent state to:

```text
editable == False
well_formed == False
open_problem is not None
```

and the valid part remains editable.

### 13.3 Document open regression

Create or derive a test `.ppam` with a real parseable VBA project and deliberately malformed/empty outer XML.

Recommended deterministic fixture construction:

1. start from a known pyOpenVBA PowerPoint macro-enabled template or an existing deterministic test fixture,
2. write/copy it with `.ppam` extension using the app's PPAM compatibility layer,
3. replace `ppt/presentation.xml` with the exact malformed bytes needed for the regression,
4. leave `ppt/vbaProject.bin` intact.

Then:

```python
draft = DocumentService().open(path)
```

must succeed and expose VBA modules.

### 13.4 VBA-only save preservation regression

This is mandatory.

Test steps:

1. Open malformed-outer-XML PPAM.
2. Capture:

```python
before = zipfile.ZipFile(path).read("ppt/presentation.xml")
```

3. Edit only a standard VBA module.
4. Run the full real save pipeline on a disposable test copy.
5. Read the resulting member:

```python
after = zipfile.ZipFile(path).read("ppt/presentation.xml")
```

6. Assert:

```python
assert after == before
```

7. Reopen with `DocumentService` and assert the VBA edit landed.

This test catches both the open bug and accidental XML normalization during save.

### 13.5 Save/reload regression

Explicitly exercise the internal path that calls `_open_full_snapshot()` or equivalent after candidate creation/commit. The test must fail on the old implementation and pass after the adapter-level fix.

### 13.6 Invalid edited XML blocks save

Start from a valid editable XML part.

1. Change it to malformed XML.
2. Attempt save.
3. Assert save is blocked before original file replacement.
4. Assert original file bytes are unchanged.
5. Assert no candidate is committed.

This proves the change did not weaken validation for user-created errors.

### 13.7 Valid edited XML still saves

Edit a valid XML part to another valid XML document and verify:

- save succeeds,
- exact intended part changes,
- unrelated parts remain byte-identical,
- candidate reopens.

### 13.8 Verifier detects accidental mutation of malformed untouched XML

Create a candidate where an untouched malformed part differs by one byte. Candidate verification must fail with a message identifying unexpected non-VBA/non-edited package mutation.

This is critical because simply “ignoring malformed XML” without enforcing byte preservation would weaken the safety model.

### 13.9 GUI regression

If GUI tests already exist, add one that:

1. opens the malformed-XML PPAM,
2. confirms VBA module tree is populated,
3. selects malformed XML part,
4. confirms XML editor is read-only / warning visible,
5. returns to VBA tab,
6. edits VBA,
7. saves successfully.

Avoid pixel-based testing; assert widget state/model state.

---

## 14. Live regression using the exact failing PPAM — mandatory release gate

Synthetic tests are necessary but not sufficient.

Add a live test, gitignored fixture, or local qualification script for the exact PPAM that produced the screenshot.

Suggested marker:

```python
@pytest.mark.live
@pytest.mark.parametrize(...)
def test_existing_malformed_xml_ppam_can_open_edit_vba_save_and_reopen(...):
    ...
```

Required assertions:

1. `DocumentService.open()` succeeds.
2. At least the expected VBA module(s) can be read.
3. Baseline bytes of `ppt/presentation.xml` are captured.
4. VBA body is changed on a working copy.
5. Full in-place save pipeline succeeds.
6. Backup is created according to existing product contract.
7. Saved add-in reopens in the app.
8. VBA edit is present.
9. `ppt/presentation.xml` after save equals baseline bytes exactly.
10. No unrelated ZIP member changed except those the save pipeline explicitly allows.

Then manually verify the saved copy still opens in:

- Office RibbonX Editor,
- PowerPoint, if available/appropriate.

If PowerPoint shows a repair dialog, the issue is **not closed** even if Python tests pass.

---

## 15. Regression for XLAM too

The bug is package-policy related, not inherently PowerPoint-specific.

Add equivalent coverage for `.xlam` where practical:

- malformed unrelated XML does not prevent VBA open,
- VBA-only save preserves it byte-for-byte,
- changed invalid XML is still blocked.

This prevents the fix from being coded only for PPAM.

---

## 16. Error handling and logging

Improve diagnostics so future package quirks are actionable.

When a baseline XML part is malformed, log/retain at least:

```text
part name
byte length
SHA-256 (optional but useful)
decode status
validation problem
editable/read-only status
```

Do **not** dump full customer XML content into logs by default.

The user-facing Open dialog should not appear for a nonfatal per-part XML problem.

Fatal package errors should still be translated to a plain-language dialog.

---

## 17. Security/safety constraints

This fix must not introduce unsafe XML parsing behavior.

- Continue using the existing safe standard-library parser or the project's chosen safe parser.
- Do not enable external entity/network resolution.
- Do not attempt automatic XML repair.
- Do not execute VBA.
- Do not weaken password/signature save gates.
- Do not rewrite unrelated package members.
- Preserve the existing original-file fingerprint/external-change checks.
- Preserve mandatory backup and atomic replacement behavior.

---

## 18. Do not implement these tempting shortcuts

### Wrong: skip only `ppt/presentation.xml`

```python
if name == "ppt/presentation.xml":
    continue
```

Reason: another malformed XML part will fail next, and the XML tab loses visibility of a real package member.

### Wrong: treat empty XML as valid

```python
if not raw:
    return []
```

Reason: empty XML is not well-formed XML. It should remain marked malformed; it simply should not prevent VBA editing when untouched.

### Wrong: globally disable XML validation

Reason: user-edited invalid XML must still be blocked.

### Wrong: parse and reserialize malformed XML to “fix” it

Reason: changes the customer's package without explicit intent and may break Office behavior.

### Wrong: catch `Exception` around the entire open flow and ignore it

Reason: this can hide actual ZIP corruption, I/O failures, missing VBA, and other conditions where safe editing is impossible.

### Wrong: fix only the GUI

Reason: headless open, save verification, reload, and tests would retain inconsistent behavior.

### Wrong: fix only initial Open

Reason: `_open_full_snapshot()` / candidate reload can trigger the same failure after Save.

---

## 19. Suggested implementation sequence

Implement in this order so failures remain easy to localize:

1. **Reproduce and inspect exact PPAM.** Record raw `ppt/presentation.xml` state.
2. **Add failing regression tests first** for package snapshot + `DocumentService.open()`.
3. Extend XML part snapshot/domain model with per-part health/editability state.
4. Change `snapshot_xml_parts()` so per-part decode/parse failures are nonfatal.
5. Update XML UI to display malformed baseline parts read-only.
6. Add/adjust dirty tracking so viewing malformed parts never marks them changed.
7. Verify save writer only replaces explicitly changed XML parts.
8. Keep strict validation in `_validate_draft_parts()` for changed XML.
9. Update candidate verification to byte-compare every untouched part, malformed or not.
10. Add VBA-only save preservation test.
11. Add invalid-edited-XML save-block test.
12. Add post-save `_open_full_snapshot()` regression.
13. Run all existing unit/integration/GUI tests.
14. Run live test with exact failing PPAM.
15. Manually open resulting saved copy in RibbonX Editor and PowerPoint.
16. Rebuild packaged EXE and repeat the exact original GUI reproduction.
17. Update `HANDOFF.md`/README with the new package policy and qualification result.

---

## 20. Acceptance criteria — all must pass

Do not mark implementation complete until every applicable item below is checked.

- [ ] Exact original failing PPAM opens in VBA Add-in Editor without the package-XML fatal dialog.
- [ ] VBA tab/module tree is usable.
- [ ] `ppt/presentation.xml` is still represented in XML UI or diagnostics.
- [ ] Existing malformed part is clearly marked and read-only (recommended v1 behavior).
- [ ] Merely selecting/viewing malformed part does not mark document dirty.
- [ ] VBA-only edit can be saved.
- [ ] Mandatory backup is still produced.
- [ ] Saved PPAM reopens in VBA Add-in Editor.
- [ ] VBA edit is present after reopen.
- [ ] Original and saved `ppt/presentation.xml` payload bytes are exactly equal when XML was untouched.
- [ ] Every other untouched ZIP member remains byte-identical according to existing verifier policy.
- [ ] Invalid XML introduced by the user blocks save before commit.
- [ ] Valid XML edits still save successfully.
- [ ] Post-save verification/reload does not fail on baseline malformed XML.
- [ ] Existing VBA password/signature/lock/external-change safety tests still pass.
- [ ] Existing XLAM behavior does not regress.
- [ ] New synthetic PPAM regression tests pass.
- [ ] New live exact-PPAM regression passes.
- [ ] Saved copy opens in Office RibbonX Editor.
- [ ] Saved copy opens in PowerPoint without repair dialog, if PowerPoint is available for qualification.
- [ ] Packaged EXE is rebuilt from the fixed source and the original user workflow passes in the packaged build.

---

## 21. Definition of “fixed for sure” for this issue

For this bug, “fixed” does **not** mean merely suppressing the dialog.

The fix is proven only when all three properties hold simultaneously:

### A. Tolerance

A pre-existing malformed/empty outer-package XML part does not prevent VBA editing.

### B. Preservation

If the user does not edit that XML part, its bytes survive a VBA-only save exactly unchanged.

### C. Strictness for new edits

If the user edits XML and makes it invalid, the application refuses to save it.

The exact failing PPAM must demonstrate A + B in a live regression. Automated invalid-edit tests must demonstrate C.

---

## 22. Recommended comments to leave in code

At the nonfatal baseline XML handling point, add a concise rationale such as:

```python
# Existing OOXML packages can contain XML parts that our editor cannot parse
# even though Office/other tooling tolerates them. Baseline parse failure is
# therefore per-part state, not a document-open failure. Untouched parts are
# preserved byte-for-byte; only user-edited XML is required to validate.
```

At candidate verification:

```python
# Do not parse untouched baseline XML here. A malformed pre-existing part is
# allowed, but it must remain byte-identical. Changed XML is validated separately.
```

These comments are important because a future refactor may otherwise “simplify” the code back into eager global validation and reintroduce the bug.

---

## 23. Handoff note for the next agent

The current public repository state previously described PPAM support as essentially an extension-whitelist adaptation over pyOpenVBA, while real Office-authored PPAM qualification was still a release gate. The executable associated with this bug contains a newer XML editing layer beyond that public state.

Therefore:

1. identify the exact source revision/worktree used to build SHA-256 `892bcc...`,
2. implement this fix in that newer source,
3. do not accidentally discard the XML feature by reverting to public `main`,
4. merge/push the newer source and tests so the binary is reproducible from Git,
5. update the project handoff with the new commit SHA and packaged EXE SHA-256 after qualification.

The source-of-truth test for this bug is the exact real PPAM plus byte-preservation assertions, not merely a green synthetic test suite.
