# HANDOFF — VBA Add-in Editor project state

_Date: 2026-08-30. Implementation of
`VBA_Addin_Editor_XML_and_Context_Menu_Implementation_Plan.md` (3169 lines) is
complete (except the human live-Office gate); this document reflects the
verified current state. The earlier `VBA_Addin_Editor_Implementation_Plan.md`
(VBA core) remains the baseline reference._

## What exists

Single app in this directory (`vba-addin-editor/`), Windows-only, offline,
Tkinter/ttk GUI. Edits VBA source inside installed Excel `.xlam` /
PowerPoint `.ppam` add-ins **in place** (same path) with mandatory backup,
candidate verification, and `ReplaceFileW` commit. `.pptm` is a first-class
document type. `.ppam`/`.pptm` also expose an **XML editor** for existing
package parts (`.xml`, `.rels`, `[Content_Types].xml`) flowing through the same
safe-save transaction as VBA changes. Right-click Cut/Copy/Paste is shared
text-widget behavior in both editors. No COM, no Python install needed at
runtime.

| Area | Location |
|---|---|
| Entry / headless self-tests | `src/vba_addin_editor/__main__.py` (`--self-test`, `--self-roundtrip`), `app.py` |
| Domain (snapshot/draft/changes, XML part models) | `src/vba_addin_editor/domain/` |
| pyOpenVBA adapter (VBA seam) | `src/vba_addin_editor/adapters/pyopenvba_adapter.py` |
| OOXML package adapter (ZIP/XML seam) — **new** | `src/vba_addin_editor/adapters/ooxml_package_adapter.py` |
| XML text codec (decode/encode/validate) — **new** | `src/vba_addin_editor/adapters/xml_codec.py` |
| Codepage + header/body codec | `src/vba_addin_editor/adapters/source_codec.py` |
| Save pipeline (composite VBA/XML matrix) | `src/vba_addin_editor/services/save_service.py` |
| Open workflow (attaches XML parts + package safety) | `src/vba_addin_editor/services/document_service.py` |
| Import/export + backup restore | `src/vba_addin_editor/services/import_export_service.py` |
| ReplaceFileW / lock probe / process detect | `src/vba_addin_editor/platform/` |
| GUI (notebook: VBA + XML tabs) | `src/vba_addin_editor/ui/main_window.py`, `ui/code_editor.py` |
| XML editor — **new** | `src/vba_addin_editor/ui/xml_editor.py` |
| Shared right-click context menu — **new** | `src/vba_addin_editor/ui/text_context_menu.py` |
| Tests (125: 121 green + 4 live-skips) | `tests/` (unit + integration; GUI drives real Text widgets) |
| Live release gate (needs real fixtures) | `tests/integration/test_live_office_fixtures.py` |
| Packaging | `packaging/VBAAddinEditor.spec`, `scripts/build.ps1`, `scripts/test.ps1` |

Dependency pin: `pyopenvba==3.4.0` exact. Python 3.10.10 used for dev/build.
No new third-party dependency was added (zipfile/codecs/ElementTree only).

## Verified state

- `python -m pytest tests -q` → 121 passed, 4 skips (live fixtures absent).
  Coverage includes: the full XLAM/PPAM regression matrix; PPTM VBA-only,
  XML-only, and combined saves with payload isolation; XML codec round-trips
  (UTF-8/UTF-16 BOMs, declarations, newlines); OOXML adapter discovery,
  duplicate rejection, signature detection, and limits; composite candidate
  verification; failure paths leaving the original byte-identical; GUI XML +
  combined round-trips; context-menu behavior on both editors (28 cases).
- **Onefile release built and smoke-tested:**
  `dist\VBAAddinEditor.exe` + `dist\VBAAddinEditor.exe.sha256` (authoritative).
  SHA-256: `892BCCFF558843E5EEF8D1E83D6F76E11B87DB572B313D8E665B789EE3352239`.
  Packaged `--self-test` and `--self-roundtrip` exit 0 on `.xlam` and on a
  synthetic `.pptm` (`modules=1 xml_parts=36`).
- **Onedir build** `dist\VBAAddinEditor\` rebuilt fresh from final sources,
  also smoke-tested. Debugging artifact only, gitignored.
- `ruff check .` → clean.
- CLI: `--self-test` reports `modules=N xml_parts=M` and works for
  `.xlam/.ppam/.pptm`; `--self-roundtrip` runs a combined VBA+XML candidate
  cycle via `save_copy` and reopens the result (verified on synthetic
  PPTM and XLAM).

## Key decisions (why the code looks like this)

- **One composite transaction.** VBA candidate is built first (unchanged
  pyOpenVBA path), then the XML patch is applied to the disposable candidate
  (`candidate.xmlpatch.<ext>` → `replace`). XML-only saves never call
  pyOpenVBA mutation APIs; the package adapter rewrites the ZIP preserving
  entry order/metadata and copying `vbaProject.bin` bytes untouched.
- **Allowed-difference invariant:** `verify_candidate(...,
  allowed_non_vba_changes=<changed XML paths>)` accepts differences only on
  the exact paths in the change set; `verify_candidate_xml` additionally
  requires each changed part's bytes to equal `encode_xml(draft text)` exactly
  and compares the saved file against the **pre-save backup** post-commit.
- **XML is written as user-typed bytes**: editor text (LF) is encoded with the
  part's original codec + BOM + dominant newline; ElementTree is used only for
  well-formedness/root-shape validation; DOCTYPE/ENTITY are rejected.
- **Policy split by mutation type:** VBA signature confirmation only when VBA
  changes; password-protected VBA blocks only VBA mutations (XML-only saves
  are allowed and verified to keep `vbaProject.bin` byte-identical);
  OPC/package signature blocks XML edits in v1.
- **Draft name validation (Stage E0) and code-page validation gate only VBA
  mutations**, so pre-existing VBA naming oddities cannot block XML-only saves.
- **Rename replay / import CRLF / ambiguous-module rules unchanged** from the
  prior phase (see prior handoff decisions).
- Context menu uses Tk virtual events; `_paste` first replaces an existing
  selection explicitly because Tk's `<<Paste>>` binding does not remove a
  programmatic selection and inserts at the insert mark.

## Safety model (enforced, tested)

No-change → no write. Office running / locked / external fingerprint → save
blocked. Malformed or encoding-conflicting XML → `error`, original untouched.
Candidate fully verified (VBA layer + XML layer) before anything is replaced.
One save → one backup → one `ReplaceFileW`. Post-commit verification compares
the saved original against the backup; failure → `recovery_required` with
Restore Backup. Backup restore remains whole-file recovery.

## Remaining work

1. **Live Office release qualification** — **requires a human with real
   Office**; the agent cannot do this:
   - Drop authentic fixtures:
     `tests/fixtures/xlam/RealAddin.xlam`,
     `tests/fixtures/ppam/RealAddin.ppam` (with a `customUI` part containing
     `VBAAE_XML_ORIGINAL`),
     `tests/fixtures/pptm/RealPresentation.pptm` (same marker contract).
   - Run `pytest tests -m live -v` and complete the human PowerPoint gate:
     open saved file with **no repair dialog**, edited macro runs, Ribbon/XML
     marker reflects the edit, restore backup works.
2. Rebuild onefile/onedir (`scripts/build.ps1 [-Onefile]`), re-run packaged
   `--self-test`/`--self-roundtrip` on `.xlam`, `.ppam`, `.pptm`, record the
   new SHA-256.
3. Update README "Release gating" after the gate passes.

## Maintenance notes

- The OOXML package adapter is the only place ZIP/XML bytes are manipulated;
  GUI never touches package entries. Keep it that way.
- XML size limits live in `ooxml_package_adapter.py`
  (`MAX_XML_PART_BYTES`/`MAX_TOTAL_XML_BYTES`/`MAX_XML_PART_COUNT`).
- `ui/text_context_menu.py` exposes `prepare_menu()` + `menu_states()` so
  tests never need `tk_popup()` (which blocks in withdrawn windows).
- If pyopenvba ever bumps off 3.4.0: re-read installed `_host.py` save path
  and `vba.py` mutation methods first; the internals import in
  `pyopenvba_adapter.py` remains the only VBA upgrade seam.
- After adding new modules to `src/`, re-run `pip install -e .`.
- Windows quirk: some shells can't execute `dist\...exe` directly; run via
  `python -c "subprocess.run([...])"` or `cmd /c`.
- Synthetic fixtures (`tests/conftest.py` → `build_xlam`/`build_pptm`) are
  acceptable for parser/save-pipeline tests but are NOT a substitute for the
  live-Office gate.
