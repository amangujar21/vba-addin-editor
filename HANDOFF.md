# HANDOFF — VBA Add-in Editor project state

_Date: 2026-08-30. Implementation of
`VBA_Addin_Editor_XML_and_Context_Menu_Implementation_Plan.md` (3169 lines) is
complete, including the source fix specified by
`XML_MALFORMED_PART_FIX_HANDOFF.md`. Automated qualification is complete;
the exact-customer-file and human live-Office gates remain. The earlier
`VBA_Addin_Editor_Implementation_Plan.md` (VBA core) remains the baseline
reference._

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
| Tests (149: 144 green + 5 live-skips) | `tests/` (unit + integration; GUI drives real Text widgets) |
| Live release gate (needs real fixtures) | `tests/integration/test_live_office_fixtures.py` |
| Packaging | `packaging/VBAAddinEditor.spec`, `scripts/build.ps1`, `scripts/test.ps1` |

Dependency pin: `pyopenvba==3.4.0` exact. Python 3.10.10 used for dev/build.
No new third-party dependency was added (zipfile/codecs/ElementTree only).

## Verified state

- `python -m pytest tests -q` → 144 passed, 5 skips (live fixtures absent).
  Coverage includes: the full XLAM/PPAM regression matrix; PPTM VBA-only,
  XML-only, and combined saves with payload isolation; XML codec round-trips
  (UTF-8/UTF-16 BOMs, declarations, newlines); OOXML adapter discovery,
  duplicate rejection, signature detection, and limits; composite candidate
  verification; failure paths leaving the original byte-identical; GUI XML +
  combined round-trips; context-menu behavior on both editors (28 cases).
  Malformed-baseline regressions cover empty/malformed/undecodable parts,
  synthetic PPAM open/VBA-save/reopen, byte preservation, changed-XML
  validation, untouched-part tamper detection, XLAM tolerance, and the real Tk
  XML read-only/VBA-save flow.
- **Onefile release:**
  `dist\VBAAddinEditor.exe` + `dist\VBAAddinEditor.exe.sha256` (authoritative).
  SHA-256: `840647C7464805322C918F08776CA031609F09D7B42B2235FA73AB7C9DC9C32A`.
  Rebuilt after malformed-part commits. The scripted XLAM `--self-test` and
  `--self-roundtrip` exit 0. Additional packaged smoke tests exit 0 for both
  modes on a synthetic PPTM and a synthetic PPAM with empty
  `ppt/presentation.xml`; the PPAM member remains byte-identical (`b""`).
- **Onedir build** `dist\VBAAddinEditor\` predates the malformed-part rebuild.
  It is a debugging artifact only; rebuild it before using it for diagnosis.
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
- **Baseline XML tolerance:** an existing empty, malformed, or undecodable XML
  member is per-part diagnostic state instead of a fatal Open error. It stays
  visible but read-only and is never reconstructed. VBA editing/saving remains
  available; every untouched malformed member must remain byte-identical, and
  only XML in the explicit change set is parsed during candidate verification.
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
blocked. Pre-existing malformed XML is tolerated only while untouched and is
preserved byte-for-byte; invalid or encoding-conflicting XML introduced by the
editor → `error`, no candidate/backup/commit, original untouched. Candidate is
fully verified (VBA + XML) before replacement. One save → one backup → one
`ReplaceFileW`. Post-commit verification compares the saved original against
the backup; failure → `recovery_required` with Restore Backup.

## Malformed-part fix state

Implemented and reviewed on `main`:

- `5993412` — tolerate malformed baseline OOXML parts;
- `167f2b7` — strengthen malformed OOXML regression coverage;
- `63030fa` — finish malformed PPAM review refinements.

The two-axis review found no remaining standards violation or automated spec
gap. Requirements remain in `XML_MALFORMED_PART_FIX_HANDOFF.md`; this handoff
records only current status. The optional exact-file gate is
`test_exact_malformed_ppam_vba_save_preserves_baseline_part` in
`tests/integration/test_live_office_fixtures.py`.

## Release provenance

Decoded-CMG protection-state handling is committed as `4fd9434`; malformed or
unexpected CMG data fails closed. The onefile EXE above was rebuilt from the
app source at that commit after the malformed-part commits. The subsequent
handoff/specification commit changes documentation only, so the packaged app
source is reproducible from Git.

## Remaining work

1. **Live Office release qualification** — **requires a human with real
   Office**; the agent cannot do this:
   - Drop authentic fixtures:
     `tests/fixtures/xlam/RealAddin.xlam`,
     `tests/fixtures/ppam/RealAddin.ppam` (with a `customUI` part containing
     `VBAAE_XML_ORIGINAL`),
     `tests/fixtures/pptm/RealPresentation.pptm` (same marker contract), and
     the exact reproducer at
     `tests/fixtures/live/ppam/ExistingMalformedPart.ppam`.
   - Run `pytest tests -m live -v` and complete the human PowerPoint gate:
     open the saved malformed-part PPAM in RibbonX Editor and PowerPoint with
     **no repair dialog**, confirm the VBA edit, and confirm
     `ppt/presentation.xml` is byte-identical. For generic fixtures, also
     confirm the macro runs, Ribbon/XML marker changed, and backup restore works.
2. Rebuild the debugging onedir artifact if needed, then repeat the exact
   malformed-PPAM GUI workflow in the onefile EXE.
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
- Never reintroduce eager parse-all verification. Baseline parse failure is
  per-part state; untouched bytes are the invariant, while changed XML remains
  strict.

## Suggested skills

- Use `diagnose` if the exact customer PPAM still fails Open, Save, reload, or
  packaged execution; reproduce against the untouched fixture first.
- Use `code-review` for future changes to CMG protection-state handling.
- Use `implement` for the remaining live qualification/reproducibility work,
  and `handoff` again after the Office and packaged-EXE gates are recorded.
