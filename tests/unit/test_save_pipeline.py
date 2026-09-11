from __future__ import annotations

from vba_addin_editor.domain.document import ModuleDisplayKind, ModuleDraft, new_module_id
from vba_addin_editor.platform import paths
from vba_addin_editor.platform import windows_file_ops as wfo
from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.services.save_service import SaveService


def _adapter():
    from vba_addin_editor.adapters.pyopenvba_adapter import PyOpenVBAAdapter

    return PyOpenVBAAdapter()


def make_service(**overrides):
    base = {
        "adapter": _adapter(),
        "process_probe": lambda p: False,
        "exclusive_probe": lambda p: None,
        "commit": None,
        "progress": None,
    }
    base.update(overrides)
    return SaveService(**base)


def edit_module1(draft) -> None:
    m1 = next(m for m in draft.modules if m.current_name == "Module1")
    m1.body = m1.body.replace("ORIGINAL", "V2")


def test_no_change_save_does_not_write(work_xlam):
    draft = DocumentService().open(work_xlam)
    before = paths.fingerprint(work_xlam)
    result = make_service().save_addin(draft)
    assert result.kind == "no_changes"
    assert paths.fingerprint(work_xlam) == before


def test_edit_save_replaces_in_place_with_backup(work_xlam):
    draft = DocumentService().open(work_xlam)
    before = work_xlam.read_bytes()
    edit_module1(draft)
    result = make_service().save_addin(draft)
    assert result.kind == "success", result
    assert result.backup_path is not None and result.backup_path.exists()
    assert result.backup_path.read_bytes() == before
    # Same path now holds the edit.
    reopened = DocumentService().open(work_xlam)
    m1 = next(m for m in reopened.modules if m.current_name == "Module1")
    assert "V2" in m1.body and "ORIGINAL" not in m1.body
    # No leftover candidates.
    assert not list(work_xlam.parent.glob(".*candidate*"))
    # Non-VBA entries preserved implicitly by verifier; assert clean draft state.
    assert not draft.is_dirty()


def test_save_blocked_when_office_running(work_xlam):
    draft = DocumentService().open(work_xlam)
    edit_module1(draft)
    before = work_xlam.read_bytes()
    result = make_service(process_probe=lambda p: True).save_addin(draft)
    assert result.kind == "blocked" and result.reason == "office_running"
    assert work_xlam.read_bytes() == before


def test_save_blocked_when_locked(work_xlam):
    draft = DocumentService().open(work_xlam)
    edit_module1(draft)
    result = make_service(exclusive_probe=lambda p: "in use").save_addin(draft)
    assert result.kind == "blocked" and result.reason == "file_in_use"


def test_save_blocked_on_external_change(work_xlam):
    draft = DocumentService().open(work_xlam)
    edit_module1(draft)
    work_xlam.write_bytes(work_xlam.read_bytes() + b"\x00")
    result = make_service().save_addin(draft)
    assert result.kind == "blocked" and result.reason == "external_change"


def test_candidate_failure_leaves_original_untouched(work_xlam):
    draft = DocumentService().open(work_xlam)
    before = work_xlam.read_bytes()
    draft.modules.append(
        ModuleDraft(
            id=new_module_id(),
            origin_name=None,
            current_name="Bad Name!",  # adapter raises during replay
            body="Sub X()\r\nEnd Sub\r\n",
            kind=ModuleDisplayKind.STANDARD,
            pyopenvba_kind="standard",
            is_new=True,
            is_deleted=False,
            destructive_ops_safe=True,
        )
    )
    result = make_service().save_addin(draft)
    assert result.kind == "error"
    assert work_xlam.read_bytes() == before
    assert list(work_xlam.parent.glob(".*candidate*")) == []


def test_commit_failure_reports_error_and_keeps_original(work_xlam):
    def failing_commit(replaced, replacement, backup):
        raise wfo.FileOpsError("boom", win32_error=5)

    draft = DocumentService().open(work_xlam)
    edit_module1(draft)
    before = work_xlam.read_bytes()
    result = make_service(commit=failing_commit).save_addin(draft)
    assert result.kind == "error" and result.reason == "commit_failed"
    assert work_xlam.read_bytes() == before
    assert result.backup_path is None


def test_add_rename_delete_replay(work_xlam):
    draft = DocumentService().open(work_xlam)
    edit_module1(draft)
    m1 = next(m for m in draft.modules if m.current_name == "Module1")
    m1.current_name = "Renamed1"
    new = ModuleDraft(
        id=new_module_id(),
        origin_name=None,
        current_name="AddedModule",
        body="Public Sub FromNew()\r\nEnd Sub\r\n",
        kind=ModuleDisplayKind.STANDARD,
        pyopenvba_kind="standard",
        is_new=True,
        is_deleted=False,
        destructive_ops_safe=True,
    )
    draft.modules.append(new)
    result = make_service().save_addin(draft)
    assert result.kind == "success", result
    reopened = DocumentService().open(work_xlam)
    names = {m.current_name for m in reopened.modules}
    assert "Renamed1" in names and "AddedModule" in names
    assert "Module1" not in names
    assert "Sheet1" in names
    m = next(m for m in reopened.modules if m.current_name == "Renamed1")
    assert "V2" in m.body


def test_rename_chain_a_b_c(work_xlam):
    draft = DocumentService().open(work_xlam)
    m1 = next(m for m in draft.modules if m.current_name == "Module1")
    m1.current_name = "StepB"
    # Now rename again in the same session: StepB -> StepC (pending chain).
    m1.current_name = "StepC"
    result = make_service().save_addin(draft)
    assert result.kind == "success", result
    reopened = DocumentService().open(work_xlam)
    assert {m.current_name for m in reopened.modules} >= {"StepC"}
    assert "Module1" not in {m.current_name for m in reopened.modules}


def test_delete_then_readd_same_name(work_xlam):
    draft = DocumentService().open(work_xlam)
    m1 = next(m for m in draft.modules if m.current_name == "Module1")
    m1.is_deleted = True
    draft.modules.append(
        ModuleDraft(
            id=new_module_id(),
            origin_name=None,
            current_name="Module1",
            body="Public Const TEST_BUILD As String = \"REBORN\"\r\n",
            kind=ModuleDisplayKind.STANDARD,
            pyopenvba_kind="standard",
            is_new=True,
            is_deleted=False,
            destructive_ops_safe=True,
        )
    )
    result = make_service().save_addin(draft)
    assert result.kind == "success", result
    reopened = DocumentService().open(work_xlam)
    m = next(m for m in reopened.modules if m.current_name == "Module1")
    assert "REBORN" in m.body


def test_signature_confirmation_required(work_xlam):
    draft = DocumentService().open(work_xlam)
    edit_module1(draft)
    from dataclasses import replace

    draft.baseline = replace(
        draft.baseline,
        safety=replace(draft.baseline.safety, signature_present=True, signature_kinds=("legacy",)),
    )
    edit_module1(draft)
    result = make_service().save_addin(draft)
    assert result.kind == "needs_signature_confirmation"
    draft.signed_save_confirmed = True
    result = make_service().save_addin(draft)
    assert result.kind == "success", result


def test_save_copy_does_not_touch_original(work_xlam, tmp_path):
    draft = DocumentService().open(work_xlam)
    edit_module1(draft)
    dest = tmp_path / "copy.xlam"
    result = make_service().save_copy(draft, dest)
    assert result.kind == "success", result
    reopened = DocumentService().open(dest)
    m1 = next(m for m in reopened.modules if m.current_name == "Module1")
    assert "V2" in m1.body


# -- PPTM composite pipeline (XML plan 12, 23.5) -----------------------------

XML_MARKER = "<!--VBAAE_XML_EDITED-->"


def _edit_core_xml(draft) -> str:
    part = draft.xml_part_by_path("docProps/core.xml")
    assert part is not None
    part.text = part.text + "\n" + XML_MARKER
    return part.path


def _payloads(path):
    import zipfile

    with zipfile.ZipFile(path) as zf:
        return {i.filename: zf.read(i.filename) for i in zf.infolist()}


def test_pptm_xml_only_save(work_pptm):
    draft = DocumentService().open(work_pptm)
    before = _payloads(work_pptm)
    changed = _edit_core_xml(draft)
    result = make_service().save_addin(draft)
    assert result.kind == "success", result
    backup = result.backup_path
    assert backup is not None and backup.exists()
    # Backup holds the pre-save bytes.
    assert _payloads(backup) == before
    after = _payloads(work_pptm)
    # VBA project payload is byte-identical on an XML-only save.
    assert after["ppt/vbaProject.bin"] == before["ppt/vbaProject.bin"]
    # All non-changed payloads identical; changed part differs.
    for name, data in before.items():
        if name == changed:
            assert after[name] != data
        else:
            assert after[name] == data, name
    # Reopen shows the XML change and a clean draft.
    reopened = DocumentService().open(work_pptm)
    assert XML_MARKER in reopened.xml_part_by_path(changed).text
    assert not reopened.is_dirty()


def test_pptm_combined_vba_and_xml_save(work_pptm):
    draft = DocumentService().open(work_pptm)
    before = _payloads(work_pptm)
    edit_module1(draft)
    changed = _edit_core_xml(draft)
    result = make_service().save_addin(draft)
    assert result.kind == "success", result
    assert result.backup_path.exists()
    after = _payloads(work_pptm)
    assert after["ppt/vbaProject.bin"] != before["ppt/vbaProject.bin"]
    assert after[changed] != before[changed]
    for name, data in before.items():
        if name in ("ppt/vbaProject.bin", changed):
            continue
        assert after[name] == data, name
    reopened = DocumentService().open(work_pptm)
    m1 = next(m for m in reopened.modules if m.current_name == "Module1")
    assert "V2" in m1.body
    assert XML_MARKER in reopened.xml_part_by_path(changed).text
    assert not reopened.is_dirty()


def test_pptm_vba_only_save_keeps_xml_payloads(work_pptm):
    draft = DocumentService().open(work_pptm)
    before = _payloads(work_pptm)
    edit_module1(draft)
    result = make_service().save_addin(draft)
    assert result.kind == "success", result
    after = _payloads(work_pptm)
    assert after["ppt/vbaProject.bin"] != before["ppt/vbaProject.bin"]
    for name, data in before.items():
        if name == "ppt/vbaProject.bin":
            continue
        assert after[name] == data, name


def test_pptm_vba_only_save_preserves_malformed_baseline_xml(
    work_pptm, replace_package_payload
):
    malformed_path = "ppt/presentation.xml"
    replace_package_payload(work_pptm, malformed_path, b"")

    draft = DocumentService().open(work_pptm)
    malformed = draft.xml_part_by_path(malformed_path)
    assert malformed is not None
    assert not malformed.editable
    assert malformed.open_problem
    assert not draft.is_dirty()
    before = _payloads(work_pptm)

    edit_module1(draft)
    result = make_service().save_addin(draft)

    assert result.kind == "success", result
    assert result.backup_path is not None and result.backup_path.exists()
    after = _payloads(work_pptm)
    assert after[malformed_path] == before[malformed_path] == b""
    for name, data in before.items():
        if name != "ppt/vbaProject.bin":
            assert after[name] == data, name

    reopened = DocumentService().open(work_pptm)
    module = next(m for m in reopened.modules if m.current_name == "Module1")
    assert "V2" in module.body
    assert not reopened.xml_part_by_path(malformed_path).editable
    assert not reopened.is_dirty()


def test_ppam_vba_only_save_preserves_malformed_baseline_xml(
    work_ppam, replace_package_payload
):
    malformed_path = "ppt/presentation.xml"
    replace_package_payload(work_ppam, malformed_path, b"")
    before = _payloads(work_ppam)

    draft = DocumentService().open(work_ppam)
    assert draft.baseline.extension == ".ppam"
    assert draft.modules
    assert not draft.xml_part_by_path(malformed_path).editable
    edit_module1(draft)

    result = make_service().save_addin(draft)

    assert result.kind == "success", result
    assert result.backup_path is not None and result.backup_path.exists()
    after = _payloads(work_ppam)
    assert after[malformed_path] == before[malformed_path] == b""
    reopened = DocumentService().open(work_ppam)
    module = next(m for m in reopened.modules if m.current_name == "Module1")
    assert "V2" in module.body
    assert not reopened.xml_part_by_path(malformed_path).editable


def test_xlam_vba_only_save_tolerates_malformed_outer_xml(
    work_xlam, replace_package_payload
):
    malformed_path = "xl/workbook.xml"
    replace_package_payload(work_xlam, malformed_path, b"")
    before = _payloads(work_xlam)

    draft = DocumentService().open(work_xlam)
    edit_module1(draft)
    result = make_service().save_addin(draft)

    assert result.kind == "success", result
    after = _payloads(work_xlam)
    assert after[malformed_path] == before[malformed_path] == b""
    reopened = DocumentService().open(work_xlam)
    module = next(m for m in reopened.modules if m.current_name == "Module1")
    assert "V2" in module.body


def test_pptm_malformed_xml_leaves_original_untouched(work_pptm):
    draft = DocumentService().open(work_pptm)
    before = work_pptm.read_bytes()
    part = draft.xml_part_by_path("docProps/core.xml")
    part.text = part.text + "\n<unclosed"
    result = make_service().save_addin(draft)
    assert result.kind == "error"
    assert result.reason == "invalid_xml"
    assert work_pptm.read_bytes() == before
    assert not list(work_pptm.parent.glob(".*vbaae-candidate-*"))
    assert not list(
        work_pptm.parent.glob(f"{work_pptm.stem} - backup *{work_pptm.suffix}")
    )


def test_pptm_xml_encoding_conflict_blocked(work_pptm):
    draft = DocumentService().open(work_pptm)
    before = work_pptm.read_bytes()
    part = draft.xml_part_by_path("docProps/core.xml")
    part.text = part.text.replace(
        'encoding="UTF-8"', 'encoding="UTF-16"', 1
    ) if 'encoding="UTF-8"' in part.text else part.text + '<?xml encoding="UTF-16"?>'
    result = make_service().save_addin(draft)
    assert result.kind == "error"
    assert work_pptm.read_bytes() == before


def test_pptm_external_change_blocks_xml_save(work_pptm):
    draft = DocumentService().open(work_pptm)
    before = work_pptm.read_bytes()
    _edit_core_xml(draft)
    work_pptm.write_bytes(before + b"tampered")
    result = make_service().save_addin(draft)
    assert result.kind == "blocked" and result.reason == "external_change"


def test_pptm_office_running_blocks_xml_save(work_pptm):
    draft = DocumentService().open(work_pptm)
    before = work_pptm.read_bytes()
    _edit_core_xml(draft)
    result = make_service(process_probe=lambda p: True).save_addin(draft)
    assert result.kind == "blocked" and result.reason == "office_running"
    assert work_pptm.read_bytes() == before


def test_pptm_package_rewrite_failure_leaves_original(work_pptm):
    draft = DocumentService().open(work_pptm)
    before = work_pptm.read_bytes()
    _edit_core_xml(draft)

    class ExplodingAdapter:
        def validate_draft_part(self, part):
            return ()

        def write_xml_candidate(self, *a, **k):
            raise OSError("disk exploded")

    svc = make_service()
    svc.package_adapter = ExplodingAdapter()
    result = svc.save_addin(draft)
    assert result.kind == "error"
    assert work_pptm.read_bytes() == before


def test_package_signed_xml_edit_blocked(work_pptm):
    draft = DocumentService().open(work_pptm)
    from dataclasses import replace

    from vba_addin_editor.domain.document import PackageSafetyInfo

    draft.baseline = replace(
        draft.baseline,
        package_safety=PackageSafetyInfo(
            opc_signature_present=True,
            signature_part_names=("_xmlsignatures/sig1.xml",),
        ),
    )
    before = work_pptm.read_bytes()
    _edit_core_xml(draft)
    result = make_service().save_addin(draft)
    assert result.kind == "blocked" and result.reason == "package_signed"
    assert work_pptm.read_bytes() == before


def test_xml_only_save_allowed_on_password_protected_vba(work_pptm):
    """XML-only save must not rewrite the protected VBA project (plan 12.4)."""
    draft = DocumentService().open(work_pptm)
    from dataclasses import replace


    draft.baseline = replace(
        draft.baseline,
        safety=replace(draft.baseline.safety, password_protected=True),
    )
    before = _payloads(work_pptm)
    _edit_core_xml(draft)
    result = make_service().save_addin(draft)
    assert result.kind == "success", result
    assert _payloads(work_pptm)["ppt/vbaProject.bin"] == before["ppt/vbaProject.bin"]


def test_vba_change_on_password_protected_project_blocked(work_pptm):
    draft = DocumentService().open(work_pptm)
    from dataclasses import replace


    draft.baseline = replace(
        draft.baseline,
        safety=replace(draft.baseline.safety, password_protected=True),
    )
    edit_module1(draft)
    result = make_service().save_addin(draft)
    assert result.kind == "blocked" and result.reason == "password_protected"


def test_pptm_post_commit_xml_verification_failure_triggers_recovery(work_pptm):
    draft = DocumentService().open(work_pptm)
    _edit_core_xml(draft)

    calls = {"n": 0}
    real_verify = SaveService._verify_candidate

    def flaky(self, reference, candidate, d, changes):
        calls["n"] += 1
        if calls["n"] == 2:  # post-commit check
            return False, ("injected post-commit XML mismatch",)
        return real_verify(self, reference, candidate, d, changes)

    svc = make_service()
    svc._verify_candidate = flaky.__get__(svc)  # type: ignore[method-assign]
    result = svc.save_addin(draft)
    assert result.kind == "recovery_required"
    assert result.backup_path.exists()


def test_pptm_save_copy_extension_mismatch_rejected(work_pptm, tmp_path):
    draft = DocumentService().open(work_pptm)
    _edit_core_xml(draft)
    result = make_service().save_copy(draft, tmp_path / "out.ppam")
    assert result.kind == "error" and result.reason == "unsupported_extension"


def test_pptm_save_copy_success(work_pptm, tmp_path):
    draft = DocumentService().open(work_pptm)
    _edit_core_xml(draft)
    dest = tmp_path / "copy.pptm"
    result = make_service().save_copy(draft, dest)
    assert result.kind == "success", result
    reopened = DocumentService().open(dest)
    assert XML_MARKER in reopened.xml_part_by_path("docProps/core.xml").text
