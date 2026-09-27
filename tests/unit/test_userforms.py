"""UserForm support: read-only layout view, delete-only capability, verification."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from conftest import add_signature_parts, build_xlam_with_form
from helpers import edit_module1, make_service
from pyopenvba.cfb import CFB
from pyopenvba.vba import parse_project_stream

from vba_addin_editor.adapters.pyopenvba_adapter import (
    AdapterError,
    PyOpenVBAAdapter,
    vba_entry_for,
)
from vba_addin_editor.domain.capabilities import (
    DESIGNER,
    USERFORM,
    build_project_index,
    classify_component,
)
from vba_addin_editor.domain.document import ModuleDraft, new_module_id
from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.services.validation_service import validate_draft, validate_module_name


def _vba_cfb(path: Path) -> CFB:
    with zipfile.ZipFile(path) as package:
        return CFB.from_bytes(package.read(vba_entry_for(path)))


def _form(draft):
    return next(m for m in draft.modules if m.current_name == "EntryForm")


# -- classification ------------------------------------------------------------


def _index():
    return build_project_index(
        standard_modules=[], class_modules=[], document_modules=[],
        base_classes=["EntryForm"], lossy=False,
    )


def test_userform_is_deletable_but_not_renamable():
    caps = classify_component(
        logical_name="EntryForm", stream_name="EntryForm", dir_kind="other",
        is_read_only=False, index=_index(),
        designer_storages=frozenset({"entryform"}), form_storages=frozenset({"entryform"}),
    )
    assert caps.can_delete and not caps.can_rename
    assert caps.project_item_kind == "userform"
    assert caps.restriction_reason == USERFORM


@pytest.mark.parametrize(
    ("stream_name", "read_only", "form_storages"),
    [
        ("EntryForm", False, frozenset()),  # BaseClass but no form storage
        ("EntryForm", True, frozenset({"entryform"})),  # read-only
        ("OtherStream", False, frozenset({"entryform"})),  # stream/name mismatch
    ],
)
def test_userform_fails_closed_without_every_signal(stream_name, read_only, form_storages):
    caps = classify_component(
        logical_name="EntryForm", stream_name=stream_name, dir_kind="other",
        is_read_only=read_only, index=_index(),
        designer_storages=frozenset({"entryform"}), form_storages=form_storages,
    )
    assert not caps.can_delete and not caps.can_rename
    assert caps.restriction_reason == DESIGNER


# -- snapshot ------------------------------------------------------------------


def test_open_reads_form_layout(work_xlam_with_form: Path):
    draft = DocumentService().open(work_xlam_with_form)
    form = _form(draft)
    assert form.project_item_kind == "userform"
    assert form.can_delete and not form.can_rename
    assert "cmdOK_Click" in form.body

    (design,) = draft.baseline.forms
    assert design.name == "EntryForm" and design.problem is None
    assert dict(design.properties)["Caption"] == "Entry"
    assert [c.name for c in design.controls] == ["fraMain", "cmdOK"]
    frame = design.controls[0]
    assert frame.kind == "MSForms.Frame"
    assert [c.name for c in frame.children] == ["txtName"]
    assert dict(design.controls[1].properties)["Caption"] == "OK"
    assert design.control_count() == 3


def test_unreadable_layout_is_reported_not_fatal(work_xlam_with_form: Path, replace_package_payload):
    cfb = _vba_cfb(work_xlam_with_form)
    cfb.write_stream_at(["EntryForm"], "o", b"\x00")  # sites no longer reconcile
    replace_package_payload(work_xlam_with_form, "xl/vbaProject.bin", cfb.to_bytes())
    draft = DocumentService().open(work_xlam_with_form)
    (design,) = draft.baseline.forms
    assert design.problem and not design.controls
    assert _form(draft).body  # code still editable


# -- delete ----------------------------------------------------------------------


def test_delete_userform_removes_code_layout_and_declaration(work_xlam_with_form: Path):
    draft = DocumentService().open(work_xlam_with_form)
    _form(draft).is_deleted = True
    result = make_service().save_addin(draft)
    assert result.kind == "success", result

    reopened = DocumentService().open(work_xlam_with_form)
    assert "EntryForm" not in {m.current_name for m in reopened.modules}
    assert reopened.baseline.forms == ()
    cfb = _vba_cfb(work_xlam_with_form)
    assert "entryform" not in {s.casefold() for s in cfb.list_storages_at(())}
    parsed = parse_project_stream(cfb.get_stream("PROJECT"), code_page=reopened.baseline.code_page)
    assert "EntryForm" not in parsed.base_classes
    assert "EntryForm" not in cfb.get_stream("PROJECT").decode("cp1252")


def test_code_edit_keeps_form_layout_byte_identical(work_xlam_with_form: Path):
    before = _vba_cfb(work_xlam_with_form)
    draft = DocumentService().open(work_xlam_with_form)
    form = _form(draft)
    form.body = form.body.replace("Unload Me", "Me.Hide")
    edit_module1(draft)
    result = make_service().save_addin(draft)
    assert result.kind == "success", result

    after = _vba_cfb(work_xlam_with_form)
    from vba_addin_editor.adapters.pyopenvba_adapter import _storage_tree

    assert _storage_tree(before, ["EntryForm"]) == _storage_tree(after, ["EntryForm"])
    assert "Me.Hide" in _form(DocumentService().open(work_xlam_with_form)).body


def test_rename_userform_blocked(work_xlam_with_form: Path):
    draft = DocumentService().open(work_xlam_with_form)
    _form(draft).current_name = "RenamedForm"
    assert any("Cannot rename" in p for p in validate_draft(draft))
    before = work_xlam_with_form.read_bytes()
    result = make_service().save_addin(draft)
    assert result.kind == "error"
    assert work_xlam_with_form.read_bytes() == before


def test_deleted_form_name_cannot_be_reused_before_save(work_xlam_with_form: Path):
    draft = DocumentService().open(work_xlam_with_form)
    _form(draft).is_deleted = True
    assert "Save the deletion" in (validate_module_name("EntryForm", draft) or "")
    assert "Save the deletion" in (validate_module_name("entryform", draft) or "")

    draft.modules.append(
        ModuleDraft(
            id=new_module_id(), origin_name=None, current_name="EntryForm", body="",
            kind="class", pyopenvba_kind="other", is_new=True, is_deleted=False,
            destructive_ops_safe=True, can_delete=True, can_rename=True,
            project_item_kind="class",
        )
    )
    before = work_xlam_with_form.read_bytes()
    result = make_service().save_addin(draft)
    assert result.kind == "error"
    assert work_xlam_with_form.read_bytes() == before
    # The adapter refuses independently of validation.
    with pytest.raises(AdapterError, match="Save the deletion"):
        PyOpenVBAAdapter().build_candidate(
            work_xlam_with_form, draft, work_xlam_with_form.with_name("cand.xlam"),
            allow_signature_removal=False,
        )


def test_verification_rejects_changed_layout(tmp_path: Path, work_xlam_with_form: Path):
    draft = DocumentService().open(work_xlam_with_form)
    edit_module1(draft)
    candidate = tmp_path / "cand.xlam"
    adapter = PyOpenVBAAdapter()
    adapter.build_candidate(work_xlam_with_form, draft, candidate, allow_signature_removal=False)
    assert adapter.verify_candidate(work_xlam_with_form, candidate, draft).ok

    # Tamper with the form's layout inside the candidate.
    with zipfile.ZipFile(candidate) as package:
        cfb = CFB.from_bytes(package.read("xl/vbaProject.bin"))
    f_stream = bytearray(cfb.get_stream_at(["EntryForm"], "f"))
    f_stream[-1] ^= 0xFF
    cfb.write_stream_at(["EntryForm"], "f", bytes(f_stream))
    adapter.package_adapter.replace_members(candidate, {"xl/vbaProject.bin": cfb.to_bytes()})
    result = adapter.verify_candidate(work_xlam_with_form, candidate, draft)
    assert not result.ok
    assert any("UserForm layout changed" in p for p in result.problems)


# -- signature parts -------------------------------------------------------------


def test_signature_in_package_parts_is_detected_and_removed_on_confirm(tmp_path: Path):
    path = build_xlam_with_form(tmp_path / "Signed.xlam")
    add_signature_parts(path)
    draft = DocumentService().open(path)
    assert draft.baseline.safety.signature_present
    assert "legacy" in draft.baseline.safety.signature_kinds

    edit_module1(draft)
    result = make_service().save_addin(draft)
    assert result.kind == "needs_signature_confirmation"
    draft.signed_save_confirmed = True
    result = make_service().save_addin(draft)
    assert result.kind == "success", result

    with zipfile.ZipFile(path) as package:
        names = set(package.namelist())
        types = package.read("[Content_Types].xml").decode("utf-8")
    assert "xl/vbaProjectSignature.bin" not in names
    assert "xl/_rels/vbaProject.bin.rels" not in names
    assert "vbaProjectSignature" not in types
    reopened = DocumentService().open(path)
    assert not reopened.baseline.safety.signature_present
    assert "V2" in next(m for m in reopened.modules if m.current_name == "Module1").body
