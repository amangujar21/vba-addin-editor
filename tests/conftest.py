"""Shared fixtures: synthetic XLAM built from the Excel-authored baked template."""

from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

import pytest
from pyopenvba import ExcelFile, PowerPointFile
from pyopenvba.vba import VBAModuleKind


def build_xlam(path: Path) -> Path:
    """Create a real-template XLAM with synthetic modules for round-trip tests."""
    with ExcelFile.create_new(path) as host:
        host.set_module(
            "Module1",
            'Public Const TEST_BUILD As String = "ORIGINAL"\r\n'
            "Public Sub Hello()\r\n    MsgBox TEST_BUILD\r\nEnd Sub\r\n",
        )
        host.save()
    return path


def build_xlam_with_class(path: Path, *, stream_name: str | None = None) -> Path:
    """Synthetic XLAM with Module1 plus an ordinary class OrdinaryClass."""
    with ExcelFile.create_new(path) as host:
        host.set_module(
            "Module1",
            'Public Const TEST_BUILD As String = "ORIGINAL"\r\n'
            "Public Sub Hello()\r\n    MsgBox TEST_BUILD\r\nEnd Sub\r\n",
        )
        project = host.vba_project()
        project.add_module(
            "OrdinaryClass",
            "Public X As Long\r\nPublic Sub Ping()\r\nEnd Sub\r\n",
            kind=VBAModuleKind.other,
            stream_name=stream_name,
        )
        host.save()
    return path


def build_xlam_with_form(path: Path) -> Path:
    """Synthetic XLAM with Module1 plus UserForm EntryForm (a frame holding a textbox, a button)."""
    with ExcelFile.create_new(path) as host:
        host.set_module(
            "Module1",
            'Public Const TEST_BUILD As String = "ORIGINAL"\r\n'
            "Public Sub Hello()\r\n    EntryForm.Show\r\nEnd Sub\r\n",
        )
        form = host.add_form("EntryForm", caption="Entry")
        form.add_control("Frame", "fraMain")
        form.add_control("TextBox", "txtName", container="fraMain")
        form.add_control("CommandButton", "cmdOK")
        form.control("cmdOK").set_property("Caption", "OK")
        host.set_module("EntryForm", "Private Sub cmdOK_Click()\r\n    Unload Me\r\nEnd Sub\r\n")
        host.save()
    return path


_SIG_REL = "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature"


def add_signature_parts(path: Path, vba_entry: str = "xl/vbaProject.bin") -> None:
    """Lay out a (fake) VBA signature in package parts, as Office stores one."""
    folder, _, name = vba_entry.rpartition("/")
    sig_part = f"{folder}/vbaProjectSignature.bin"
    rels_part = f"{folder}/_rels/{name}.rels"
    with zipfile.ZipFile(path) as package:
        members = {info.filename: package.read(info) for info in package.infolist()}
    assert rels_part not in members
    members[sig_part] = b"\x00" * 64
    members[rels_part] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="{_SIG_REL}" Target="vbaProjectSignature.bin"/>'
        "</Relationships>"
    ).encode()
    types = members["[Content_Types].xml"].decode("utf-8")
    override = (
        f'<Override PartName="/{sig_part}" '
        'ContentType="application/vnd.ms-office.vbaProjectSignature"/>'
    )
    members["[Content_Types].xml"] = types.replace("</Types>", override + "</Types>").encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
        for member, data in members.items():
            package.writestr(member, data)


RIBBON = """<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="tabTools" label="Tools">
        <group id="grpMain" label="Main">
          <button id="btnHello" label="Say Hello" size="large" onAction="Hello" />
          <!-- <button id="btnOld" onAction="Ghost" /> -->
          <button id="btnMissing" label="Broken" onAction="NoSuchMacro" />
        </group>
        <group id="grpMenus" label="Menus">
          <menu id="mnuMore" label="More">
            <toggleButton id="tglX" label="Toggle" onAction="OnToggle" getPressed="GetToggle" />
          </menu>
        </group>
        <group id="grpEmpty" label="Empty"></group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
"""


def build_xlam_with_ribbon(path: Path, xml: str = RIBBON) -> Path:
    build_xlam_with_class(path)
    with zipfile.ZipFile(path, "a") as package:
        package.writestr("customUI/customUI14.xml", xml.encode("utf-8"))
    return path


def build_pptm_with_class(path: Path) -> Path:
    with PowerPointFile.create_new(path) as host:
        host.set_module(
            "Module1",
            'Public Const TEST_BUILD As String = "ORIGINAL"\r\n'
            "Public Sub Hello()\r\n    MsgBox TEST_BUILD\r\nEnd Sub\r\n",
        )
        host.vba_project().add_module(
            "OrdinaryClass",
            "Public X As Long\r\n",
            kind=VBAModuleKind.other,
        )
        host.save()
    return path


@pytest.fixture()
def xlam_path(tmp_path: Path) -> Path:
    return build_xlam(tmp_path / "TestAddin.xlam")


@pytest.fixture()
def work_xlam(tmp_path: Path, xlam_path: Path) -> Path:
    """A disposable copy per test."""
    dest = tmp_path / "work" / "TestAddin.xlam"
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(xlam_path, dest)
    return dest


def build_pptm(path: Path) -> Path:
    """Create a synthetic PPTM with one module for VBA round-trip tests."""
    with PowerPointFile.create_new(path) as host:
        host.set_module(
            "Module1",
            'Public Const TEST_BUILD As String = "ORIGINAL"\r\n'
            "Public Sub Hello()\r\n    MsgBox TEST_BUILD\r\nEnd Sub\r\n",
        )
        host.save()
    return path


@pytest.fixture()
def pptm_path(tmp_path: Path) -> Path:
    return build_pptm(tmp_path / "TestPresentation.pptm")


@pytest.fixture()
def work_pptm(tmp_path: Path, pptm_path: Path) -> Path:
    """A disposable copy per test."""
    dest = tmp_path / "work" / "TestPresentation.pptm"
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(pptm_path, dest)
    return dest


@pytest.fixture()
def work_xlam_with_form(tmp_path: Path) -> Path:
    return build_xlam_with_form(tmp_path / "FormAddin.xlam")


@pytest.fixture
def work_xlam_with_class(tmp_path: Path) -> Path:
    dest = tmp_path / "work" / "ClassAddin.xlam"
    dest.parent.mkdir(parents=True, exist_ok=True)
    return build_xlam_with_class(dest)


@pytest.fixture()
def work_pptm_with_class(tmp_path: Path) -> Path:
    dest = tmp_path / "work" / "ClassDeck.pptm"
    dest.parent.mkdir(parents=True, exist_ok=True)
    return build_pptm_with_class(dest)


@pytest.fixture()
def work_ppam_with_class(tmp_path: Path) -> Path:
    dest = tmp_path / "work" / "ClassAddin.ppam"
    dest.parent.mkdir(parents=True, exist_ok=True)
    build_pptm_with_class(dest.with_suffix(".pptm"))
    shutil.copy2(dest.with_suffix(".pptm"), dest)
    return dest


@pytest.fixture()
def work_ppam(tmp_path: Path, pptm_path: Path) -> Path:
    """A disposable PPAM-compatible copy with a real parseable VBA project."""
    dest = tmp_path / "work" / "TestAddin.ppam"
    dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(pptm_path, dest)
    return dest


@pytest.fixture()
def replace_package_payload():
    """Return a helper that replaces one ZIP member while preserving metadata."""

    def replace(path: Path, member: str, replacement: bytes) -> None:
        temp = path.with_suffix(".malformed" + path.suffix)
        with zipfile.ZipFile(path) as src, zipfile.ZipFile(temp, "w") as dst:
            for info in src.infolist():
                data = replacement if info.filename == member else src.read(info.filename)
                dst.writestr(info, data)
            dst.comment = src.comment
        temp.replace(path)

    return replace
