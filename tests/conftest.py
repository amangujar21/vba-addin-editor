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
