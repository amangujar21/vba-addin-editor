"""Shared fixtures: synthetic XLAM built from the Excel-authored baked template."""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest
from pyopenvba import ExcelFile, PowerPointFile


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
