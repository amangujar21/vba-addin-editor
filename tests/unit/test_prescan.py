"""Background pre-read: antivirus scans a saved file before Office asks for it."""

from __future__ import annotations

import threading
from pathlib import Path

from vba_addin_editor.platform import prescan
from vba_addin_editor.platform import windows_file_ops as wfo


def test_open_for_read_shares_the_file_and_reports_missing(tmp_path: Path):
    path = tmp_path / "Addin.xlam"
    assert prescan.open_for_read(path) is False
    path.write_bytes(b"data")
    assert prescan.open_for_read(path) is True
    assert wfo.probe_exclusive_access(path).ok  # handle closed again


def test_state_goes_from_scanning_to_ready(tmp_path: Path):
    release = threading.Event()
    opened: list[Path] = []

    def opener(path: Path) -> bool:
        opened.append(path)
        release.wait(5)
        return True

    scan = prescan.Prescan(opener)
    assert (scan.state, scan.busy) == (prescan.IDLE, False)
    scan.start(tmp_path / "Addin.xlam")
    assert (scan.state, scan.busy) == (prescan.SCANNING, True)
    release.set()
    scan.settle()
    assert (scan.state, scan.busy) == (prescan.READY, False)
    assert opened == [tmp_path / "Addin.xlam"]


def test_a_file_that_cannot_be_opened_is_failed_not_ready(tmp_path: Path):
    scan = prescan.Prescan()
    scan.start(tmp_path / "missing.xlam")
    scan.settle()
    assert scan.state == prescan.FAILED

    def broken(path: Path) -> bool:
        raise OSError("gone")

    scan = prescan.Prescan(broken)
    scan.start(tmp_path / "missing.xlam")
    scan.settle()
    assert scan.state == prescan.FAILED


def test_settle_waits_for_the_open_handle(tmp_path: Path):
    release = threading.Event()
    scan = prescan.Prescan(lambda path: release.wait(5))
    scan.start(tmp_path / "Addin.xlam")
    threading.Timer(0.1, release.set).start()
    scan.settle()
    assert not scan.busy and release.is_set()


def test_starting_again_settles_the_previous_pre_read(tmp_path: Path):
    release = threading.Event()
    order: list[str] = []

    def opener(path: Path) -> bool:
        release.wait(5)
        order.append(path.name)
        return True

    scan = prescan.Prescan(opener)
    scan.start(tmp_path / "one.xlam")
    threading.Timer(0.1, release.set).start()
    scan.start(tmp_path / "two.xlam")
    scan.settle()
    assert order == ["one.xlam", "two.xlam"]
    assert scan.path == tmp_path / "two.xlam"
