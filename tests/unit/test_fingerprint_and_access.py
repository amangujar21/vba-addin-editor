from __future__ import annotations

import os
from pathlib import Path

from vba_addin_editor.domain.document import FileFingerprint
from vba_addin_editor.platform import paths
from vba_addin_editor.platform import windows_file_ops as wfo
from vba_addin_editor.platform import windows_processes as wp


def test_timestamp_only_change_is_not_content_conflict(tmp_path: Path):
    target = tmp_path / "a.xlam"
    target.write_bytes(b"same-bytes")
    fp = paths.fingerprint(target)
    os.utime(target, (1_700_000_000, 1_700_000_000))
    assert paths.fingerprint_matches(target, fp)
    assert paths.fingerprint_status(target, fp) == "match"


def test_same_size_content_change_is_conflict(tmp_path: Path):
    target = tmp_path / "a.xlam"
    target.write_bytes(b"AAAA")
    fp = paths.fingerprint(target)
    target.write_bytes(b"BBBB")
    os.utime(target, (fp.mtime_ns / 1_000_000_000, fp.mtime_ns / 1_000_000_000))
    assert paths.fingerprint_status(target, fp) == "external_change"


def test_missing_source_status(tmp_path: Path):
    missing = tmp_path / "gone.xlam"
    fp = FileFingerprint(sha256="abc", size=1, mtime_ns=0)
    assert paths.fingerprint_status(missing, fp) == "source_missing"


def test_exclusive_probe_distinguishes_sharing(tmp_path: Path):
    target = tmp_path / "b.xlam"
    target.write_bytes(b"x")
    assert wfo.probe_exclusive_access(target).ok
    with open(target, "rb"):
        probe = wfo.probe_exclusive_access(target)
        assert not probe.ok
        assert probe.reason == "file_in_use"
        assert probe.win32_error in {
            wfo.ERROR_SHARING_VIOLATION,
            wfo.ERROR_LOCK_VIOLATION,
        }


def test_missing_file_probe_is_not_in_use(tmp_path: Path):
    probe = wfo.probe_exclusive_access(tmp_path / "missing.xlam")
    assert not probe.ok
    assert probe.reason == "source_missing"


def test_host_label_is_host_specific(tmp_path: Path):
    assert wp.corresponding_host_label(tmp_path / "a.xlam") == "Excel"
    assert wp.corresponding_host_label(tmp_path / "a.ppam") == "PowerPoint"
    assert wp.corresponding_host_label(tmp_path / "a.pptm") == "PowerPoint"


def test_same_file_identity_is_case_insensitive(tmp_path: Path):
    target = tmp_path / "AddIn.xlam"
    target.write_bytes(b"x")
    assert paths.paths_are_same_file(target, tmp_path / "addin.xlam")


def _hold_in_child(target: Path, seconds: float):
    """A separate process that keeps ``target`` open for ``seconds``, then exits."""
    import subprocess
    import sys
    import time

    code = (
        "import sys, time\n"
        "f = open(sys.argv[1], 'rb')\n"
        "print('held', flush=True)\n"
        "time.sleep(float(sys.argv[2]))\n"
    )
    child = subprocess.Popen(
        [sys.executable, "-c", code, str(target), str(seconds)],
        stdout=subprocess.PIPE, text=True,
    )
    assert child.stdout is not None and child.stdout.readline().strip() == "held"
    time.sleep(0.05)
    return child


def test_in_use_message_names_the_program_holding_the_file(tmp_path: Path):
    target = tmp_path / "Held.xlam"
    target.write_bytes(b"x")
    child = _hold_in_child(target, 30)
    try:
        assert wfo.file_holders(target), "Restart Manager should see the child process"
        probe = wfo.probe_exclusive_access(target)
        assert probe.reason == "file_in_use"
        assert "Another program" not in (probe.message or "")
        assert "is using this add-in file" in (probe.message or "")
    finally:
        child.kill()
        child.wait()


def test_file_holders_ignores_this_process(tmp_path: Path):
    target = tmp_path / "Mine.xlam"
    target.write_bytes(b"x")
    with open(target, "rb"):
        assert wfo.file_holders(target) == []
        assert "Another program" in (wfo.probe_exclusive_access(target).message or "")


def test_probe_waits_for_a_brief_lock(tmp_path: Path):
    target = tmp_path / "Brief.xlam"
    target.write_bytes(b"x")
    child = _hold_in_child(target, 0.8)
    try:
        assert not wfo.probe_exclusive_access(target).ok
        assert wfo.probe_exclusive_access(target, wait=10).ok
    finally:
        child.kill()
        child.wait()
