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
