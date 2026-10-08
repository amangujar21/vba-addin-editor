"""Open and save never read back a file they wrote; commits keep the file's metadata.

Antivirus scans a freshly written macro package on its first open for read
(seconds for a large VBA project), so the pipeline works from bytes in memory.
"""

from __future__ import annotations

import os
import time
from pathlib import Path

import pytest
from helpers import edit_module1, make_service

from vba_addin_editor.adapters import pyopenvba_adapter as pa
from vba_addin_editor.platform import paths
from vba_addin_editor.platform import windows_file_ops as wfo
from vba_addin_editor.platform.settings import Settings, load_settings, save_settings
from vba_addin_editor.services.backup_service import BackupService
from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.services.session_service import SessionService


@pytest.fixture()
def reads(monkeypatch):
    """Names of every file read through Path.read_bytes while the fixture is live."""
    seen: list[str] = []
    real = Path.read_bytes

    def tracking(self):
        seen.append(self.name)
        return real(self)

    monkeypatch.setattr(Path, "read_bytes", tracking)
    return seen


# -- commit primitive ---------------------------------------------------------


def _pair(tmp_path: Path) -> tuple[Path, Path, Path]:
    original = tmp_path / "Addin.xlam"
    original.write_bytes(b"old")
    candidate = tmp_path / ".Addin.candidate.xlam"
    wfo.write_verified(candidate, b"new")
    return original, candidate, tmp_path / "Addin - backup.xlam"


def test_rename_commit_keeps_identity_backup_and_creation_time(tmp_path: Path):
    original, candidate, backup = _pair(tmp_path)
    time.sleep(0.02)
    created = os.stat(original).st_ctime_ns
    old_identity = paths.file_identity(original)
    new_identity = paths.file_identity(candidate)
    assert wfo.commit_replace(original, candidate, backup) == "rename"
    assert (original.read_bytes(), backup.read_bytes()) == (b"new", b"old")
    assert not candidate.exists()
    assert paths.file_identity(original) == new_identity
    assert paths.file_identity(backup) == old_identity
    assert os.stat(original).st_ctime_ns == created


def test_rename_commit_without_backup_leaves_one_file(tmp_path: Path):
    original, candidate, _backup = _pair(tmp_path)
    assert wfo.commit_replace(original, candidate, None) == "rename"
    assert original.read_bytes() == b"new"
    assert [p.name for p in tmp_path.iterdir()] == ["Addin.xlam"]


def test_alternate_data_stream_forces_replacefile_which_keeps_it(tmp_path: Path):
    original, candidate, backup = _pair(tmp_path)
    with open(f"{original}:Zone.Identifier", "w") as stream:
        stream.write("[ZoneTransfer]")
    assert not wfo.rename_commit_preserves_metadata(original, candidate)
    assert wfo.commit_replace(original, candidate, backup) == "replace_file"
    assert original.read_bytes() == b"new"
    with open(f"{original}:Zone.Identifier") as stream:
        assert stream.read() == "[ZoneTransfer]"


def test_differing_attributes_force_replacefile(tmp_path: Path):
    original, candidate, _backup = _pair(tmp_path)
    hidden = 0x2
    wfo._KERNEL32.SetFileAttributesW(str(original), hidden)
    assert not wfo.rename_commit_preserves_metadata(original, candidate)


def test_backup_name_collision_fails_before_anything_moves(tmp_path: Path):
    original, candidate, backup = _pair(tmp_path)
    backup.write_bytes(b"in the way")
    with pytest.raises(wfo.FileOpsError):
        wfo._rename_commit(original, candidate, backup)
    assert (original.read_bytes(), candidate.read_bytes(), backup.read_bytes()) == (
        b"old",
        b"new",
        b"in the way",
    )


def test_write_verified_refuses_to_overwrite(tmp_path: Path):
    target = tmp_path / "x.bin"
    wfo.write_verified(target, b"abc")
    with pytest.raises(FileExistsError):
        wfo.write_verified(target, b"def")
    assert target.read_bytes() == b"abc"


def test_exclusive_probe_sees_other_handles_and_read_only(tmp_path: Path):
    target = tmp_path / "Addin.xlam"
    target.write_bytes(b"data")
    assert wfo.probe_exclusive_access(target).ok
    with open(target, "rb"):
        assert wfo.probe_exclusive_access(target).reason == "file_in_use"
    os.chmod(target, 0o444)
    try:
        assert wfo.probe_exclusive_access(target).reason == "read_only"
    finally:
        os.chmod(target, 0o666)
    assert wfo.probe_exclusive_access(tmp_path / "missing.xlam").reason == "source_missing"


# -- fingerprints ---------------------------------------------------------------


def test_written_identity_skips_the_hash_only_while_the_file_is_untouched(tmp_path: Path, reads):
    target = tmp_path / "Addin.xlam"
    wfo.write_verified(target, b"payload")
    fp = paths.written_fingerprint(b"payload", target)
    assert paths.fingerprint_status(target, fp) == "match"
    assert reads == []
    time.sleep(0.02)
    target.write_bytes(b"PAYLOAD")  # same size, new mtime
    assert paths.fingerprint_status(target, fp) == "external_change"
    assert reads  # fell back to hashing the content
    target.unlink()
    assert paths.fingerprint_status(target, fp) == "source_missing"


def test_a_fingerprint_from_reading_always_hashes(tmp_path: Path, reads):
    target = tmp_path / "Addin.xlam"
    target.write_bytes(b"payload")
    data, fp = paths.read_stable(target)
    assert data == b"payload" and fp.written_identity is None
    reads.clear()
    assert paths.fingerprint_status(target, fp) == "match"
    assert reads


# -- open / save ------------------------------------------------------------------


def test_session_open_reads_only_the_original(work_xlam: Path, tmp_path: Path, reads):
    session = SessionService(session_root=tmp_path / "sessions").open(work_xlam)
    assert set(reads) == {work_xlam.name}
    assert session.draft.baseline.package_bytes == session.captured_path.read_bytes()
    assert session.captured_fingerprint.sha256 == paths.fingerprint(work_xlam).sha256


def test_save_and_publish_read_nothing_back(work_xlam: Path, tmp_path: Path, reads):
    sessions = SessionService(session_root=tmp_path / "sessions")
    session = sessions.open(work_xlam)
    edit_module1(session.draft)
    reads.clear()
    result = make_service().save_addin(session.draft)
    assert result.kind == "success", result
    committed = session.draft.baseline.package_bytes
    assert sessions.publish_verified_baseline(session, work_xlam, committed) is None
    catalog = BackupService(catalog_path=tmp_path / "catalog.json")
    assert (
        catalog.record_success(
            original=work_xlam,
            backup=result.backup_path,
            operation_type="save",
            sha256=result.details["hash_backup"],
        )
        is None
    )
    # Only the pre-commit checks read the original, which was already read at open.
    assert set(reads) <= {work_xlam.name}
    reads.clear()
    assert work_xlam.read_bytes() == committed == session.captured_path.read_bytes()
    assert catalog.list_records(work_xlam)[0].sha256 == paths.sha256_bytes(
        result.backup_path.read_bytes()
    )
    assert session.baseline_sha256 == paths.sha256_bytes(committed)


def test_draft_without_package_bytes_still_saves_from_disk(work_xlam: Path):
    from dataclasses import replace

    draft = DocumentService().open(work_xlam)
    draft.baseline = replace(draft.baseline, package_bytes=None)
    edit_module1(draft)
    assert make_service().save_addin(draft).kind == "success"
    assert "V2" in DocumentService().open(work_xlam).find_current("Module1").body


def test_stale_package_bytes_block_the_save(work_xlam: Path):
    from dataclasses import replace

    draft = DocumentService().open(work_xlam)
    draft.baseline = replace(draft.baseline, package_bytes=draft.baseline.package_bytes + b"\x00")
    edit_module1(draft)
    before = work_xlam.read_bytes()
    result = make_service().save_addin(draft)
    assert result.kind == "blocked" and result.reason == "external_change"
    assert work_xlam.read_bytes() == before


def test_capture_only_save_cannot_reach_the_disk():
    """If the capture hook ever stopped applying, save() would fail, not write a file."""
    assert not pa._CAPTURE_ONLY.parent.exists()
    with pytest.raises(OSError):
        pa._CAPTURE_ONLY.write_bytes(b"x")


# -- module source cache ------------------------------------------------------------


def test_sources_are_decompressed_once_and_come_back_from_disk(
    work_xlam: Path, tmp_path: Path, monkeypatch
):
    from pyopenvba.vba import VBAModule

    big = "\r\n".join(f"Public Sub P{i}()\r\n    MsgBox {i}\r\nEnd Sub" for i in range(400)) + "\r\n"
    draft = DocumentService().open(work_xlam)
    draft.find_current("Module1").body = big
    assert make_service().save_addin(draft).kind == "success"

    loads: list[str] = []
    real_source = VBAModule.source.fget

    def counting(self):
        if not self.source_loaded:
            loads.append(self.name)
        return real_source(self)

    monkeypatch.setattr(VBAModule, "source", property(counting, VBAModule.source.fset))
    cache = pa._SourceCache()
    cache.use_disk(tmp_path / "cache")
    monkeypatch.setattr(pa, "_SOURCES", cache)

    first = DocumentService().open(work_xlam)
    assert loads == ["Module1"]  # the only module too large for pyOpenVBA's first chunk
    DocumentService().open(work_xlam)
    assert loads == ["Module1"]

    # A new process: empty memory, same disk folder.
    fresh = pa._SourceCache()
    fresh.use_disk(tmp_path / "cache")
    monkeypatch.setattr(pa, "_SOURCES", fresh)
    again = DocumentService().open(work_xlam)
    assert loads == ["Module1"]
    assert again.find_current("Module1").body == first.find_current("Module1").body == big.replace(
        "\r\n", "\n"
    )

    # A damaged cache file is ignored, not trusted.
    (cached,) = (tmp_path / "cache").glob("*/*.z")
    cached.write_bytes(cached.read_bytes()[:-3] + b"xyz")
    empty = pa._SourceCache()
    empty.use_disk(tmp_path / "cache")
    monkeypatch.setattr(pa, "_SOURCES", empty)
    assert DocumentService().open(work_xlam).find_current("Module1").body == big.replace("\r\n", "\n")
    assert loads == ["Module1", "Module1"]


def test_cache_prune_drops_oldest_first(tmp_path: Path):
    root = tmp_path / "cache" / "ab"
    root.mkdir(parents=True)
    for index in range(4):
        file = root / f"{index}.z"
        file.write_bytes(b"x" * 100)
        os.utime(file, (1000 + index, 1000 + index))
    pa._prune_source_cache(tmp_path / "cache", max_bytes=300)
    assert sorted(p.name for p in root.iterdir()) == ["2.z", "3.z"]


# -- settings -----------------------------------------------------------------------


def test_settings_round_trip_and_tolerate_bad_files(tmp_path: Path):
    path = tmp_path / "settings.json"
    assert load_settings(path) == Settings(auto_backup=True)
    assert save_settings(Settings(auto_backup=False), path)
    assert load_settings(path).auto_backup is False
    path.write_text('{"auto_backup": "no", "unknown": 1}', encoding="utf-8")
    assert load_settings(path).auto_backup is True
    path.write_text("not json", encoding="utf-8")
    assert load_settings(path).auto_backup is True
