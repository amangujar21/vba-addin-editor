"""Saving while Office has the add-in loaded: SaveService with a HostSwap.

The swap here is a fake that records what the pipeline asked of it; the real
one (``platform.office_live.HostSwap``) is exercised against real Excel and
PowerPoint in ``tests/integration/test_live_reload_office.py``.
"""

from __future__ import annotations

from pathlib import Path

from helpers import edit_module1, make_service

from vba_addin_editor.platform import office_live
from vba_addin_editor.platform import windows_file_ops as wfo
from vba_addin_editor.services.document_service import DocumentService


class FakeSwap:
    label = "Excel"

    def __init__(self, path: Path, *, loaded: bool = True, unload_error: str | None = None,
                 reload_error: str | None = None) -> None:
        self.path = path
        self.loaded = loaded
        self.unload_error = unload_error
        self.reload_error = reload_error
        self.calls: list[str] = []
        self.content_at_reload: bytes | None = None

    def unload(self) -> bool:
        self.calls.append("unload")
        if self.unload_error:
            raise office_live.LiveHostError(self.unload_error)
        return self.loaded

    def reload(self) -> None:
        self.calls.append("reload")
        self.content_at_reload = self.path.read_bytes()
        if self.reload_error:
            raise office_live.LiveHostError(self.reload_error)


def _body(path: Path) -> str:
    draft = DocumentService().open(path)
    return next(m.body for m in draft.modules if m.current_name == "Module1")


def test_save_with_office_running_closes_commits_and_reloads(work_xlam: Path):
    draft = DocumentService().open(work_xlam)
    edit_module1(draft)
    swap = FakeSwap(work_xlam)
    # Office running would block a plain save; with a swap it is expected.
    result = make_service(process_probe=lambda p: True).save_addin(draft, swap=swap)
    assert result.kind == "success", result
    assert swap.calls == ["unload", "reload"]
    assert "V2" in _body(work_xlam)
    assert result.details["live_host"] == "Excel" and result.details["live_reloaded"] is True


def test_plain_save_still_blocked_while_office_runs(work_xlam: Path):
    draft = DocumentService().open(work_xlam)
    edit_module1(draft)
    result = make_service(process_probe=lambda p: True).save_addin(draft)
    assert result.kind == "blocked" and result.reason == "office_running"


def test_not_loaded_swap_saves_without_reload(work_xlam: Path):
    """Office runs but does not have this file: the save goes through, nothing to reload."""
    draft = DocumentService().open(work_xlam)
    edit_module1(draft)
    swap = FakeSwap(work_xlam, loaded=False)
    result = make_service(process_probe=lambda p: True).save_addin(draft, swap=swap)
    assert result.kind == "success", result
    assert swap.calls == ["unload"]
    assert "live_host" not in result.details


def test_office_refusing_to_close_saves_nothing(work_xlam: Path):
    before = work_xlam.read_bytes()
    draft = DocumentService().open(work_xlam)
    edit_module1(draft)
    swap = FakeSwap(work_xlam, unload_error="Excel is busy")
    result = make_service(process_probe=lambda p: True).save_addin(draft, swap=swap)
    assert result.kind == "blocked" and result.reason == "live_unload_failed"
    assert "Excel is busy" in (result.message or "")
    assert swap.calls == ["unload"]
    assert work_xlam.read_bytes() == before
    assert not list(work_xlam.parent.glob(".*candidate*"))
    assert draft.is_dirty()


def test_file_still_locked_after_close_reloads_the_old_file(work_xlam: Path):
    before = work_xlam.read_bytes()
    draft = DocumentService().open(work_xlam)
    edit_module1(draft)
    swap = FakeSwap(work_xlam)
    service = make_service(exclusive_probe=lambda p: "Excel is using this add-in file.")
    service.lock_wait_seconds = 0
    result = service.save_addin(draft, swap=swap)
    assert result.kind == "blocked" and result.reason == "file_in_use"
    assert swap.calls == ["unload", "reload"]
    assert swap.content_at_reload == before  # Office gets its add-in back as it was
    assert work_xlam.read_bytes() == before


def test_failed_commit_still_reloads(work_xlam: Path):
    before = work_xlam.read_bytes()
    draft = DocumentService().open(work_xlam)
    edit_module1(draft)
    swap = FakeSwap(work_xlam)

    def broken_commit(replaced, replacement, backup):
        raise wfo.FileOpsError("injected", win32_error=5)

    result = make_service(commit=broken_commit).save_addin(draft, swap=swap)
    assert result.kind == "error" and result.reason == "commit_failed"
    assert swap.calls == ["unload", "reload"]
    assert swap.content_at_reload == before


def test_reload_failure_is_reported_on_a_successful_save(work_xlam: Path):
    draft = DocumentService().open(work_xlam)
    edit_module1(draft)
    swap = FakeSwap(work_xlam, reload_error="Excel could not reopen it")
    result = make_service().save_addin(draft, swap=swap)
    assert result.kind == "success"
    assert result.details["live_reloaded"] is False
    assert "could not reopen" in result.details["live_reload_error"]
    assert "V2" in _body(work_xlam)


def test_host_of():
    assert office_live.host_of(Path("a.XLAM")) == office_live.EXCEL
    assert office_live.host_of(Path("a.ppam")) == office_live.POWERPOINT
    assert office_live.host_of(Path("a.pptm")) == office_live.POWERPOINT
    assert office_live.host_of(Path("a.xlsm")) is None
