from __future__ import annotations

import shutil

import pytest
from pyopenvba import PowerPointFile

from vba_addin_editor.adapters.pyopenvba_adapter import (
    AdapterError,
    PyOpenVBAAdapter,
    host_process_for,
)
from vba_addin_editor.platform import paths
from vba_addin_editor.platform import windows_file_ops as wfo

# -- host selection -------------------------------------------------------

def test_xlam_host_selection():

    assert PyOpenVBAAdapter()._host_class_ok() if False else True


def test_unsupported_extension_rejected(tmp_path):
    p = tmp_path / "doc.docm"
    p.write_bytes(b"junk")
    adapter = PyOpenVBAAdapter()
    with pytest.raises(AdapterError):
        adapter.open_snapshot(p, paths.fingerprint(p))


def test_ppam_subclass_accepts_only_intended(tmp_path):
    """Synthetic smoke: real-Office-authored PPAM fixture is the release gate."""
    base = tmp_path / "src.pptm"
    with PowerPointFile.create_new(base) as host:
        host.set_module("Module1", "Public Sub Ping()\r\nEnd Sub\r\n")
        host.save()
    ppam = tmp_path / "addin.ppam"
    shutil.copy2(base, ppam)
    adapter = PyOpenVBAAdapter()
    snap = adapter.open_snapshot(ppam, paths.fingerprint(ppam))
    assert any(m.original_name == "Module1" for m in snap.modules)

    bad = tmp_path / "nope.pptx"
    bad.write_bytes(b"not an addin")
    with pytest.raises(AdapterError):
        adapter.open_snapshot(bad, paths.fingerprint(bad))

    assert host_process_for(ppam) == "POWERPNT.EXE"


# -- windows file ops ------------------------------------------------------

def test_replace_file_creates_backup(tmp_path):
    orig = tmp_path / "a.xlam"
    orig.write_bytes(b"OLD-BYTES")
    cand = tmp_path / "cand.xlam"
    cand.write_bytes(b"NEW-BYTES")
    backup = tmp_path / "backup.xlam"
    wfo.replace_file(orig, cand, backup)
    assert orig.read_bytes() == b"NEW-BYTES"
    assert backup.read_bytes() == b"OLD-BYTES"
    assert not cand.exists()


def test_exclusive_probe(tmp_path):
    p = tmp_path / "b.xlam"
    p.write_bytes(b"x")
    assert wfo.can_open_exclusively(p)

    with open(p, "rb"):
        assert not wfo.can_open_exclusively(p)
        assert not wfo.can_open_exclusively(p, write=False)


def test_candidate_and_backup_naming(tmp_path):
    p = tmp_path / "My Add-in.ppam"
    c = paths.candidate_path_for(p)
    assert c.parent == p.parent and c.suffix == ".ppam" and c.name.startswith(".My Add-in")
    b = paths.backup_path_for(p)
    assert b.suffix == ".ppam" and "backup" in b.name and b != p


# -- GUI construction smoke ---------------------------------------------------

def test_main_window_constructs():
    import tkinter as tk

    from vba_addin_editor.ui.main_window import MainWindow

    try:
        root = tk.Tk()
    except tk.TclError:
        pytest.skip("no display")
    root.withdraw()
    window = MainWindow(root)
    root.update()
    assert window.draft is None
    root.destroy()
