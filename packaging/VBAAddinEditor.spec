# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for VBA Add-in Editor (plan 33).

Build onedir first for debugging:
    pyinstaller packaging/VBAAddinEditor.spec
Final release (onefile):
    pyinstaller packaging/VBAAddinEditor.spec -- --onefile
"""

import json
import os
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

block_cipher = None
ONEFILE = "--onefile" in sys.argv
ROOT = Path(SPECPATH).resolve().parent.parent
SRC = ROOT / "src" / "vba_addin_editor"
BUILD_INFO = SRC / "_build_info.py"


def _git(args: list[str]) -> str:
    try:
        return subprocess.check_output(["git", *args], cwd=ROOT, text=True).strip()
    except Exception:
        return "unbuilt"


commit = _git(["rev-parse", "HEAD"])
dirty = bool(_git(["status", "--porcelain"]))
release_mode = os.environ.get("VBAAE_RELEASE") == "1"
if release_mode and dirty:
    raise SystemExit("Release builds require a clean git tree.")
mode = "onefile" if ONEFILE else "onedir"
identity = {
    "version": "0.1.0",
    "source_commit": commit,
    "dirty_tree": dirty,
    "build_time_utc": datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z"),
    "python_version": sys.version.split()[0],
    "python_architecture": "64bit" if sys.maxsize > 2**32 else "32bit",
    "packaged_mode": mode if not dirty or not release_mode else f"{mode}-dirty",
    "pyopenvba_pin": "3.4.0",
    "dependencies": {"pyopenvba": "3.4.0"},
}
if dirty and not release_mode:
    identity["packaged_mode"] = f"{mode}-development-dirty"
BUILD_INFO.write_text("BUILD_INFO = " + repr(identity) + "\n", encoding="utf-8")
manifest_dir = ROOT / "dist"
manifest_dir.mkdir(exist_ok=True)
(manifest_dir / "build-identity.json").write_text(
    json.dumps(identity, indent=2, sort_keys=True) + "\n", encoding="utf-8"
)

a = Analysis(
    ["../src/vba_addin_editor/__main__.py"],
    pathex=["../src"],
    binaries=[],
    datas=[("../THIRD_PARTY_NOTICES.md", "."), ("../README.md", ".")],
    hiddenimports=["pyopenvba", "pyopenvba._templates"],
    hookspath=[],
    runtime_hooks=[],
    excludes=["tkinter.test", "test", "unittest"],
    cipher=block_cipher,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

if ONEFILE:
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.zipfiles,
        a.datas,
        name="VBAAddinEditor",
        console=False,
        upx=False,
    )
else:
    exe = EXE(
        pyz,
        a.scripts,
        exclude_binaries=True,
        name="VBAAddinEditor",
        console=False,
        upx=False,
    )
    coll = COLLECT(exe, a.binaries, a.zipfiles, a.datas, name="VBAAddinEditor")
