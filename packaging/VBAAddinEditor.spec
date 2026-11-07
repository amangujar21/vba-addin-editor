# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for VBA Add-in Editor (plan 33).

Build onedir first for debugging:
    pyinstaller packaging/VBAAddinEditor.spec
Final release (onefile):
    pyinstaller packaging/VBAAddinEditor.spec -- --onefile
"""

import json
import os
from importlib import metadata
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

block_cipher = None
ONEFILE = "--onefile" in sys.argv
# SPECPATH is the spec's directory; locate the repo root from it.
_SPEC_DIR = Path(SPECPATH).resolve()
ROOT = next(p for p in (_SPEC_DIR, *_SPEC_DIR.parents) if (p / "pyproject.toml").exists())
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
    "version": "0.9.0",
    "source_commit": commit,
    "dirty_tree": dirty,
    "build_time_utc": datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z"),
    "python_version": sys.version.split()[0],
    "python_architecture": "64bit" if sys.maxsize > 2**32 else "32bit",
    "packaged_mode": mode if not dirty or not release_mode else f"{mode}-dirty",
    "pyopenvba_pin": "6.2.0",
    "dependencies": {"pyopenvba": "6.2.0", "pillow": metadata.version("pillow")},
}
if dirty and not release_mode:
    identity["packaged_mode"] = f"{mode}-development-dirty"
# Pillow's LICENSE also carries the licences of the image libraries it bundles.
_pillow = metadata.distribution("pillow")
_license = next(
    (f for f in _pillow.files or () if f.name.upper() in ("LICENSE", "LICENSE.TXT") and "dist-info" in str(f)),
    None,
)
if _license is None:
    raise SystemExit("Pillow's licence file was not found in its installed distribution.")
PILLOW_LICENSE = Path(str(_pillow.locate_file(_license)))
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
    datas=[
        ("../THIRD_PARTY_NOTICES.md", "."),
        ("../README.md", "."),
        (str(PILLOW_LICENSE), "licenses/pillow"),
    ],
    hiddenimports=["pyopenvba", "pyopenvba._templates"],
    hookspath=[],
    runtime_hooks=[],
    # Icons only need Pillow's core, PNG/JPEG/BMP/GIF/ICO readers and ImageGrab:
    # leave out fonts, colour management, WebP/AVIF and the Tk bridge (~2.5 MB),
    # and the optional packages Pillow imports only if present: NumPy (~40 MB
    # with OpenBLAS), defusedxml (XMP metadata; drags in xmlrpc, pydoc,
    # asyncio, multiprocessing) and olefile (FlashPix/MIC formats).
    excludes=[
        "tkinter.test",
        "test",
        "unittest",
        "numpy",
        "yaml",
        "defusedxml",
        "olefile",
        "PIL.FpxImagePlugin",
        "PIL.MicImagePlugin",
        "PIL.ImageFont",
        "PIL._imagingft",
        "PIL.ImageCms",
        "PIL._imagingcms",
        "PIL.WebPImagePlugin",
        "PIL._webp",
        "PIL.AvifImagePlugin",
        "PIL._avif",
        "PIL.ImageTk",
        "PIL._imagingtk",
        "PIL.ImageQt",
        "PIL.ImageShow",
    ],
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
