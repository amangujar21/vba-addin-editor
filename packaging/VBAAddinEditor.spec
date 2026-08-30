# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for VBA Add-in Editor (plan 33).

Build onedir first for debugging:
    pyinstaller packaging/VBAAddinEditor.spec
Final release (onefile):
    pyinstaller packaging/VBAAddinEditor.spec -- --onefile
"""

import sys

block_cipher = None
ONEFILE = "--onefile" in sys.argv

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
