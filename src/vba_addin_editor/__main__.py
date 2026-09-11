"""CLI entry: GUI by default; --self-test runs headless verification (plan 36)."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


def self_test(path: Path) -> int:
    """Open, parse, list, validate an add-in without any GUI. Exit 0 on success."""
    from vba_addin_editor.adapters.pyopenvba_adapter import PyOpenVBAAdapter
    from vba_addin_editor.platform import paths

    try:
        fp = paths.fingerprint(path)
    except OSError as exc:
        print(f"FAIL: cannot read {path}: {exc}", file=sys.stderr)
        return 2
    adapter = PyOpenVBAAdapter()
    package = None
    from vba_addin_editor.adapters.ooxml_package_adapter import (
        XML_EDITABLE_EXTENSIONS,
        OoxmlPackageAdapter,
    )

    if path.suffix.lower() in XML_EDITABLE_EXTENSIONS:
        package = OoxmlPackageAdapter()
    try:
        snapshot = adapter.open_snapshot(path, fp)
        if package is not None:
            from dataclasses import replace

            snapshot = replace(
                snapshot,
                xml_parts=package.snapshot_xml_parts(path),
                package_safety=package.inspect_package_signature(path),
            )
    except Exception as exc:  # noqa: BLE001 - single CLI failure boundary
        from vba_addin_editor.adapters.ooxml_package_adapter import PackageError

        msg = str(exc) if isinstance(exc, PackageError) else f"unexpected failure: {exc!r}"
        print(f"FAIL: {msg}", file=sys.stderr)
        return 2
    print(f"OK {path.name}: project={snapshot.project_name!r} "
          f"code_page={snapshot.code_page} modules={len(snapshot.modules)} "
          f"xml_parts={len(snapshot.xml_parts)}")
    for m in snapshot.modules:
        print(f"  {m.pyopenvba_kind:9} {m.original_name}")
    return 0

def self_roundtrip(path: Path) -> int:
    """Developer-only candidate round-trip on a COPY of the fixture (plan 36)."""
    import shutil
    import tempfile

    from vba_addin_editor.adapters.pyopenvba_adapter import (
        AdapterError,
        PyOpenVBAAdapter,
    )
    from vba_addin_editor.domain.changes import compute_changes
    from vba_addin_editor.services.document_service import DocumentService
    from vba_addin_editor.services.save_service import SaveService

    with tempfile.TemporaryDirectory() as td:
        work = Path(td) / path.name
        shutil.copy2(path, work)
        try:
            draft = DocumentService().open(work)
        except AdapterError as exc:
            print(f"FAIL: {exc}", file=sys.stderr)
            return 2
        target = next((m for m in draft.modules if m.pyopenvba_kind == "standard"), None)
        if target is None:
            print("FAIL: no standard module to edit", file=sys.stderr)
            return 2
        target.body = target.body + "\r\n' vbaae round-trip marker\r\n"
        if draft.xml_parts:
            part = next(
                (p for p in draft.xml_parts if p.path == "docProps/core.xml"),
                draft.xml_parts[0],
            )
            part.text = part.text + "\n<!-- vbaae round-trip marker -->"
        changes = compute_changes(draft)
        if changes.is_empty:
            print("FAIL: no changes staged", file=sys.stderr)
            return 2
        dest = Path(td) / f"copy{path.suffix}"
        result = SaveService(adapter=PyOpenVBAAdapter()).save_copy(draft, dest)
        if result.kind != "success":
            print(
                f"FAIL: save_copy: kind={result.kind} reason={result.reason} "
                f"problems={result.problems}",
                file=sys.stderr,
            )
            return 2
        try:
            reopened = DocumentService().open(dest)
        except AdapterError as exc:
            print(f"FAIL: reopen: {exc}", file=sys.stderr)
            return 2
        if reopened.is_dirty():
            print("FAIL: reopened draft unexpectedly dirty", file=sys.stderr)
            return 2
    print(f"OK round-trip {path.name}")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="VBAAddinEditor",
        description="Edit VBA source inside Office VBA files (.xlam/.ppam/.pptm).",
    )
    parser.add_argument("file", nargs="?", help="Office VBA file (.xlam/.ppam/.pptm) to open")
    parser.add_argument("--self-test", metavar="ADDIN", help="headless open+parse+validate")
    parser.add_argument("--self-roundtrip", metavar="ADDIN", help="headless edit-candidate round-trip on a copy")
    parser.add_argument("--version-json", action="store_true", help="print machine-readable build identity")
    args = parser.parse_args(argv)

    if args.version_json:
        import json

        from vba_addin_editor.version import build_identity

        print(json.dumps(build_identity(), indent=2, sort_keys=True))
        return 0
    if args.self_test:
        return self_test(Path(args.self_test))
    if args.self_roundtrip:
        return self_roundtrip(Path(args.self_roundtrip))

    from vba_addin_editor.ui.main_window import run_gui

    if args.file:
        import tkinter as tk

        root = tk.Tk()
        root.withdraw()
        window = None

        from vba_addin_editor.ui.main_window import MainWindow

        window = MainWindow(root)
        root.deiconify()
        window.load_path(Path(args.file))
        root.mainloop()
        return 0

    run_gui()
    return 0


if __name__ == "__main__":
    sys.exit(main())
