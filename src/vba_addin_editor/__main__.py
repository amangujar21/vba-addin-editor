"""CLI entry: GUI by default; --self-test runs headless verification (plan 36)."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


def self_test(path: Path) -> int:
    """Open, parse, list, validate an add-in without any GUI. Exit 0 on success."""
    from vba_addin_editor.adapters.pyopenvba_adapter import AdapterError, PyOpenVBAAdapter
    from vba_addin_editor.platform import paths

    try:
        fp = paths.fingerprint(path)
    except OSError as exc:
        print(f"FAIL: cannot read {path}: {exc}", file=sys.stderr)
        return 2
    adapter = PyOpenVBAAdapter()
    try:
        snapshot = adapter.open_snapshot(path, fp)
    except AdapterError as exc:
        print(f"FAIL: {exc}", file=sys.stderr)
        return 2
    print(f"OK {path.name}: project={snapshot.project_name!r} "
          f"code_page={snapshot.code_page} modules={len(snapshot.modules)}")
    for m in snapshot.modules:
        print(f"  {m.pyopenvba_kind:9} {m.original_name}")
    return 0


def self_roundtrip(path: Path) -> int:
    """Developer-only candidate round-trip on a COPY of the fixture (plan 36)."""
    import shutil
    import tempfile

    from vba_addin_editor.adapters.pyopenvba_adapter import AdapterError, PyOpenVBAAdapter
    from vba_addin_editor.domain.document import draft_from_snapshot
    from vba_addin_editor.platform import paths

    adapter = PyOpenVBAAdapter()
    with tempfile.TemporaryDirectory() as td:
        work = Path(td) / path.name
        shutil.copy2(path, work)
        try:
            draft = draft_from_snapshot(adapter.open_snapshot(work, paths.fingerprint(work)))
        except AdapterError as exc:
            print(f"FAIL: {exc}", file=sys.stderr)
            return 2
        target = next((m for m in draft.modules if m.pyopenvba_kind == "standard"), None)
        if target is None:
            print("FAIL: no standard module to edit", file=sys.stderr)
            return 2
        target.body = target.body + "\r\n' vbaae round-trip marker\r\n"
        candidate = Path(td) / f"cand{path.suffix}"
        try:
            adapter.build_candidate(work, draft, candidate, allow_signature_removal=False)
        except AdapterError as exc:
            print(f"FAIL: build: {exc}", file=sys.stderr)
            return 2
        verification = adapter.verify_candidate(work, candidate, draft)
        if not verification.ok:
            print(f"FAIL: verify: {verification.problems}", file=sys.stderr)
            return 2
    print(f"OK round-trip {path.name}")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="VBAAddinEditor")
    parser.add_argument("file", nargs="?", help="add-in (.xlam/.ppam) to open")
    parser.add_argument("--self-test", metavar="ADDIN", help="headless open+parse+validate")
    parser.add_argument("--self-roundtrip", metavar="ADDIN", help="headless edit-candidate round-trip on a copy")
    args = parser.parse_args(argv)

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
