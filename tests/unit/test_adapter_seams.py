from __future__ import annotations

import ast
from pathlib import Path

from helpers import make_service
from pyopenvba.cfb import CFB
from pyopenvba.vba import parse_project_stream, parse_projectwm

from vba_addin_editor.adapters.pyopenvba_adapter import vba_entry_for
from vba_addin_editor.services.document_service import DocumentService


def test_pyopenvba_adapter_does_not_rewrite_zip():
    source = Path("src/vba_addin_editor/adapters/pyopenvba_adapter.py").read_text(encoding="utf-8")
    assert "_rewrite_zip_member" not in source
    tree = ast.parse(source)
    for node in ast.walk(tree):
        if (
            isinstance(node, ast.Call)
            and isinstance(node.func, ast.Attribute)
            and node.func.attr == "ZipFile"
        ):
            for arg in node.args[1:]:
                if isinstance(arg, ast.Constant) and arg.value in {"w", "a", "x"}:
                    raise AssertionError("pyopenvba_adapter opened ZipFile for write")
            for kw in node.keywords:
                if kw.arg == "mode" and isinstance(kw.value, ast.Constant) and kw.value.value in {"w", "a", "x"}:
                    raise AssertionError("pyopenvba_adapter opened ZipFile for write")
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute) and node.func.attr == "writestr":
            raise AssertionError("pyopenvba_adapter wrote a ZIP member")


def test_deleted_class_absent_from_stream_dir_project_and_projectwm(work_xlam_with_class: Path):
    draft = DocumentService().open(work_xlam_with_class)
    cls = next(m for m in draft.modules if m.current_name == "OrdinaryClass")
    stream = cls.stream_name or cls.current_name
    cls.is_deleted = True
    result = make_service().save_addin(draft)
    assert result.kind == "success", result
    entry = vba_entry_for(work_xlam_with_class)
    import zipfile

    with zipfile.ZipFile(work_xlam_with_class) as package:
        raw = package.read(entry)
    cfb = CFB.from_bytes(raw)
    reopened = DocumentService().open(work_xlam_with_class)
    names = {m.current_name.casefold() for m in reopened.modules}
    assert "ordinaryclass" not in names
    parsed = parse_project_stream(cfb.get_stream("PROJECT"), code_page=reopened.baseline.code_page)
    assert "OrdinaryClass" not in parsed.class_modules
    streams = {item.casefold() for item in cfb.list_streams_in_storage("VBA")}
    assert stream.casefold() not in streams
    wm = parse_projectwm(cfb.get_stream("PROJECTwm"), code_page=reopened.baseline.code_page)
    wm_names = {item[0].casefold() for item in wm} if wm else set()
    assert "ordinaryclass" not in wm_names
    assert "Module1" in {m.current_name for m in reopened.modules}
