"""Module import/export (.bas/.cls only; no .frm/.frx) and backup restore (plan 11, 16)."""

from __future__ import annotations

from pathlib import Path

from vba_addin_editor.adapters.pyopenvba_adapter import AdapterError, PyOpenVBAAdapter
from vba_addin_editor.domain.changes import compute_changes
from vba_addin_editor.domain.document import (
    DocumentDraft,
    ModuleDisplayKind,
    ModuleDraft,
    new_module_id,
)

_IMPORT_EXTS = {".bas": "standard", ".cls": "other"}
_MAX_IMPORT_BYTES = 8 * 1024 * 1024


class ImportExportService:
    def __init__(self, adapter: PyOpenVBAAdapter | None = None) -> None:
        self.adapter = adapter or PyOpenVBAAdapter()

    # -- export ----------------------------------------------------------

    def export_module(self, draft: DocumentDraft, module_id: str, dest: Path) -> str:
        mod = draft.module_by_id(module_id)
        if mod is None:
            raise KeyError(module_id)
        baseline = next((m for m in draft.baseline.modules if m.id == module_id), None)
        header = mod.hidden_header or (baseline.hidden_header if baseline is not None else "")
        from vba_addin_editor.adapters.source_codec import compose_module_source

        text = compose_module_source(
            body=mod.body,
            hidden_header=header,
            current_name=mod.current_name,
            kind=mod.kind,
            pyopenvba_kind=mod.pyopenvba_kind,
        )
        ext = ".bas" if mod.pyopenvba_kind == "standard" else ".cls"
        dest = dest.with_suffix(ext) if dest.suffix.lower() not in (".bas", ".cls") else dest
        dest.write_bytes(
            to_crlf_bytes(text)
        )
        return str(dest)

    def export_all(self, draft: DocumentDraft, folder: Path) -> list[str]:
        written: list[str] = []
        for mod in draft.final_module_state():
            suffix = ".bas" if mod.pyopenvba_kind == "standard" else ".cls"
            written.append(self.export_module(draft, mod.id, folder / f"{mod.current_name}{suffix}"))
        return written

    # -- import ----------------------------------------------------------

    def read_import(self, path: Path) -> tuple[str, str, bool]:
        """Return (suggested_name, LF text, is_class). Rejects .frm/.frx and junk."""
        path = Path(path)
        if path.suffix.lower() not in _IMPORT_EXTS:
            raise AdapterError("Only .bas and .cls source files can be imported.")
        raw = path.read_bytes()
        if len(raw) > _MAX_IMPORT_BYTES:
            raise AdapterError("Import file is unreasonably large.")
        text = raw.decode("utf-8", errors="strict")
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        is_class = path.suffix.lower() == ".cls"
        return path.stem, text, is_class

    def import_as_new(self, draft: DocumentDraft, path: Path, name: str) -> ModuleDraft:
        _, text, is_class = self.read_import(path)
        mod = ModuleDraft(
            id=new_module_id(),
            origin_name=None,
            current_name=name,
            body=text,
            kind=ModuleDisplayKind.CLASS if is_class else ModuleDisplayKind.STANDARD,
            pyopenvba_kind="other" if is_class else "standard",
            is_new=True,
            is_deleted=False,
            destructive_ops_safe=True,
        )
        draft.modules.append(mod)
        return mod

    def replace_body_from_file(self, draft: DocumentDraft, module_id: str, path: Path) -> None:
        mod = draft.module_by_id(module_id)
        if mod is None:
            raise KeyError(module_id)
        _suggested, text, _is_class = self.read_import(path)
        mod.body = text


def to_crlf_bytes(text: str) -> bytes:
    return (
        text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n").encode("utf-8")
    )


def has_changes(draft: DocumentDraft) -> bool:
    return not compute_changes(draft).is_empty
