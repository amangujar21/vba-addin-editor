"""Managed folder export/import for ordinary VBA components (IMP-09)."""

from __future__ import annotations

import hashlib
import json
import os
import re
from dataclasses import dataclass
from pathlib import Path

from vba_addin_editor.adapters.pyopenvba_adapter import AdapterError
from vba_addin_editor.adapters.source_codec import compose_module_source, split_attribute_header
from vba_addin_editor.domain.document import (
    ModuleDisplayKind,
    ModuleDraft,
    new_module_id,
)
from vba_addin_editor.domain.session import DocumentSession, ProposalBinding
from vba_addin_editor.platform.paths import (
    is_inside,
    path_or_ancestor_is_reparse_point,
    sha256_bytes,
)
from vba_addin_editor.services.validation_service import validate_module_name

SCHEMA_VERSION = 1
MANIFEST_NAME = "vbaae-project.json"
BASELINE_DIR = ".vbaae-baseline"
MAX_FILE_BYTES = 8 * 1024 * 1024
MAX_AGGREGATE = 64 * 1024 * 1024
MAX_MANIFEST = 4 * 1024 * 1024
_WINDOWS_RESERVED = {
    "con", "prn", "aux", "nul",
    *(f"com{i}" for i in range(1, 10)),
    *(f"lpt{i}" for i in range(1, 10)),
}


@dataclass
class FolderChange:
    export_id: str
    operation: str  # edit | add | delete | rename | conflict | noop
    logical_name: str
    selected: bool
    folder_text: str | None = None
    draft_text: str | None = None
    ancestor_text: str | None = None
    resolution: str | None = None
    new_name: str | None = None
    kind: str = "standard"


@dataclass
class FolderPreview:
    binding: ProposalBinding
    changes: list[FolderChange]
    problems: list[str]


def _safe_filename(name: str, export_id: str, ext: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip(" .")
    if not cleaned or cleaned.casefold() in _WINDOWS_RESERVED:
        cleaned = "module"
    if cleaned.endswith("."):
        cleaned = cleaned[:-1] + "_"
    return f"{cleaned}__{export_id}{ext}"


class FolderSyncService:
    def export_folder(self, session: DocumentSession, dest: Path) -> Path:
        dest = Path(dest)
        if path_or_ancestor_is_reparse_point(dest):
            raise AdapterError("Choose a regular folder. Linked paths are not supported.")
        dest.mkdir(parents=True, exist_ok=True)
        baseline_dir = dest / BASELINE_DIR
        baseline_dir.mkdir(exist_ok=True)
        draft = session.draft
        entries = []
        used_names: set[str] = set()
        for module in draft.final_module_state():
            export_id = hashlib.sha1(module.id.encode("utf-8")).hexdigest()[:12]
            if module.project_item_kind == "class" or module.kind == ModuleDisplayKind.CLASS:
                ext = ".cls"
                kind = "class"
            elif module.pyopenvba_kind == "standard":
                ext = ".bas"
                kind = "standard"
            else:
                ext = ".vba.txt"
                kind = module.project_item_kind or "unknown"
            rel = _safe_filename(module.current_name, export_id, ext)
            if rel.casefold() in used_names:
                raise AdapterError(f"Export produced a case-colliding filename: {rel}")
            used_names.add(rel.casefold())
            abs_path = dest / rel
            if not is_inside(abs_path, dest):
                raise AdapterError("Export path escaped the destination folder.")
            text = compose_module_source(
                body=module.body,
                hidden_header=module.hidden_header,
                current_name=module.current_name,
                kind=module.kind,
                pyopenvba_kind=module.pyopenvba_kind,
            )
            payload = text.replace("\r\n", "\n").replace("\r", "\n").encode("utf-8")
            staging = abs_path.with_suffix(abs_path.suffix + ".tmp")
            staging.write_bytes(payload)
            os.replace(staging, abs_path)
            ancestor = baseline_dir / f"{export_id}.txt"
            ancestor.write_bytes(payload)
            entries.append(
                {
                    "export_id": export_id,
                    "logical_name": module.current_name,
                    "kind": kind,
                    "relative_path": rel,
                    "stream_name": module.stream_name,
                    "module_id": module.id,
                    "source_sha256": sha256_bytes(payload),
                    "can_delete": module.can_delete if not module.is_new else True,
                    "can_rename": module.can_rename if not module.is_new else True,
                }
            )
        manifest = {
            "schema_version": SCHEMA_VERSION,
            "session_id": session.session_id,
            "revision": session.revision,
            "baseline_sha256": session.baseline_sha256,
            "extension": draft.baseline.extension,
            "project_name": draft.baseline.project_name,
            "encoding": "utf-8",
            "newline": "lf",
            "entries": entries,
        }
        encoded = json.dumps(manifest, indent=2, sort_keys=True).encode("utf-8")
        tmp = dest / (MANIFEST_NAME + ".tmp")
        tmp.write_bytes(encoded)
        os.replace(tmp, dest / MANIFEST_NAME)
        return dest / MANIFEST_NAME

    def preview(self, session: DocumentSession, folder: Path) -> FolderPreview:
        folder = Path(folder)
        manifest_path = folder / MANIFEST_NAME
        problems: list[str] = []
        if not manifest_path.exists():
            return FolderPreview(session.binding(), [], ["Folder is missing vbaae-project.json."])
        raw = manifest_path.read_bytes()
        if len(raw) > MAX_MANIFEST:
            return FolderPreview(session.binding(), [], ["Manifest exceeds 4 MiB."])
        try:
            manifest = json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            return FolderPreview(session.binding(), [], ["Manifest is not valid UTF-8 JSON."])
        if manifest.get("schema_version") != SCHEMA_VERSION:
            return FolderPreview(session.binding(), [], ["Unknown folder-export schema."])
        if manifest.get("baseline_sha256") != session.baseline_sha256:
            return FolderPreview(
                session.binding(),
                [],
                ["Folder export belongs to a different baseline. Export again or import modules individually."],
            )
        if path_or_ancestor_is_reparse_point(folder):
            return FolderPreview(session.binding(), [], ["Linked folders cannot be imported."])
        changes: list[FolderChange] = []
        listed_ids = {entry["export_id"] for entry in manifest.get("entries", [])}
        module_by_id = {m.id: m for m in session.draft.modules}
        aggregate = 0
        for entry in manifest.get("entries", []):
            rel = entry.get("relative_path", "")
            if not rel or Path(rel).is_absolute() or ".." in Path(rel).parts:
                problems.append(f"Unsafe path in manifest: {rel!r}")
                continue
            path = folder / rel
            if not is_inside(path, folder):
                problems.append(f"Path escaped the folder: {rel!r}")
                continue
            module = module_by_id.get(entry.get("module_id", ""))
            ancestor_path = folder / BASELINE_DIR / f"{entry['export_id']}.txt"
            ancestor = ancestor_path.read_text(encoding="utf-8") if ancestor_path.exists() else None
            if not path.exists():
                changes.append(
                    FolderChange(
                        export_id=entry["export_id"],
                        operation="delete",
                        logical_name=entry["logical_name"],
                        selected=False,
                        ancestor_text=ancestor,
                        kind=entry.get("kind", "standard"),
                    )
                )
                continue
            data = path.read_bytes()
            if len(data) > MAX_FILE_BYTES:
                problems.append(f"{rel} exceeds 8 MiB.")
                continue
            aggregate += len(data)
            text = data.decode("utf-8")
            draft_text = None
            if module is not None:
                draft_text = compose_module_source(
                    body=module.body,
                    hidden_header=module.hidden_header,
                    current_name=module.current_name,
                    kind=module.kind,
                    pyopenvba_kind=module.pyopenvba_kind,
                )
            if ancestor is not None and text == ancestor and draft_text == ancestor:
                changes.append(
                    FolderChange(
                        export_id=entry["export_id"],
                        operation="noop",
                        logical_name=entry["logical_name"],
                        selected=False,
                        folder_text=text,
                        draft_text=draft_text,
                        ancestor_text=ancestor,
                    )
                )
            elif ancestor is not None and text != ancestor and draft_text == ancestor:
                changes.append(
                    FolderChange(
                        export_id=entry["export_id"],
                        operation="edit",
                        logical_name=entry["logical_name"],
                        selected=True,
                        folder_text=text,
                        draft_text=draft_text,
                        ancestor_text=ancestor,
                    )
                )
            elif ancestor is not None and text != ancestor and draft_text != ancestor and text != draft_text:
                changes.append(
                    FolderChange(
                        export_id=entry["export_id"],
                        operation="conflict",
                        logical_name=entry["logical_name"],
                        selected=False,
                        folder_text=text,
                        draft_text=draft_text,
                        ancestor_text=ancestor,
                        resolution=None,
                    )
                )
            else:
                changes.append(
                    FolderChange(
                        export_id=entry["export_id"],
                        operation="noop",
                        logical_name=entry["logical_name"],
                        selected=False,
                        folder_text=text,
                        draft_text=draft_text,
                        ancestor_text=ancestor,
                    )
                )
        for path in folder.glob("*.bas"):
            self._maybe_add_unlisted(path, listed_ids, changes, problems, kind="standard")
        for path in folder.glob("*.cls"):
            self._maybe_add_unlisted(path, listed_ids, changes, problems, kind="class")
        if aggregate > MAX_AGGREGATE:
            problems.append("Folder source exceeds 64 MiB.")
        return FolderPreview(session.binding(), changes, problems)

    def _maybe_add_unlisted(
        self, path: Path, listed_ids: set[str], changes: list[FolderChange], problems: list[str], *, kind: str
    ) -> None:
        name = path.name
        export_id = ""
        if "__" in path.stem:
            export_id = path.stem.rsplit("__", 1)[-1]
        if export_id in listed_ids:
            return
        if path.suffix.lower() in {".frm", ".frx"}:
            problems.append(f"{name}: UserForm files cannot be imported.")
            return
        text = path.read_text(encoding="utf-8")
        changes.append(
            FolderChange(
                export_id=export_id or path.stem,
                operation="add",
                logical_name=path.stem.split("__")[0],
                selected=False,
                folder_text=text,
                kind=kind,
            )
        )

    def apply(self, session: DocumentSession, preview: FolderPreview) -> None:
        if not session.matches_proposal(preview.binding):
            raise AdapterError("This folder preview is out of date.")
        if preview.problems:
            raise AdapterError("; ".join(preview.problems))
        draft = session.draft
        for change in preview.changes:
            if change.operation == "noop" or not change.selected:
                continue
            if change.operation == "conflict" and change.resolution != "use_folder":
                continue
            module = next(
                (
                    item
                    for item in draft.modules
                    if hashlib.sha1(item.id.encode("utf-8")).hexdigest()[:12] == change.export_id
                    or item.current_name.casefold() == change.logical_name.casefold()
                ),
                None,
            )
            if change.operation == "delete":
                if module is None or module.is_new:
                    if module is not None:
                        draft.modules.remove(module)
                    continue
                if not module.can_delete:
                    raise AdapterError(f"Cannot delete {module.current_name} from a folder import.")
                module.is_deleted = True
                continue
            if change.operation == "add":
                header, body = split_attribute_header(change.folder_text or "")
                name = change.new_name or change.logical_name
                problem = validate_module_name(name, draft)
                if problem:
                    raise AdapterError(problem)
                draft.modules.append(
                    ModuleDraft(
                        id=new_module_id(),
                        origin_name=None,
                        current_name=name,
                        body=body,
                        kind=ModuleDisplayKind.CLASS if change.kind == "class" else ModuleDisplayKind.STANDARD,
                        pyopenvba_kind="other" if change.kind == "class" else "standard",
                        is_new=True,
                        is_deleted=False,
                        destructive_ops_safe=True,
                        can_delete=True,
                        can_rename=True,
                        hidden_header=header,
                        project_item_kind=change.kind,
                    )
                )
                continue
            if module is None:
                continue
            header, body = split_attribute_header(change.folder_text or "")
            module.body = body
            if header and module.hidden_header:
                # Preserve existing attributes other than a consistent VB_Name rename.
                pass
        session.bump()
