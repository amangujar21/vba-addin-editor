"""Managed folder export/import for ordinary VBA components (IMP-09)."""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
from dataclasses import dataclass, field
from pathlib import Path

from vba_addin_editor.adapters.pyopenvba_adapter import AdapterError
from vba_addin_editor.adapters.source_codec import compose_module_source, split_attribute_header
from vba_addin_editor.domain.document import (
    ModuleDisplayKind,
    ModuleDraft,
    clone_draft,
    new_module_id,
)
from vba_addin_editor.domain.history import HistoryCommand
from vba_addin_editor.domain.session import DocumentSession, ProposalBinding
from vba_addin_editor.platform.paths import (
    is_inside,
    path_or_ancestor_is_reparse_point,
    sha256_bytes,
)
from vba_addin_editor.services.history_service import snapshot_modules
from vba_addin_editor.services.validation_service import validate_draft, validate_module_name

SCHEMA_VERSION = 1
MANIFEST_NAME = "vbaae-project.json"
BASELINE_DIR = ".vbaae-baseline"
STAGING_DIR = ".vbaae-staging"
MAX_FILE_BYTES = 8 * 1024 * 1024
MAX_AGGREGATE = 64 * 1024 * 1024
MAX_MANIFEST = 4 * 1024 * 1024
_WINDOWS_RESERVED = {
    "con", "prn", "aux", "nul",
    *(f"com{i}" for i in range(1, 10)),
    *(f"lpt{i}" for i in range(1, 10)),
}
_KIND_EXT = {"standard": ".bas", "class": ".cls"}


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


@dataclass
class FolderExportResult:
    ok: bool
    manifest_path: Path | None = None
    problems: list[str] = field(default_factory=list)
    conflicts: list[str] = field(default_factory=list)


def _safe_filename(name: str, export_id: str, ext: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip(" .")
    if not cleaned or cleaned.casefold() in _WINDOWS_RESERVED:
        cleaned = "module"
    if cleaned.endswith("."):
        cleaned = cleaned[:-1] + "_"
    return f"{cleaned}__{export_id}{ext}"


def export_id_for(module: ModuleDraft) -> str:
    return hashlib.sha1(module.id.encode("utf-8")).hexdigest()[:12]


class FolderSyncService:
    def export_folder(self, session: DocumentSession, dest: Path) -> Path:
        result = self.export_folder_result(session, dest)
        if not result.ok:
            raise AdapterError("; ".join(result.problems + result.conflicts) or "Export failed.")
        assert result.manifest_path is not None
        return result.manifest_path

    def export_folder_result(self, session: DocumentSession, dest: Path) -> FolderExportResult:
        dest = Path(dest)
        problems: list[str] = []
        conflicts: list[str] = []
        for probe in (dest, dest / BASELINE_DIR, dest / STAGING_DIR, dest / MANIFEST_NAME):
            if path_or_ancestor_is_reparse_point(probe):
                return FolderExportResult(False, problems=["Linked paths are not supported."])
        dest.mkdir(parents=True, exist_ok=True)
        previous = self._load_previous_manifest(dest, problems)
        if problems and previous is None and (dest / MANIFEST_NAME).exists():
            return FolderExportResult(False, problems=problems)

        managed_hashes: dict[str, str] = {}
        if previous is not None:
            for entry in previous.get("entries", []):
                rel = entry.get("relative_path")
                digest = entry.get("source_sha256")
                if isinstance(rel, str) and isinstance(digest, str):
                    managed_hashes[rel.replace("\\", "/")] = digest

        draft = session.draft
        planned: list[tuple[str, bytes, dict]] = []
        used_names: set[str] = set()
        for module in draft.final_module_state():
            export_id = export_id_for(module)
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
                return FolderExportResult(False, problems=[f"Export produced a case-colliding filename: {rel}"])
            used_names.add(rel.casefold())
            abs_path = dest / rel
            if not is_inside(abs_path, dest) or path_or_ancestor_is_reparse_point(abs_path):
                return FolderExportResult(False, problems=["Export path escaped the destination folder."])
            text = compose_module_source(
                body=module.body,
                hidden_header=module.hidden_header,
                current_name=module.current_name,
                kind=module.kind,
                pyopenvba_kind=module.pyopenvba_kind,
            )
            payload = text.replace("\r\n", "\n").replace("\r", "\n").encode("utf-8")
            header, _body = split_attribute_header(text.replace("\r\n", "\n").replace("\r", "\n"))
            ancestor_name = f"{export_id}.txt"
            entry = {
                "export_id": export_id,
                "logical_name": module.current_name,
                "kind": kind,
                "relative_path": rel,
                "stream_name": module.stream_name,
                "module_id": module.id,
                "source_sha256": sha256_bytes(payload),
                "attribute_sha256": sha256_bytes(header.encode("utf-8")),
                "ancestor_sha256": sha256_bytes(payload),
                "ancestor_path": f"{BASELINE_DIR}/{ancestor_name}",
                "can_delete": module.can_delete if not module.is_new else True,
                "can_rename": module.can_rename if not module.is_new else True,
            }
            planned.append((rel, payload, entry))

            existing = dest / rel
            if existing.exists():
                rel_key = rel.replace("\\", "/")
                current_digest = sha256_bytes(existing.read_bytes())
                recorded = managed_hashes.get(rel_key)
                if recorded is None:
                    problems.append(f"Refusing to overwrite unmanaged file: {rel}")
                elif current_digest != recorded:
                    conflicts.append(f"Managed file changed externally: {rel}")

        if problems or conflicts:
            return FolderExportResult(False, problems=problems, conflicts=conflicts)

        staging = dest / STAGING_DIR
        if staging.exists():
            shutil.rmtree(staging)
        staging.mkdir()
        baseline_stage = staging / BASELINE_DIR
        baseline_stage.mkdir()
        try:
            entries = []
            for rel, payload, entry in planned:
                out = staging / rel
                if not is_inside(out, staging):
                    raise AdapterError("Staging path escaped the export folder.")
                out.write_bytes(payload)
                ancestor = baseline_stage / f"{entry['export_id']}.txt"
                ancestor.write_bytes(payload)
                if sha256_bytes(out.read_bytes()) != entry["source_sha256"]:
                    raise AdapterError("Staged source hash did not verify.")
                if sha256_bytes(ancestor.read_bytes()) != entry["ancestor_sha256"]:
                    raise AdapterError("Staged ancestor hash did not verify.")
                entries.append(entry)
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
            if len(encoded) > MAX_MANIFEST:
                raise AdapterError("Manifest exceeds 4 MiB.")
            (staging / MANIFEST_NAME).write_bytes(encoded)

            dest_baseline = dest / BASELINE_DIR
            dest_baseline.mkdir(exist_ok=True)
            for rel, _payload, entry in planned:
                src = staging / rel
                target = dest / rel
                os.replace(src, target)
                ancestor_src = staging / BASELINE_DIR / f"{entry['export_id']}.txt"
                ancestor_dest = dest_baseline / f"{entry['export_id']}.txt"
                os.replace(ancestor_src, ancestor_dest)
            os.replace(staging / MANIFEST_NAME, dest / MANIFEST_NAME)
        except (OSError, AdapterError) as exc:
            return FolderExportResult(False, problems=[str(exc) or type(exc).__name__])
        finally:
            if staging.exists():
                shutil.rmtree(staging, ignore_errors=True)
        return FolderExportResult(True, manifest_path=dest / MANIFEST_NAME)

    def _load_previous_manifest(self, dest: Path, problems: list[str]) -> dict | None:
        path = dest / MANIFEST_NAME
        if not path.exists():
            return None
        try:
            raw = path.read_bytes()
        except OSError as exc:
            problems.append(f"Previous manifest could not be read: {type(exc).__name__}")
            return None
        if len(raw) > MAX_MANIFEST:
            problems.append("Previous manifest exceeds 4 MiB.")
            return None
        try:
            manifest = json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            problems.append("Previous manifest is not valid UTF-8 JSON.")
            return None
        if not isinstance(manifest, dict):
            problems.append("Previous manifest is not an object.")
            return None
        if manifest.get("schema_version") != SCHEMA_VERSION:
            problems.append("Previous manifest has an unknown schema.")
            return None
        entries = manifest.get("entries")
        if not isinstance(entries, list):
            problems.append("Previous manifest is missing entries.")
            return None
        return manifest

    def preview(self, session: DocumentSession, folder: Path) -> FolderPreview:
        folder = Path(folder)
        problems: list[str] = []
        if path_or_ancestor_is_reparse_point(folder):
            return FolderPreview(session.binding(), [], ["Linked folders cannot be imported."])
        manifest_path = folder / MANIFEST_NAME
        if not manifest_path.exists():
            return FolderPreview(session.binding(), [], ["Folder is missing vbaae-project.json."])
        try:
            raw = manifest_path.read_bytes()
        except OSError as exc:
            return FolderPreview(session.binding(), [], [f"Manifest could not be read: {type(exc).__name__}"])
        if len(raw) > MAX_MANIFEST:
            return FolderPreview(session.binding(), [], ["Manifest exceeds 4 MiB."])
        try:
            manifest = json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            return FolderPreview(session.binding(), [], ["Manifest is not valid UTF-8 JSON."])
        if not isinstance(manifest, dict):
            return FolderPreview(session.binding(), [], ["Manifest is not an object."])
        if manifest.get("schema_version") != SCHEMA_VERSION:
            return FolderPreview(session.binding(), [], ["Unknown folder-export schema."])
        if manifest.get("baseline_sha256") != session.baseline_sha256:
            return FolderPreview(
                session.binding(),
                [],
                ["Folder export belongs to a different baseline. Export again or import modules individually."],
            )
        entries = manifest.get("entries")
        if not isinstance(entries, list):
            return FolderPreview(session.binding(), [], ["Manifest entries are missing."])

        changes: list[FolderChange] = []
        listed_ids: set[str] = set()
        listed_names: set[str] = set()
        listed_paths: set[str] = set()
        module_by_id = {m.id: m for m in session.draft.modules}
        aggregate = 0
        for index, entry in enumerate(entries):
            parsed = self._parse_entry(entry, index, problems)
            if parsed is None:
                continue
            export_id, rel, logical_name, kind, module_id = parsed
            if export_id in listed_ids:
                problems.append(f"Duplicate export id: {export_id}")
                continue
            listed_ids.add(export_id)
            name_key = logical_name.casefold()
            if name_key in listed_names:
                problems.append(f"Duplicate logical name: {logical_name}")
                continue
            listed_names.add(name_key)
            path_key = rel.casefold()
            if path_key in listed_paths:
                problems.append(f"Duplicate path: {rel}")
                continue
            listed_paths.add(path_key)
            if kind in _KIND_EXT:
                expected_ext = _KIND_EXT[kind]
                ok_ext = Path(rel).suffix.lower() == expected_ext
            else:
                ok_ext = rel.lower().endswith(".vba.txt")
            if not ok_ext:
                problems.append(f"Type/extension mismatch for {logical_name}: {rel}")
                continue
            if Path(rel).is_absolute() or ".." in Path(rel).parts or rel.startswith("/") or "\\" in Path(rel).anchor:
                problems.append(f"Unsafe path in manifest: {rel!r}")
                continue
            path = folder / rel
            if not is_inside(path, folder) or path_or_ancestor_is_reparse_point(path):
                problems.append(f"Path escaped the folder: {rel!r}")
                continue
            ancestor_rel = entry.get("ancestor_path") or f"{BASELINE_DIR}/{export_id}.txt"
            if not isinstance(ancestor_rel, str) or ".." in Path(ancestor_rel).parts:
                problems.append(f"Unsafe ancestor path for {export_id}")
                continue
            ancestor_path = folder / ancestor_rel
            if not is_inside(ancestor_path, folder):
                problems.append(f"Ancestor path escaped the folder: {ancestor_rel}")
                continue
            ancestor = None
            if ancestor_path.exists():
                try:
                    ancestor_bytes = ancestor_path.read_bytes()
                except OSError as exc:
                    problems.append(f"Ancestor could not be read: {type(exc).__name__}")
                    continue
                if len(ancestor_bytes) > MAX_FILE_BYTES:
                    problems.append(f"Ancestor for {export_id} exceeds 8 MiB.")
                    continue
                recorded_ancestor = entry.get("ancestor_sha256")
                if isinstance(recorded_ancestor, str) and sha256_bytes(ancestor_bytes) != recorded_ancestor:
                    problems.append(f"Corrupt ancestor hash for {export_id}")
                    continue
                try:
                    ancestor = ancestor_bytes.decode("utf-8")
                except UnicodeDecodeError:
                    problems.append(f"Ancestor for {export_id} is not valid UTF-8.")
                    continue
            module = module_by_id.get(module_id) if module_id else None
            if not path.exists():
                changes.append(
                    FolderChange(
                        export_id=export_id,
                        operation="delete",
                        logical_name=logical_name,
                        selected=False,
                        ancestor_text=ancestor,
                        kind=kind,
                    )
                )
                continue
            try:
                data = path.read_bytes()
            except OSError as exc:
                problems.append(f"{rel} could not be read: {type(exc).__name__}")
                continue
            if len(data) > MAX_FILE_BYTES:
                problems.append(f"{rel} exceeds 8 MiB.")
                continue
            aggregate += len(data)
            recorded_source = entry.get("source_sha256")
            if isinstance(recorded_source, str) and ancestor is not None:
                # Source hash is the last export; folder edits make it differ from disk. Do not require match.
                pass
            try:
                text = data.decode("utf-8")
            except UnicodeDecodeError:
                problems.append(f"{rel} is not valid UTF-8.")
                continue
            _header, _body = split_attribute_header(text)
            recorded_attr = entry.get("attribute_sha256")
            if isinstance(recorded_attr, str) and ancestor is not None:
                pass
            draft_text = None
            if module is not None:
                draft_text = compose_module_source(
                    body=module.body,
                    hidden_header=module.hidden_header,
                    current_name=module.current_name,
                    kind=module.kind,
                    pyopenvba_kind=module.pyopenvba_kind,
                )
            new_name = None
            operation = "noop"
            selected = False
            resolution = None
            if module is not None and logical_name.casefold() != module.current_name.casefold():
                operation = "rename"
                new_name = logical_name
                selected = False
            if ancestor is not None and text == ancestor and draft_text == ancestor:
                operation = "noop"
                selected = False
            elif ancestor is not None and text != ancestor and draft_text == ancestor:
                operation = "edit" if operation != "rename" else "rename"
                selected = operation == "edit"
            elif ancestor is not None and text != ancestor and draft_text != ancestor and text != draft_text:
                operation = "conflict"
                selected = False
                resolution = None
            changes.append(
                FolderChange(
                    export_id=export_id,
                    operation=operation,
                    logical_name=logical_name,
                    selected=selected,
                    folder_text=text,
                    draft_text=draft_text,
                    ancestor_text=ancestor,
                    resolution=resolution,
                    new_name=new_name,
                    kind=kind,
                )
            )
        for path in list(folder.glob("*.bas")) + list(folder.glob("*.cls")):
            self._maybe_add_unlisted(
                path, listed_ids, changes, problems, kind="class" if path.suffix.lower() == ".cls" else "standard"
            )
        for path in folder.glob("*.frm"):
            problems.append(f"{path.name}: UserForm files cannot be imported.")
        if aggregate > MAX_AGGREGATE:
            problems.append("Folder source exceeds 64 MiB.")
        return FolderPreview(session.binding(), changes, problems)

    def _parse_entry(self, entry, index: int, problems: list[str]) -> tuple[str, str, str, str, str] | None:
        if not isinstance(entry, dict):
            problems.append(f"Manifest entry {index} is not an object.")
            return None
        export_id = entry.get("export_id")
        rel = entry.get("relative_path")
        logical_name = entry.get("logical_name")
        kind = entry.get("kind") or "standard"
        module_id = entry.get("module_id") or ""
        if not isinstance(export_id, str) or not export_id:
            problems.append(f"Manifest entry {index} is missing export_id.")
            return None
        if not isinstance(rel, str) or not rel:
            problems.append(f"Manifest entry {index} is missing relative_path.")
            return None
        if not isinstance(logical_name, str) or not logical_name:
            problems.append(f"Manifest entry {index} is missing logical_name.")
            return None
        if not isinstance(kind, str):
            problems.append(f"Manifest entry {index} has an invalid kind.")
            return None
        if module_id and not isinstance(module_id, str):
            problems.append(f"Manifest entry {index} has an invalid module_id.")
            return None
        stem = Path(rel).stem.casefold()
        if any(stem == reserved or stem.startswith(reserved + ".") for reserved in _WINDOWS_RESERVED):
            problems.append(f"Reserved filename: {rel}")
            return None
        return export_id, rel, logical_name, kind, module_id

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
        try:
            text = path.read_text(encoding="utf-8")
        except (OSError, UnicodeDecodeError) as exc:
            problems.append(f"{name} could not be read: {type(exc).__name__}")
            return
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

    def apply(self, session: DocumentSession, preview: FolderPreview, *, history_service=None) -> None:
        if not session.matches_proposal(preview.binding):
            raise AdapterError("This folder preview is out of date.")
        if preview.problems:
            raise AdapterError("; ".join(preview.problems))
        for change in preview.changes:
            if change.operation == "conflict" and change.selected and change.resolution not in {"use_folder", "keep_draft"}:
                raise AdapterError(f"Conflict for {change.logical_name} needs an explicit choice.")
        before = snapshot_modules(session.draft)
        clone = clone_draft(session.draft)
        self._apply_to_draft(clone, preview)
        problems = validate_draft(clone)
        if problems:
            raise AdapterError("; ".join(problems))
        session.draft.modules[:] = clone.modules
        session.draft.xml_parts[:] = clone.xml_parts
        if history_service is not None:
            history_service.record_structural(
                session,
                HistoryCommand(
                    op="folder_sync",
                    target_id="*",
                    before=before,
                    after=snapshot_modules(session.draft),
                ),
            )
        else:
            session.bump()

    def _apply_to_draft(self, draft, preview: FolderPreview) -> None:
        for change in preview.changes:
            if change.operation == "noop":
                continue
            if change.operation == "conflict" and change.resolution != "use_folder":
                continue
            if not change.selected and change.operation != "conflict":
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
            if change.new_name and change.new_name != module.current_name:
                if not module.can_rename and not module.is_new:
                    raise AdapterError(f"Cannot rename {module.current_name} from a folder import.")
                problem = validate_module_name(change.new_name, draft, exclude_id=module.id)
                if problem:
                    raise AdapterError(problem)
                module.current_name = change.new_name
            if header and module.hidden_header:
                pass
