"""Draft autosave and crash recovery. Never writes the add-in (IMP-01)."""

from __future__ import annotations

import json
import os
import shutil
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

from vba_addin_editor.domain.document import (
    DocumentDraft,
    ModuleDraft,
    XmlPartDraft,
)
from vba_addin_editor.domain.results import SaveResult
from vba_addin_editor.domain.session import DocumentSession
from vba_addin_editor.platform.paths import sessions_root
from vba_addin_editor.platform.session_store import (
    is_owned,
    read_json,
    write_json_atomic,
)

SCHEMA_VERSION = 1
METADATA_LIMIT = 64 * 1024 * 1024
STORAGE_WARN_BYTES = 500 * 1024 * 1024
IDLE_MS = 1000
MAX_INTERVAL_MS = 10000
GENERATIONS = ("gen_a", "gen_b")


class Clock:
    def now(self) -> datetime:
        return datetime.now(timezone.utc)

    def monotonic_ms(self) -> int:
        return int(os.times().elapsed * 1000) if hasattr(os, "times") else 0


@dataclass
class RecoveryListing:
    session_id: str
    directory: Path
    source_path: str
    filename: str
    checkpoint_utc: str
    revision: int
    dirty: bool
    invalid: bool
    reason: str | None = None


class RecoveryService:
    def __init__(self, *, session_root: Path | None = None, clock: Clock | None = None) -> None:
        self.session_root = session_root or sessions_root()
        self.clock = clock or Clock()

    def checkpoint(self, session: DocumentSession, *, cursor: dict | None = None) -> SaveResult | None:
        """Write a new generation then atomically publish the pointer."""
        try:
            payload = self._serialize(session, cursor=cursor)
            encoded = json.dumps(payload, indent=2, sort_keys=True).encode("utf-8")
            if len(encoded) > METADATA_LIMIT:
                return SaveResult.error(
                    "Draft recovery is larger than 64 MiB and was not written.",
                    reason="recovery_write_failed",
                    operation_type="recovery",
                )
            current = self._current_generation(session.session_dir)
            nxt = "gen_b" if current == "gen_a" else "gen_a"
            gen_dir = session.session_dir / nxt
            if gen_dir.exists():
                shutil.rmtree(gen_dir)
            gen_dir.mkdir(parents=True)
            checkpoint_path = gen_dir / "checkpoint.json"
            tmp = checkpoint_path.with_suffix(".json.tmp")
            with open(tmp, "wb") as handle:
                handle.write(encoded)
                handle.flush()
                os.fsync(handle.fileno())
            os.replace(tmp, checkpoint_path)
            write_json_atomic(session.session_dir / "current.json", {"generation": nxt})
            session.last_checkpoint_revision = session.revision
            session.recovery_generation = nxt
        except OSError as exc:
            return SaveResult.error(
                "Draft recovery could not be written. Editing stays enabled.",
                reason="recovery_write_failed",
                details={"exception_type": type(exc).__name__},
                operation_type="recovery",
            )
        return None

    def mark_complete(self, session: DocumentSession) -> None:
        write_json_atomic(
            session.session_dir / "complete.marker",
            {"utc": self.clock.now().isoformat(), "revision": session.revision},
        )

    def mark_discard(self, session: DocumentSession) -> None:
        write_json_atomic(
            session.session_dir / "discard.marker",
            {"utc": self.clock.now().isoformat(), "revision": session.revision},
        )

    def list_recoverable(self) -> list[RecoveryListing]:
        root = self.session_root
        if not root.exists():
            return []
        listings: list[RecoveryListing] = []
        for directory in sorted(root.iterdir()):
            if not directory.is_dir():
                continue
            if is_owned(directory):
                continue
            if (directory / "complete.marker").exists() or (directory / "discard.marker").exists():
                continue
            listing = self._inspect(directory)
            if listing is not None:
                listings.append(listing)
        return listings

    def load_checkpoint(self, directory: Path) -> dict:
        gen = self._current_generation(directory)
        if gen is None:
            raise ValueError("No recovery generation is published.")
        path = directory / gen / "checkpoint.json"
        data = json.loads(path.read_text(encoding="utf-8"))
        self._validate_checkpoint(data, directory)
        return data

    def apply_checkpoint(self, session: DocumentSession, data: dict) -> None:
        draft = session.draft
        restored_modules = [self._module_from_json(item) for item in data["draft"]["modules"]]
        restored_xml = [self._xml_from_json(item) for item in data["draft"].get("xml_parts", [])]
        draft.modules[:] = restored_modules
        draft.xml_parts[:] = restored_xml
        draft.signed_save_confirmed = False
        session.revision = int(data["revision"])
        session.history.clear()

    def reconstruct_if_disk_matches(self, session: DocumentSession, data: dict) -> str:
        """Return 'clean_reconstruct' | 'needs_conflict' | 'missing_source'."""
        source = Path(data["source_locator"]["path"])
        if not source.exists():
            return "missing_source"
        from vba_addin_editor.platform.paths import fingerprint

        current = fingerprint(source)
        if current.sha256 == data["baseline_sha256"]:
            self.apply_checkpoint(session, data)
            return "clean_reconstruct"
        self.apply_checkpoint(session, data)
        return "needs_conflict"

    def total_storage_bytes(self) -> int:
        root = self.session_root
        if not root.exists():
            return 0
        total = 0
        for dirpath, _names, filenames in os.walk(root):
            for name in filenames:
                try:
                    total += (Path(dirpath) / name).stat().st_size
                except OSError:
                    continue
        return total

    def delete_recovery(self, directory: Path) -> None:
        shutil.rmtree(directory, ignore_errors=True)

    def _serialize(self, session: DocumentSession, *, cursor: dict | None) -> dict:
        draft = session.draft
        now = self.clock.now().replace(microsecond=0).isoformat().replace("+00:00", "Z")
        return {
            "schema_version": SCHEMA_VERSION,
            "session_id": session.session_id,
            "revision": session.revision,
            "created_utc": now,
            "checkpoint_utc": now,
            "source_locator": {"path": str(session.original_path)},
            "baseline_sha256": session.baseline_sha256,
            "baseline_size": session.captured_fingerprint.size,
            "draft": {
                "signed_save_confirmed": False,
                "modules": [self._module_to_json(module) for module in draft.modules],
                "xml_parts": [self._xml_to_json(part) for part in draft.xml_parts],
                "cursor": cursor or {},
            },
        }

    def _module_to_json(self, module: ModuleDraft) -> dict:
        return {
            "id": module.id,
            "origin_name": module.origin_name,
            "current_name": module.current_name,
            "body": module.body,
            "kind": module.kind,
            "pyopenvba_kind": module.pyopenvba_kind,
            "is_new": module.is_new,
            "is_deleted": module.is_deleted,
            "original_body": module.original_body,
            "stream_name": module.stream_name,
            "hidden_header": module.hidden_header,
            "project_item_kind": module.project_item_kind,
        }

    def _xml_to_json(self, part: XmlPartDraft) -> dict:
        return {
            "path": part.path,
            "text": part.text,
            "original_text": part.original_text,
            "encoding": part.encoding,
            "bom": part.bom.hex(),
            "newline": part.newline,
            "is_relationships_part": part.is_relationships_part,
            "is_content_types_part": part.is_content_types_part,
            "open_problem": part.open_problem,
        }

    def _module_from_json(self, item: dict) -> ModuleDraft:
        # Capabilities are revalidated from the trusted snapshot, never from JSON.
        return ModuleDraft(
            id=item["id"],
            origin_name=item.get("origin_name"),
            current_name=item["current_name"],
            body=item.get("body") or "",
            kind=item.get("kind") or "standard",
            pyopenvba_kind=item.get("pyopenvba_kind") or "standard",
            is_new=bool(item.get("is_new")),
            is_deleted=bool(item.get("is_deleted")),
            destructive_ops_safe=False,
            original_body=item.get("original_body"),
            can_delete=False,
            can_rename=False,
            stream_name=item.get("stream_name") or "",
            hidden_header=item.get("hidden_header") or "",
            project_item_kind=item.get("project_item_kind") or "unknown",
        )

    def _xml_from_json(self, item: dict) -> XmlPartDraft:
        return XmlPartDraft(
            path=item["path"],
            text=item.get("text"),
            original_text=item.get("original_text"),
            encoding=item.get("encoding"),
            bom=bytes.fromhex(item["bom"]) if item.get("bom") else b"",
            newline=item.get("newline") or "\n",
            is_relationships_part=bool(item.get("is_relationships_part")),
            is_content_types_part=bool(item.get("is_content_types_part")),
            open_problem=item.get("open_problem"),
        )

    def revalidate_capabilities(self, draft: DocumentDraft) -> None:
        trusted = {module.id: module for module in draft.baseline.modules}
        for module in draft.modules:
            snap = trusted.get(module.id)
            if snap is None:
                if module.is_new:
                    module.can_delete = True
                    module.can_rename = True
                    module.destructive_ops_safe = True
                else:
                    module.can_delete = False
                    module.can_rename = False
                    module.destructive_ops_safe = False
                continue
            module.can_delete = snap.can_delete
            module.can_rename = snap.can_rename
            module.destructive_ops_safe = snap.destructive_ops_safe
            module.restriction_reason = snap.restriction_reason
            module.stream_name = snap.stream_name
            module.hidden_header = snap.hidden_header
            module.project_item_kind = snap.project_item_kind

    def _validate_checkpoint(self, data: dict, directory: Path) -> None:
        if data.get("schema_version") != SCHEMA_VERSION:
            raise ValueError("unknown schema")
        if "session_id" not in data or "revision" not in data:
            raise ValueError("missing ids")
        if "draft" not in data or "modules" not in data["draft"]:
            raise ValueError("missing draft")
        ids = [item["id"] for item in data["draft"]["modules"]]
        if len(ids) != len(set(ids)):
            raise ValueError("duplicate ids")

    def _current_generation(self, directory: Path) -> str | None:
        pointer = directory / "current.json"
        if not pointer.exists():
            return None
        try:
            data = read_json(pointer)
        except (OSError, json.JSONDecodeError, ValueError):
            return None
        gen = data.get("generation")
        if gen not in GENERATIONS:
            return None
        if not (directory / gen / "checkpoint.json").exists():
            return None
        return gen

    def _inspect(self, directory: Path) -> RecoveryListing | None:
        try:
            meta = read_json(directory / "session.json")
        except (OSError, json.JSONDecodeError, ValueError):
            return RecoveryListing(
                session_id=directory.name,
                directory=directory,
                source_path="",
                filename=directory.name,
                checkpoint_utc="",
                revision=0,
                dirty=True,
                invalid=True,
                reason="invalid metadata",
            )
        try:
            data = self.load_checkpoint(directory)
            source = data.get("source_locator", {}).get("path") or meta.get("source_path", "")
            return RecoveryListing(
                session_id=meta.get("session_id", directory.name),
                directory=directory,
                source_path=source,
                filename=Path(source).name if source else directory.name,
                checkpoint_utc=data.get("checkpoint_utc", ""),
                revision=int(data.get("revision", 0)),
                dirty=True,
                invalid=False,
            )
        except (OSError, ValueError, json.JSONDecodeError, KeyError) as exc:
            return RecoveryListing(
                session_id=meta.get("session_id", directory.name),
                directory=directory,
                source_path=meta.get("source_path", ""),
                filename=Path(meta.get("source_path", directory.name)).name,
                checkpoint_utc="",
                revision=0,
                dirty=True,
                invalid=True,
                reason=str(type(exc).__name__),
            )
