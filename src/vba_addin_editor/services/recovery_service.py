"""Draft autosave and crash recovery. Never writes the add-in (IMP-01)."""

from __future__ import annotations

import json
import os
import shutil
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import NoReturn

from vba_addin_editor.domain.document import (
    DocumentDraft,
    ModuleDraft,
    XmlPartDraft,
)
from vba_addin_editor.domain.results import SaveResult
from vba_addin_editor.domain.session import DocumentSession
from vba_addin_editor.platform import paths
from vba_addin_editor.platform.paths import sessions_root
from vba_addin_editor.platform.session_store import (
    baseline_package_path,
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
KNOWN_XML_ENCODINGS = {"utf-8", "utf-8-sig", "utf-16", "utf-16-le", "utf-16-be", "utf-16le", "utf-16be"}
ALLOWED_NEWLINES = {"\n", "\r\n"}


def _invalid(message: str) -> NoReturn:
    raise ValueError(message)


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


@dataclass
class RecoveryOpenResult:
    status: str  # clean | needs_conflict | missing_source | invalid
    session: DocumentSession | None
    checkpoint: dict | None
    listing: RecoveryListing | None
    reason: str | None = None
    external: object | None = None
    problems: tuple[str, ...] = ()
    used_fallback: bool = False


class RecoveryService:
    def __init__(self, *, session_root: Path | None = None, clock: Clock | None = None) -> None:
        self.session_root = session_root or sessions_root()
        self.clock = clock or Clock()

    def checkpoint(self, session: DocumentSession, *, cursor: dict | None = None) -> SaveResult | None:
        """Write a new generation then atomically publish the pointer."""
        try:
            if session.draft.is_dirty():
                self.clear_completion_marker(session)
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

    def clear_completion_marker(self, session: DocumentSession) -> None:
        marker = session.session_dir / "complete.marker"
        if marker.exists():
            try:
                marker.unlink()
            except OSError:
                pass

    def retain_named_checkpoint(self, session: DocumentSession, name: str = "pre_resolution") -> None:
        gen = self._current_generation(session.session_dir)
        if gen is None:
            return
        dest = session.session_dir / name
        if dest.exists():
            shutil.rmtree(dest, ignore_errors=True)
        src = session.session_dir / gen
        if src.exists():
            shutil.copytree(src, dest)

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
        data, _fallback = self.load_checkpoint_with_fallback(directory)
        return data

    def load_checkpoint_with_fallback(self, directory: Path) -> tuple[dict, bool]:
        pointer_gen = self._current_generation(directory)
        order: list[str] = []
        if pointer_gen:
            order.append(pointer_gen)
        for gen in GENERATIONS:
            if gen not in order:
                order.append(gen)
        last_error: Exception | None = None
        for index, gen in enumerate(order):
            path = directory / gen / "checkpoint.json"
            if not path.exists():
                continue
            try:
                raw = path.read_bytes()
                if len(raw) > METADATA_LIMIT:
                    raise ValueError("oversized metadata")
                data = json.loads(raw.decode("utf-8"))
                self._validate_checkpoint(data, directory)
                return data, index > 0
            except (OSError, ValueError, json.JSONDecodeError, KeyError, TypeError, UnicodeError) as exc:
                last_error = exc
                continue
        raise ValueError(
            f"No valid recovery generation is published ({type(last_error).__name__ if last_error else 'missing'})."
        )

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
        current = paths.fingerprint(source)
        self.apply_checkpoint(session, data)
        if current.sha256 == data["baseline_sha256"]:
            return "clean_reconstruct"
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
        bom_hex = item.get("bom") or ""
        return XmlPartDraft(
            path=item["path"],
            text=item.get("text"),
            original_text=item.get("original_text"),
            encoding=item.get("encoding"),
            bom=bytes.fromhex(bom_hex) if bom_hex else b"",
            newline=item.get("newline") or "\n",
            is_relationships_part=bool(item.get("is_relationships_part")),
            is_content_types_part=bool(item.get("is_content_types_part")),
            open_problem=item.get("open_problem"),
        )

    def revalidate_capabilities(self, draft: DocumentDraft) -> None:
        trusted = {module.id: module for module in draft.baseline.modules}
        trusted_stream = {
            (module.stream_name or module.original_name).casefold(): module
            for module in draft.baseline.modules
        }
        for module in draft.modules:
            snap = trusted.get(module.id)
            if snap is None and module.stream_name:
                snap = trusted_stream.get(module.stream_name.casefold())
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
            if not module.stream_name:
                module.stream_name = snap.stream_name
            if not module.hidden_header:
                module.hidden_header = snap.hidden_header
            if module.project_item_kind == "unknown":
                module.project_item_kind = snap.project_item_kind

    def _validate_session_meta(self, meta: dict, directory: Path) -> None:
        if not isinstance(meta, dict):
            _invalid("session metadata is not an object")
        if meta.get("schema_version") not in (None, SCHEMA_VERSION, 1):
            _invalid("unknown schema")
        session_id = meta.get("session_id")
        if not isinstance(session_id, str) or not session_id:
            _invalid("missing ids")
        for key in ("baseline_sha256", "source_path"):
            if key in meta and not isinstance(meta[key], str):
                _invalid(f"malformed {key}")
        if "baseline_size" in meta and not isinstance(meta["baseline_size"], int):
            _invalid("malformed baseline_size")
        if "baseline_sha256" in meta and (
            len(meta["baseline_sha256"]) != 64
            or any(c not in "0123456789abcdefABCDEF" for c in meta["baseline_sha256"])
        ):
            _invalid("wrong hashes")

    def _validate_checkpoint(self, data: dict, directory: Path) -> None:
        if not isinstance(data, dict):
            _invalid("checkpoint is not an object")
        if data.get("schema_version") != SCHEMA_VERSION:
            _invalid("unknown schema")
        if "session_id" not in data or "revision" not in data:
            _invalid("missing ids")
        if not isinstance(data["session_id"], str) or not isinstance(data["revision"], int):
            _invalid("missing ids")
        if data["revision"] < 0:
            _invalid("missing ids")
        if "draft" not in data or not isinstance(data["draft"], dict) or "modules" not in data["draft"]:
            _invalid("missing draft")
        modules = data["draft"]["modules"]
        if not isinstance(modules, list):
            _invalid("missing draft")
        ids = []
        for item in modules:
            if not isinstance(item, dict) or "id" not in item or "current_name" not in item:
                _invalid("missing ids")
            if not isinstance(item["id"], str):
                _invalid("missing ids")
            ids.append(item["id"])
            body = item.get("body")
            if body is not None and not isinstance(body, str):
                _invalid("invalid encoding")
        if len(ids) != len(set(ids)):
            _invalid("duplicate ids")
        xml_parts = data["draft"].get("xml_parts", [])
        if xml_parts is None:
            xml_parts = []
        if not isinstance(xml_parts, list):
            _invalid("missing draft")
        xml_paths = []
        for item in xml_parts:
            if not isinstance(item, dict) or "path" not in item:
                _invalid("missing draft")
            path = item["path"]
            if not isinstance(path, str) or not path or path.startswith("/") or ".." in Path(path).parts:
                _invalid("path traversal")
            xml_paths.append(path)
            encoding = item.get("encoding")
            if encoding is not None and not isinstance(encoding, str):
                _invalid("invalid BOM/encoding metadata")
            if encoding and encoding.lower() not in KNOWN_XML_ENCODINGS:
                _invalid("invalid BOM/encoding metadata")
            newline = item.get("newline")
            if newline is not None and newline not in ALLOWED_NEWLINES:
                _invalid("invalid BOM/encoding metadata")
            bom = item.get("bom") or ""
            if bom and (not isinstance(bom, str) or len(bom) % 2 or any(c not in "0123456789abcdefABCDEF" for c in bom)):
                _invalid("invalid BOM/encoding metadata")
        if len(xml_paths) != len(set(xml_paths)):
            _invalid("duplicate ids")
        sha = data.get("baseline_sha256")
        if not isinstance(sha, str) or len(sha) != 64:
            _invalid("wrong hashes")
        size = data.get("baseline_size")
        if not isinstance(size, int) or size < 0:
            _invalid("wrong lengths")
        locator = data.get("source_locator")
        if locator is not None and (
            not isinstance(locator, dict) or not isinstance(locator.get("path", ""), str)
        ):
            _invalid("path traversal")
        for payload_key in ("payloads", "files", "binaries"):
            refs = data.get(payload_key) or data.get("draft", {}).get(payload_key)
            if not refs:
                continue
            if not isinstance(refs, list):
                _invalid("path traversal")
            for ref in refs:
                rel = ref.get("path") if isinstance(ref, dict) else ref
                if not isinstance(rel, str):
                    _invalid("path traversal")
                resolved = (directory / rel).resolve()
                if not paths.is_inside(resolved, directory.resolve()):
                    _invalid("path traversal")

    def validate_captured_package(self, directory: Path, data: dict, extension: str) -> None:
        captured = baseline_package_path(directory, extension)
        if not captured.exists():
            raise ValueError("mismatched captured package")
        fp = paths.fingerprint(captured)
        if fp.sha256 != data["baseline_sha256"]:
            raise ValueError("mismatched captured package")
        if fp.size != int(data["baseline_size"]):
            raise ValueError("wrong lengths")

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
            data, _fallback = self.load_checkpoint_with_fallback(directory)
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
        except (OSError, ValueError, json.JSONDecodeError, KeyError, TypeError) as exc:
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


class RecoveryOpenService:
    """Reconstruct a draft from the checkpoint's captured baseline (IMP-01 / repair §2)."""

    def __init__(self, recovery: RecoveryService, sessions) -> None:
        self.recovery = recovery
        self.sessions = sessions

    def open_recoverable(self, directory: Path) -> RecoveryOpenResult:
        directory = Path(directory)
        listing = self.recovery._inspect(directory)
        try:
            meta = read_json(directory / "session.json")
            self.recovery._validate_session_meta(meta, directory)
            data, used_fallback = self.recovery.load_checkpoint_with_fallback(directory)
            if meta.get("session_id") and data.get("session_id") != meta["session_id"]:
                raise ValueError("missing ids")
            if meta.get("baseline_sha256") and meta["baseline_sha256"] != data.get("baseline_sha256"):
                raise ValueError("wrong hashes")
            extension = meta.get("extension") or Path(
                data.get("source_locator", {}).get("path") or meta.get("source_path") or ".xlam"
            ).suffix.lower()
            if not extension:
                extension = ".xlam"
            self.recovery.validate_captured_package(directory, data, extension)
            source = Path(data.get("source_locator", {}).get("path") or meta.get("source_path") or "")
            original_exists = bool(source.parts) and source.exists()
            status = "missing_source"
            external = None
            if original_exists:
                current = paths.fingerprint(source)
                if current.sha256 == data["baseline_sha256"] and current.size == int(data["baseline_size"]):
                    status = "clean"
                else:
                    status = "needs_conflict"
            session = self.sessions.attach_captured(
                session_id=str(meta.get("session_id") or data["session_id"]),
                directory=directory,
                source_path=source if source.parts else directory / f"missing{extension}",
                extension=extension,
                revision=int(data["revision"]),
                original_exists=status != "missing_source",
                baseline_generation=int(meta.get("baseline_generation") or 1),
            )
            self.recovery.apply_checkpoint(session, data)
            self.recovery.revalidate_capabilities(session.draft)
            if status == "needs_conflict":
                session.conflicts_pending = True
                try:
                    external = self.sessions.document_service.open(source).baseline
                except Exception as exc:  # noqa: BLE001 - unparseable E is a blocked outcome
                    return RecoveryOpenResult(
                        status="needs_conflict",
                        session=session,
                        checkpoint=data,
                        listing=listing,
                        reason="unparseable_external",
                        problems=(f"The current file could not be parsed ({type(exc).__name__}).",),
                        used_fallback=used_fallback,
                    )
            return RecoveryOpenResult(
                status=status,
                session=session,
                checkpoint=data,
                listing=listing,
                external=external,
                used_fallback=used_fallback,
            )
        except (OSError, ValueError, json.JSONDecodeError, KeyError, TypeError, UnicodeError) as exc:
            return RecoveryOpenResult(
                status="invalid",
                session=None,
                checkpoint=None,
                listing=listing,
                reason=str(exc) if str(exc) else type(exc).__name__,
                problems=(str(exc) or type(exc).__name__,),
            )
