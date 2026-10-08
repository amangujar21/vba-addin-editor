"""Single seam between the application and pyOpenVBA.

All pyopenvba imports (public and internal) live here. GUI and services
must never import pyopenvba directly (plan sections 2.9, 5.3, 24).

Internals used and why (pinned pyopenvba==6.2.0):
- ``_encoding_for_codepage``: maps the project's dir-stream code page to a
  Python codec for strict save-time validation; no public equivalent exists.
- ``CFB``: reads streams and designer storages without mutating the host.
- ``_package_signature.without_signature``: the exact package edits
  ``host.save()`` makes when it drops a signature kept in parts beside
  ``vbaProject.bin``; verification allows those edits and nothing else.
- ``forms.form_names`` / ``forms.read_form``: UserForm designer storages,
  read for display and structural verification only. The editor never
  writes form layouts; ``host.forms()`` is never called on a host that
  saves, so ``save()`` cannot write a designer back.
- ``VBAHostFile._open_zip`` / ``_write_container``: overridden so a host opens
  from bytes already in memory and returns what ``save()`` produced instead of
  writing a file. Antivirus scans a freshly written package the first time it
  is opened for read, which takes seconds for a large project, so the save
  pipeline never reads back a file it wrote.
"""

from __future__ import annotations

import hashlib
import io
import os
import threading
import warnings
import zipfile
import zlib
from collections import OrderedDict
from pathlib import Path
from typing import Any

from pyopenvba import ExcelFile, PowerPointFile
from pyopenvba._host import VBAHostFile
from pyopenvba._package_signature import without_signature
from pyopenvba.cfb import CFB
from pyopenvba.exceptions import UnsupportedFormatError, VBAProjectError
from pyopenvba.forms import FormControl, form_names, read_form
from pyopenvba.vba import (
    VBAModuleKind,
    _encoding_for_codepage,
    parse_project_stream,
    parse_projectwm,
    serialize_project_stream,
)

from vba_addin_editor.adapters.ooxml_package_adapter import (
    OoxmlPackageAdapter,
    PackageError,
    PackageSource,
    open_package,
)
from vba_addin_editor.adapters.source_codec import (
    split_attribute_header,
    to_editor_text,
    to_vba_crlf,
)
from vba_addin_editor.domain.capabilities import (
    ComponentCapabilities,
    build_project_index,
    classify_component,
    empty_project_index,
)
from vba_addin_editor.domain.changes import compute_changes
from vba_addin_editor.domain.document import (
    DocumentDraft,
    DocumentSnapshot,
    FileFingerprint,
    FormControlSnapshot,
    FormDesignSnapshot,
    HostKind,
    ModuleDisplayKind,
    ModuleSnapshot,
    ProjectSafetyInfo,
)
from vba_addin_editor.domain.results import CandidateVerificationResult

SUPPORTED_EXTENSIONS = (".xlam", ".ppam", ".pptm")
_VBA_ENTRY = {
    ".xlam": "xl/vbaProject.bin",
    ".ppam": "ppt/vbaProject.bin",
    ".pptm": "ppt/vbaProject.bin",
}
_HOST_PROCESS = {
    ".xlam": "EXCEL.EXE",
    ".ppam": "POWERPNT.EXE",
    ".pptm": "POWERPNT.EXE",
}


class AdapterError(Exception):
    """Plain-language adapter failure; .details carries technical info."""

    def __init__(self, message: str, details: dict[str, Any] | None = None) -> None:
        super().__init__(message)
        self.details = details or {}


def _decrypt_project_data(value: str) -> bytes:
    """Decode an MS-OVBA Data Encryption value from CMG/DPB/GC.

    The encoded record carries its seed, version, project key, ignored-byte
    count, four-byte little-endian data length, and data. Protection state
    must be read from decoded CMG bits; raw DPB length is not a protection
    signal.
    """
    try:
        encoded = bytes.fromhex(value)
    except ValueError as exc:
        raise AdapterError("The VBA project protection metadata is malformed.") from exc
    if len(encoded) < 7:
        raise AdapterError("The VBA project protection metadata is truncated.")

    seed = encoded[0]
    if encoded[1] ^ seed != 2:
        raise AdapterError("The VBA project protection metadata has an unsupported version.")

    ignored_length = (seed & 0x06) // 2
    encrypted_byte_2 = encoded[1]
    encrypted_byte_1 = encoded[2]
    unencrypted_byte_1 = encoded[2] ^ seed
    decoded = bytearray()
    for encrypted_byte in encoded[3:]:
        byte = encrypted_byte ^ ((encrypted_byte_2 + unencrypted_byte_1) & 0xFF)
        decoded.append(byte)
        encrypted_byte_2 = encrypted_byte_1
        encrypted_byte_1 = encrypted_byte
        unencrypted_byte_1 = byte

    length_offset = ignored_length
    data_offset = length_offset + 4
    if len(decoded) < data_offset:
        raise AdapterError("The VBA project protection metadata has no data length.")
    data_length = int.from_bytes(decoded[length_offset:data_offset], "little")
    if len(decoded) != data_offset + data_length:
        raise AdapterError("The VBA project protection metadata has an invalid data length.")
    return bytes(decoded[data_offset:])


def _has_active_project_protection(protection: Any) -> bool:
    """Return whether decoded CMG says user, host, or VBE protection is active.

    pyOpenVBA (still in 6.2.0) derives ``has_password`` from raw DPB string length. DPB
    describes password material, while CMG is the authoritative active
    protection state. Fall back to pyOpenVBA only when CMG is absent; malformed
    or unexpected CMG data fails closed because its protection flags are unknown.
    """
    if protection is None:
        return False
    cmg = getattr(protection, "cmg", "")
    if not cmg:
        return bool(getattr(protection, "has_password", False))
    try:
        state = _decrypt_project_data(cmg)
    except AdapterError:
        return True
    if len(state) != 4:
        return True
    flags = int.from_bytes(state, "little")
    return bool(flags & 0x07)


class PPAMPowerPointFile(PowerPointFile):
    """Application-local PPAM compatibility subclass (plan 4.2, 24.2).

    PPAM is a ZIP OOXML package with ``ppt/vbaProject.bin``, same as PPTM;
    only the extension dispatch whitelist differs. Proven against authentic
    Office fixtures by tests/test_ppam_live.py (release gate).
    """

    _zip_formats = frozenset({".pptm", ".potm", ".ppam"})
    _no_vba_hint = "Make sure the PowerPoint add-in contains a VBA project."


def _host_class(path: Path) -> type[VBAHostFile]:
    suffix = path.suffix.lower()
    if suffix == ".xlam":
        return ExcelFile
    if suffix == ".ppam":
        return PPAMPowerPointFile
    if suffix == ".pptm":
        return PowerPointFile
    raise AdapterError("Only .xlam, .ppam, and .pptm files are supported.")


# A destination save() can never write: its folder does not exist. The
# in-memory host recognises it and returns the saved parts instead.
_CAPTURE_ONLY = Path("vbaae-capture-only") / "never-written"


class _InMemoryHost:
    """Mixin for the pyOpenVBA hosts: open from bytes, and capture a save."""

    def __init__(self, path: Path, data: bytes | None = None) -> None:
        self._preloaded = data
        self.saved_project: bytes | None = None
        self.saved_package_edits: dict[str, bytes | None] = {}
        super().__init__(path)  # type: ignore[call-arg]

    def _open_zip(self) -> None:
        if self._preloaded is None:
            super()._open_zip()  # type: ignore[misc]
            return
        self._container_raw = self._preloaded
        self._zip = zipfile.ZipFile(io.BytesIO(self._preloaded), mode="r")
        self._has_project = self._vba_entry in self._zip.namelist()  # type: ignore[attr-defined]

    def _write_container(self, dest, new_cfb_bytes, package_edits) -> None:
        if dest != _CAPTURE_ONLY:
            super()._write_container(dest, new_cfb_bytes, package_edits)  # type: ignore[misc]
            return
        self.saved_project = new_cfb_bytes
        self.saved_package_edits = dict(package_edits)


_MEMORY_HOSTS: dict[type, type] = {}


def _open_host(path: Path, data: bytes | None = None) -> Any:
    """Open the host for ``path``; with ``data``, parse those bytes and read no file."""
    base = _host_class(path)
    cls = _MEMORY_HOSTS.get(base)
    if cls is None:
        cls = _MEMORY_HOSTS[base] = type(f"InMemory{base.__name__}", (_InMemoryHost, base), {})
    return cls(path, data)


def _package_bytes(source: PackageSource) -> bytes:
    return source if isinstance(source, bytes) else Path(source).read_bytes()


class _SourceCache:
    """Decoded module sources, keyed by their compressed stream bytes.

    pyOpenVBA's decompressor is pure Python (seconds for a large project), and
    a save used to run it over every module three times. Decompression depends
    on nothing but the stream, so each distinct stream is decoded once: per
    run in memory, and across runs when a disk folder is set, which is what
    makes reopening an add-in fast.
    """

    def __init__(self, max_chars: int = 64_000_000) -> None:
        self._items: OrderedDict[tuple[bytes, str], str] = OrderedDict()
        self._chars = 0
        self._max_chars = max_chars
        self._lock = threading.Lock()
        self._disk: Path | None = None

    def use_disk(self, root: Path | None) -> None:
        self._disk = root

    def get(self, key: tuple[bytes, str]) -> str | None:
        with self._lock:
            text = self._items.get(key)
            if text is not None:
                self._items.move_to_end(key)
                return text
        if self._disk is None:
            return None
        text = _read_cached_source(self._disk_path(key))
        if text is not None:
            self._remember(key, text)
        return text

    def put(self, key: tuple[bytes, str], text: str) -> None:
        if self._remember(key, text) and self._disk is not None:
            _write_cached_source(self._disk_path(key), text)

    def _remember(self, key: tuple[bytes, str], text: str) -> bool:
        with self._lock:
            if key in self._items:
                return False
            self._items[key] = text
            self._chars += len(text)
            while self._chars > self._max_chars and len(self._items) > 1:
                _old, dropped = self._items.popitem(last=False)
                self._chars -= len(dropped)
            return True

    def _disk_path(self, key: tuple[bytes, str]) -> Path:
        assert self._disk is not None
        name = key[0].hex()
        return self._disk / name[:2] / f"{name}.{key[1]}.z"


def _read_cached_source(path: Path) -> str | None:
    """A cached source, or None when absent or not exactly what was stored."""
    try:
        blob = path.read_bytes()
        encoded = zlib.decompress(blob[32:])
        if hashlib.sha256(encoded).digest() != blob[:32]:
            return None
        return encoded.decode("utf-8", errors="surrogatepass")
    except (OSError, zlib.error, UnicodeError):
        return None


def _write_cached_source(path: Path, text: str) -> None:
    try:
        encoded = text.encode("utf-8", errors="surrogatepass")
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_name(f"{path.name}.{threading.get_ident()}.tmp")
        tmp.write_bytes(hashlib.sha256(encoded).digest() + zlib.compress(encoded, 1))
        os.replace(tmp, path)
    except (OSError, UnicodeError):
        pass  # a cache that cannot be written only costs time


def _prune_source_cache(root: Path, max_bytes: int) -> None:
    """Drop the oldest cached sources once the folder outgrows ``max_bytes``."""
    try:
        stats = [(f, f.stat()) for f in root.glob("*/*.z")]
    except OSError:
        return
    files = sorted((st.st_mtime, st.st_size, f) for f, st in stats)
    total = sum(size for _mtime, size, _f in files)
    for _mtime, size, file in files:
        if total <= max_bytes * 3 // 4:
            break
        try:
            file.unlink()
            total -= size
        except OSError:
            pass


_SOURCES = _SourceCache()


def enable_source_disk_cache(root: Path, max_bytes: int = 200 * 1024 * 1024) -> None:
    """Keep decoded module sources under ``root`` between runs."""
    _SOURCES.use_disk(Path(root))
    threading.Thread(
        target=_prune_source_cache, args=(Path(root), max_bytes), name="vbaae-cache-prune", daemon=True
    ).start()


def _module_source(cfb: CFB, module: Any, encoding: str, *, load: bool = True) -> str | None:
    """Full source of a parsed module; ``load=False`` answers only from the cache."""
    if module.source_loaded:
        return module.source
    stream_name = module.stream_name or module.name
    try:
        try:
            stream = cfb.get_stream_in_storage("VBA", stream_name)
        except KeyError:
            stream = cfb.get_stream(stream_name)
    except KeyError:
        return module.source if load else None
    key = (hashlib.sha256(stream[module.text_offset:]).digest(), encoding)
    text = _SOURCES.get(key)
    if text is not None:
        module.source = text  # spares pyOpenVBA its own decompression
        return text
    if not load:
        return None
    text = module.source
    _SOURCES.put(key, text)
    return text


def vba_entry_for(path: Path) -> str:
    return _VBA_ENTRY[path.suffix.lower()]


def host_process_for(path: Path) -> str:
    return _HOST_PROCESS[path.suffix.lower()]


def host_kind_for(path: Path) -> HostKind:
    suffix = path.suffix.lower()
    if suffix == ".xlam":
        return "excel"
    if suffix in {".ppam", ".pptm"}:
        return "powerpoint"
    raise AdapterError("Only .xlam, .ppam, and .pptm files are supported.")


def make_display_kind(pyopenvba_kind: str, *, is_new_app_class: bool = False) -> str:
    if pyopenvba_kind == "standard":
        return ModuleDisplayKind.STANDARD
    return ModuleDisplayKind.CLASS if is_new_app_class else ModuleDisplayKind.AMBIGUOUS


def _project_bytes_lossy(raw: bytes, code_page: int) -> tuple[object | None, bool]:
    """Parse PROJECT and report whether decoding required replacement characters."""
    encoding = _encoding_for_codepage(code_page)
    lossy = False
    try:
        raw.decode(encoding, errors="strict")
    except UnicodeDecodeError:
        lossy = True
    try:
        parsed = parse_project_stream(raw, code_page=code_page)
    except Exception:  # noqa: BLE001 - fail closed at the classifier
        return None, True
    return parsed, lossy


def _designer_storage_names(cfb: CFB) -> frozenset[str]:
    try:
        storages = cfb.list_storages()
    except Exception:  # noqa: BLE001
        return frozenset()
    return frozenset(name.casefold() for name in storages if name.upper() != "VBA")


def _form_storage_names(cfb: CFB) -> list[str]:
    """Top-level designer storages holding a UserForm ``f`` stream."""
    try:
        return form_names(cfb)
    except Exception:  # noqa: BLE001 - unreadable => no deletable forms
        return []


def _display_value(value: object) -> str:
    size = getattr(value, "width", None), getattr(value, "height", None)
    if None not in size:
        return f"{size[0]} x {size[1]} (HIMETRIC)"
    return str(value)


def _display_properties(props: dict[str, object]) -> tuple[tuple[str, str], ...]:
    return tuple((name, _display_value(value)) for name, value in props.items())


def _control_snapshot(control: FormControl) -> FormControlSnapshot:
    return FormControlSnapshot(
        name=control.name,
        kind=control.kind,
        properties=_display_properties(control.properties()),
        children=tuple(_control_snapshot(child) for child in control.children),
    )


def _read_form_designs(cfb: CFB, names: list[str], code_page: int) -> tuple[FormDesignSnapshot, ...]:
    """Display-only form layouts; a form that fails to parse reports why."""
    designs: list[FormDesignSnapshot] = []
    for name in names:
        try:
            form = read_form(cfb, name, code_page=code_page)
            designs.append(
                FormDesignSnapshot(
                    name=name,
                    properties=_display_properties(form.properties()),
                    controls=tuple(_control_snapshot(c) for c in form.controls),
                )
            )
        except Exception as exc:  # noqa: BLE001 - shown to the user, never fatal
            designs.append(FormDesignSnapshot(name=name, problem=f"{type(exc).__name__}: {exc}"))
    return tuple(designs)


def _storage_tree(cfb: CFB, path: list[str]) -> dict[str, bytes]:
    """Every stream under the storage at ``path``, keyed by relative path."""
    out: dict[str, bytes] = {}
    for stream in cfb.list_streams_at(path):
        out["/".join([*path, stream])] = cfb.get_stream_at(path, stream)
    for storage in cfb.list_storages_at(path):
        out.update(_storage_tree(cfb, [*path, storage]))
    return out


def _index_from_project(parsed: object | None, *, lossy: bool, missing: bool) -> object:
    if parsed is None:
        return empty_project_index(missing=missing, lossy=lossy)
    documents = [name for name, _hid in getattr(parsed, "document_modules", [])]
    return build_project_index(
        standard_modules=list(getattr(parsed, "standard_modules", [])),
        class_modules=list(getattr(parsed, "class_modules", [])),
        document_modules=documents,
        base_classes=list(getattr(parsed, "base_classes", [])),
        lossy=lossy,
        missing=missing,
    )


def classify_host_module(
    module: object,
    *,
    index,
    designer_storages: frozenset[str],
    form_storages: frozenset[str] = frozenset(),
) -> ComponentCapabilities:
    return classify_component(
        logical_name=module.name,
        stream_name=module.stream_name or module.name,
        dir_kind=module.kind.name,
        is_read_only=bool(module.is_read_only),
        index=index,
        designer_storages=designer_storages,
        form_storages=form_storages,
    )


class PyOpenVBAAdapter:
    def __init__(self, package_adapter: OoxmlPackageAdapter | None = None) -> None:
        self.package_adapter = package_adapter or OoxmlPackageAdapter()

    # -- safety inspection ------------------------------------------------

    def inspect_signature(self, host: VBAHostFile) -> tuple[bool, tuple[str, ...]]:
        # Streams inside the project plus the package parts beside it, which
        # is where Office itself keeps a VBA signature.
        info = host.vba_signature()
        return info.present, tuple(info.kinds)

    # -- snapshot ---------------------------------------------------------

    def open_snapshot(
        self,
        path: Path,
        fingerprint: FileFingerprint,
        host_process_running: bool = False,
        *,
        data: bytes | None = None,
    ) -> DocumentSnapshot:
        """Parse the package at ``path``, or ``data`` when the caller already
        holds the bytes ``fingerprint`` describes (nothing is read from disk)."""
        path = Path(path)
        suffix = path.suffix.lower()
        if suffix not in SUPPORTED_EXTENSIONS:
            raise AdapterError("Only .xlam, .ppam, and .pptm files are supported.")
        try:
            host = _open_host(path, data)
        except UnsupportedFormatError as exc:
            raise AdapterError(
                "Only .xlam, .ppam, and .pptm files are supported.", {"exception": repr(exc)}
            ) from exc
        except VBAProjectError as exc:
            raise AdapterError(
                "This add-in does not contain a readable VBA project. The file has not been changed.",
                {"exception": repr(exc)},
            ) from exc
        except Exception as exc:
            raise AdapterError(
                "The Office add-in package appears damaged or is not a valid add-in file.",
                {"exception": repr(exc)},
            ) from exc
        with host:
            try:
                project = host.vba_project()
            except Exception as exc:
                raise AdapterError(
                    "The VBA project inside this add-in could not be read.",
                    {"exception": repr(exc)},
                ) from exc
            sig_present, sig_kinds = self.inspect_signature(host)
            cfb = CFB.from_bytes(host.vba_project_bytes())
            try:
                project_raw = cfb.get_stream("PROJECT")
                parsed_project, lossy = _project_bytes_lossy(project_raw, project.code_page)
                missing_project = parsed_project is None
            except KeyError:
                parsed_project, lossy, missing_project = None, False, True
            type_index = _index_from_project(
                parsed_project, lossy=lossy, missing=missing_project
            )
            designer_storages = _designer_storage_names(cfb)
            form_storage_list = _form_storage_names(cfb)
            form_storages = frozenset(name.casefold() for name in form_storage_list)
            forms = _read_form_designs(cfb, form_storage_list, project.code_page)
            safety = ProjectSafetyInfo(
                password_protected=_has_active_project_protection(project.protection),
                signature_present=sig_present,
                signature_kinds=sig_kinds,
                writable=True,  # refined by save service preflight
                host_process_running=host_process_running,
            )
            modules: list[ModuleSnapshot] = []
            encoding = _encoding_for_codepage(project.code_page)
            for m in project.modules:
                source = to_editor_text(_module_source(cfb, m, encoding) or "")
                header, body = split_attribute_header(source)
                caps = classify_host_module(
                    m,
                    index=type_index,
                    designer_storages=designer_storages,
                    form_storages=form_storages,
                )
                modules.append(
                    ModuleSnapshot(
                        id=f"{m.stream_name}:{len(modules)}",
                        original_name=m.name,
                        kind=caps.display_kind,
                        pyopenvba_kind=m.kind.name,
                        full_source=source,
                        hidden_header=header,
                        body=body,
                        is_read_only=m.is_read_only,
                        is_private=m.is_private,
                        destructive_ops_safe=caps.destructive_ops_safe,
                        ends_with_newline=body.endswith("\n"),
                        can_delete=caps.can_delete,
                        can_rename=caps.can_rename,
                        restriction_reason=caps.restriction_reason,
                        stream_name=m.stream_name or m.name,
                        project_item_kind=caps.project_item_kind,
                    )
                )
            return DocumentSnapshot(
                path=path,
                host_kind=host_kind_for(path),
                extension=suffix,  # type: ignore[arg-type]
                file_fingerprint=fingerprint,
                project_name=project.name,
                code_page=project.code_page,
                safety=safety,
                modules=tuple(modules),
                forms=forms,
                package_bytes=data,
            )

    def codepage_encoding(self, code_page: int) -> str:
        return _encoding_for_codepage(code_page)

    # -- candidate build (plan 14 Stage G/H/I, replay algorithm 25) -------

    def build_candidate(
        self,
        original_path: Path,
        draft: DocumentDraft,
        candidate_path: Path,
        *,
        allow_signature_removal: bool,
    ) -> None:
        """Open the ORIGINAL fresh, replay the draft onto it, serialize to candidate."""
        built = self.build_candidate_bytes(
            Path(original_path), draft, allow_signature_removal=allow_signature_removal
        )
        Path(candidate_path).write_bytes(built)

    def build_candidate_bytes(
        self,
        source: PackageSource,
        draft: DocumentDraft,
        *,
        allow_signature_removal: bool,
        xml_replacements: dict[str, bytes] | None = None,
    ) -> bytes:
        """Replay the draft onto a fresh parse of ``source``; return the new package.

        ``xml_replacements`` (member name -> encoded bytes) are written in the
        same pass. Nothing is written to disk.
        """
        _assert_destructive_ops_against_snapshot(draft)
        host_path = draft.baseline.path if isinstance(source, bytes) else Path(source)
        data = _package_bytes(source)
        host = _open_host(host_path, data)
        with host:
            project = host.vba_project()
            encoding = _encoding_for_codepage(project.code_page)
            for module in project.modules:
                _module_source(host._get_cfb(), module, encoding, load=False)
            try:
                # 1. Delete permitted original modules gone in final state.
                for mod in draft.deleted_original_modules():
                    project.delete_module(mod.origin_name or mod.current_name)

                # 2. Renames via unique temporary names to avoid collisions.
                rename_items = draft.changed_names()
                temp_map: dict[str, str] = {}
                for item in rename_items:
                    temp = _make_unique_temp_name(project)
                    project.rename_module(item.origin_name or item.current_name, temp)
                    temp_map[item.id] = temp
                for item in rename_items:
                    project.rename_module(temp_map[item.id], item.current_name)

                # 3. Add new modules.
                for item in draft.new_modules():
                    kind = (
                        VBAModuleKind.standard
                        if item.pyopenvba_kind == "standard"
                        else VBAModuleKind.other
                    )
                    project.add_module(item.current_name, to_vba_crlf(item.body), kind=kind)

                # 4. Body changes on surviving originals.
                for item in draft.changed_existing_modules():
                    host.set_module(item.current_name, to_vba_crlf(item.body))

                with warnings.catch_warnings(record=True) as caught:
                    warnings.simplefilter("always")
                    host.save(
                        _CAPTURE_ONLY,
                        allow_invalidate_signature=allow_signature_removal,
                        # pyOpenVBA (3.4.0-6.2.0) can falsely infer protection from raw
                        # DPB length. Override only when decoded CMG proves the
                        # project is unlocked; active/malformed protection stays
                        # blocked by both the service preflight and the library.
                        allow_protected=not _has_active_project_protection(project.protection),
                    )
                unexpected = [str(w.message) for w in caught if not allow_signature_removal]
                if unexpected:
                    raise AdapterError(
                        "The VBA project signature was dropped without confirmation; save aborted.",
                        {"warnings": unexpected},
                    )
            except ValueError as exc:
                raise AdapterError(str(exc)) from exc
        project_bytes = host.saved_project
        if project_bytes is None:
            raise AdapterError("pyOpenVBA did not return a VBA project to save.")
        surviving = {mod.current_name.casefold() for mod in draft.final_module_state()}
        logical_deletes = [
            mod.origin_name or mod.current_name
            for mod in draft.deleted_original_modules()
            if (mod.origin_name or mod.current_name).casefold() not in surviving
        ]
        if logical_deletes:
            project_bytes = _scrub_project_declarations(
                project_bytes, logical_deletes, draft.baseline.code_page
            )
        replacements: dict[str, bytes | None] = {
            **host.saved_package_edits,
            vba_entry_for(host_path): project_bytes,
            **(xml_replacements or {}),
        }
        try:
            return self.package_adapter.build_package(data, replacements)
        except PackageError as exc:
            raise AdapterError(str(exc), getattr(exc, "details", None)) from exc

    # -- candidate verification (plan 14 Stage J, 77) ----------------------

    def verify_candidate(
        self,
        reference_path: PackageSource,
        candidate_path: PackageSource,
        expected: DocumentDraft,
        *,
        allowed_non_vba_changes: frozenset[str] = frozenset(),
    ) -> CandidateVerificationResult:
        """Verify the candidate against a reference package.

        Either may be a path or the package bytes; the save pipeline passes
        bytes so nothing it wrote is read back from disk.
        """
        problems: list[str] = []
        details: dict[str, Any] = {}
        host_path = expected.baseline.path
        if isinstance(candidate_path, Path):
            if not candidate_path.exists() or candidate_path.stat().st_size == 0:
                return CandidateVerificationResult(False, ("Candidate file missing or empty.",))
            host_path = candidate_path
            if (
                isinstance(reference_path, Path)
                and candidate_path.suffix.lower() != reference_path.suffix.lower()
            ):
                problems.append("Candidate extension does not match original.")
        entry = vba_entry_for(reference_path if isinstance(reference_path, Path) else host_path)
        cand_data = _package_bytes(candidate_path)
        ref_data = _package_bytes(reference_path)
        if not cand_data:
            return CandidateVerificationResult(False, ("Candidate file missing or empty.",))
        ref_vba: bytes | None = None
        try:
            with open_package(cand_data) as cand, open_package(ref_data) as ref:
                if cand.testzip() is not None:
                    problems.append("Candidate ZIP has a corrupt entry.")
                cand_names = set(cand.namelist())
                ref_names = set(ref.namelist())
                ref_vba = ref.read(entry) if entry in ref_names else None
                # An approved signature drop removes the signature parts and
                # rewrites their relationships / content-type overrides, as
                # Office does. Exactly those edits are allowed; nothing else.
                sig_edits = _signature_removal_edits(ref, entry, expected)
                removed = {name for name, data in sig_edits.items() if data is None}
                expected_names = ref_names - removed
                if cand_names != expected_names:
                    problems.append(
                        f"ZIP entry set changed: only-in-original={sorted(expected_names - cand_names)} "
                        f"only-in-candidate={sorted(cand_names - expected_names)}"
                    )
                for name, data in sig_edits.items():
                    if data is None or name not in cand_names:
                        continue
                    if name in allowed_non_vba_changes:
                        problems.append(
                            f"{name} was edited and is also rewritten by removing the VBA "
                            "signature; save the XML and VBA changes separately."
                        )
                    elif cand.read(name) != data:
                        problems.append(f"Signature removal changed {name} unexpectedly.")
                differing: list[str] = []
                for name in sorted(ref_names & cand_names):
                    if name == entry or name in allowed_non_vba_changes or name in sig_edits:
                        continue
                    if (
                        hashlib.sha256(ref.read(name)).digest()
                        != hashlib.sha256(cand.read(name)).digest()
                    ):
                        differing.append(name)
                if differing:
                    problems.append(f"Non-VBA package payload changed: {differing}")
        except zipfile.BadZipFile as exc:
            return CandidateVerificationResult(
                False, ("Candidate is not a valid Office package.", {"exception": repr(exc)})
            )

        expected_final = {m.current_name.casefold(): m for m in expected.final_module_state()}
        try:
            cand_host = _open_host(host_path, cand_data)
        except Exception as exc:  # noqa: BLE001 - translated below
            return CandidateVerificationResult(
                False, ("Candidate could not be reopened.", {"exception": repr(exc)})
            )
        with cand_host:
            project = cand_host.vba_project()
            cand_cfb = cand_host._get_cfb()
            encoding = _encoding_for_codepage(project.code_page)
            structural = project.validate(cand_cfb)
            if structural:
                problems.extend(f"Structural validation: {s}" for s in structural)
            if project.code_page != expected.baseline.code_page:
                problems.append(
                    f"Code page changed: {expected.baseline.code_page} -> {project.code_page}"
                )
            actual = {m.name.casefold(): m for m in project.modules}
            for name in sorted(set(actual) - set(expected_final)):
                problems.append(f"Unexpected module in candidate: {name!r}")
            for name, mod in expected_final.items():
                if name not in actual:
                    problems.append(f"Expected module missing in candidate: {mod.current_name!r}")
                    continue
                got = _module_source(cand_cfb, actual[name], encoding) or ""
                got_body = to_vba_crlf(split_attribute_header(to_editor_text(got))[1])
                want = to_vba_crlf(mod.body)
                if got_body != want:
                    problems.append(f"Body mismatch for module {mod.current_name!r}.")
                expected_kind = (
                    "standard" if mod.pyopenvba_kind == "standard" else "other"
                )
                if actual[name].kind.name != expected_kind:
                    problems.append(
                        f"Module kind mismatch for {mod.current_name!r}: "
                        f"{actual[name].kind.name} != {expected_kind}."
                    )
            problems.extend(_verify_project_cleanup(cand_host, expected))
            if ref_vba is not None:
                problems.extend(
                    _verify_form_storages(
                        CFB.from_bytes(ref_vba),
                        CFB.from_bytes(cand_host.vba_project_bytes()),
                        expected,
                    )
                )
            sig_present, kinds = self.inspect_signature(cand_host)
            if expected.baseline.safety.signature_present:
                if sig_present and allow_removal_expected(expected):
                    problems.append("Signature still present after approved removal.")
                details["candidate_signature_kinds"] = list(kinds)
            else:
                if sig_present:
                    problems.append("Candidate unexpectedly gained a signature.")
        return CandidateVerificationResult(not problems, tuple(problems), details)

    def signature_removal_edits(
        self, reference_path: PackageSource, draft: DocumentDraft
    ) -> dict[str, bytes | None]:
        """Package edits an approved signature drop makes (None = part removed)."""
        entry_path = reference_path if isinstance(reference_path, Path) else draft.baseline.path
        try:
            with open_package(reference_path) as ref:
                return _signature_removal_edits(ref, vba_entry_for(entry_path), draft)
        except (zipfile.BadZipFile, OSError):
            return {}

    # -- import/export -----------------------------------------------------

    def export_module_source(self, snapshot: DocumentSnapshot, module_name: str) -> str:
        mod = next(
            (m for m in snapshot.modules if m.original_name.casefold() == module_name.casefold()),
            None,
        )
        if mod is None:
            raise KeyError(module_name)
        return to_vba_crlf(mod.body if mod.kind == ModuleDisplayKind.STANDARD else mod.full_source)

    def parse_import_file(self, path: Path) -> tuple[str, str, bool]:
        """Return (suggested_name, body_or_full_source, is_class_form)."""
        raw = path.read_bytes()
        if len(raw) > 8 * 1024 * 1024:
            raise AdapterError("Import file is unreasonably large.")
        text = to_editor_text(raw.decode("utf-8", errors="strict"))
        header, _ = split_attribute_header(text)
        is_class = "VERSION 1.0 CLASS" in header
        name = path.stem
        return name, text, is_class


def allow_removal_expected(draft: DocumentDraft) -> bool:
    return draft.signed_save_confirmed


def _signature_removal_edits(
    ref: zipfile.ZipFile, entry: str, draft: DocumentDraft
) -> dict[str, bytes | None]:
    """What host.save() does to the package when it drops a signature kept in
    parts beside vbaProject.bin, as Office does; empty unless approved."""
    if not (draft.baseline.safety.signature_present and allow_removal_expected(draft)):
        return {}
    if not compute_changes(draft).has_vba_changes:
        return {}  # XML-only saves never rewrite the VBA project
    return without_signature(ref.namelist(), ref.read, entry)


def _assert_destructive_ops_against_snapshot(draft: DocumentDraft) -> None:
    trusted = {module.id: module for module in draft.baseline.modules}
    for mod in draft.deleted_original_modules():
        snap = trusted.get(mod.id)
        if snap is None or not snap.can_delete:
            reason = getattr(snap, "restriction_reason", None) if snap is not None else "unknown"
            raise AdapterError(
                f"Deleting {mod.origin_name or mod.current_name!r} is not permitted "
                f"({reason or 'unverified component'}).",
                {"module_id": mod.id, "reason": reason},
            )
    # pyOpenVBA recycles a deleted stream when a module of the same name is
    # added in the same save, which would leave a deleted form's layout
    # storage in place under the new module.
    final_names = {m.current_name.casefold() for m in draft.final_module_state()}
    for mod in draft.deleted_original_modules():
        snap = trusted.get(mod.id)
        name = mod.origin_name or mod.current_name
        if snap is not None and snap.project_item_kind == "userform" and name.casefold() in final_names:
            raise AdapterError(
                f"The UserForm {name!r} is deleted in this draft and its name is reused. "
                "Save the deletion before reusing the name.",
                {"module_id": mod.id},
            )
    for mod in draft.changed_names():
        snap = trusted.get(mod.id)
        if snap is None or not snap.can_rename:
            reason = getattr(snap, "restriction_reason", None) if snap is not None else "unknown"
            raise AdapterError(
                f"Renaming {mod.origin_name or mod.current_name!r} is not permitted "
                f"({reason or 'unverified component'}).",
                {"module_id": mod.id, "reason": reason},
            )


def _scrub_project_declarations(
    project_bytes: bytes, logical_names: list[str], code_page: int
) -> bytes:
    """Shim pyOpenVBA 3.4.0: save() deletes PROJECT lines by stream name.

    serialize_project_stream expects logical names. When they differ, Class=
    declarations leak. Rewrite PROJECT with the logical names after save.
    """
    cfb = CFB.from_bytes(project_bytes)
    try:
        project_raw = cfb.get_stream("PROJECT")
    except KeyError:
        return project_bytes
    rewritten = serialize_project_stream(
        project_raw,
        {},
        delete_names=set(logical_names),
        code_page=code_page,
    )
    cfb.write_stream("PROJECT", rewritten)
    return cfb.to_bytes()


def _verify_project_cleanup(host: VBAHostFile, expected: DocumentDraft) -> list[str]:
    problems: list[str] = []
    cfb = CFB.from_bytes(host.vba_project_bytes())
    expected_names = {
        module.current_name.casefold() for module in expected.final_module_state()
    }
    surviving_streams = set(expected_names)
    for module in expected.final_module_state():
        if module.stream_name:
            surviving_streams.add(module.stream_name.casefold())
    deleted_names = {
        (module.origin_name or module.current_name).casefold()
        for module in expected.deleted_original_modules()
        if (module.origin_name or module.current_name).casefold() not in expected_names
    }
    deleted_streams = {
        (module.stream_name or module.origin_name or module.current_name)
        for module in expected.deleted_original_modules()
        if (module.stream_name or module.origin_name or module.current_name).casefold()
        not in surviving_streams
    }
    try:
        parsed, lossy = _project_bytes_lossy(cfb.get_stream("PROJECT"), expected.baseline.code_page)
    except KeyError:
        problems.append("Candidate PROJECT stream is missing.")
        return problems
    if parsed is None:
        problems.append("Candidate PROJECT stream could not be parsed.")
        return problems
    if lossy:
        problems.append("Candidate PROJECT stream decoded with replacement characters.")
    declared = {name.casefold() for name in parsed.standard_modules}
    declared.update(name.casefold() for name in parsed.class_modules)
    declared.update(name.casefold() for name, _hid in parsed.document_modules)
    declared.update(name.casefold() for name in parsed.base_classes)
    for name in sorted(deleted_names):
        if name in declared:
            problems.append(f"Deleted module still declared in PROJECT: {name!r}")
    dir_names = {module.name.casefold() for module in host.vba_project().modules}
    for name in sorted(deleted_names):
        if name in dir_names:
            problems.append(f"Deleted module still listed in dir: {name!r}")
    try:
        vba_streams = {item.casefold() for item in cfb.list_streams_in_storage("VBA")}
    except Exception as exc:  # noqa: BLE001 - inability to verify is failure
        problems.append(f"Could not enumerate VBA streams ({type(exc).__name__}).")
        return problems
    for stream in deleted_streams:
        if stream and stream.casefold() in vba_streams:
            problems.append(f"Deleted module stream still present: {stream!r}")
    try:
        wm_raw = cfb.get_stream("PROJECTwm")
    except KeyError:
        if deleted_names:
            problems.append("Candidate PROJECTwm stream is missing.")
        return problems
    try:
        wm = parse_projectwm(wm_raw, code_page=expected.baseline.code_page)
    except (OSError, UnicodeError, ValueError, TypeError) as exc:
        problems.append(f"Candidate PROJECTwm stream could not be parsed ({type(exc).__name__}).")
        return problems
    wm_names = {item[0].casefold() for item in wm} if wm else set()
    for name in sorted(deleted_names):
        if name in wm_names:
            problems.append(f"Deleted module still listed in PROJECTwm: {name!r}")
    return problems


def _verify_form_storages(ref_cfb: CFB, cand_cfb: CFB, expected: DocumentDraft) -> list[str]:
    """Deleted forms leave no storage; every other form layout is byte-identical.

    The editor never writes layouts, so any designer difference is a defect.
    """
    problems: list[str] = []
    surviving = {m.current_name.casefold() for m in expected.final_module_state()}
    ref_forms = {name.casefold(): name for name in _form_storage_names(ref_cfb)}
    try:
        cand_storages = {name.casefold(): name for name in cand_cfb.list_storages_at(())}
    except Exception as exc:  # noqa: BLE001 - inability to verify is failure
        return [f"Could not enumerate candidate designer storages ({type(exc).__name__})."]
    for folded, name in sorted(ref_forms.items()):
        if folded not in surviving:
            if folded in cand_storages:
                problems.append(f"Deleted UserForm layout still present: {name!r}")
            continue
        if folded not in cand_storages:
            problems.append(f"UserForm layout missing in candidate: {name!r}")
            continue
        try:
            before = _storage_tree(ref_cfb, [name])
            after = _storage_tree(cand_cfb, [cand_storages[folded]])
        except Exception as exc:  # noqa: BLE001
            problems.append(f"Could not read UserForm layout {name!r} ({type(exc).__name__}).")
            continue
        if before != after:
            problems.append(f"UserForm layout changed: {name!r}")
    for folded in sorted(set(_form_storage_names_folded(cand_cfb)) - set(ref_forms)):
        problems.append(f"Unexpected UserForm layout in candidate: {folded!r}")
    return problems


def _form_storage_names_folded(cfb: CFB) -> list[str]:
    return [name.casefold() for name in _form_storage_names(cfb)]


def _make_unique_temp_name(project: Any) -> str:
    """ASCII temp component name guaranteed unique and short (plan 14 Stage H)."""
    import uuid

    existing = {m.name.casefold() for m in project.modules}
    for _ in range(100):
        candidate = f"VBAAE_TMP_{uuid.uuid4().hex[:6].upper()}"
        if candidate.casefold() not in existing:
            return candidate
    raise AdapterError("Could not generate a temporary module name for rename.")
