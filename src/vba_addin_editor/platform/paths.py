"""Filesystem helpers: fingerprinting, candidate/backup naming, log directory."""

from __future__ import annotations

import hashlib
import os
from datetime import datetime, timezone
from pathlib import Path

from vba_addin_editor.domain.document import FileFingerprint

_FINGERPRINT_RETRIES = 2
FILE_ATTRIBUTE_REPARSE_POINT = 0x400


class UnstableSourceError(OSError):
    """The file's bytes changed while they were being read."""


def fingerprint(path: Path, *, retries: int = _FINGERPRINT_RETRIES) -> FileFingerprint:
    """Hash file contents. Size/mtime never skip the hash.

    Reads twice per attempt so a change during the read is detected. A bounded
    retry covers brief writers; remaining instability raises UnstableSourceError.
    """
    path = Path(path)
    last_reason = "content changed during read"
    attempts = max(0, retries) + 1
    for _ in range(attempts):
        data = path.read_bytes()
        digest = hashlib.sha256(data).hexdigest()
        size = len(data)
        stat = path.stat()
        if stat.st_size != size:
            last_reason = "size changed during read"
            continue
        again = path.read_bytes()
        if hashlib.sha256(again).hexdigest() != digest or len(again) != size:
            last_reason = "content changed during read"
            continue
        return FileFingerprint(sha256=digest, size=size, mtime_ns=stat.st_mtime_ns)
    raise UnstableSourceError(
        f"The file changed while it was being read ({last_reason})."
    )


def fingerprint_status(path: Path, fp: FileFingerprint) -> str:
    """Return match | external_change | source_missing | source_unstable | access_denied."""
    try:
        current = fingerprint(path)
    except FileNotFoundError:
        return "source_missing"
    except PermissionError:
        return "access_denied"
    except UnstableSourceError:
        return "source_unstable"
    except OSError as exc:
        if getattr(exc, "winerror", None) in {2, 3}:
            return "source_missing"
        if getattr(exc, "winerror", None) in {5, 19}:
            return "access_denied"
        return "access_denied"
    if current.content_equal(fp):
        return "match"
    return "external_change"


def fingerprint_matches(path: Path, fp: FileFingerprint) -> bool:
    """True when SHA-256 and size match. Timestamp-only changes do not conflict."""
    return fingerprint_status(path, fp) == "match"


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def candidate_path_for(original: Path) -> Path:
    """Same directory, same volume, hidden collision-resistant name (plan 14 F)."""
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    token = os.urandom(3).hex()
    return original.with_name(
        f".{original.stem}.vbaae-candidate-{stamp}-{token}{original.suffix}"
    )


def sibling_temp_path(dest: Path, *, kind: str = "copy") -> Path:
    """Unique temporary sibling of dest, never dest itself."""
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    token = os.urandom(3).hex()
    return dest.with_name(f".{dest.stem}.vbaae-{kind}-{stamp}-{token}{dest.suffix}")


def xml_patch_path_for(candidate: Path) -> Path:
    """Sibling temp name for the XML-patched copy of a disposable candidate."""
    return candidate.with_name(candidate.stem + ".xmlpatch" + candidate.suffix)


def backup_path_for(original: Path, label: str = "backup") -> Path:
    """Visible same-extension backup name: 'Name - backup 2026-08-30 180212.ppam'."""
    stamp = datetime.now().strftime("%Y-%m-%d %H%M%S")  # noqa: DTZ005 -- local time is intentional: user-facing backup filename
    base = f"{original.stem} - {label} {stamp}{original.suffix}"
    candidate = original.with_name(base)
    counter = 2
    while candidate.exists():
        candidate = original.with_name(
            f"{original.stem} - {label} {stamp} ({counter}){original.suffix}"
        )
        counter += 1
    return candidate


def log_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    return Path(base) / "VBAAddinEditor" / "logs"


def appdata_root() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    return Path(base) / "VBAAddinEditor"


def sessions_root() -> Path:
    return appdata_root() / "sessions"


def backups_catalog_path() -> Path:
    return appdata_root() / "backups" / "catalog.json"


def paths_are_same_file(left: Path, right: Path) -> bool:
    """True when normalized Windows paths or file identity refer to one file."""
    a = Path(left)
    b = Path(right)
    try:
        a_res = a.resolve()
        b_res = b.resolve()
    except OSError:
        a_res, b_res = a, b
    if os.path.normcase(str(a_res)) == os.path.normcase(str(b_res)):
        return True
    try:
        if a.exists() and b.exists() and os.path.samefile(a, b):
            return True
    except OSError:
        return False
    return False


def path_or_ancestor_is_reparse_point(path: Path) -> bool:
    """True when path or an ancestor has FILE_ATTRIBUTE_REPARSE_POINT."""
    current = Path(path)
    seen: set[str] = set()
    while True:
        key = os.path.normcase(str(current))
        if key in seen:
            break
        seen.add(key)
        try:
            attrs = current.lstat().st_file_attributes  # type: ignore[attr-defined]
        except AttributeError:
            try:
                import ctypes

                GetFileAttributesW = ctypes.windll.kernel32.GetFileAttributesW  # type: ignore[attr-defined]
                GetFileAttributesW.restype = ctypes.c_uint32
                raw = GetFileAttributesW(str(current))
                if raw == 0xFFFFFFFF:
                    attrs = 0
                else:
                    attrs = raw
            except Exception:  # noqa: BLE001
                attrs = 0
        except OSError:
            attrs = 0
        if attrs & FILE_ATTRIBUTE_REPARSE_POINT:
            return True
        if current.parent == current:
            break
        current = current.parent
    return False


def is_inside(child: Path, parent: Path) -> bool:
    try:
        child.resolve().relative_to(parent.resolve())
        return True
    except (ValueError, OSError):
        return False
