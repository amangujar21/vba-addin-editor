"""Application-owned session directories and exclusive ownership handles."""

from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import json
import os
from dataclasses import dataclass
from pathlib import Path

from vba_addin_editor.platform.paths import sessions_root

_KERNEL32 = ctypes.WinDLL("kernel32", use_last_error=True)
GENERIC_READ = 0x80000000
GENERIC_WRITE = 0x40000000
CREATE_ALWAYS = 2
OPEN_EXISTING = 3
INVALID_HANDLE_VALUE = wt.HANDLE(-1).value
FILE_FLAG_DELETE_ON_CLOSE = 0x04000000
ERROR_SHARING_VIOLATION = 32

_CREATEFILEW = _KERNEL32.CreateFileW
_CREATEFILEW.argtypes = [
    wt.LPCWSTR, wt.DWORD, wt.DWORD, wt.LPVOID, wt.DWORD, wt.DWORD, wt.HANDLE,
]
_CREATEFILEW.restype = wt.HANDLE
_CLOSEHANDLE = _KERNEL32.CloseHandle
_CLOSEHANDLE.argtypes = [wt.HANDLE]
_CLOSEHANDLE.restype = wt.BOOL

LOCK_NAME = "owner.lock"
METADATA_NAME = "session.json"
BASELINE_DIR = "baseline"
PACKAGE_NAME = "package"


class SessionStoreError(OSError):
    pass


@dataclass
class SessionLock:
    path: Path
    handle: int

    def release(self) -> None:
        if self.handle:
            _CLOSEHANDLE(self.handle)
            self.handle = 0


def session_dir_for(session_id: str, *, root: Path | None = None) -> Path:
    return (root or sessions_root()) / session_id


def create_session_dir(session_id: str, *, root: Path | None = None) -> Path:
    directory = session_dir_for(session_id, root=root)
    directory.mkdir(parents=True, exist_ok=True)
    (directory / BASELINE_DIR).mkdir(exist_ok=True)
    return directory


def acquire_lock(directory: Path) -> SessionLock:
    lock_path = directory / LOCK_NAME
    handle = _CREATEFILEW(
        str(lock_path),
        GENERIC_READ | GENERIC_WRITE,
        0,
        None,
        CREATE_ALWAYS,
        0,
        None,
    )
    if handle == INVALID_HANDLE_VALUE:
        err = ctypes.get_last_error()
        raise SessionStoreError(f"Could not take session ownership (Win32 {err}).")
    return SessionLock(path=lock_path, handle=int(handle))


def is_owned(directory: Path) -> bool:
    """True when another process holds the exclusive owner.lock handle."""
    lock_path = directory / LOCK_NAME
    if not lock_path.exists():
        return False
    handle = _CREATEFILEW(
        str(lock_path),
        GENERIC_READ | GENERIC_WRITE,
        0,
        None,
        OPEN_EXISTING,
        0,
        None,
    )
    if handle == INVALID_HANDLE_VALUE:
        err = ctypes.get_last_error()
        return err == ERROR_SHARING_VIOLATION
    _CLOSEHANDLE(handle)
    return False


def write_json_atomic(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    data = json.dumps(payload, indent=2, sort_keys=True).encode("utf-8")
    with open(tmp, "wb") as handle:
        handle.write(data)
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(tmp, path)


def read_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def baseline_package_path(directory: Path, extension: str) -> Path:
    return directory / BASELINE_DIR / f"{PACKAGE_NAME}{extension}"
