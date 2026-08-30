"""Win32 file operations: ReplaceFileW commit and exclusive-access probe.

Plan sections 8.3, 15, 76. Fail closed; structured diagnostics; no flags that
ignore metadata-merge errors.
"""

from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
from dataclasses import dataclass
from pathlib import Path

_KERNEL32 = ctypes.WinDLL("kernel32", use_last_error=True)

GENERIC_READ = 0x80000000
GENERIC_WRITE = 0x40000000
OPEN_EXISTING = 3
INVALID_HANDLE_VALUE = wt.HANDLE(-1).value
ERROR_SHARING_VIOLATION = 32
ERROR_LOCK_VIOLATION = 33
ERROR_UNABLE_TO_REMOVE_REPLACED = 1175
ERROR_UNABLE_TO_MOVE_REPLACEMENT = 1176
ERROR_UNABLE_TO_MOVE_REPLACEMENT_2 = 1177

_REPLACEFILEW = _KERNEL32.ReplaceFileW
_REPLACEFILEW.argtypes = [
    wt.LPCWSTR, wt.LPCWSTR, wt.LPCWSTR, wt.DWORD, wt.LPVOID, wt.LPVOID,
]
_REPLACEFILEW.restype = wt.BOOL

_CREATEFILEW = _KERNEL32.CreateFileW
_CREATEFILEW.argtypes = [
    wt.LPCWSTR, wt.DWORD, wt.DWORD, wt.LPVOID, wt.DWORD, wt.DWORD, wt.HANDLE,
]
_CREATEFILEW.restype = wt.HANDLE

_CLOSEHANDLE = _KERNEL32.CloseHandle
_CLOSEHANDLE.argtypes = [wt.HANDLE]
_CLOSEHANDLE.restype = wt.BOOL


class FileOpsError(Exception):
    def __init__(self, message: str, win32_error: int | None = None,
                 details: dict[str, object] | None = None) -> None:
        super().__init__(message)
        self.win32_error = win32_error
        self.details = details or {}


@dataclass(frozen=True)
class ReplaceResult:
    ok: bool
    win32_error: int | None = None


def can_open_exclusively(path: Path, *, write: bool = True) -> bool:
    """True when the file can be opened with NO sharing flags (plan 8.3)."""
    access = GENERIC_READ | (GENERIC_WRITE if write else 0)
    handle = _CREATEFILEW(
        str(path), access, 0, None, OPEN_EXISTING, 0, None
    )
    if handle == INVALID_HANDLE_VALUE:
        return False
    _CLOSEHANDLE(handle)
    return True


def exclusive_access_error(path: Path) -> str | None:
    """Return a plain-language blocker string or None when access is free."""
    if not can_open_exclusively(path):
        return "Another program is using this add-in file. Close it and try again."
    return None


def replace_file(
    replaced: Path,
    replacement: Path,
    backup: Path | None = None,
) -> ReplaceResult:
    """ReplaceFileW commit: preserves ACLs/streams of the replaced file (plan 15).

    No REPLACEFILE_* flags: fail closed on metadata-merge problems.
    """
    ok = _REPLACEFILEW(
        str(replaced),
        str(replacement),
        str(backup) if backup is not None else None,
        0,  # flags
        None,
        None,
    )
    if ok:
        return ReplaceResult(ok=True)
    err = ctypes.get_last_error()
    state = {
        "replaced_exists": replaced.exists(),
        "replacement_exists": replacement.exists(),
        "backup_exists": backup.exists() if backup is not None else None,
    }
    raise FileOpsError(
        "Windows could not safely replace the add-in file.", win32_error=err, details=state
    )


def describe_win32_error(err: int) -> str:
    names = {
        ERROR_SHARING_VIOLATION: "sharing violation",
        ERROR_LOCK_VIOLATION: "lock violation",
        ERROR_UNABLE_TO_REMOVE_REPLACED: "could not remove the replaced file",
        ERROR_UNABLE_TO_MOVE_REPLACEMENT: "could not move the replacement file",
        ERROR_UNABLE_TO_MOVE_REPLACEMENT_2: "could not move the replacement file (2)",
    }
    return names.get(err, f"Win32 error {err}")
