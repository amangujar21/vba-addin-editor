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
FILE_ATTRIBUTE_READONLY = 0x1
FILE_ATTRIBUTE_NORMAL = 0x80
ERROR_FILE_NOT_FOUND = 2
ERROR_PATH_NOT_FOUND = 3
ERROR_ACCESS_DENIED = 5
ERROR_WRITE_PROTECT = 19
ERROR_SHARING_VIOLATION = 32
ERROR_LOCK_VIOLATION = 33
ERROR_ALREADY_EXISTS = 183
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

_GETFILEATTRIBUTESW = _KERNEL32.GetFileAttributesW
_GETFILEATTRIBUTESW.argtypes = [wt.LPCWSTR]
_GETFILEATTRIBUTESW.restype = wt.DWORD

_MOVEFILEEXW = _KERNEL32.MoveFileExW
_MOVEFILEEXW.argtypes = [wt.LPCWSTR, wt.LPCWSTR, wt.DWORD]
_MOVEFILEEXW.restype = wt.BOOL

MOVEFILE_COPY_ALLOWED = 0x2


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


@dataclass(frozen=True)
class AccessProbe:
    ok: bool
    reason: str | None = None
    win32_error: int | None = None
    win32_name: str | None = None
    message: str | None = None


def can_open_exclusively(path: Path, *, write: bool = True) -> bool:
    """True when the file can be opened with NO sharing flags (plan 8.3)."""
    return probe_exclusive_access(path, write=write).ok


def exclusive_access_error(path: Path) -> str | None:
    """Return a plain-language blocker string or None when access is free."""
    probe = probe_exclusive_access(path)
    if probe.ok:
        return None
    return probe.message or "The file is not available for exclusive write access."


def probe_exclusive_access(path: Path, *, write: bool = True) -> AccessProbe:
    """Distinguish in-use, access denied, read-only, and missing files."""
    access = GENERIC_READ | (GENERIC_WRITE if write else 0)
    handle = _CREATEFILEW(
        str(path), access, 0, None, OPEN_EXISTING, 0, None
    )
    if handle != INVALID_HANDLE_VALUE:
        _CLOSEHANDLE(handle)
        return AccessProbe(ok=True)
    err = ctypes.get_last_error()
    name = describe_win32_error(err)
    reason, message = _classify_createfile_error(path, err)
    return AccessProbe(
        ok=False,
        reason=reason,
        win32_error=err,
        win32_name=name,
        message=message,
    )


def _classify_createfile_error(path: Path, err: int) -> tuple[str, str]:
    if err in {ERROR_SHARING_VIOLATION, ERROR_LOCK_VIOLATION}:
        return (
            "file_in_use",
            "Another program is using this add-in file. Close it and try again.",
        )
    if err in {ERROR_FILE_NOT_FOUND, ERROR_PATH_NOT_FOUND}:
        return (
            "source_missing",
            "The add-in file could not be found.",
        )
    if err in {ERROR_ACCESS_DENIED, ERROR_WRITE_PROTECT} or _is_readonly(path):
        if _is_readonly(path) or err == ERROR_WRITE_PROTECT:
            return (
                "read_only",
                "The add-in file is read-only, so it cannot be replaced.",
            )
        return (
            "access_denied",
            "Windows denied write access to this add-in file.",
        )
    return (
        "unknown",
        f"The add-in file could not be opened for exclusive write access ({describe_win32_error(err)}).",
    )


def _is_readonly(path: Path) -> bool:
    attrs = _GETFILEATTRIBUTESW(str(path))
    if attrs == 0xFFFFFFFF:
        return False
    return bool(attrs & FILE_ATTRIBUTE_READONLY)


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


def move_new_file(source: Path, dest: Path) -> None:
    """Same-volume atomic move that fails if destination already exists."""
    ok = _MOVEFILEEXW(str(source), str(dest), 0)
    if ok:
        return
    err = ctypes.get_last_error()
    raise FileOpsError(
        "Windows could not move the new file into place.",
        win32_error=err,
        details={
            "source_exists": source.exists(),
            "dest_exists": dest.exists(),
        },
    )


def describe_win32_error(err: int) -> str:
    names = {
        ERROR_FILE_NOT_FOUND: "file not found",
        ERROR_PATH_NOT_FOUND: "path not found",
        ERROR_ACCESS_DENIED: "access denied",
        ERROR_WRITE_PROTECT: "write protect",
        ERROR_SHARING_VIOLATION: "sharing violation",
        ERROR_LOCK_VIOLATION: "lock violation",
        ERROR_ALREADY_EXISTS: "already exists",
        ERROR_UNABLE_TO_REMOVE_REPLACED: "could not remove the replaced file",
        ERROR_UNABLE_TO_MOVE_REPLACEMENT: "could not move the replacement file",
        ERROR_UNABLE_TO_MOVE_REPLACEMENT_2: "could not move the replacement file (2)",
    }
    return names.get(err, f"Win32 error {err}")
