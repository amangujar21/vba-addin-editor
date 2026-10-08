"""Win32 file operations: ReplaceFileW commit and exclusive-access probe.

Plan sections 8.3, 15, 76. Fail closed; structured diagnostics; no flags that
ignore metadata-merge errors.
"""

from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import os
from dataclasses import dataclass
from pathlib import Path

_KERNEL32 = ctypes.WinDLL("kernel32", use_last_error=True)
_ADVAPI32 = ctypes.WinDLL("advapi32", use_last_error=True)

GENERIC_READ = 0x80000000
GENERIC_WRITE = 0x40000000
DELETE = 0x00010000
FILE_WRITE_ATTRIBUTES = 0x0100
FILE_SHARE_ALL = 0x7
FILE_ATTRIBUTE_ARCHIVE = 0x20
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

MOVEFILE_REPLACE_EXISTING = 0x1
MOVEFILE_COPY_ALLOWED = 0x2
MOVEFILE_WRITE_THROUGH = 0x8

_SETFILETIME = _KERNEL32.SetFileTime
_SETFILETIME.argtypes = [wt.HANDLE, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p]
_SETFILETIME.restype = wt.BOOL


class _WIN32_FIND_STREAM_DATA(ctypes.Structure):
    _fields_ = [("StreamSize", ctypes.c_longlong), ("cStreamName", wt.WCHAR * 296)]


_FINDFIRSTSTREAMW = _KERNEL32.FindFirstStreamW
_FINDFIRSTSTREAMW.argtypes = [wt.LPCWSTR, ctypes.c_int, ctypes.c_void_p, wt.DWORD]
_FINDFIRSTSTREAMW.restype = wt.HANDLE
_FINDNEXTSTREAMW = _KERNEL32.FindNextStreamW
_FINDNEXTSTREAMW.argtypes = [wt.HANDLE, ctypes.c_void_p]
_FINDNEXTSTREAMW.restype = wt.BOOL
_FINDCLOSE = _KERNEL32.FindClose
_FINDCLOSE.argtypes = [wt.HANDLE]
_FINDCLOSE.restype = wt.BOOL

_GETFILESECURITYW = _ADVAPI32.GetFileSecurityW
_GETFILESECURITYW.argtypes = [
    wt.LPCWSTR, wt.DWORD, ctypes.c_void_p, wt.DWORD, ctypes.POINTER(wt.DWORD),
]
_GETFILESECURITYW.restype = wt.BOOL
_OWNER_GROUP_DACL = 0x1 | 0x2 | 0x4


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
    """Distinguish in-use, access denied, read-only, and missing files.

    No sharing flags, so any other open handle is a sharing violation. Asks
    for write and delete access, never read: opening for read makes antivirus
    scan a package nobody has read since it was written, which takes seconds.
    """
    access = (GENERIC_WRITE | DELETE) if write else GENERIC_READ
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


def write_verified(path: Path, data: bytes) -> None:
    """Create ``path`` with ``data``, flushed to disk and read back intact.

    The read-back goes through the writing handle. Reopening the file would
    make antivirus scan it before returning.
    """
    with open(path, "x+b") as handle:
        handle.write(data)
        handle.flush()
        os.fsync(handle.fileno())
        handle.seek(0)
        written = handle.read()
    if written != data:
        raise OSError(f"{Path(path).name} did not read back as written.")


def _stream_names(path: Path) -> list[str] | None:
    data = _WIN32_FIND_STREAM_DATA()
    handle = _FINDFIRSTSTREAMW(str(path), 0, ctypes.byref(data), 0)
    if handle == INVALID_HANDLE_VALUE:
        return None
    names: list[str] = []
    try:
        while True:
            names.append(data.cStreamName)
            if not _FINDNEXTSTREAMW(handle, ctypes.byref(data)):
                break
    finally:
        _FINDCLOSE(handle)
    return names


def _security_descriptor(path: Path) -> bytes | None:
    needed = wt.DWORD(0)
    _GETFILESECURITYW(str(path), _OWNER_GROUP_DACL, None, 0, ctypes.byref(needed))
    if not needed.value:
        return None
    buffer = ctypes.create_string_buffer(needed.value)
    if not _GETFILESECURITYW(
        str(path), _OWNER_GROUP_DACL, buffer, needed.value, ctypes.byref(needed)
    ):
        return None
    return buffer.raw[: needed.value]


def _copy_creation_time(source: Path, dest: Path) -> bool:
    created = int(source.stat().st_ctime_ns // 100) + 116444736000000000  # FILETIME
    filetime = wt.FILETIME(created & 0xFFFFFFFF, created >> 32)
    handle = _CREATEFILEW(
        str(dest), FILE_WRITE_ATTRIBUTES, FILE_SHARE_ALL, None, OPEN_EXISTING, 0, None
    )
    if handle == INVALID_HANDLE_VALUE:
        return False
    try:
        return bool(_SETFILETIME(handle, ctypes.byref(filetime), None, None))
    finally:
        _CLOSEHANDLE(handle)


def rename_commit_preserves_metadata(replaced: Path, replacement: Path) -> bool:
    """True when renaming ``replacement`` over ``replaced`` loses nothing.

    ReplaceFileW carries the replaced file's security descriptor, attributes
    and alternate data streams onto the replacement. A plain rename does not,
    so it is only equivalent when both files already agree on all three.
    Anything that cannot be established counts as a difference.
    """
    try:
        if _stream_names(replaced) != ["::$DATA"]:
            return False
        ignore = FILE_ATTRIBUTE_ARCHIVE | FILE_ATTRIBUTE_NORMAL
        before = _GETFILEATTRIBUTESW(str(replaced))
        after = _GETFILEATTRIBUTESW(str(replacement))
        if 0xFFFFFFFF in (before, after) or (before | ignore) != (after | ignore):
            return False
        security = _security_descriptor(replaced)
        return security is not None and security == _security_descriptor(replacement)
    except OSError:
        return False


def _rename_commit(replaced: Path, replacement: Path, backup: Path | None) -> None:
    def state() -> dict[str, object]:
        return {
            "replaced_exists": replaced.exists(),
            "replacement_exists": replacement.exists(),
            "backup_exists": backup.exists() if backup is not None else None,
        }

    if backup is None:
        if not _MOVEFILEEXW(
            str(replacement), str(replaced), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH
        ):
            raise FileOpsError(
                "Windows could not safely replace the add-in file.",
                win32_error=ctypes.get_last_error(),
                details=state(),
            )
        return
    if not _MOVEFILEEXW(str(replaced), str(backup), 0):
        raise FileOpsError(
            "Windows could not move the add-in file to its backup name.",
            win32_error=ctypes.get_last_error(),
            details=state(),
        )
    if not _MOVEFILEEXW(str(replacement), str(replaced), MOVEFILE_WRITE_THROUGH):
        err = ctypes.get_last_error()
        rolled_back = bool(_MOVEFILEEXW(str(backup), str(replaced), 0))
        raise FileOpsError(
            "Windows could not move the edited file into place.",
            win32_error=err,
            details={**state(), "rolled_back": rolled_back},
        )


def commit_replace(replaced: Path, replacement: Path, backup: Path | None = None) -> str:
    """Put ``replacement`` at ``replaced``, keeping the old file as ``backup`` if given.

    ReplaceFileW opens the replacement for read, so antivirus scans the whole
    new package before the call returns: seconds for a large VBA project. A
    rename opens nothing. The rename is used when it leaves the file exactly as
    ReplaceFileW would (same security, attributes, streams, creation time);
    otherwise ReplaceFileW does the commit. Returns "rename" or "replace_file".
    """
    if rename_commit_preserves_metadata(replaced, replacement) and _copy_creation_time(
        replaced, replacement
    ):
        _rename_commit(replaced, replacement, backup)
        return "rename"
    replace_file(replaced, replacement, backup)
    return "replace_file"


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
