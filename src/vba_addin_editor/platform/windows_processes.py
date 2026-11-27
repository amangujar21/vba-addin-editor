"""Office host process detection via Win32 Toolhelp (plan 8.1). No psutil."""

from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
from dataclasses import dataclass
from pathlib import Path

_TH32CS_SNAPPROCESS = 0x00000002
_INVALID_HANDLE_VALUE = wt.HANDLE(-1).value


class PROCESSENTRY32W(ctypes.Structure):
    _fields_ = [
        ("dwSize", wt.DWORD),
        ("cntUsage", wt.DWORD),
        ("th32ProcessID", wt.DWORD),
        ("th32DefaultHeapID", ctypes.POINTER(ctypes.c_ulong)),
        ("th32ModuleID", wt.DWORD),
        ("cntThreads", wt.DWORD),
        ("th32ParentProcessID", wt.DWORD),
        ("pcPriClassBase", ctypes.c_long),
        ("dwFlags", wt.DWORD),
        ("szExeFile", wt.WCHAR * 260),
    ]


class ProcessEnumerationError(OSError):
    def __init__(self, win32_error: int) -> None:
        super().__init__(f"Windows could not list running processes (Win32 error {win32_error}).")
        self.win32_error = win32_error


_PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
_STILL_ACTIVE = 259


def _has_exited(kernel32, pid: int) -> bool:
    """True only when Windows says the process has ended.

    An exited process stays in the snapshot while another process still holds a
    handle to it (an Office instance a COM add-in kept from unloading cleanly).
    Anything that cannot be queried counts as running, so a failed query never
    lets a save through.
    """
    kernel32.OpenProcess.argtypes = [wt.DWORD, wt.BOOL, wt.DWORD]
    kernel32.OpenProcess.restype = wt.HANDLE
    kernel32.GetExitCodeProcess.argtypes = [wt.HANDLE, ctypes.POINTER(wt.DWORD)]
    kernel32.GetExitCodeProcess.restype = wt.BOOL
    handle = kernel32.OpenProcess(_PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not handle:
        return False
    try:
        code = wt.DWORD()
        if not kernel32.GetExitCodeProcess(handle, ctypes.byref(code)):
            return False
        return code.value != _STILL_ACTIVE
    finally:
        kernel32.CloseHandle(handle)


def running_process_names() -> set[str]:
    """Set of executable base names (upper-case) currently running.

    Raises ProcessEnumerationError when the snapshot cannot be taken. An empty
    set means no matching processes, not a failed probe.
    """
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    CreateToolhelp32Snapshot = kernel32.CreateToolhelp32Snapshot
    CreateToolhelp32Snapshot.argtypes = [wt.DWORD, wt.DWORD]
    CreateToolhelp32Snapshot.restype = wt.HANDLE
    Process32FirstW = kernel32.Process32FirstW
    Process32FirstW.argtypes = [wt.HANDLE, ctypes.POINTER(PROCESSENTRY32W)]
    Process32FirstW.restype = wt.BOOL
    Process32NextW = kernel32.Process32NextW
    Process32NextW.argtypes = [wt.HANDLE, ctypes.POINTER(PROCESSENTRY32W)]
    Process32NextW.restype = wt.BOOL

    snap = CreateToolhelp32Snapshot(_TH32CS_SNAPPROCESS, 0)
    if snap == _INVALID_HANDLE_VALUE:
        raise ProcessEnumerationError(ctypes.get_last_error())
    names: set[str] = set()
    entry = PROCESSENTRY32W()
    entry.dwSize = ctypes.sizeof(PROCESSENTRY32W)
    try:
        if Process32FirstW(snap, ctypes.byref(entry)):
            while True:
                name = entry.szExeFile.upper()
                if name not in names and not _has_exited(kernel32, entry.th32ProcessID):
                    names.add(name)
                if not Process32NextW(snap, ctypes.byref(entry)):
                    break
    finally:
        kernel32.CloseHandle(snap)
    return names


_HOST_PROCESS = {
    ".xlam": "EXCEL.EXE",
    ".ppam": "POWERPNT.EXE",
    ".pptm": "POWERPNT.EXE",
}

_HOST_LABEL = {
    ".xlam": "Excel",
    ".ppam": "PowerPoint",
    ".pptm": "PowerPoint",
}


@dataclass(frozen=True)
class HostProcessProbe:
    corresponding_host_running: bool
    enumeration_failed: bool
    win32_error: int | None = None
    exe_name: str | None = None
    label: str | None = None


def corresponding_host_label(path: Path) -> str:
    return _HOST_LABEL.get(path.suffix.lower(), "Office")


def corresponding_host_exe(path: Path) -> str | None:
    return _HOST_PROCESS.get(path.suffix.lower())


def probe_host_process(path: Path) -> HostProcessProbe:
    """Authoritative host-process probe. Enumeration failure is not 'not running'."""
    target = corresponding_host_exe(path)
    label = corresponding_host_label(path)
    if target is None:
        return HostProcessProbe(
            corresponding_host_running=False,
            enumeration_failed=False,
            label=label,
        )
    try:
        names = running_process_names()
    except ProcessEnumerationError as exc:
        return HostProcessProbe(
            corresponding_host_running=False,
            enumeration_failed=True,
            win32_error=exc.win32_error,
            exe_name=target,
            label=label,
        )
    return HostProcessProbe(
        corresponding_host_running=target in names,
        enumeration_failed=False,
        exe_name=target,
        label=label,
    )


def host_process_running(path: Path) -> bool:
    """True when the corresponding Office host for this file is running.

    If process enumeration fails, this returns True (fail closed for callers
    that only consume a boolean). Prefer probe_host_process for diagnostics.
    """
    probe = probe_host_process(path)
    return probe.corresponding_host_running or probe.enumeration_failed
