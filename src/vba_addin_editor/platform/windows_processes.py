"""Office host process detection via Win32 Toolhelp (plan 8.1). No psutil."""

from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
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


def running_process_names() -> set[str]:
    """Set of executable base names (upper-case) currently running."""
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
        return set()
    names: set[str] = set()
    entry = PROCESSENTRY32W()
    entry.dwSize = ctypes.sizeof(PROCESSENTRY32W)
    try:
        if Process32FirstW(snap, ctypes.byref(entry)):
            while True:
                names.add(entry.szExeFile.upper())
                if not Process32NextW(snap, ctypes.byref(entry)):
                    break
    finally:
        kernel32.CloseHandle(snap)
    return names


def host_process_running(path: Path) -> bool:
    """True when the corresponding Office host for this add-in is running."""
    target = "EXCEL.EXE" if path.suffix.lower() == ".xlam" else "POWERPNT.EXE"
    return target in running_process_names()
