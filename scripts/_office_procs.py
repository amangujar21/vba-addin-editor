"""Which Office processes are alive, so a script only ever stops the one it started."""

from __future__ import annotations

import time

_STILL_ACTIVE = 259


def live_pids(image: str) -> set[int]:
    """PIDs of running ``image`` (e.g. POWERPNT.EXE); exited ones left listed are skipped."""
    import win32api
    import win32con
    import win32process

    found: set[int] = set()
    for pid in win32process.EnumProcesses():
        try:
            handle = win32api.OpenProcess(
                win32con.PROCESS_QUERY_LIMITED_INFORMATION if hasattr(win32con, "PROCESS_QUERY_LIMITED_INFORMATION") else 0x1000,
                False,
                pid,
            )
        except Exception:  # noqa: BLE001, S112 - not ours to inspect
            continue
        try:
            name = win32process.GetModuleFileNameEx(handle, 0).rsplit("\\", 1)[-1]
            if name.upper() == image.upper() and win32process.GetExitCodeProcess(handle) == _STILL_ACTIVE:
                found.add(pid)
        except Exception:  # noqa: BLE001, S110 - gone or not ours
            pass
        finally:
            win32api.CloseHandle(handle)
    return found


def stop_new(image: str, before: set[int], wait: float = 10.0) -> list[int]:
    """Stop instances of ``image`` that were not running before, once they had ``wait`` s to quit."""
    import win32api
    import win32con

    deadline = time.monotonic() + wait
    left = live_pids(image) - before
    while left and time.monotonic() < deadline:
        time.sleep(0.5)
        left = live_pids(image) - before
    for pid in left:
        try:
            handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid)
            win32api.TerminateProcess(handle, 1)
            win32api.CloseHandle(handle)
        except Exception:  # noqa: BLE001, S110 - gone or not ours
            pass
    return sorted(left)
