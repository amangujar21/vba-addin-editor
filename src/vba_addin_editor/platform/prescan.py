"""Background pre-read of a file the editor just wrote.

Antivirus scans a freshly written macro package the first time it is opened
for read, and the save pipeline never opens its own output. Left alone, that
scan lands on Office when it next loads the add-in. Opening the file here, on
a thread nobody waits for, spends it between the save and the next Office
start instead.
"""

from __future__ import annotations

import threading
from collections.abc import Callable
from pathlib import Path

from vba_addin_editor.platform import windows_file_ops as wfo

IDLE = "idle"
SCANNING = "scanning"
READY = "ready"
FAILED = "failed"


def open_for_read(path: Path) -> bool:
    """Open ``path`` for read and close it again; True when the open succeeded.

    The open is what waits for the scan. Every sharing flag is set, so the
    handle never keeps Office or anything else from the file.
    """
    handle = wfo._CREATEFILEW(
        str(path), wfo.GENERIC_READ, wfo.FILE_SHARE_ALL, None, wfo.OPEN_EXISTING, 0, None
    )
    if handle == wfo.INVALID_HANDLE_VALUE:
        return False
    wfo._CLOSEHANDLE(handle)
    return True


class Prescan:
    """One pre-read at a time; ``state`` is safe to poll from the UI thread."""

    def __init__(self, opener: Callable[[Path], bool] = open_for_read) -> None:
        self._opener = opener
        self._thread: threading.Thread | None = None
        self.path: Path | None = None
        self.state = IDLE

    @property
    def busy(self) -> bool:
        thread = self._thread
        return thread is not None and thread.is_alive()

    def start(self, path: Path) -> None:
        self.settle()
        self.path = path
        self.state = SCANNING
        # Daemon: a read-only open is safe to abandon when the app exits.
        thread = threading.Thread(
            target=self._run, args=(path,), name="vbaae-prescan", daemon=True
        )
        self._thread = thread
        thread.start()

    def _run(self, path: Path) -> None:
        try:
            opened = self._opener(path)
        except OSError:
            opened = False
        self.state = READY if opened else FAILED

    def settle(self) -> None:
        """Wait until a pre-read still in flight has closed its handle.

        The exclusive-access probe and the commit both need the file free of
        other handles, including this one. A pending open cannot be cut short
        (Defender ignores CancelSynchronousIo), so this waits for the scan.
        """
        thread = self._thread
        if thread is not None and thread.is_alive():
            thread.join()
