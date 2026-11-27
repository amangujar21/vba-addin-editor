"""Live reload: save an add-in while Excel or PowerPoint has it loaded.

Office keeps a loaded add-in open with a sharing lock, so the file cannot be
replaced while it is loaded. Office also builds an add-in's ribbon only when
it loads it, so new buttons need a reload anyway. ``HostSwap`` therefore
closes just this add-in in the running Office, lets the save commit, and
loads it again. Nothing else in Office is touched: other workbooks,
presentations and add-ins stay open.

This is the only module that talks to Office. It uses COM late binding
through pywin32 (imported lazily; without it ``available()`` is False and
saves behave as before: close Office first).

- Excel: an open workbook registers its full path in the Running Object
  Table, so the instance holding the file is found by path even with several
  Excel instances open. Unload = ``Auto_Close`` + ``Workbook.Close(False)``;
  reload = ``Workbooks.Open`` + ``Auto_Open``, which is what Excel runs when it
  loads an add-in at startup. Closing an installed add-in's workbook does not
  uninstall it.
- PowerPoint (one instance): a ``.ppam`` stays locked after
  ``AddIn.Loaded = False`` until it is removed from ``AddIns``, so unload =
  unload + remove and reload = add + load, restoring ``Registered`` and
  ``AutoLoad``. A ``.pptm`` is closed and opened again with a window.

COM objects belong to the thread that made them: create, use and drop a
``HostSwap`` on one thread (the save worker); ``locate`` is self-contained.
Never call ``GetObject(path)`` with a path that might not be open: it would
start Office and open the file. The ROT is searched instead.
"""

from __future__ import annotations

import time
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

EXCEL = "excel"
POWERPOINT = "powerpoint"

_HOST_OF = {".xlam": EXCEL, ".ppam": POWERPOINT, ".pptm": POWERPOINT}
_LABEL = {EXCEL: "Excel", POWERPOINT: "PowerPoint"}

# Office refuses calls while a cell is being edited or a modal dialog is up.
_RPC_E_CALL_REJECTED = -2147418111
_RPC_E_SERVERCALL_RETRYLATER = -2147417846
_BUSY = frozenset({_RPC_E_CALL_REJECTED, _RPC_E_SERVERCALL_RETRYLATER})
BUSY_WAIT_SECONDS = 15.0

_XL_AUTO_OPEN = 1
_XL_AUTO_CLOSE = 2


class LiveHostError(Exception):
    """Office could not close or reopen the add-in; the message is for the user."""


def available() -> bool:
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    return True


def host_of(path: Path) -> str | None:
    return _HOST_OF.get(path.suffix.lower())


@dataclass(frozen=True)
class LiveState:
    """Where ``path`` is loaded, as plain data (safe to hand to the Tk thread)."""

    host: str | None  # EXCEL / POWERPOINT, None for other files
    loaded: bool
    unsaved: bool = False  # Office holds changes to it that were never saved
    error: str | None = None  # COM failed: nothing can be said about the file

    @property
    def label(self) -> str:
        return _LABEL.get(self.host or "", "Office")


def _same(a: str, b: Path) -> bool:
    try:
        return Path(a).resolve() == b.resolve() or a.casefold() == str(b).casefold()
    except OSError:
        return a.casefold() == str(b).casefold()


def _com_error_code(exc: BaseException) -> int | None:
    code = getattr(exc, "hresult", None)
    if code is None and getattr(exc, "args", None):
        code = exc.args[0]
    return code if isinstance(code, int) else None


def _retry(call: Callable[[], Any], *, what: str, label: str, wait: float = BUSY_WAIT_SECONDS) -> Any:
    """Run one COM call, retrying while Office says it is busy."""
    import pywintypes

    deadline = time.monotonic() + wait
    while True:
        try:
            return call()
        except pywintypes.com_error as exc:
            if _com_error_code(exc) in _BUSY and time.monotonic() < deadline:
                time.sleep(0.25)
                continue
            if _com_error_code(exc) in _BUSY:
                raise LiveHostError(
                    f"{label} is busy (a cell is being edited or a dialog is open), "
                    f"so it could not {what}."
                ) from exc
            raise LiveHostError(f"{label} could not {what}: {_describe(exc)}") from exc


def _describe(exc: BaseException) -> str:
    args = getattr(exc, "args", ())
    # pywintypes.com_error: (hresult, text, (wcode, source, description, ...), argerr)
    if len(args) >= 3 and isinstance(args[2], tuple) and len(args[2]) >= 3 and args[2][2]:
        return str(args[2][2]).strip()
    if len(args) >= 2 and args[1]:
        return str(args[1])
    return str(exc)


def _dispatch(obj: Any) -> Any:
    import pythoncom
    import win32com.client.dynamic as dyn

    return dyn.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))


def _excel_workbook(path: Path) -> Any | None:
    """The open Excel workbook for ``path`` from the ROT, or None. Never opens anything."""
    import pythoncom

    # pywin32's stubs miss that the ROT iterates monikers and CreateBindCtx's argument.
    rot: Any = pythoncom.GetRunningObjectTable()
    ctx: Any = pythoncom.CreateBindCtx(0)  # pyright: ignore[reportCallIssue]
    for moniker in rot:
        try:
            name = moniker.GetDisplayName(ctx, None)
        except pythoncom.com_error:
            continue
        if not name or name.startswith("!") or not _same(name, path):
            continue
        try:
            book = _dispatch(rot.GetObject(moniker))
            # Only an Excel workbook has these; anything else registered under
            # the path (another program) is not ours to close.
            if str(book.Application.Name).startswith("Microsoft Excel") and _same(book.FullName, path):
                return book
        except Exception:  # noqa: BLE001, S112 - not an Excel workbook
            continue
    return None


def _powerpoint_app() -> Any | None:
    import pythoncom

    try:
        return _dispatch(pythoncom.GetActiveObject("PowerPoint.Application"))
    except pythoncom.com_error:
        return None


def _powerpoint_item(app: Any, path: Path) -> Any | None:
    """The loaded AddIn (.ppam) or open Presentation (.pptm) for ``path``."""
    items = app.AddIns if path.suffix.lower() == ".ppam" else app.Presentations
    for i in range(1, int(items.Count) + 1):
        item = items.Item(i)
        if _same(str(item.FullName), path):
            if path.suffix.lower() == ".ppam" and not int(item.Loaded):
                return None  # listed, not loaded: it does not hold the file
            return item
    return None


def _with_com(work: Callable[[], LiveState], host: str) -> LiveState:
    import pythoncom

    pythoncom.CoInitialize()
    try:
        return work()
    except LiveHostError as exc:
        return LiveState(host, loaded=False, error=str(exc))
    except Exception as exc:  # noqa: BLE001 - COM reports anything at all
        return LiveState(host, loaded=False, error=_describe(exc))
    finally:
        pythoncom.CoUninitialize()


def locate(path: Path) -> LiveState:
    """Is ``path`` loaded in a running Office, and does Office hold unsaved changes to it?

    Runs on any thread (it initialises COM itself). A file Office has not
    loaded gives ``loaded=False``; whether something else holds it is for the
    exclusive-access probe to say.
    """
    host = host_of(path)
    if host is None:
        return LiveState(None, loaded=False)
    if not available():
        return LiveState(host, loaded=False, error="pywin32 is not installed")

    def work() -> LiveState:
        if host == EXCEL:
            book = _excel_workbook(path)
            if book is None:
                return LiveState(host, loaded=False)
            saved = _retry(lambda: bool(book.Saved), what="report on the add-in", label="Excel")
            return LiveState(host, loaded=True, unsaved=not saved)
        app = _powerpoint_app()
        if app is None:
            return LiveState(host, loaded=False)
        item = _retry(lambda: _powerpoint_item(app, path), what="list its add-ins", label="PowerPoint")
        if item is None:
            return LiveState(host, loaded=False)
        unsaved = path.suffix.lower() == ".pptm" and not int(item.Saved)
        return LiveState(host, loaded=True, unsaved=unsaved)

    return _with_com(work, host)


class HostSwap:
    """Close ``path`` in the Office that has it loaded, then load it again.

    ``unload()`` before the commit, ``reload()`` after it, whatever the commit
    did (a failed save reloads the old file). Both are no-ops when Office did
    not have the file loaded. Make, use and drop it on one thread.
    """

    def __init__(self, path: Path, *, discard_unsaved: bool = False) -> None:
        self.path = path
        self.host = host_of(path)
        self.label = _LABEL.get(self.host or "", "Office")
        self.discard_unsaved = discard_unsaved
        self.unloaded = False
        self.reloaded = False
        self.seconds_unloaded = 0.0
        self._app: Any = None
        self._ppam_flags: tuple[bool, bool] | None = None
        self._started = 0.0
        self._com = False

    # -- COM lifetime -----------------------------------------------------

    def __enter__(self):  # returns self (typing.Self needs 3.11)
        import pythoncom

        pythoncom.CoInitialize()
        self._com = True
        return self

    def __exit__(self, *_exc) -> None:
        self._app = None
        if self._com:
            import pythoncom

            self._com = False
            pythoncom.CoUninitialize()

    # -- the swap ---------------------------------------------------------

    def unload(self) -> bool:
        """Close the add-in in Office. True when it was loaded (and is now closed)."""
        if self.host == EXCEL:
            done = self._unload_excel()
        elif self.host == POWERPOINT:
            done = self._unload_powerpoint()
        else:
            done = False
        if done:
            self.unloaded = True
            self._started = time.monotonic()
        return done

    def reload(self) -> None:
        """Load the add-in again; raises LiveHostError when Office refuses."""
        if not self.unloaded or self.reloaded:
            return
        try:
            if self.host == EXCEL:
                self._reload_excel()
            else:
                self._reload_powerpoint()
        finally:
            self.seconds_unloaded = time.monotonic() - self._started
        self.reloaded = True

    # Excel

    def _unload_excel(self) -> bool:
        book = _excel_workbook(self.path)
        if book is None:
            return False
        label = self.label
        saved = _retry(lambda: bool(book.Saved), what="report on the add-in", label=label)
        if not saved and not self.discard_unsaved:
            raise LiveHostError(
                f"Excel holds unsaved changes to {self.path.name} (edits in the VBA "
                "editor?). Save or discard them in Excel first."
            )
        self._app = _retry(lambda: book.Application, what="report on the add-in", label=label)
        # Auto_Close runs when Excel unloads an add-in, not on Workbook.Close from COM.
        _retry(lambda: book.RunAutoMacros(_XL_AUTO_CLOSE), what="run Auto_Close", label=label)
        _retry(lambda: book.Close(False), what=f"close {self.path.name}", label=label)
        if _excel_workbook(self.path) is not None:
            raise LiveHostError(f"Excel kept {self.path.name} open (its BeforeClose cancelled the close).")
        return True

    def _reload_excel(self) -> None:
        app = self._app
        label = self.label
        if app is None:
            raise LiveHostError("Excel is no longer running.")
        book = _retry(
            lambda: app.Workbooks.Open(str(self.path), UpdateLinks=0, AddToMru=False),
            what=f"reopen {self.path.name}",
            label=label,
        )
        _retry(lambda: book.RunAutoMacros(_XL_AUTO_OPEN), what="run Auto_Open", label=label)

    # PowerPoint

    def _unload_powerpoint(self) -> bool:
        app = _powerpoint_app()
        if app is None:
            return False
        label = self.label
        item = _retry(lambda: _powerpoint_item(app, self.path), what="list its add-ins", label=label)
        if item is None:
            return False
        self._app = app
        if self.path.suffix.lower() == ".ppam":
            name = str(item.Name)
            self._ppam_flags = (bool(int(item.Registered)), bool(int(item.AutoLoad)))
            _retry(lambda: setattr(item, "Loaded", 0), what=f"unload {self.path.name}", label=label)
            # PowerPoint keeps an unloaded add-in's file open until it leaves AddIns.
            try:
                _retry(lambda: app.AddIns.Remove(name), what=f"unload {self.path.name}", label=label)
            except LiveHostError:
                try:  # the save stops here: give the user the add-in back
                    item.Loaded = -1
                except Exception:  # noqa: BLE001, S110 - the error below says enough
                    pass
                raise
            return True
        if not int(item.Saved) and not self.discard_unsaved:
            raise LiveHostError(
                f"PowerPoint holds unsaved changes to {self.path.name}. Save or discard them first."
            )
        if not int(item.Saved):
            item.Saved = -1  # discard: Close would otherwise ask
        _retry(lambda: item.Close(), what=f"close {self.path.name}", label=label)
        return True

    def _reload_powerpoint(self) -> None:
        app = self._app
        label = self.label
        if self.path.suffix.lower() == ".ppam":
            addin = _retry(lambda: app.AddIns.Add(str(self.path)), what=f"reload {self.path.name}", label=label)
            _retry(lambda: setattr(addin, "Loaded", -1), what=f"reload {self.path.name}", label=label)
            if self._ppam_flags is not None:
                registered, autoload = self._ppam_flags
                try:
                    addin.Registered = -1 if registered else 0
                    addin.AutoLoad = -1 if autoload else 0
                except Exception:  # noqa: BLE001, S110 - loaded; the flags are a nicety
                    pass
            return
        _retry(
            lambda: app.Presentations.Open(str(self.path), ReadOnly=0, Untitled=0, WithWindow=-1),
            what=f"reopen {self.path.name}",
            label=label,
        )
