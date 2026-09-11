"""Application version and build metadata."""

from __future__ import annotations

from typing import Any

APP_NAME = "VBA Add-in Editor"
APP_EXE_BASENAME = "VBAAddinEditor"
VERSION = "0.1.0"

# Pinned dependency, recorded for THIRD_PARTY_NOTICES and diagnostics.
PYOPENVBA_PIN = "3.4.0"
PYOPENVBA_URL = "https://pypi.org/project/pyOpenVBA/"


def _installed_pyopenvba() -> str | None:
    try:
        import pyopenvba

        return getattr(pyopenvba, "__version__", None) or PYOPENVBA_PIN
    except Exception:  # noqa: BLE001 - diagnostics must not fail import
        return None


def build_identity() -> dict[str, Any]:
    """Runtime build identity. Packaged builds embed `_build_info` at freeze time."""
    try:
        from vba_addin_editor._build_info import BUILD_INFO  # type: ignore[import-not-found]

        identity = dict(BUILD_INFO)
    except Exception:  # noqa: BLE001 - source runs have no embedded metadata
        identity = {
            "version": VERSION,
            "source_commit": "unbuilt",
            "dirty_tree": None,
            "build_time_utc": None,
            "python_version": None,
            "python_architecture": None,
            "packaged_mode": "development",
            "pyopenvba_pin": PYOPENVBA_PIN,
            "dependencies": {"pyopenvba": PYOPENVBA_PIN},
        }
    identity.setdefault("version", VERSION)
    identity.setdefault("pyopenvba_pin", PYOPENVBA_PIN)
    identity["pyopenvba_installed"] = _installed_pyopenvba()
    return identity
