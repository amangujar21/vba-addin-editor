"""Filesystem helpers: fingerprinting, candidate/backup naming, log directory."""

from __future__ import annotations

import hashlib
import os
from datetime import datetime, timezone
from pathlib import Path

from vba_addin_editor.domain.document import FileFingerprint


def fingerprint(path: Path) -> FileFingerprint:
    data = Path(path).read_bytes()
    stat = Path(path).stat()
    return FileFingerprint(
        sha256=hashlib.sha256(data).hexdigest(),
        size=len(data),
        mtime_ns=stat.st_mtime_ns,
    )


def fingerprint_matches(path: Path, fp: FileFingerprint) -> bool:
    try:
        return fingerprint(path) == fp
    except (FileNotFoundError, PermissionError, OSError):
        return False


def candidate_path_for(original: Path) -> Path:
    """Same directory, same volume, hidden collision-resistant name (plan 14 F)."""
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    token = os.urandom(3).hex()
    return original.with_name(
        f".{original.stem}.vbaae-candidate-{stamp}-{token}{original.suffix}"
    )


def xml_patch_path_for(candidate: Path) -> Path:
    """Sibling temp name for the XML-patched copy of a disposable candidate."""
    return candidate.with_name(candidate.stem + ".xmlpatch" + candidate.suffix)


def backup_path_for(original: Path, label: str = "backup") -> Path:
    """Visible same-extension backup name: 'Name - backup 2026-08-30 180212.ppam'."""
    stamp = datetime.now().strftime("%Y-%m-%d %H%M%S")  # noqa: DTZ005 -- local time is intentional: user-facing backup filename
    base = f"{original.stem} - {label} {stamp}{original.suffix}"
    candidate = original.with_name(base)
    counter = 2
    while candidate.exists():
        candidate = original.with_name(
            f"{original.stem} - {label} {stamp} ({counter}){original.suffix}"
        )
        counter += 1
    return candidate


def log_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    return Path(base) / "VBAAddinEditor" / "logs"
