"""User preferences kept between runs (%LOCALAPPDATA%\\VBAAddinEditor\\settings.json)."""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field, fields
from pathlib import Path

from vba_addin_editor.platform.paths import appdata_root
from vba_addin_editor.platform.session_store import write_json_atomic


@dataclass
class Settings:
    # Keep the replaced file as a same-folder backup on every in-place save.
    auto_backup: bool = True
    theme: str = "dark"
    editor_font_size: int = 11
    # Most recently opened first.
    recent_files: list[str] = field(default_factory=list)


def settings_path() -> Path:
    return appdata_root() / "settings.json"


def load_settings(path: Path | None = None) -> Settings:
    """Saved settings; defaults for anything missing, unreadable or of the wrong type."""
    settings = Settings()
    try:
        data = json.loads((path or settings_path()).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return settings
    if isinstance(data, dict):
        for item in fields(Settings):
            value = data.get(item.name)
            default = getattr(settings, item.name)
            # bool is an int subclass; a flag must not become a font size, nor the reverse.
            if isinstance(value, type(default)) and isinstance(value, bool) == isinstance(default, bool):
                setattr(settings, item.name, value)
    return settings


def save_settings(settings: Settings, path: Path | None = None) -> bool:
    try:
        write_json_atomic(path or settings_path(), asdict(settings))
    except OSError:
        return False
    return True
