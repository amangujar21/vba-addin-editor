"""Settings file: new fields load, and values of the wrong type fall back to defaults."""

from __future__ import annotations

import json
from pathlib import Path

from vba_addin_editor.platform.settings import Settings, load_settings, save_settings


def test_round_trip(tmp_path: Path):
    path = tmp_path / "settings.json"
    saved = Settings(auto_backup=False, theme="light", editor_font_size=14, recent_files=["C:\\a.xlam"])
    assert save_settings(saved, path)
    assert load_settings(path) == saved


def test_wrong_types_fall_back_to_defaults(tmp_path: Path):
    path = tmp_path / "settings.json"
    path.write_text(
        json.dumps({"auto_backup": 0, "theme": 3, "editor_font_size": True, "recent_files": "C:\\a.xlam"}),
        encoding="utf-8",
    )
    assert load_settings(path) == Settings()
