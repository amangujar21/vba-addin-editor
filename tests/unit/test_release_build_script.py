from __future__ import annotations

from pathlib import Path


def test_release_mode_is_a_strict_gate():
    text = Path("scripts/build.ps1").read_text(encoding="utf-8")
    assert "param(" in text
    assert "$Release" in text
    assert "Release builds require a clean git tree." in text
    assert "python -m pytest tests -q" in text
    assert "python -m ruff check ." in text
    assert "python -m pyright" in text
    assert "RealAddin.xlam" in text
    assert "RealAddin.ppam" in text
    assert "RealPresentation.pptm" in text
    assert "throw \"Release builds require qualification fixtures" in text or "Release builds require qualification fixtures" in text
    assert "release-manifest.json" in text
    assert "development-dirty" in text
    # Build OK for release only after both modes and smoke tests.
    assert "qualified onedir and onefile" in text
    assert "Not a qualified release" in text
