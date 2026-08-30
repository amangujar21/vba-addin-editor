"""Release gate: full app-level cycle against authentic Office-authored fixtures.

Plan §4.1/§4.3/§4.5. These tests run the real DocumentService + SaveService
pipeline (real probes, real ReplaceFileW commit) on copies of genuine
Office-authored add-ins created via "Save As -> Add-in" in Excel/PowerPoint.
They prove parser-level round-trip; the final PowerPoint/Excel load-and-run
check remains a human step (plan §29) and cannot be automated here.

Fixtures are NEVER committed (binary, machine-specific). When absent, every
test here skips, so `pytest tests` stays green on machines without fixtures:

    tests/fixtures/xlam/RealAddin.xlam
    tests/fixtures/ppam/RealAddin.ppam

Fixture content contract (plan §4.1):
    ModuleKeep   - standard module, must survive untouched
    ModuleEdit   - standard module containing TEST_BUILD = "ORIGINAL"
    ModuleDelete - standard module, deleted by the cycle
"""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from vba_addin_editor.domain.document import (
    ModuleDisplayKind,
    ModuleDraft,
    new_module_id,
)
from vba_addin_editor.services.document_service import DocumentService
from vba_addin_editor.services.save_service import SaveService

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"

LIVE_PARAMS = [
    pytest.param(
        FIXTURES / "xlam" / "RealAddin.xlam", "excel", id="xlam-live"
    ),
    pytest.param(
        FIXTURES / "ppam" / "RealAddin.ppam", "powerpoint", id="ppam-live"
    ),
]


def _live_service() -> SaveService:
    """Real SaveService: real process/lock probes and ReplaceFileW commit."""
    return SaveService()


def _require_fixture(path: Path) -> None:
    if not path.exists():
        pytest.skip(
            f"Live fixture missing: {path}. Create it via Office Save As "
            "(plan §4.1) to run the release-gate cycle."
        )


def _require_modules(draft, names: tuple[str, ...]) -> None:
    present = {m.current_name for m in draft.modules if not m.is_deleted}
    missing = [n for n in names if n not in present]
    if missing:
        pytest.skip(
            f"Fixture does not match the plan §4.1 module contract; missing {missing}."
        )


@pytest.mark.live
@pytest.mark.parametrize(("fixture_path", "host_kind"), LIVE_PARAMS)
def test_full_mutation_cycle(fixture_path: Path, host_kind: str, tmp_path: Path):
    """Edit + add + rename + delete + save in place + reopen (plan §4.3, parser level)."""
    _require_fixture(fixture_path)
    work = tmp_path / f"Live_{fixture_path.stem}{fixture_path.suffix}"
    work.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(fixture_path, work)

    service = DocumentService()
    draft = service.open(work)
    assert draft.baseline.host_kind == host_kind
    _require_modules(draft, ("ModuleKeep", "ModuleEdit", "ModuleDelete"))

    # Edit ModuleEdit (plan step 6).
    edit = draft.find_current("ModuleEdit")
    assert edit is not None and edit.destructive_ops_safe
    assert 'TEST_BUILD As String = "ORIGINAL"' in edit.body
    edit.body = edit.body.replace('TEST_BUILD As String = "ORIGINAL"', 'TEST_BUILD As String = "EDITED"')

    # Add a standard module (plan step 11).
    draft.modules.append(
        ModuleDraft(
            id=new_module_id(),
            origin_name=None,
            current_name="LiveAdded",
            body='Public Function LiveAddedProbe() As String\r\n    LiveAddedProbe = "ADDED_OK"\r\nEnd Function\r\n',
            kind=ModuleDisplayKind.STANDARD,
            pyopenvba_kind="standard",
            is_new=True,
            is_deleted=False,
            destructive_ops_safe=True,
        )
    )

    # Rename ModuleEdit (plan step 15).
    edit.current_name = "ModuleEditRenamed"

    # Delete ModuleDelete (plan step 18).
    doomed = draft.find_current("ModuleDelete")
    assert doomed is not None
    doomed.is_deleted = True

    # Real in-place save: probe, fingerprint, candidate, verify, backup, ReplaceFileW.
    result = _live_service().save_addin(draft)
    assert result.kind == "success", getattr(result, "detail", result)

    # Same path preserved (plan constraint 6).
    assert work.exists()
    backup = result.backup_path
    assert backup is not None and backup.exists(), "mandatory backup missing"

    # Reopen and verify every mutation landed (plan steps 8-9, parser level).
    reopened = service.open(work)
    names = {m.current_name for m in reopened.modules if not m.is_deleted}
    assert "ModuleEditRenamed" in names
    assert "ModuleKeep" in names
    assert "LiveAdded" in names
    assert "ModuleEdit" not in names
    assert "ModuleDelete" not in names
    renamed = reopened.find_current("ModuleEditRenamed")
    assert renamed is not None and 'TEST_BUILD As String = "EDITED"' in renamed.body
    keep = reopened.find_current("ModuleKeep")
    snapshot_keep = next(
        m for m in reopened.baseline.modules if m.original_name == "ModuleKeep"
    )
    assert keep is not None and keep.body == snapshot_keep.body, "ModuleKeep must be untouched"


@pytest.mark.live
@pytest.mark.parametrize(("fixture_path", "host_kind"), LIVE_PARAMS)
def test_no_change_save_writes_nothing(fixture_path: Path, host_kind: str, tmp_path: Path):
    """Plan constraint 5: open-close cycle must not mutate the authentic file."""
    _require_fixture(fixture_path)
    work = tmp_path / f"Live_{fixture_path.stem}{fixture_path.suffix}"
    work.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(fixture_path, work)
    before = work.read_bytes()

    draft = DocumentService().open(work)
    _require_modules(draft, ("ModuleKeep",))
    result = _live_service().save_addin(draft)
    assert result.kind == "no_changes", result
    assert work.read_bytes() == before
