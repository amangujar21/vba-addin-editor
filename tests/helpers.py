from __future__ import annotations

from vba_addin_editor.adapters.pyopenvba_adapter import PyOpenVBAAdapter
from vba_addin_editor.services.save_service import SaveService


def adapter():
    return PyOpenVBAAdapter()


def make_service(**overrides):
    base = {
        "adapter": adapter(),
        "process_probe": lambda p: False,
        "exclusive_probe": lambda p: None,
        "commit": None,
        "progress": None,
    }
    base.update(overrides)
    return SaveService(**base)


def edit_module1(draft) -> None:
    m1 = next(m for m in draft.modules if m.current_name == "Module1")
    m1.body = m1.body.replace("ORIGINAL", "V2")
