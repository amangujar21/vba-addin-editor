"""Open workflow: preflight, snapshot, fingerprint (plan 7.3)."""

from __future__ import annotations

import zipfile
from dataclasses import replace
from pathlib import Path

from vba_addin_editor.adapters.ooxml_package_adapter import (
    XML_EDITABLE_EXTENSIONS,
    OoxmlPackageAdapter,
    PackageError,
)
from vba_addin_editor.adapters.pyopenvba_adapter import (
    SUPPORTED_EXTENSIONS,
    AdapterError,
    PyOpenVBAAdapter,
)
from vba_addin_editor.domain.document import DocumentDraft, DocumentSnapshot, draft_from_snapshot
from vba_addin_editor.platform import paths
from vba_addin_editor.platform import windows_processes as wp


class DocumentService:
    def __init__(
        self,
        adapter: PyOpenVBAAdapter | None = None,
        package_adapter: OoxmlPackageAdapter | None = None,
    ) -> None:
        self.adapter = adapter or PyOpenVBAAdapter()
        self.package_adapter = package_adapter or OoxmlPackageAdapter()

    def open(self, path: Path) -> DocumentDraft:
        path = Path(path).resolve()
        if not path.is_file():
            raise AdapterError("The selected add-in file could not be found.")
        if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            raise AdapterError("Only .xlam, .ppam, and .pptm files are supported.")
        fp = paths.fingerprint(path)
        snapshot = self.adapter.open_snapshot(
            path,
            fp,
            host_process_running=wp.host_process_running(path),
        )
        snapshot = self._with_package_state(path, snapshot)
        return draft_from_snapshot(snapshot)

    def _with_package_state(self, path: Path, snapshot: DocumentSnapshot) -> DocumentSnapshot:
        """Attach XML part snapshots and package safety for editable packages."""
        if path.suffix.lower() not in XML_EDITABLE_EXTENSIONS:
            return snapshot
        try:
            xml_parts = self.package_adapter.snapshot_xml_parts(path)
            package_safety = self.package_adapter.inspect_package_signature(path)
        except PackageError as exc:
            raise AdapterError(str(exc), exc.details) from exc
        except (OSError, zipfile.BadZipFile) as exc:
            raise AdapterError(
                "This Office package could not be read; the file has not been changed.",
                {"exception": repr(exc)},
            ) from exc
        return replace(snapshot, xml_parts=xml_parts, package_safety=package_safety)

    def reload(self, draft: DocumentDraft) -> DocumentDraft:
        return self.open(draft.baseline.path)

    def is_current_on_disk(self, draft: DocumentDraft) -> bool:
        return paths.fingerprint_matches(draft.baseline.path, draft.baseline.file_fingerprint)


def snapshot_report(snapshot: DocumentSnapshot) -> str:
    """Human-readable one-line status used by banners."""
    bits = [f"{snapshot.extension} ", snapshot.project_name or "(unnamed project)"]
    flags: list[str] = []
    if snapshot.safety.password_protected:
        flags.append("password-protected")
    if snapshot.safety.signature_present:
        flags.append("digitally signed")
    if snapshot.package_safety.opc_signature_present:
        flags.append("package digitally signed")
    if snapshot.safety.host_process_running:
        flags.append("Office is running")
    return "".join(bits) + (" — " + ", ".join(flags) if flags else "")
