"""Start from nothing: a new blank add-in file, and a first ribbon for a file without one.

A new file is the Office-authored empty package pyOpenVBA ships (Excel's
add-in for .xlam, PowerPoint's macro-enabled presentation for .pptm, and the
same presentation with the add-in main content type for .ppam). It holds no
ribbon: ``add_starter_ribbon`` adds one to any draft as ordinary unsaved
changes (a new customUI part, its root relationship and, when needed, a
content type), so the next save writes and verifies it like any other edit.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from vba_addin_editor.adapters.ooxml_package_adapter import (
    CONTENT_TYPES_NAME,
    OoxmlPackageAdapter,
    open_package,
    package_member_names,
)
from vba_addin_editor.adapters.pyopenvba_adapter import SUPPORTED_EXTENSIONS, blank_template_bytes
from vba_addin_editor.domain.document import DocumentDraft, FileFingerprint, XmlPartDraft
from vba_addin_editor.platform import paths
from vba_addin_editor.platform.windows_file_ops import write_verified
from vba_addin_editor.services import ribbon_service as rs

RIBBON_PART_PATH = "customUI/customUI14.xml"
ROOT_RELATIONSHIPS_PATH = "_rels/.rels"
# customUI14.xml (the 2009/07 schema, Office 2010+) is linked with the 2007
# type; customUI.xml (2006/01) with the 2006 one.
UI_EXTENSIBILITY_REL_TYPE = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
UI_EXTENSIBILITY_REL_TYPES = frozenset(
    {
        UI_EXTENSIBILITY_REL_TYPE,
        "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility",
    }
)
CUSTOMUI_NAMESPACE = "http://schemas.microsoft.com/office/2009/07/customui"
_XML_CONTENT_TYPE = "application/xml"

_PRESENTATION_PART = "/ppt/presentation.xml"
_PPTM_MAIN_TYPE = "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
_PPAM_MAIN_TYPE = "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml"

DEFAULT_TAB_LABEL = "My Add-in"
DEFAULT_GROUP_LABEL = "Macros"
_ID_RE = re.compile(r"^[A-Za-z_][\w.-]*$")


class StarterError(ValueError):
    """A new file or ribbon could not be made; the message is for the user."""


# -- a first ribbon --------------------------------------------------------------


@dataclass(frozen=True)
class StarterRibbonSpec:
    tab_label: str = DEFAULT_TAB_LABEL
    group_label: str = DEFAULT_GROUP_LABEL
    tab_id: str = ""  # derived from the label when empty
    group_id: str = ""

    def ids(self) -> tuple[str, str]:
        tab_id = self.tab_id or rs.unique_control_id(set(), self.tab_label, "tab")
        group_id = self.group_id or rs.unique_control_id({tab_id}, self.group_label, "grp")
        return tab_id, group_id


def starter_ribbon_xml(spec: StarterRibbonSpec) -> str:
    """customUI with one custom tab holding one empty group, ready for buttons."""
    tab_id, group_id = spec.ids()
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<customUI xmlns="{CUSTOMUI_NAMESPACE}">\n'
        "  <ribbon>\n"
        "    <tabs>\n"
        f'      <tab id="{rs.xml_attr(tab_id)}" label="{rs.xml_attr(spec.tab_label)}">\n'
        f'        <group id="{rs.xml_attr(group_id)}" label="{rs.xml_attr(spec.group_label)}">\n'
        "        </group>\n"
        "      </tab>\n"
        "    </tabs>\n"
        "  </ribbon>\n"
        "</customUI>\n"
    )


def validate_spec(spec: StarterRibbonSpec) -> str | None:
    """Why ``spec`` cannot be used, or None."""
    if not spec.tab_label.strip():
        return "Enter a label for the tab."
    if not spec.group_label.strip():
        return "Enter a label for the group."
    tab_id, group_id = spec.ids()
    for value in (tab_id, group_id):
        if not _ID_RE.match(value):
            return f'"{value}" is not a valid control id: start with a letter or underscore and use letters, digits, _ . -'
    if tab_id.casefold() == group_id.casefold():
        return "The tab and the group need different ids."
    return None


def _package_members(draft: DocumentDraft) -> set[str]:
    source = draft.baseline.package_bytes
    names = set(package_member_names(source if source is not None else draft.baseline.path))
    return {name.lower() for name in names}


def ribbon_blocker(draft: DocumentDraft) -> str | None:
    """Why a first ribbon cannot be added to ``draft``, or None when it can."""
    if any(rs.is_ribbon_part(p.path) for p in draft.xml_parts):
        return "This file already has ribbon XML (customUI)."
    if any(rs.is_ribbon_part(name) for name in _package_members(draft)):
        return "This file already has ribbon XML (customUI), but it cannot be edited here."
    if draft.baseline.package_safety.opc_signature_present:
        return "This file has a package digital signature; adding a ribbon would break it."
    for path in (ROOT_RELATIONSHIPS_PATH, CONTENT_TYPES_NAME):
        part = draft.xml_part_by_path(path)
        if part is None or part.text is None or not part.editable:
            return f"{path} is missing or not editable, so no ribbon can be added."
    rels = draft.xml_part_by_path(ROOT_RELATIONSHIPS_PATH)
    assert rels is not None and rels.text is not None
    for _rid, kind, target in rs.relationships(rels.text):
        if kind in UI_EXTENSIBILITY_REL_TYPES and target.lstrip("/") != RIBBON_PART_PATH:
            return (
                f"_rels/.rels already links a ribbon part ({target}) that is not in the file. "
                "Fix or remove that relationship on the XML tab first."
            )
    return None


def add_starter_ribbon(draft: DocumentDraft, spec: StarterRibbonSpec) -> XmlPartDraft:
    """Give ``draft`` a ribbon: a new customUI14.xml linked from the package root.

    Edits ``_rels/.rels`` and, only when .xml files are not already typed
    application/xml, ``[Content_Types].xml``. Returns the new ribbon part.
    Raises StarterError (draft untouched) when it cannot be done.
    """
    problem = ribbon_blocker(draft) or validate_spec(spec)
    if problem:
        raise StarterError(problem)
    rels = draft.xml_part_by_path(ROOT_RELATIONSHIPS_PATH)
    types = draft.xml_part_by_path(CONTENT_TYPES_NAME)
    assert rels is not None and rels.text is not None
    assert types is not None and types.text is not None
    try:
        rels_text = rels.text
        linked = any(
            kind in UI_EXTENSIBILITY_REL_TYPES and target.lstrip("/") == RIBBON_PART_PATH
            for _rid, kind, target in rs.relationships(rels_text)
        )
        if not linked:
            rels_text, _rid = rs.add_relationship(
                rels_text, UI_EXTENSIBILITY_REL_TYPE, RIBBON_PART_PATH, id_prefix="rIdUI"
            )
        types_text = types.text
        part_name = "/" + RIBBON_PART_PATH
        if rs.override_content_type(types_text, part_name) is None:
            default = rs.default_content_type(types_text, "xml")
            if default is None:
                types_text = rs.ensure_default_content_type(types_text, "xml", _XML_CONTENT_TYPE)
            elif default != _XML_CONTENT_TYPE:
                types_text = rs.add_override_content_type(types_text, part_name, _XML_CONTENT_TYPE)
    except rs.RibbonEditError as exc:
        raise StarterError(str(exc)) from exc

    rels.text = rels_text
    types.text = types_text
    part = XmlPartDraft(
        path=RIBBON_PART_PATH,
        text=starter_ribbon_xml(spec),
        original_text=None,
        encoding="utf-8",
        bom=b"",
        newline="\r\n",
        is_relationships_part=False,
        is_content_types_part=False,
        is_new=True,
    )
    draft.xml_parts.append(part)
    return part


# -- a new file ------------------------------------------------------------------


def blank_package(extension: str) -> bytes:
    """The bytes of a new, empty macro file of ``extension`` (.xlam, .ppam, .pptm)."""
    suffix = extension.lower()
    if suffix not in SUPPORTED_EXTENSIONS:
        raise StarterError("Only .xlam, .ppam, and .pptm files can be created.")
    data = blank_template_bytes(suffix)
    if suffix != ".ppam":
        return data
    # A PowerPoint add-in is a macro-enabled presentation whose main part is
    # typed as an add-in; nothing else in the package differs.
    with open_package(data) as package:
        types = package.read(CONTENT_TYPES_NAME).decode("utf-8")
    if rs.override_content_type(types, _PRESENTATION_PART) != _PPTM_MAIN_TYPE:
        raise StarterError("The PowerPoint template is not the one this editor expects.")
    types = types.replace(
        f'ContentType="{_PPTM_MAIN_TYPE}"', f'ContentType="{_PPAM_MAIN_TYPE}"', 1
    )
    return OoxmlPackageAdapter().build_package(data, {CONTENT_TYPES_NAME: types.encode("utf-8")})


def create_blank_file(path: Path) -> tuple[bytes, FileFingerprint]:
    """Write a new empty macro file at ``path``; never replaces an existing file.

    Returns the bytes written and their fingerprint, so the file can be opened
    without reading it back (a read of a fresh macro file waits for antivirus).
    """
    path = Path(path)
    data = blank_package(path.suffix)
    if path.exists():
        raise StarterError(f"{path.name} already exists. Choose a new file name.")
    try:
        write_verified(path, data)
    except FileExistsError as exc:
        raise StarterError(f"{path.name} already exists. Choose a new file name.") from exc
    except OSError:
        try:
            path.unlink()
        except OSError:
            pass
        raise
    return data, paths.written_fingerprint(data, path)
