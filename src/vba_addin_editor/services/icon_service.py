"""Custom ribbon icons: turn a picture into a small PNG and wire it into the draft.

Pillow is imported on first use only, so opening the editor does not load it.
Images become package members of the draft (``BinaryPartDraft``) and are
written by the next save like any other change.
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass
from pathlib import Path

from vba_addin_editor.adapters.ooxml_package_adapter import (
    CONTENT_TYPES_NAME,
    MAX_ADDED_MEMBER_BYTES,
    is_removable_member,
    package_member_names,
    read_member_from_bytes,
)
from vba_addin_editor.domain.document import BinaryPartDraft, DocumentDraft, XmlPartDraft
from vba_addin_editor.services import ribbon_service as rs

ICON_SIZES = (16, 32)
DEFAULT_ICON_SIZE = 32
# A 32x32 PNG is a few KB; anything near this has not been shrunk.
MAX_ICON_BYTES = 64 * 1024
MAX_SOURCE_PIXELS = 40_000_000  # refuse decompression bombs before resizing
PNG_CONTENT_TYPE = "image/png"
_RELS_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml"


class IconError(ValueError):
    """An icon could not be made or added; the message is for the user."""


@dataclass(frozen=True)
class IconImage:
    png: bytes
    size: int  # square edge in pixels
    source: str  # where it came from, for status text

    def describe(self) -> str:
        return f"{self.size}×{self.size} PNG, {len(self.png) / 1024:.1f} KB"


# -- making the PNG ---------------------------------------------------------------


def _pil():
    try:
        from PIL import Image
    except ImportError as exc:  # pragma: no cover - Pillow is a declared dependency
        raise IconError("Custom icons need the Pillow package, which is not installed.") from exc
    return Image


def make_icon(image, size: int = DEFAULT_ICON_SIZE, source: str = "") -> IconImage:
    """A square, transparent-padded PNG of ``image`` (a PIL image) at ``size``."""
    Image = _pil()
    if size not in ICON_SIZES:
        raise IconError(f"Icon size must be one of {ICON_SIZES}.")
    width, height = image.size
    if width < 1 or height < 1:
        raise IconError("The image is empty.")
    if width * height > MAX_SOURCE_PIXELS:
        raise IconError("The image is too large to use as an icon.")
    picture = image.convert("RGBA")
    # Trim fully transparent borders so the visible shape fills the icon.
    box = picture.getchannel("A").getbbox()
    if box is not None:
        picture = picture.crop(box)
    width, height = picture.size
    edge = max(width, height)
    square = Image.new("RGBA", (edge, edge), (0, 0, 0, 0))
    square.paste(picture, ((edge - width) // 2, (edge - height) // 2))
    icon = square.resize((size, size), Image.Resampling.LANCZOS)
    out = io.BytesIO()
    icon.save(out, format="PNG", optimize=True)
    data = out.getvalue()
    if len(data) > MAX_ICON_BYTES:  # pragma: no cover - a 32x32 PNG cannot get this big
        raise IconError("The icon came out larger than expected; try a simpler image.")
    return IconImage(png=data, size=size, source=source)


def icon_from_bytes(data: bytes, size: int = DEFAULT_ICON_SIZE, source: str = "") -> IconImage:
    Image = _pil()
    Image.MAX_IMAGE_PIXELS = MAX_SOURCE_PIXELS
    try:
        with Image.open(io.BytesIO(data)) as image:
            image.load()
            return make_icon(image, size, source)
    except IconError:
        raise
    except Exception as exc:  # PIL raises many types for unreadable input
        raise IconError(f"That file is not an image this editor can read ({exc}).") from exc


def icon_from_file(path: Path, size: int = DEFAULT_ICON_SIZE) -> IconImage:
    try:
        data = Path(path).read_bytes()
    except OSError as exc:
        raise IconError(f"The image could not be read: {exc}") from exc
    return icon_from_bytes(data, size, source=Path(path).name)


def icon_from_clipboard(size: int = DEFAULT_ICON_SIZE) -> IconImage:
    """The image on the clipboard: a copied picture, or a copied image file."""
    _pil()
    from PIL import ImageGrab

    try:
        grabbed = ImageGrab.grabclipboard()
    except Exception as exc:  # pragma: no cover - depends on clipboard owner
        raise IconError(f"The clipboard could not be read ({exc}).") from exc
    if isinstance(grabbed, list):
        files = [Path(name) for name in grabbed if name]
        if not files:
            raise IconError("The clipboard holds no image.")
        return icon_from_file(files[0], size)
    if grabbed is None:
        raise IconError(
            "The clipboard holds no image. In your browser, right-click a picture "
            "and choose Copy image, then try again."
        )
    grabbed.load()
    return make_icon(grabbed, size, source="clipboard")


# -- adding it to the draft ------------------------------------------------------------


def existing_member_names(draft: DocumentDraft) -> set[str]:
    """Lower-cased names of every member in the package as opened, plus draft additions."""
    source = draft.baseline.package_bytes
    names = set(package_member_names(source if source is not None else draft.baseline.path))
    names |= {p.path for p in draft.xml_parts}
    names |= {p.path for p in draft.binary_parts}
    return {name.lower() for name in names}


@dataclass(frozen=True)
class ImageLink:
    ribbon_path: str
    rels_path: str
    rid: str
    member: str  # package member the relationship points at
    used: bool  # some control in the ribbon part has image="rid"


def image_links(draft: DocumentDraft) -> list[ImageLink]:
    """Every image relationship of every ribbon part, as the draft has them."""
    links: list[ImageLink] = []
    for part in rs.ribbon_parts(draft):
        rels = draft.xml_part_by_path(rs.relationships_path(part.path))
        if rels is None or rels.text is None:
            continue
        used_ids = {
            e.attr("image") for e in rs.element_tree(part.text or "") if e.attr("image")
        }
        for rid, kind, target in rs.relationships(rels.text):
            if kind not in rs.IMAGE_REL_TYPES:
                continue
            member = rs.relationship_target_path(part.path, target)
            if member is not None:
                links.append(ImageLink(part.path, rels.path, rid, member, rid in used_ids))
    return links


def linked_image_paths(draft: DocumentDraft) -> set[str]:
    """Image members some ribbon control still shows."""
    return {link.member for link in image_links(draft) if link.used}


def _drop_unused_link(draft: DocumentDraft, link: ImageLink, keep: set[str]) -> str | None:
    """Remove ``link``'s relationship, and its image when nothing else shows it.

    Returns the image member dropped, if any. ``keep`` = members still shown.
    """
    rels = draft.xml_part_by_path(link.rels_path)
    if rels is None or rels.text is None or not rels.editable:
        return None
    rels.text = rs.remove_relationship(rels.text, link.rid)
    if link.member in keep or not is_removable_member(link.member):
        return None
    if any(l.member == link.member for l in image_links(draft)):
        return None  # another relationship still points at it
    pending = draft.binary_part_by_path(link.member)
    if pending is not None:
        draft.binary_parts.remove(pending)
        return link.member
    source = draft.baseline.package_bytes
    present = set(package_member_names(source if source is not None else draft.baseline.path))
    if link.member in present and link.member not in draft.removed_parts:
        draft.removed_parts.append(link.member)
        return link.member
    return None


def remove_unused_images(draft: DocumentDraft) -> list[str]:
    """Drop image relationships no control uses, and images nothing shows any more.

    Returns the image members removed. Signed packages and read-only parts are
    left alone.
    """
    if draft.baseline.package_safety.opc_signature_present:
        return []
    removed: list[str] = []
    while True:
        keep = linked_image_paths(draft)
        unused = [
            link for link in image_links(draft)
            if not link.used and (p := draft.xml_part_by_path(link.rels_path)) is not None
            and p.editable
        ]
        if not unused:
            return removed
        dropped = _drop_unused_link(draft, unused[0], keep)
        if dropped is not None:
            removed.append(dropped)


def _release_image(draft: DocumentDraft, ribbon_path: str, rid: str) -> None:
    """After a control stopped using ``rid``: drop it if no other control uses it."""
    link = next(
        (l for l in image_links(draft) if l.ribbon_path == ribbon_path and l.rid == rid), None
    )
    if link is not None and not link.used:
        _drop_unused_link(draft, link, linked_image_paths(draft))


IMAGE_CONTENT_TYPES = {
    "png": PNG_CONTENT_TYPE,
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "ico": "image/x-icon",
    "tif": "image/tiff",
    "tiff": "image/tiff",
    "emf": "image/x-emf",
    "wmf": "image/x-wmf",
}


def _image_path(folder: str, control_id: str, taken: set[str], extension: str = "png") -> str:
    stem = re.sub(r"[^A-Za-z0-9_.-]", "", control_id).strip(".") or "icon"
    prefix = f"{folder}/images/" if folder else "images/"
    candidate, n = f"{prefix}{stem}.{extension}", 2
    while candidate.lower() in taken:
        candidate, n = f"{prefix}{stem}{n}.{extension}", n + 1
    return candidate


def add_image_member(
    draft: DocumentDraft, part_path: str, data: bytes, *, extension: str, stem: str
) -> tuple[str, str]:
    """Add ``data`` as an image of ribbon part ``part_path``; returns (member path, rId).

    Edits the part's relationships (created when missing) and
    [Content_Types].xml and adds the image; the ribbon part itself is not
    touched: the caller points a control's image= at the returned rId.
    Raises IconError (draft untouched) when it cannot be done.
    """
    extension = extension.lower().lstrip(".")
    content_type = IMAGE_CONTENT_TYPES.get(extension)
    if content_type is None:
        raise IconError(f"A .{extension} file cannot be used as a ribbon image.")
    if len(data) > MAX_ADDED_MEMBER_BYTES:
        raise IconError("The image is too large.")
    if draft.baseline.package_safety.opc_signature_present:
        raise IconError("This file has a package digital signature; adding an image would break it.")
    types = draft.xml_part_by_path(CONTENT_TYPES_NAME)
    if types is None or types.text is None or not types.editable:
        raise IconError("[Content_Types].xml is missing or not editable, so no image can be added.")
    rels_path = rs.relationships_path(part_path)
    rels = draft.xml_part_by_path(rels_path)
    if rels is not None and (rels.text is None or not rels.editable):
        raise IconError(f"{rels_path} is not editable, so no image can be linked to the ribbon.")
    rels_text = rels.text if rels is not None and rels.text is not None else rs.new_relationships_text()
    rels_text = rs.repair_image_relationships(rels_text)  # links an older version wrote
    folder = part_path.rpartition("/")[0]
    try:
        types_text = rs.ensure_default_content_type(types.text, extension, content_type)
        if rels is None:
            types_text = rs.ensure_default_content_type(types_text, "rels", _RELS_CONTENT_TYPE)
        path = _image_path(folder, stem, existing_member_names(draft), extension)
        target = path[len(folder) + 1 :] if folder else path
        rels_text, rid = rs.add_relationship(rels_text, rs.IMAGE_REL_TYPE, target)
    except rs.RibbonEditError as exc:
        raise IconError(str(exc)) from exc

    types.text = types_text
    if rels is None:
        draft.xml_parts.append(
            XmlPartDraft(
                path=rels_path,
                text=rels_text,
                original_text=None,
                encoding="utf-8",
                bom=b"",
                newline="\n",
                is_relationships_part=True,
                is_content_types_part=False,
                is_new=True,
            )
        )
    else:
        rels.text = rels_text
    draft.binary_parts.append(BinaryPartDraft(path, data, content_type))
    return path, rid


def set_custom_icon(
    draft: DocumentDraft, part_path: str, control_start: int, icon: IconImage
) -> str:
    """Make the control whose tag starts at ``control_start`` show ``icon``.

    Edits the ribbon part, its relationships part (created when missing) and
    [Content_Types].xml in the draft, and adds the image. Returns the image's
    package path. Raises IconError (draft untouched) when it cannot be done.
    """
    if len(icon.png) > min(MAX_ICON_BYTES, MAX_ADDED_MEMBER_BYTES):
        raise IconError("The icon is too large.")
    ribbon = draft.xml_part_by_path(part_path)
    if ribbon is None or ribbon.text is None or not ribbon.editable:
        raise IconError("The ribbon XML part is not editable.")
    element = next(
        (e for e in rs.element_tree(ribbon.text) if e.start == control_start and e.complete), None
    )
    if element is None or element.name.rsplit(":", 1)[-1] not in rs.IMAGE_ELEMENTS:
        raise IconError("Put the cursor on a button, menu or other control that can show an icon.")
    if draft.baseline.package_safety.opc_signature_present:
        raise IconError(
            "This file has a package digital signature; adding an icon would break it."
        )
    types = draft.xml_part_by_path(CONTENT_TYPES_NAME)
    if types is None or types.text is None or not types.editable:
        raise IconError("[Content_Types].xml is missing or not editable, so no image can be added.")

    rels_path = rs.relationships_path(part_path)
    rels = draft.xml_part_by_path(rels_path)
    if rels is not None and (rels.text is None or not rels.editable):
        raise IconError(f"{rels_path} is not editable, so no image can be linked to the ribbon.")
    folder = part_path.rpartition("/")[0]

    # Replacing an icon this draft added: overwrite the image, keep the link.
    current = element.attr("image")
    if current and rels is not None and rels.text is not None:
        rels_text = rs.repair_image_relationships(rels.text)
        targets = {rid: target for rid, kind, target in rs.relationships(rels_text) if kind in rs.IMAGE_REL_TYPES}
        target = targets.get(current)
        if target is not None:
            path = f"{folder}/{target}" if folder else target
            added = draft.binary_part_by_path(path)
            if added is not None:
                added.data = icon.png
                rels.text = rels_text
                return path

    control_id = element.attr("id") or element.attr("idQ") or "icon"
    path, rid = add_image_member(
        draft, part_path, icon.png, extension="png", stem=control_id.rsplit(":", 1)[-1]
    )
    ribbon.text = rs.set_control_image(ribbon.text, element, rid)
    if current:
        # The old picture: gone from the file too, unless another control shows it.
        _release_image(draft, part_path, current)
    return path


def clear_custom_icon(
    draft: DocumentDraft, part_path: str, control_start: int, image_mso: str = ""
) -> None:
    """Take the custom picture off a control, optionally giving it a built-in icon instead.

    The picture leaves the file too, unless another control still shows it.
    """
    ribbon = draft.xml_part_by_path(part_path)
    if ribbon is None or ribbon.text is None or not ribbon.editable:
        raise IconError("The ribbon XML part is not editable.")
    element = next(
        (e for e in rs.element_tree(ribbon.text) if e.start == control_start and e.complete), None
    )
    if element is None:
        raise IconError("That control is no longer in the ribbon XML.")
    current = element.attr("image")
    ribbon.text = rs.set_element_attrs(
        ribbon.text, element, {"image": None, "imageMso": image_mso or None}
    )
    if current and not draft.baseline.package_safety.opc_signature_present:
        _release_image(draft, part_path, current)


def image_bytes(draft: DocumentDraft, part_path: str, rid: str) -> bytes | None:
    """The picture behind a control's image="rid", when it is at hand in memory.

    Pictures this draft added come from the draft; pictures already in the file
    come from the package bytes kept since it was opened (the file is not read).
    """
    link = next((l for l in image_links(draft) if l.ribbon_path == part_path and l.rid == rid), None)
    if link is None or link.member in draft.removed_parts:
        return None
    pending = draft.binary_part_by_path(link.member)
    if pending is not None:
        return pending.data
    package = draft.baseline.package_bytes
    return read_member_from_bytes(package, link.member) if package is not None else None
