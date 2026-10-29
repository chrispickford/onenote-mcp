"""Targeted item edits on a page tree read with GetPageContent(page_id, 1).

OneNote replaces an outline's whole child list when the outline is sent with
its objectID: OEs sent with their objectID keep it, OEs without one are new,
and OEs left out are deleted (verified against OneNote 16.0.20527).  So each
edit modifies the page tree in place and returns the complete outlines it
touched, to be sent in one UpdatePageContent.  The binary read matters:
re-sent images keep their bytes only if <one:Data> is present.

Handwriting, recordings and content from newer OneNote versions cannot be
re-sent, so an outline that still holds any is refused.  Pure functions; no COM.
"""

from dataclasses import dataclass, field
import xml.etree.ElementTree as ET

from onenote_mcp.builders import (
    HEADING_QUICK_STYLE_INDEX, TAG_LOCAL_INDEX, apply_cdata, build_item_elements, oe_content_kind,
    quick_style_def_el, tag_def_el, tag_symbol,
)
from onenote_mcp.errors import OneNoteError
from onenote_mcp.models import ListChild, ListItem

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
_RESEND_UNSAFE = frozenset({"ink", "media", "unknown"})


def _q(tag: str) -> str:
    return f"{{{_ONE_NS}}}{tag}"


@dataclass
class Edit:
    """The result of an edit: what to send and what to delete."""

    outlines: list[ET.Element] = field(default_factory=list)
    delete_outline_ids: list[str] = field(default_factory=list)
    cdata_map: dict[int, str] = field(default_factory=dict)
    tag_defs: list[ET.Element] = field(default_factory=list)
    quick_style_defs: list[ET.Element] = field(default_factory=list)
    new_oes: list[ET.Element] = field(default_factory=list)


# ---- Lookups -----------------------------------------------------------------

def _parents(page_root: ET.Element) -> dict[ET.Element, ET.Element]:
    return {child: parent for parent in page_root.iter() for child in parent}


def find_oe(page_root: ET.Element, item_id: str) -> tuple[ET.Element, ET.Element, ET.Element]:
    """Return (OE, its parent <one:OEChildren>, its <one:Outline>) for *item_id*.

    Only OEs inside outlines count (not the title).  bad_request if absent.
    """
    parents = _parents(page_root)
    for outline in page_root.findall(_q("Outline")):
        for oe in outline.iter(_q("OE")):
            if oe.get("objectID") == item_id:
                return oe, parents[oe], outline
    raise OneNoteError(
        f"Item {item_id!r} not found on the page; re-read it with get_page for current IDs.",
        code="bad_request",
    )


def _find_outline(page_root: ET.Element, outline_id: str) -> ET.Element:
    for outline in page_root.findall(_q("Outline")):
        if outline.get("objectID") == outline_id:
            return outline
    raise OneNoteError(
        f"Outline {outline_id!r} not found on the page; re-read it with get_page for current IDs.",
        code="bad_request",
    )


def _list_style_of(oe: ET.Element) -> str:
    """'bullet' or 'numbered' for a list OE; 'bullet' for anything else."""
    list_el = oe.find(_q("List"))
    if list_el is not None and list_el.find(_q("Number")) is not None:
        return "numbered"
    return "bullet"


# ---- New content and its styles ------------------------------------------------

def _max_index(defs: list[ET.Element]) -> int:
    return max((int(d.get("index", "-1")) for d in defs), default=-1)


def _build(page_root: ET.Element, items: list[ListChild], list_style: str, edit: Edit) -> list[ET.Element]:
    """Build OEs for *items* with heading and tag indexes mapped to the page's tables.

    A page definition with the same heading name (or tag symbol) is reused;
    otherwise a new definition at the next free index goes into *edit*.
    """
    oes, cdata_map = build_item_elements(items, list_style)
    edit.cdata_map.update(cdata_map)
    edit.new_oes.extend(oes)
    page_qsds = page_root.findall(_q("QuickStyleDef")) + edit.quick_style_defs
    page_tags = page_root.findall(_q("TagDef")) + edit.tag_defs
    heading_by_local = {index: name for name, index in HEADING_QUICK_STYLE_INDEX.items()}
    tag_by_local = {index: kind for kind, index in TAG_LOCAL_INDEX.items()}
    qsd_index: dict[str, str] = {}
    tag_index: dict[str, str] = {}
    for oe in [el for root in oes for el in root.iter(_q("OE"))]:
        local = oe.get("quickStyleIndex")
        if local is not None:
            name = heading_by_local[local]
            if name not in qsd_index:
                existing = next((d for d in page_qsds if d.get("name") == name), None)
                if existing is None:
                    existing = quick_style_def_el(name)
                    existing.set("index", str(_max_index(page_qsds) + 1))
                    page_qsds.append(existing)
                    edit.quick_style_defs.append(existing)
                qsd_index[name] = existing.get("index")
            oe.set("quickStyleIndex", qsd_index[name])
        for tag in oe.findall(_q("Tag")):
            kind = tag_by_local[tag.get("index")]
            if kind not in tag_index:
                existing = next((d for d in page_tags if d.get("symbol") == tag_symbol(kind)), None)
                if existing is None:
                    existing = tag_def_el(kind, str(_max_index(page_tags) + 1))
                    page_tags.append(existing)
                    edit.tag_defs.append(existing)
                tag_index[kind] = existing.get("index")
            tag.set("index", tag_index[kind])
    return oes


# ---- Sending -------------------------------------------------------------------

def _touch(edit: Edit, outline: ET.Element) -> None:
    if outline not in edit.outlines:
        edit.outlines.append(outline)


def _finish(page_root: ET.Element, edit: Edit) -> Edit:
    """Drop emptied OEChildren, turn emptied outlines into deletions, order the
    outlines as on the page, and refuse outlines that cannot be re-sent."""
    for outline in list(edit.outlines):
        for oe_children in [el for el in outline.iter(_q("OEChildren"))]:
            if oe_children.find(_q("OE")) is None and oe_children is not outline.find(_q("OEChildren")):
                parent = _parents(outline)[oe_children]
                parent.remove(oe_children)
        top = outline.find(_q("OEChildren"))
        if top is None or top.find(_q("OE")) is None:
            edit.outlines.remove(outline)
            if outline.get("objectID"):
                edit.delete_outline_ids.append(outline.get("objectID"))
    order = page_root.findall(_q("Outline"))
    edit.outlines.sort(key=order.index)
    for outline in edit.outlines:
        unsafe = [
            (oe.get("objectID", ""), kind) for oe in outline.iter(_q("OE"))
            if (kind := oe_content_kind(oe)) in _RESEND_UNSAFE
        ]
        if unsafe:
            listed = ", ".join(f"{kind} {object_id}" for object_id, kind in unsafe)
            raise OneNoteError(
                f"This outline holds content that cannot be re-sent: {listed}. Editing items in "
                "the same outline would delete it; edit it in OneNote, or delete that content "
                "first if that is intended.",
                code="bad_request",
            )
    return edit


def edit_xml(page_id: str, page_root: ET.Element, edit: Edit) -> str:
    """Return the <one:Page> XML to send for *edit*.

    All of the page's TagDefs and QuickStyleDefs are sent too (plus the new
    ones): indexes in written XML are local to it, and the re-sent OEs refer
    to the page's own indexes.
    """
    page = ET.Element(_q("Page"), {"ID": page_id})
    page.extend(page_root.findall(_q("TagDef")) + edit.tag_defs)
    page.extend(page_root.findall(_q("QuickStyleDef")) + edit.quick_style_defs)
    page.extend(edit.outlines)
    return apply_cdata(ET.tostring(page, encoding="unicode", xml_declaration=False), edit.cdata_map)


# ---- Operations ------------------------------------------------------------------

def update_item(page_root: ET.Element, item_id: str, item: ListChild) -> Edit:
    """Replace the item *item_id* with *item*.

    A single new OE keeps the old objectID; a list item given as a plain
    ListItem keeps the list style of the item it replaces.  Content the tools
    cannot write (attachments, handwriting, ...) cannot be changed.
    """
    oe, parent, outline = find_oe(page_root, item_id)
    if oe_content_kind(oe) is not None:
        raise OneNoteError(
            f"Item {item_id!r} holds content that cannot be changed ({oe_content_kind(oe)}); "
            "it can only be moved or deleted.",
            code="bad_request",
        )
    edit = Edit()
    new = _build(page_root, [item], _list_style_of(oe), edit)
    if len(new) == 1:
        new[0].set("objectID", item_id)
    position = list(parent).index(oe)
    parent.remove(oe)
    for offset, new_oe in enumerate(new):
        parent.insert(position + offset, new_oe)
    _touch(edit, outline)
    return _finish(page_root, edit)


def _anchor(
    page_root: ET.Element,
    before_id: str | None,
    after_id: str | None,
    outline_id: str | None,
    at_end: bool,
) -> tuple[ET.Element, int, ET.Element, ET.Element | None]:
    """Resolve an insert position: (parent OEChildren, index, outline, anchor OE or None)."""
    if sum(anchor is not None for anchor in (before_id, after_id, outline_id)) != 1:
        raise OneNoteError("Give exactly one of before_id, after_id or outline_id.", code="bad_request")
    if outline_id is not None:
        outline = _find_outline(page_root, outline_id)
        top = outline.find(_q("OEChildren"))
        return top, len(top) if at_end else 0, outline, None
    oe, parent, outline = find_oe(page_root, before_id or after_id)
    index = list(parent).index(oe) + (0 if before_id is not None else 1)
    return parent, index, outline, oe


def insert_items(
    page_root: ET.Element,
    items: list[ListChild],
    before_id: str | None = None,
    after_id: str | None = None,
    outline_id: str | None = None,
    at_end: bool = True,
) -> Edit:
    """Insert *items* before or after an item, or at the start or end of an outline.

    Plain ListItems take the list style of the anchor item when it is a list item.
    """
    parent, index, outline, anchor = _anchor(page_root, before_id, after_id, outline_id, at_end)
    edit = Edit()
    new = _build(page_root, items, _list_style_of(anchor) if anchor is not None else "bullet", edit)
    for offset, new_oe in enumerate(new):
        parent.insert(index + offset, new_oe)
    _touch(edit, outline)
    return _finish(page_root, edit)


def delete_items(page_root: ET.Element, item_ids: list[str]) -> Edit:
    """Delete the items *item_ids* (with what is indented under them).

    An outline left without items is deleted as a whole.
    """
    edit = Edit()
    for item_id in item_ids:
        oe, parent, outline = find_oe(page_root, item_id)
        parent.remove(oe)
        _touch(edit, outline)
    return _finish(page_root, edit)


def move_item(
    page_root: ET.Element,
    item_id: str,
    before_id: str | None = None,
    after_id: str | None = None,
    under_id: str | None = None,
) -> Edit:
    """Move the item *item_id* (with its indented content) before or after an
    item, or to the end of the content indented under *under_id*.

    Within one outline the item keeps its objectID; moved to another outline
    it is written as new content there (OneNote gives it a new ID).
    """
    if sum(anchor is not None for anchor in (before_id, after_id, under_id)) != 1:
        raise OneNoteError("Give exactly one of before_id, after_id or under_id.", code="bad_request")
    oe, source_parent, source_outline = find_oe(page_root, item_id)
    target_id = before_id or after_id or under_id
    if target_id == item_id or any(el.get("objectID") == target_id for el in oe.iter(_q("OE"))):
        raise OneNoteError(f"Cannot move item {item_id!r} into itself.", code="bad_request")
    target, target_parent, target_outline = find_oe(page_root, target_id)
    source_parent.remove(oe)
    if under_id is not None:
        children = target.find(_q("OEChildren"))
        if children is None:
            children = ET.SubElement(target, _q("OEChildren"))
        children.append(oe)
    else:
        index = list(target_parent).index(target) + (0 if before_id is not None else 1)
        target_parent.insert(index, oe)
    if target_outline is not source_outline:
        for el in oe.iter():
            el.attrib.pop("objectID", None)
    edit = Edit()
    _touch(edit, source_outline)
    _touch(edit, target_outline)
    return _finish(page_root, edit)
