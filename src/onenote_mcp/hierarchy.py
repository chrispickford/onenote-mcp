"""Hierarchy XML (notebooks, section groups, sections, pages): parsing for the
navigation tools and lookups shared by the structure and allowlist code.

Pure functions on XML; no COM.
"""

import xml.etree.ElementTree as ET

from onenote_mcp.errors import OneNoteError

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
_NOTEBOOK = f"{{{_ONE_NS}}}Notebook"
_SECTION_GROUP = f"{{{_ONE_NS}}}SectionGroup"
_SECTION = f"{{{_ONE_NS}}}Section"


def _in_recycle_bin(el: ET.Element) -> bool:
    return el.get("isRecycleBin") == "true" or el.get("isInRecycleBin") == "true"


def _section(el: ET.Element) -> dict:
    section = {"id": el.get("ID", ""), "name": el.get("name", "")}
    if el.get("color") and el.get("color") != "none":
        section["color"] = el.get("color")
    return section


def _container(el: ET.Element) -> dict:
    """Return the sections and section groups directly in *el* (recursively for groups).

    ``section_groups`` is left out when there are none.
    """
    result: dict = {
        "sections": [_section(child) for child in el.findall(_SECTION) if not _in_recycle_bin(child)],
    }
    groups = [
        {"id": child.get("ID", ""), "name": child.get("name", ""), **_container(child)}
        for child in el.findall(_SECTION_GROUP) if not _in_recycle_bin(child)
    ]
    if groups:
        result["section_groups"] = groups
    return result


def parse_notebook_tree(xml: str) -> list[dict]:
    """Parse hierarchy XML into the get_notebooks navigation tree (no pages).

    Shape: [{"id", "name", "sections": [{"id", "name", "color"?}],
    "section_groups"?: [{"id", "name", "sections", "section_groups"?}]}].
    A notebook's ``sections`` holds only the sections directly in it, so
    walking the tree lists every section once.  ``name`` is the display name
    (OneNote's nickname, falling back to the folder name).  OneNote's recycle
    bin and the sections in it are left out.
    """
    root = ET.fromstring(xml)
    return [
        {"id": nb.get("ID", ""), "name": nb.get("nickname") or nb.get("name", ""), **_container(nb)}
        for nb in root.findall(_NOTEBOOK)
    ]


def find_object(root: ET.Element, object_id: str) -> ET.Element | None:
    """Return the element of *root* whose ID is *object_id*, or None."""
    return next((el for el in root.iter() if el.get("ID") == object_id), None)


def parent_map(root: ET.Element) -> dict[ET.Element, ET.Element]:
    return {child: parent for parent in root.iter() for child in parent}


def notebook_of(root: ET.Element, el: ET.Element) -> ET.Element | None:
    """Return the <one:Notebook> containing *el* (or *el* itself), or None."""
    parents = parent_map(root)
    cur: ET.Element | None = el
    while cur is not None and cur.tag != _NOTEBOOK:
        cur = parents.get(cur)
    return cur


# ---- UpdateHierarchy XML ------------------------------------------------------
# UpdateHierarchy takes a parent element listing children by ID: renames set
# attributes on a child; moves and reorders list the parent's children in the
# new order (a listed child from elsewhere is moved there).

def rename_xml(root: ET.Element, object_id: str, name: str) -> str:
    """XML renaming a section or section group (attribute name on its parent),
    or a notebook (its display name, nickname, under <one:Notebooks>)."""
    el = find_object(root, object_id)
    if el.tag == _NOTEBOOK:
        parent = ET.Element(f"{{{_ONE_NS}}}Notebooks")
        ET.SubElement(parent, _NOTEBOOK, {"ID": object_id, "nickname": name})
    else:
        source_parent = parent_map(root)[el]
        parent = ET.Element(source_parent.tag, {"ID": source_parent.get("ID", "")})
        ET.SubElement(parent, el.tag, {"ID": object_id, "name": name})
    return ET.tostring(parent, encoding="unicode", xml_declaration=False)


def order_xml(parent_el: ET.Element, child_tag: str, ordered_ids: list[str], levels: dict[str, int] | None = None) -> str:
    """XML listing *ordered_ids* as <child_tag> children of *parent_el*, in order.

    With *levels*, each child gets its pageLevel (pages only).
    """
    parent = ET.Element(parent_el.tag, {"ID": parent_el.get("ID", "")})
    for object_id in ordered_ids:
        child = ET.SubElement(parent, f"{{{_ONE_NS}}}{child_tag}", {"ID": object_id})
        if levels is not None:
            child.set("pageLevel", str(levels[object_id]))
    return ET.tostring(parent, encoding="unicode", xml_declaration=False)


MAX_PAGE_LEVEL = 3


def check_page_levels(levels: list[int]) -> None:
    """Raise bad_request unless *levels* (a section's pages in order) is valid:
    the first page at level 1, each level 1..3, at most one deeper than the page before."""
    previous = 0
    for position, level in enumerate(levels):
        if not 1 <= level <= MAX_PAGE_LEVEL or level > previous + 1:
            raise OneNoteError(
                f"Invalid page level {level} at position {position + 1}: levels run 1 to "
                f"{MAX_PAGE_LEVEL}, the first page is level 1, and a page is at most one level "
                "deeper than the page before it.",
                code="bad_request",
            )
        previous = level
