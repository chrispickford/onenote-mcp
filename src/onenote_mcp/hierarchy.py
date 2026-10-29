"""Hierarchy XML (notebooks, section groups, sections, pages): parsing for the
navigation tools and lookups shared by the structure and allowlist code.

Pure functions on XML; no COM.
"""

import xml.etree.ElementTree as ET

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
