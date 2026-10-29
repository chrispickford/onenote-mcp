"""Notebook structure: create, open and close notebooks; create section groups
and sections.

COM calls (OneNote.Application):
  OpenHierarchy(path, relative_to_id, create_type) -> object ID
      create_type: 0 = open only, 1 = notebook, 2 = section group (folder), 3 = section
  UpdateHierarchy(xml)            rename, colour, move, reorder
  CloseNotebook(notebook_id, force)

Every object must resolve to a notebook allowed by ONENOTE_ALLOWED_NOTEBOOKS.
Runs through com._run_com like every other COM call.
"""

import os
import re
import xml.etree.ElementTree as ET

import comtypes

from onenote_mcp import com, hierarchy
from onenote_mcp.errors import OneNoteError

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
_CFT_NONE, _CFT_NOTEBOOK, _CFT_FOLDER, _CFT_SECTION = 0, 1, 2, 3
_FORBIDDEN_NAME_CHARS = re.compile(r'[\\/:*?"<>|]')
_HEX_COLOR = re.compile(r"^#[0-9A-Fa-f]{6}$")


def _q(tag: str) -> str:
    return f"{{{_ONE_NS}}}{tag}"


def _is_url(location: str) -> bool:
    return location.lower().startswith(("http://", "https://"))


def _check_name(name: str, what: str) -> str:
    """Return *name* stripped, or raise bad_request for names OneNote cannot use."""
    stripped = name.strip()
    if not stripped or set(stripped) == {"."} or _FORBIDDEN_NAME_CHARS.search(stripped):
        raise OneNoteError(
            f"Invalid {what} name {name!r}: it must not be empty, only dots, or contain \\ / : * ? \" < > |.",
            code="bad_request",
        )
    return stripped


def _tree(app) -> ET.Element:
    return ET.fromstring(app.GetHierarchy("", 4))


def _container(app, parent_id: str) -> ET.Element:
    """Return the notebook or section group *parent_id*, checked against the allowlist."""
    com._ensure_allowed(app, parent_id, "Parent")
    parent = hierarchy.find_object(_tree(app), parent_id)
    if parent is None or parent.tag not in (_q("Notebook"), _q("SectionGroup")):
        raise OneNoteError(
            f"Parent {parent_id!r} must be a notebook or section group id from get_notebooks.",
            code="bad_request",
        )
    return parent


def _refuse_duplicate(parent: ET.Element, tag: str, name: str, what: str) -> None:
    if any(child.get("name", "").lower() == name.lower() for child in parent.findall(_q(tag))):
        raise OneNoteError(
            f"{parent.get('name')!r} already has a {what} named {name!r}; OneNote would open it "
            "instead of creating a new one.",
            code="bad_request",
        )


def _run(fn, *args, what: str):
    try:
        return com._run_com(fn, *args, timeout=com.WRITE_TIMEOUT)
    except comtypes.COMError as exc:
        raise OneNoteError(f"{what} failed: hresult={exc.hresult:#010x}") from exc


# ---- Sections and section groups ---------------------------------------------------

def _create_section_impl(parent_id: str, name: str, color: str | None) -> dict:
    app = com._app()
    parent = _container(app, parent_id)
    _refuse_duplicate(parent, "Section", name, "section")
    section_id = app.OpenHierarchy(f"{name}.one", parent_id, _CFT_SECTION)
    if color is not None:
        parent_el = ET.Element(parent.tag, {"ID": parent_id})
        ET.SubElement(parent_el, _q("Section"), {"ID": section_id, "color": color})
        app.UpdateHierarchy(ET.tostring(parent_el, encoding="unicode", xml_declaration=False))
    return {"id": section_id, "name": name}


def create_section(parent_id: str, name: str, color: str | None = None) -> dict:
    """Create a section in a notebook or section group; returns {"id", "name"}."""
    name = _check_name(name, "section")
    if name.lower().endswith(".one"):
        name = name[:-4].rstrip()
        name = _check_name(name, "section")
    if color is not None and not _HEX_COLOR.match(color):
        raise OneNoteError(f"color must be #RRGGBB, got {color!r}.", code="bad_request")
    return _run(_create_section_impl, parent_id, name, color, what="create_section")


def _create_section_group_impl(parent_id: str, name: str) -> dict:
    app = com._app()
    parent = _container(app, parent_id)
    _refuse_duplicate(parent, "SectionGroup", name, "section group")
    return {"id": app.OpenHierarchy(name, parent_id, _CFT_FOLDER), "name": name}


def create_section_group(parent_id: str, name: str) -> dict:
    """Create a section group in a notebook or section group; returns {"id", "name"}."""
    return _run(_create_section_group_impl, parent_id, _check_name(name, "section group"),
                what="create_section_group")


# ---- Notebooks ------------------------------------------------------------------------

def _notebook_info(app, notebook_id: str) -> dict:
    nb = hierarchy.find_object(ET.fromstring(app.GetHierarchy("", 2)), notebook_id)
    if nb is None:
        return {"id": notebook_id, "name": None, "path": None}
    return {"id": notebook_id, "name": nb.get("nickname") or nb.get("name"), "path": nb.get("path")}


def _allowed_to_add(name: str | None, path: str | None) -> bool:
    """Whether a notebook with *name* / *path* may be added under the allowlist."""
    allowed = com._allowed_notebooks()
    if allowed is None:
        return True
    probe = ET.Element(_q("Notebook"), {k: v for k, v in (("name", name), ("path", path)) if v})
    return com._notebook_allowed(probe, allowed)


def _create_notebook_impl(name: str, path: str) -> dict:
    app = com._app()
    open_paths = {
        com._normalized_location(nb.get("path", ""))
        for nb in ET.fromstring(app.GetHierarchy("", 2)).findall(_q("Notebook"))
    }
    if com._normalized_location(path) in open_paths:
        raise OneNoteError(f"A notebook is already open at {path!r}.", code="bad_request")
    return _notebook_info(app, app.OpenHierarchy(path, "", _CFT_NOTEBOOK))


def create_notebook(name: str, location: str) -> dict:
    """Create a notebook named *name* in *location* (an absolute local folder or a
    SharePoint/OneDrive URL); returns {"id", "name", "path"}."""
    name = _check_name(name, "notebook")
    location = location.strip()
    if _is_url(location):
        path = f"{location.rstrip('/')}/{name}"
    elif os.path.isabs(location):
        path = os.path.join(location.rstrip("\\/") + os.sep, name)
        if os.path.exists(path):
            raise OneNoteError(f"{path!r} already exists; open it with open_notebook instead.", code="bad_request")
    else:
        raise OneNoteError(
            f"location must be an absolute folder path or a URL, got {location!r}.", code="bad_request",
        )
    if not _allowed_to_add(name, path):
        raise OneNoteError(
            f"Notebook {name!r} is not allowed by ONENOTE_ALLOWED_NOTEBOOKS; add its name or path there first.",
            code="bad_request",
        )
    return _run(_create_notebook_impl, name, path, what="create_notebook")


def _open_notebook_impl(location: str) -> dict:
    app = com._app()
    already_open = {nb.get("ID") for nb in ET.fromstring(app.GetHierarchy("", 2)).findall(_q("Notebook"))}
    notebook_id = app.OpenHierarchy(location, "", _CFT_NONE)
    info = _notebook_info(app, notebook_id)
    if not _allowed_to_add(info["name"], info["path"]):
        if notebook_id not in already_open:
            app.CloseNotebook(notebook_id, False)
        raise OneNoteError(
            f"Notebook {info['name']!r} is not allowed by ONENOTE_ALLOWED_NOTEBOOKS.", code="bad_request",
        )
    return info


def open_notebook(location: str) -> dict:
    """Open an existing notebook from a local folder or URL; returns {"id", "name", "path"}."""
    return _run(_open_notebook_impl, location.strip(), what="open_notebook")


def _close_notebook_impl(notebook_id: str, force: bool) -> str:
    app = com._app()
    com._ensure_allowed(app, notebook_id, "Notebook")
    if hierarchy.find_object(_tree(app), notebook_id).tag != _q("Notebook"):
        raise OneNoteError(f"{notebook_id!r} is not a notebook id.", code="bad_request")
    app.CloseNotebook(notebook_id, force)
    return "ok"


def close_notebook(notebook_id: str, force: bool = False) -> str:
    """Close a notebook in OneNote (its files stay where they are)."""
    return _run(_close_notebook_impl, notebook_id, force, what="close_notebook")


# ---- Rename ---------------------------------------------------------------------------

_KIND_NAMES = {"Notebook": "notebook", "SectionGroup": "section group", "Section": "section"}


def _rename_impl(object_id: str, name: str) -> str:
    app = com._app()
    com._ensure_allowed(app, object_id, "Object")
    root = _tree(app)
    el = hierarchy.find_object(root, object_id)
    local = el.tag.split("}")[1]
    if local == "Page":
        raise OneNoteError("Use rename_page to rename a page.", code="bad_request")
    if local == "Section" and name.lower().endswith(".one"):
        name = _check_name(name[:-4], "section")
    if local != "Notebook":
        parent = hierarchy.parent_map(root)[el]
        siblings = [child for child in parent.findall(el.tag) if child is not el]
        if any(child.get("name", "").lower() == name.lower() for child in siblings):
            raise OneNoteError(
                f"{parent.get('name')!r} already has a {_KIND_NAMES[local]} named {name!r}.", code="bad_request",
            )
    app.UpdateHierarchy(hierarchy.rename_xml(root, object_id, name))
    return "ok"


def rename(object_id: str, name: str) -> str:
    """Rename a notebook (its display name), section group or section."""
    return _run(_rename_impl, object_id, _check_name(name, "new"), what="rename")


# ---- Move and reorder -------------------------------------------------------------

def _ordered_with(ids: list[str], moved: str, before_id: str | None, after_id: str | None) -> list[str]:
    """Return *ids* without *moved*, then with it inserted before/after an id (default: at the end)."""
    ordered = [object_id for object_id in ids if object_id != moved]
    anchor = before_id or after_id
    if anchor is None:
        return [*ordered, moved]
    if anchor not in ordered:
        raise OneNoteError(f"{anchor!r} is not a sibling at the target.", code="bad_request")
    index = ordered.index(anchor) + (0 if before_id is not None else 1)
    return [*ordered[:index], moved, *ordered[index:]]


def _move_container_child_impl(
    object_id: str, new_parent_id: str, before_id: str | None, after_id: str | None, tag: str,
) -> str:
    app = com._app()
    com._ensure_allowed(app, object_id, _KIND_NAMES[tag].capitalize())
    _container(app, new_parent_id)
    root = _tree(app)
    moved = hierarchy.find_object(root, object_id)
    if moved is None or moved.tag != _q(tag):
        raise OneNoteError(f"{object_id!r} is not a {_KIND_NAMES[tag]} id.", code="bad_request")
    if any(el.get("ID") == new_parent_id for el in moved.iter()):
        raise OneNoteError(f"Cannot move {_KIND_NAMES[tag]} {object_id!r} into itself.", code="bad_request")
    target = hierarchy.find_object(root, new_parent_id)
    siblings = [child for child in target.findall(_q(tag)) if child.get("ID") != object_id]
    if any(child.get("name", "").lower() == moved.get("name", "").lower() for child in siblings):
        raise OneNoteError(
            f"{target.get('name')!r} already has a {_KIND_NAMES[tag]} named {moved.get('name')!r}.",
            code="bad_request",
        )
    ordered = _ordered_with([child.get("ID") for child in target.findall(_q(tag))], object_id, before_id, after_id)
    app.UpdateHierarchy(hierarchy.order_xml(target, tag, ordered))
    return "ok"


def _check_one_anchor(before_id: str | None, after_id: str | None) -> None:
    if before_id is not None and after_id is not None:
        raise OneNoteError("Give at most one of before_id and after_id.", code="bad_request")


def move_section(section_id: str, new_parent_id: str, after_id: str | None = None, before_id: str | None = None) -> str:
    """Move a section into a notebook or section group, optionally next to a sibling section."""
    _check_one_anchor(before_id, after_id)
    return _run(_move_container_child_impl, section_id, new_parent_id, before_id, after_id, "Section",
                what="move_section")


def move_section_group(group_id: str, new_parent_id: str, after_id: str | None = None, before_id: str | None = None) -> str:
    """Move a section group into a notebook or another section group."""
    _check_one_anchor(before_id, after_id)
    return _run(_move_container_child_impl, group_id, new_parent_id, before_id, after_id, "SectionGroup",
                what="move_section_group")


def _section_pages(app, section_id: str) -> tuple[ET.Element, list[str], dict[str, int]]:
    section = ET.fromstring(app.GetHierarchy(section_id, 4))
    if section.tag != _q("Section"):
        raise OneNoteError(f"{section_id!r} is not a section id.", code="bad_request")
    pages = section.findall(_q("Page"))
    return section, [p.get("ID") for p in pages], {p.get("ID"): int(p.get("pageLevel", "1")) for p in pages}


def _section_of_page(app, page_id: str) -> str:
    root = _tree(app)
    page = hierarchy.find_object(root, page_id)
    if page is None or page.tag != _q("Page"):
        raise OneNoteError(f"{page_id!r} is not a page id.", code="bad_request")
    return hierarchy.parent_map(root)[page].get("ID")


def _write_order(app, section: ET.Element, ordered: list[str], levels: dict[str, int]) -> None:
    hierarchy.check_page_levels([levels[page_id] for page_id in ordered])
    app.UpdateHierarchy(hierarchy.order_xml(section, "Page", ordered, levels))


def _reorder_pages_impl(section_id: str, page_ids: list[str]) -> str:
    app = com._app()
    com._ensure_allowed(app, section_id, "Section")
    section, current, levels = _section_pages(app, section_id)
    if sorted(page_ids) != sorted(current):
        raise OneNoteError(
            "page_ids must list exactly the section's pages, each once (see list_pages).", code="bad_request",
        )
    _write_order(app, section, page_ids, levels)
    return "ok"


def reorder_pages(section_id: str, page_ids: list[str]) -> str:
    """Put a section's pages in the order of *page_ids*, keeping their levels."""
    return _run(_reorder_pages_impl, section_id, page_ids, what="reorder_pages")


def _set_page_level_impl(page_id: str, level: int) -> str:
    app = com._app()
    com._ensure_allowed(app, page_id, "Page")
    section, current, levels = _section_pages(app, _section_of_page(app, page_id))
    _write_order(app, section, current, {**levels, page_id: level})
    return "ok"


def set_page_level(page_id: str, level: int) -> str:
    """Make a page a top-level page (1) or a subpage (2 or 3) of the page above it."""
    return _run(_set_page_level_impl, page_id, level, what="set_page_level")


def _move_page_impl(page_id: str, target_section_id: str, after_page_id: str | None, level: int) -> dict:
    app = com._app()
    com._ensure_allowed(app, page_id, "Page")
    com._ensure_allowed(app, target_section_id, "Section")
    source_section_id = _section_of_page(app, page_id)
    section, current, levels = _section_pages(app, target_section_id)
    ordered = _ordered_with(current, page_id, None, after_page_id)
    _write_order(app, section, ordered, {**levels, page_id: level})
    if source_section_id == target_section_id:
        return {"page_id": page_id, "id_changed": False}
    # Page IDs embed their section: OneNote gives a moved page a new ID.
    _, after, _ = _section_pages(app, target_section_id)
    new_ids = [candidate for candidate in after if candidate not in current]
    new_id = new_ids[0] if len(new_ids) == 1 else page_id
    return {"page_id": new_id, "id_changed": new_id != page_id}


def move_page(page_id: str, target_section_id: str, after_page_id: str | None = None, level: int = 1) -> dict:
    """Move a page into a section (the same one to reorder), after *after_page_id* or at the end.

    Its subpages are not moved along.  Returns {"page_id", "id_changed"}:
    moving to another section changes the page's ID.
    """
    return _run(_move_page_impl, page_id, target_section_id, after_page_id, level, what="move_page")
