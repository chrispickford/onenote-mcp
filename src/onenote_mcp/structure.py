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
