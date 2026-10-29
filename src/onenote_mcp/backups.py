"""Local backups of pages, taken before every write that changes or removes
existing content, so any such write can be undone with restore_page.

Layout under the backup root (config.backup_dir):

    <backup_id>/page.xml    the page as GetPageContent(page_id, 1) returned it
                            (binary data included, so images survive)
    <backup_id>/meta.json   page, notebook and section IDs and names, title,
                            tool, creation time, attachments
    <backup_id>/files/      copies of the page's attachments

A backup_id is "<UTC yyyymmddTHHMMSSffffff>-<sha1(page_id)[:8]>", so sorting
backup IDs sorts them by time.  Backups hold notebook content: they never
leave this machine.
"""

import hashlib
import json
import re
import shutil
import threading
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone
from pathlib import Path

from onenote_mcp.errors import OneNoteError

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
_BACKUP_ID = re.compile(r"^\d{8}T\d{12}-[0-9a-f]{8}$")
_UNSAFE_NAME = re.compile(r'[\\/:*?"<>|]')
_TIME_FORMAT = "%Y-%m-%dT%H:%M:%SZ"  # meta.json "created"

_LAST_STAMP = [datetime.min.replace(tzinfo=timezone.utc)]
_STAMP_LOCK = threading.Lock()


def _q(tag: str) -> str:
    return f"{{{_ONE_NS}}}{tag}"


def _next_stamp() -> datetime:
    """Return the current UTC time, strictly after the previous stamp."""
    with _STAMP_LOCK:
        now = datetime.now(timezone.utc)
        if now <= _LAST_STAMP[0]:
            now = _LAST_STAMP[0] + timedelta(microseconds=1)
        _LAST_STAMP[0] = now
        return now


def _location(app, page_id: str) -> dict:
    """Return the notebook and section (IDs and names) holding *page_id*."""
    root = ET.fromstring(app.GetHierarchy("", 4))
    parent_of = {child: parent for parent in root.iter() for child in parent}
    page_el = next((el for el in root.iter(_q("Page")) if el.get("ID") == page_id), None)
    section_el = notebook_el = None
    cur = parent_of.get(page_el) if page_el is not None else None
    while cur is not None:
        if cur.tag == _q("Section") and section_el is None:
            section_el = cur
        if cur.tag == _q("Notebook"):
            notebook_el = cur
            break
        cur = parent_of.get(cur)
    return {
        "notebook_id": notebook_el.get("ID") if notebook_el is not None else None,
        "notebook_name": (notebook_el.get("nickname") or notebook_el.get("name")) if notebook_el is not None else None,
        "section_id": section_el.get("ID") if section_el is not None else None,
        "section_name": section_el.get("name") if section_el is not None else None,
    }


def _title(page_root: ET.Element) -> str:
    title_el = page_root.find(_q("Title"))
    if title_el is None:
        return ""
    return "".join(t.text or "" for t in title_el.iter(_q("T")))


def _copy_attachments(page_root: ET.Element, files_dir: Path) -> tuple[dict[str, str], list[str]]:
    """Copy each attachment's cached file into *files_dir*.

    Returns ({objectID of the attachment's OE: file name}, [names that could
    not be copied]).
    """
    copied: dict[str, str] = {}
    missing: list[str] = []
    parent_of = {child: parent for parent in page_root.iter() for child in parent}
    for index, file_el in enumerate(page_root.iter(_q("InsertedFile"))):
        name = file_el.get("preferredName") or "attachment"
        cache = file_el.get("pathCache")
        owner = parent_of.get(file_el)
        object_id = (owner.get("objectID") if owner is not None and owner.tag == _q("OE") else None) or file_el.get("objectID")
        if not cache or not Path(cache).is_file() or not object_id:
            missing.append(name)
            continue
        safe = f"{index:03d}-{_UNSAFE_NAME.sub('_', name)}"
        files_dir.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(cache, files_dir / safe)
        copied[object_id] = safe
    return copied, missing


def _parse_time(value: str) -> datetime:
    """Parse an ISO 8601 time; a time without zone is UTC.  bad_request if invalid."""
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as exc:
        raise OneNoteError(f"Invalid time {value!r}; use ISO 8601, e.g. 2026-10-29T12:00:00Z.", code="bad_request") from exc
    return parsed if parsed.tzinfo else parsed.replace(tzinfo=timezone.utc)


def purge(root: Path, retention_days: float, now: datetime | None = None) -> None:
    """Delete backups created more than *retention_days* ago."""
    if not root.is_dir():
        return
    cutoff = (now or datetime.now(timezone.utc)) - timedelta(days=retention_days)
    for meta in list_backups(root):
        if _parse_time(meta["created"]) < cutoff:
            shutil.rmtree(root / meta["backup_id"], ignore_errors=True)


def save(app, page_id: str, tool: str, *, root: Path, retention_days: float) -> str:
    """Back up *page_id* (with binary data and attachments) and return the backup_id.

    Runs on the COM worker thread.  Purges expired backups first.
    """
    root = Path(root)
    purge(root, retention_days)
    page_xml = app.GetPageContent(page_id, 1)
    page_root = ET.fromstring(page_xml)
    stamp = _next_stamp()
    backup_id = f"{stamp.strftime('%Y%m%dT%H%M%S%f')}-{hashlib.sha1(page_id.encode()).hexdigest()[:8]}"
    folder = root / backup_id
    folder.mkdir(parents=True)
    (folder / "page.xml").write_text(page_xml, encoding="utf-8")
    attachments, missing = _copy_attachments(page_root, folder / "files")
    meta = {
        "page_id": page_id,
        "title": _title(page_root),
        **_location(app, page_id),
        "tool": tool,
        "created": stamp.strftime(_TIME_FORMAT),
        "attachments": attachments,
        "missing_attachments": missing,
    }
    (folder / "meta.json").write_text(json.dumps(meta, indent=2), encoding="utf-8")
    return backup_id


def list_backups(
    root: Path,
    *,
    page_id: str | None = None,
    section_id: str | None = None,
    since: str | None = None,
) -> list[dict]:
    """Return the metadata of matching backups (plus backup_id), newest first.

    *since* is an ISO 8601 UTC time such as "2026-10-29T12:00:00Z".
    """
    root = Path(root)
    if not root.is_dir():
        return []
    cutoff = _parse_time(since) if since is not None else None
    result: list[dict] = []
    for folder in sorted(root.iterdir(), reverse=True):
        meta_path = folder / "meta.json"
        if not _BACKUP_ID.match(folder.name) or not meta_path.is_file():
            continue
        meta = {**json.loads(meta_path.read_text(encoding="utf-8")), "backup_id": folder.name}
        if page_id is not None and meta.get("page_id") != page_id:
            continue
        if section_id is not None and meta.get("section_id") != section_id:
            continue
        if cutoff is not None and _parse_time(meta["created"]) < cutoff:
            continue
        result.append(meta)
    return result


def load(root: Path, backup_id: str) -> tuple[str, dict]:
    """Return (page XML, metadata) of *backup_id*; bad_request if unknown."""
    folder = Path(root) / backup_id
    if not _BACKUP_ID.match(backup_id or "") or not (folder / "meta.json").is_file():
        raise OneNoteError(f"Unknown backup {backup_id!r}; list_backups shows the available ones.", code="bad_request")
    meta = {**json.loads((folder / "meta.json").read_text(encoding="utf-8")), "backup_id": backup_id}
    return (folder / "page.xml").read_text(encoding="utf-8"), meta


def restorable_xml(
    root: Path,
    backup_id: str,
    page_xml: str,
    meta: dict,
    target_page_id: str,
    drop_tags: frozenset[str] = frozenset(),
) -> str:
    """Return the backed-up page as XML that writes it as new content on *target_page_id*.

    objectIDs are dropped, and attachments point at the files saved in the
    backup instead of OneNote's temporary cache.  Attachments whose file was
    not saved are left out, and so are direct page children whose tag is in
    *drop_tags* (canvas objects the restore keeps on the page instead).
    """
    page_root = ET.fromstring(page_xml)
    page_root.attrib.clear()
    page_root.set("ID", target_page_id)
    for child in [child for child in page_root if child.tag in drop_tags]:
        page_root.remove(child)
    files_dir = Path(root) / backup_id / "files"
    parent_of = {child: parent for parent in page_root.iter() for child in parent}
    for file_el in list(page_root.iter(_q("InsertedFile"))):
        owner = parent_of.get(file_el)
        object_id = (owner.get("objectID") if owner is not None and owner.tag == _q("OE") else None) or file_el.get("objectID")
        name = meta.get("attachments", {}).get(object_id)
        if name is None:
            parent_of[file_el].remove(file_el)
            continue
        file_el.attrib.pop("pathCache", None)
        file_el.set("pathSource", str(files_dir / name))
    for el in page_root.iter():
        el.attrib.pop("objectID", None)
    return ET.tostring(page_root, encoding="unicode", xml_declaration=False)


# ---- Manifests of deleted sections and section groups ------------------------------
#
# A manifest records what a section or section-group delete removed: the
# hierarchy XML of the deleted object and the backup IDs of its pages.  They
# live in <root>/manifests/<manifest_id>.json, apart from the page backups.

def save_manifest(root: Path, deleted_id: str, hierarchy_xml: str, page_backup_ids: list[str]) -> str:
    """Record a section or section-group delete; return the manifest_id."""
    stamp = _next_stamp()
    manifest_id = f"{stamp.strftime('%Y%m%dT%H%M%S%f')}-{hashlib.sha1(deleted_id.encode()).hexdigest()[:8]}"
    folder = Path(root) / "manifests"
    folder.mkdir(parents=True, exist_ok=True)
    (folder / f"{manifest_id}.json").write_text(json.dumps({
        "deleted_id": deleted_id,
        "created": stamp.strftime(_TIME_FORMAT),
        "hierarchy_xml": hierarchy_xml,
        "page_backups": page_backup_ids,
    }, indent=2), encoding="utf-8")
    return manifest_id


def load_manifest(root: Path, manifest_id: str) -> dict:
    """Return a manifest saved by save_manifest; bad_request if unknown."""
    path = Path(root) / "manifests" / f"{manifest_id}.json"
    if not _BACKUP_ID.match(manifest_id or "") or not path.is_file():
        raise OneNoteError(f"Unknown manifest {manifest_id!r}.", code="bad_request")
    return json.loads(path.read_text(encoding="utf-8"))
