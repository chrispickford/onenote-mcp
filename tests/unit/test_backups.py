"""Unit tests for backups.py: saving, listing, loading and purging local page
backups.  A fake app stands in for OneNote.Application; backups go to tmp_path.
"""
import json
import os
import time
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone

import pytest

from onenote_mcp import backups

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
_XMLNS = f'xmlns:one="{_ONE_NS}"'


def _hierarchy() -> str:
    return (
        f'<one:Notebooks {_XMLNS}>'
        '<one:Notebook ID="{NB}" name="Spike" nickname="Spike">'
        '<one:SectionGroup ID="{G}" name="Group">'
        '<one:Section ID="{S}" name="Sec"><one:Page ID="{P}" name="Title A"/></one:Section>'
        '</one:SectionGroup></one:Notebook></one:Notebooks>'
    )


class _FakeApp:
    def __init__(self, page_xml: str) -> None:
        self.page_xml = page_xml

    def GetPageContent(self, page_id: str, page_info: int) -> str:
        assert page_info == 1, "backups read with binary data"
        return self.page_xml

    def GetHierarchy(self, start: str, scope: int) -> str:
        return _hierarchy()


def _page(attachment_cache: str | None = None) -> str:
    file_el = (
        f'<one:OE objectID="{{F}}"><one:InsertedFile pathCache="{attachment_cache}" '
        'pathSource="D:\\gone\\notes.txt" preferredName="notes.txt"/></one:OE>'
        if attachment_cache else ""
    )
    return (
        f'<one:Page {_XMLNS} ID="{{P}}"><one:Title><one:OE><one:T>Title A</one:T></one:OE></one:Title>'
        f'<one:Outline><one:OEChildren><one:OE><one:T>body</one:T></one:OE>{file_el}'
        '</one:OEChildren></one:Outline></one:Page>'
    )


def test_save_writes_page_and_meta(tmp_path):
    backup_id = backups.save(_FakeApp(_page()), "{P}", "replace_page", root=tmp_path, retention_days=30)

    folder = tmp_path / backup_id
    assert (folder / "page.xml").read_text(encoding="utf-8") == _page()
    meta = json.loads((folder / "meta.json").read_text(encoding="utf-8"))
    assert meta["page_id"] == "{P}"
    assert meta["title"] == "Title A"
    assert meta["tool"] == "replace_page"
    assert (meta["notebook_id"], meta["notebook_name"]) == ("{NB}", "Spike")
    assert (meta["section_id"], meta["section_name"]) == ("{S}", "Sec")
    assert meta["created"].endswith("Z")
    assert meta["attachments"] == {} and meta["missing_attachments"] == []


def test_save_copies_attachments(tmp_path):
    cache = tmp_path / "cache.bin"
    cache.write_bytes(b"attached bytes")
    backup_id = backups.save(_FakeApp(_page(str(cache))), "{P}", "replace_page", root=tmp_path / "bk", retention_days=30)

    meta = json.loads((tmp_path / "bk" / backup_id / "meta.json").read_text(encoding="utf-8"))
    ((object_id, name),) = meta["attachments"].items()
    assert object_id == "{F}"
    assert (tmp_path / "bk" / backup_id / "files" / name).read_bytes() == b"attached bytes"
    assert name.endswith("notes.txt")


def test_missing_attachment_is_recorded(tmp_path):
    backup_id = backups.save(_FakeApp(_page(str(tmp_path / "nope.bin"))), "{P}", "t", root=tmp_path, retention_days=30)
    meta = json.loads((tmp_path / backup_id / "meta.json").read_text(encoding="utf-8"))
    assert meta["missing_attachments"] == ["notes.txt"]


def test_backup_id_sortable_and_unique(tmp_path):
    app = _FakeApp(_page())
    first = backups.save(app, "{P}", "t", root=tmp_path, retention_days=30)
    second = backups.save(app, "{P}", "t", root=tmp_path, retention_days=30)
    assert first != second
    assert sorted([second, first]) == [first, second]


def test_retention_purges_old(tmp_path):
    app = _FakeApp(_page())
    old = backups.save(app, "{P}", "t", root=tmp_path, retention_days=30)
    meta_path = tmp_path / old / "meta.json"
    meta = json.loads(meta_path.read_text(encoding="utf-8"))
    meta["created"] = (datetime.now(timezone.utc) - timedelta(days=31)).strftime("%Y-%m-%dT%H:%M:%SZ")
    meta_path.write_text(json.dumps(meta), encoding="utf-8")

    new = backups.save(app, "{P}", "t", root=tmp_path, retention_days=30)
    assert not (tmp_path / old).exists()
    assert (tmp_path / new).exists()


def test_list_filters_and_sorts_newest_first(tmp_path):
    app = _FakeApp(_page())
    a = backups.save(app, "{P}", "t", root=tmp_path, retention_days=30)
    b = backups.save(app, "{P}", "t", root=tmp_path, retention_days=30)
    assert [m["backup_id"] for m in backups.list_backups(tmp_path, page_id="{P}")] == [b, a]
    assert backups.list_backups(tmp_path, page_id="{OTHER}") == []
    assert len(backups.list_backups(tmp_path, section_id="{S}")) == 2


def test_list_since_filter(tmp_path):
    backups.save(_FakeApp(_page()), "{P}", "t", root=tmp_path, retention_days=30)
    assert backups.list_backups(tmp_path, since="2999-01-01T00:00:00Z") == []
    assert len(backups.list_backups(tmp_path, since="2000-01-01T00:00:00Z")) == 1


def test_list_of_missing_root_is_empty(tmp_path):
    assert backups.list_backups(tmp_path / "none") == []


def test_load_round_trip(tmp_path):
    backup_id = backups.save(_FakeApp(_page()), "{P}", "t", root=tmp_path, retention_days=30)
    xml, meta = backups.load(tmp_path, backup_id)
    assert xml == _page()
    assert meta["backup_id"] == backup_id


@pytest.mark.parametrize("bad", ["../x", "nope", "..\\x", ""])
def test_load_unknown_or_unsafe_id_is_bad_request(tmp_path, bad):
    with pytest.raises(backups.OneNoteError) as exc:
        backups.load(tmp_path, bad)
    assert exc.value.code == "bad_request"


def test_restorable_xml_points_attachments_at_backup_files(tmp_path):
    cache = tmp_path / "cache.bin"
    cache.write_bytes(b"x")
    root = tmp_path / "bk"
    backup_id = backups.save(_FakeApp(_page(str(cache))), "{P}", "t", root=root, retention_days=30)
    xml, meta = backups.load(root, backup_id)

    page = ET.fromstring(backups.restorable_xml(root, backup_id, xml, meta, "{TARGET}"))
    assert page.get("ID") == "{TARGET}"
    assert all(el.get("objectID") is None for el in page.iter())
    file_el = page.find(f".//{{{_ONE_NS}}}InsertedFile")
    assert file_el.get("pathCache") is None
    assert file_el.get("pathSource") == str(root / backup_id / "files" / meta["attachments"]["{F}"])
