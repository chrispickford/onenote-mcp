"""Unit tests for write preparation in com.py: partial image sizes, the
<one:Page> wrapper accepted by append, and the delete pass of replace.

A fake app stands in for OneNote.Application and records the XML it receives.
"""
import base64
import xml.etree.ElementTree as ET

import pytest

from tests.png import png_base64 as _png_base64

try:
    import comtypes
    from onenote_mcp import backups, com, pagestate
    from onenote_mcp.builders import UNSUPPORTED_PAGE_OBJECT_TAGS
    from onenote_mcp.com import OneNoteError, _complete_image_sizes
    _COM_AVAILABLE = True
except Exception:
    _COM_AVAILABLE = False

pytestmark = pytest.mark.skipif(
    not _COM_AVAILABLE,
    reason="com.py module-level imports (comtypes/winreg) unavailable in this environment",
)

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
_XMLNS = f'xmlns:one="{_ONE_NS}"'


def _image_page(size_attrs: str, data: str) -> str:
    return (
        f'<one:Page {_XMLNS}><one:Outline><one:OEChildren><one:OE><one:Image>'
        f'<one:Size {size_attrs} isSetByUser="true"/><one:Data>{data}</one:Data>'
        f'</one:Image></one:OE></one:OEChildren></one:Outline></one:Page>'
    )


def _size(xml: str) -> dict[str, str]:
    return ET.fromstring(xml).find(f".//{{{_ONE_NS}}}Size").attrib


def test_width_only_gets_height_from_aspect_ratio():
    xml = _complete_image_sizes(_image_page('width="240"', _png_base64(200, 100)))
    assert _size(xml) == {"width": "240", "isSetByUser": "true", "height": "120.0"}


def test_height_only_gets_width_from_aspect_ratio():
    xml = _complete_image_sizes(_image_page('height="50"', _png_base64(200, 100)))
    assert _size(xml)["width"] == "100.0"


def test_complete_size_is_unchanged():
    xml = _complete_image_sizes(_image_page('width="240" height="10"', _png_base64(200, 100)))
    assert _size(xml) == {"width": "240", "height": "10", "isSetByUser": "true"}


def test_unreadable_image_with_partial_size_is_a_bad_request():
    with pytest.raises(OneNoteError) as exc:
        _complete_image_sizes(_image_page('width="240"', base64.b64encode(b"no image").decode()))
    assert exc.value.code == "bad_request"
    assert "Pass both width and height" in str(exc.value)


class _FakeApp:
    """Stand-in for OneNote.Application recording UpdatePageContent calls."""

    def __init__(self) -> None:
        self.updates: list[str] = []

    def GetPageContent(self, page_id: str, page_info: int) -> str:
        return f'<one:Page {_XMLNS} ID="{page_id}"/>'

    def GetHierarchy(self, start: str, scope: int) -> str:
        return (f'<one:Notebooks {_XMLNS}><one:Notebook ID="{{NB}}" name="N">'
                '<one:Section ID="{SEC}" name="S"><one:Page ID="{PAGE}" name="Old"/></one:Section>'
                '</one:Notebook></one:Notebooks>')

    def UpdatePageContent(self, xml: str, date_expected_last_modified: float) -> None:
        self.updates.append(xml)


@pytest.fixture
def app(monkeypatch):
    monkeypatch.delenv("ONENOTE_ALLOWED_NOTEBOOKS", raising=False)
    fake = _FakeApp()
    monkeypatch.setattr(com, "_app", lambda: fake)
    return fake


def test_append_wraps_a_single_element_in_the_page(app):
    com._append_page_impl("{PAGE}", f'<one:Outline {_XMLNS} objectID="x"><one:OEChildren/></one:Outline>')

    (xml,) = app.updates
    page = ET.fromstring(xml)
    assert page.get("ID") == "{PAGE}"
    assert [child.tag for child in page] == [f"{{{_ONE_NS}}}Outline"]
    assert page[0].get("objectID") is None


def test_append_sends_the_children_of_a_page_wrapper(app):
    com._append_page_impl("{PAGE}", (
        f'<one:Page {_XMLNS}>'
        '<one:QuickStyleDef index="0" name="h2" font="Calibri" fontSize="14.0"/>'
        '<one:Outline><one:OEChildren><one:OE quickStyleIndex="0"><one:T>x</one:T></one:OE></one:OEChildren></one:Outline>'
        '</one:Page>'
    ))

    (xml,) = app.updates
    page = ET.fromstring(xml)
    assert page.get("ID") == "{PAGE}"
    assert [child.tag for child in page] == [f"{{{_ONE_NS}}}QuickStyleDef", f"{{{_ONE_NS}}}Outline"]


def test_append_completes_a_width_only_image(app):
    com._append_page_impl("{PAGE}", _image_page('width="240"', _png_base64(200, 100)))

    (xml,) = app.updates
    assert _size(xml)["height"] == "120.0"


# ---------------------------------------------------------------------------
# replace: delete pass, kept page objects and partial writes
# ---------------------------------------------------------------------------

_CURRENT_PAGE = (
    f'<one:Page {_XMLNS} ID="{{PAGE}}">'
    '<one:QuickStyleDef index="0" name="p"/>'
    '<one:Title objectID="title"><one:OE><one:T>Old</one:T></one:OE></one:Title>'
    '<one:Outline objectID="outline-1"><one:OEChildren><one:OE><one:T>a</one:T></one:OE></one:OEChildren></one:Outline>'
    '<one:InkDrawing objectID="ink-1"><one:CallbackID callbackID="{1}"/></one:InkDrawing>'
    '<one:Outline objectID="outline-2"><one:OEChildren><one:OE><one:T>b</one:T></one:OE></one:OEChildren></one:Outline>'
    '<one:InsertedFile objectID="file-1" preferredName="a.pdf"/>'
    '</one:Page>'
)
_NEW_PAGE = (
    f'<one:Page {_XMLNS}><one:Title><one:OE><one:T>New</one:T></one:OE></one:Title>'
    '<one:Outline><one:OEChildren><one:OE><one:T>c</one:T></one:OE></one:OEChildren></one:Outline></one:Page>'
)
_HRESULT = -2147213298  # 0x8004200E, hrPageObjectDoesNotExist


class _FakeReplaceApp:
    """Stand-in for OneNote.Application logging the replace calls in order.

    DeletePageContent raises a COMError for the ids in *undeletable*.
    """

    def __init__(self, undeletable: frozenset[str] = frozenset()) -> None:
        self.calls: list[tuple[str, str]] = []
        self._undeletable = undeletable

    def GetPageContent(self, page_id: str, page_info: int) -> str:
        return _CURRENT_PAGE

    def GetHierarchy(self, start: str, scope: int) -> str:
        return (f'<one:Notebooks {_XMLNS}><one:Notebook ID="{{NB}}" name="N">'
                '<one:Section ID="{SEC}" name="S"><one:Page ID="{PAGE}" name="Old"/></one:Section>'
                '</one:Notebook></one:Notebooks>')

    def DeletePageContent(self, page_id: str, object_id: str, date_expected_last_modified: float) -> None:
        self.calls.append(("delete", object_id))
        if object_id in self._undeletable:
            raise comtypes.COMError(_HRESULT, "locked", None)

    def UpdatePageContent(self, xml: str, date_expected_last_modified: float) -> None:
        self.calls.append(("update", xml))


def _use(monkeypatch, fake) -> None:
    monkeypatch.delenv("ONENOTE_ALLOWED_NOTEBOOKS", raising=False)
    monkeypatch.setattr(com, "_app", lambda: fake)


def _deleted(fake: _FakeReplaceApp) -> list[str]:
    return [arg for call, arg in fake.calls if call == "delete"]


def test_replace_writes_then_deletes_every_old_object_but_the_title(monkeypatch):
    fake = _FakeReplaceApp()
    _use(monkeypatch, fake)
    com._replace_page_impl("{PAGE}", _NEW_PAGE, frozenset())

    assert _deleted(fake) == ["outline-1", "ink-1", "outline-2", "file-1"]
    assert [call for call, _ in fake.calls][0] == "update"
    assert ET.fromstring(fake.calls[0][1]).find(f"{{{_ONE_NS}}}Title").get("objectID") == "title"


def test_replace_keeps_objects_whose_tag_is_kept(monkeypatch):
    fake = _FakeReplaceApp()
    _use(monkeypatch, fake)
    com._replace_page_impl("{PAGE}", _NEW_PAGE, UNSUPPORTED_PAGE_OBJECT_TAGS)

    assert _deleted(fake) == ["outline-1", "outline-2"]


def test_replace_writes_then_attempts_every_delete_before_reporting_a_partial_write(monkeypatch):
    fake = _FakeReplaceApp(undeletable=frozenset({"outline-1", "file-1"}))
    _use(monkeypatch, fake)
    with pytest.raises(OneNoteError) as exc:
        com._replace_page_impl("{PAGE}", _NEW_PAGE, frozenset())

    assert _deleted(fake) == ["outline-1", "ink-1", "outline-2", "file-1"]
    assert [call for call, _ in fake.calls][0] == "update"
    assert exc.value.code == "partial_write"
    message = str(exc.value)
    assert message.startswith("partial_write: The new content was written, but 2 old page object(s)")
    assert "Outline outline-1 (hresult=-0x7ffbdff2)" in message
    assert "InsertedFile file-1 (hresult=-0x7ffbdff2)" in message
    assert "ink-1" not in message


def test_public_replace_page_passes_the_partial_write_through(monkeypatch):
    fake = _FakeReplaceApp(undeletable=frozenset({"outline-2"}))
    _use(monkeypatch, fake)
    with pytest.raises(OneNoteError) as exc:
        com.replace_page("{PAGE}", _NEW_PAGE, expected_revision=pagestate.revision(_CURRENT_PAGE))

    assert exc.value.code == "partial_write"
    assert "Outline outline-2" in str(exc.value)


# ---------------------------------------------------------------------------
# revisions: writes refuse pages changed since they were read
# ---------------------------------------------------------------------------

@pytest.fixture
def clean_pagestate():
    pagestate.forget_all()
    yield
    pagestate.forget_all()


def test_replace_with_stale_revision_is_conflict_and_writes_nothing(monkeypatch, clean_pagestate):
    fake = _FakeReplaceApp()
    _use(monkeypatch, fake)
    with pytest.raises(OneNoteError) as exc:
        com._replace_page_impl("{PAGE}", _NEW_PAGE, frozenset(), "0000000000000000")
    assert exc.value.code == "conflict"
    assert fake.calls == []


def test_replace_with_matching_revision_writes(monkeypatch, clean_pagestate):
    fake = _FakeReplaceApp()
    _use(monkeypatch, fake)
    com._replace_page_impl("{PAGE}", _NEW_PAGE, frozenset(), pagestate.revision(_CURRENT_PAGE))
    assert ("update" in [call for call, _ in fake.calls])


def test_replace_remembers_the_revision_after_writing(monkeypatch, clean_pagestate):
    fake = _FakeReplaceApp()
    _use(monkeypatch, fake)
    com._replace_page_impl("{PAGE}", _NEW_PAGE, frozenset(), pagestate.revision(_CURRENT_PAGE))
    assert pagestate.remembered("{PAGE}") == pagestate.revision(_CURRENT_PAGE)


def test_public_replace_without_a_read_is_bad_request_before_com(monkeypatch, clean_pagestate):
    fake = _FakeReplaceApp()
    _use(monkeypatch, fake)
    with pytest.raises(OneNoteError) as exc:
        com.replace_page("{PAGE}", _NEW_PAGE)
    assert exc.value.code == "bad_request"
    assert fake.calls == []


def test_public_replace_uses_the_remembered_revision(monkeypatch, clean_pagestate):
    fake = _FakeReplaceApp()
    _use(monkeypatch, fake)
    pagestate.remember("{PAGE}", "0000000000000000")
    with pytest.raises(OneNoteError) as exc:
        com.replace_page("{PAGE}", _NEW_PAGE)
    assert exc.value.code == "conflict"


def test_get_page_remembers_the_revision(monkeypatch, clean_pagestate):
    fake = _FakeReplaceApp()
    _use(monkeypatch, fake)
    com.get_page("{PAGE}")
    assert pagestate.remembered("{PAGE}") == pagestate.revision(_CURRENT_PAGE)


def test_append_needs_no_read(monkeypatch, clean_pagestate):
    fake = _FakeReplaceApp()
    _use(monkeypatch, fake)
    com.append_page("{PAGE}", f'<one:Outline {_XMLNS}><one:OEChildren><one:OE><one:T>x</one:T></one:OE></one:OEChildren></one:Outline>')
    assert [call for call, _ in fake.calls] == ["update"]
    assert pagestate.remembered("{PAGE}") is None


def test_append_with_stale_given_revision_is_conflict(monkeypatch, clean_pagestate):
    fake = _FakeReplaceApp()
    _use(monkeypatch, fake)
    with pytest.raises(OneNoteError) as exc:
        com.append_page("{PAGE}", f'<one:Outline {_XMLNS}><one:OEChildren/></one:Outline>', "0000000000000000")
    assert exc.value.code == "conflict"


# ---------------------------------------------------------------------------
# backups and restore
# ---------------------------------------------------------------------------

def test_replace_backs_up_before_writing(monkeypatch, clean_pagestate, _backup_dir):
    fake = _FakeReplaceApp()
    _use(monkeypatch, fake)
    original_update = fake.UpdatePageContent

    def update(xml, date):
        assert len(backups.list_backups(_backup_dir, page_id="{PAGE}")) == 1, "backup must exist before the write"
        original_update(xml, date)

    fake.UpdatePageContent = update
    com._replace_page_impl("{PAGE}", _NEW_PAGE, frozenset(), pagestate.revision(_CURRENT_PAGE))
    (meta,) = backups.list_backups(_backup_dir, page_id="{PAGE}")
    assert meta["tool"] == "replace_page"


def test_conflict_takes_no_backup(monkeypatch, clean_pagestate, _backup_dir):
    fake = _FakeReplaceApp()
    _use(monkeypatch, fake)
    with pytest.raises(OneNoteError):
        com._replace_page_impl("{PAGE}", _NEW_PAGE, frozenset(), "0000000000000000")
    assert backups.list_backups(_backup_dir) == []


def test_restore_over_existing_page_backs_up_then_writes_backup_content(monkeypatch, clean_pagestate, _backup_dir):
    fake = _FakeReplaceApp()
    _use(monkeypatch, fake)
    backup_id = backups.save(fake, "{PAGE}", "replace_page", root=_backup_dir, retention_days=30)

    result = com.restore_page(backup_id)

    assert result == {"page_id": "{PAGE}", "restored_over": True, "missing_attachments": ["a.pdf"]}
    written = ET.fromstring([arg for call, arg in fake.calls if call == "update"][-1])
    assert [t.text for t in written.iter(f"{{{_ONE_NS}}}T")] == ["Old", "a", "b"]
    assert [m["tool"] for m in backups.list_backups(_backup_dir)] == ["restore_page", "replace_page"]


def test_restore_requires_target_when_page_gone(monkeypatch, clean_pagestate, _backup_dir):
    fake = _FakeReplaceApp()
    _use(monkeypatch, fake)
    backup_id = backups.save(fake, "{PAGE}", "t", root=_backup_dir, retention_days=30)
    fake.GetHierarchy = lambda start, scope: f'<one:Notebooks {_XMLNS}/>'
    with pytest.raises(OneNoteError) as exc:
        com.restore_page(backup_id)
    assert exc.value.code == "bad_request"
    assert "target_section_id" in str(exc.value)


def test_restore_into_new_page_when_page_gone(monkeypatch, clean_pagestate, _backup_dir):
    fake = _FakeReplaceApp()
    _use(monkeypatch, fake)
    backup_id = backups.save(fake, "{PAGE}", "t", root=_backup_dir, retention_days=30)
    fake.GetHierarchy = lambda start, scope: (
        f'<one:Notebooks {_XMLNS}><one:Notebook ID="{{NB}}" name="N"><one:Section ID="{{SEC2}}" name="S2"/>'
        '</one:Notebook></one:Notebooks>')
    fake.CreateNewPage = lambda section_id: "{NEWPAGE}"

    result = com.restore_page(backup_id, target_section_id="{SEC2}")

    assert result["page_id"] == "{NEWPAGE}" and result["restored_over"] is False
    written = ET.fromstring([arg for call, arg in fake.calls if call == "update"][-1])
    assert written.get("ID") == "{NEWPAGE}"


# ---------------------------------------------------------------------------
# write before delete, and content replace cannot re-create
# ---------------------------------------------------------------------------

def test_replace_writes_before_deleting(monkeypatch, clean_pagestate):
    fake = _FakeReplaceApp()
    _use(monkeypatch, fake)
    com._replace_page_impl("{PAGE}", _NEW_PAGE, frozenset(), pagestate.revision(_CURRENT_PAGE))
    calls = [call for call, _ in fake.calls]
    assert calls[0] == "update"
    assert set(calls[1:]) == {"delete"}


_PROTECTED_PAGE = (
    f'<one:Page {_XMLNS} ID="{{PAGE}}">'
    '<one:Title objectID="title"><one:OE><one:T>Old</one:T></one:OE></one:Title>'
    '<one:Outline objectID="outline-1"><one:OEChildren>'
    '<one:OE objectID="p-1"><one:T>a</one:T></one:OE>'
    '<one:OE objectID="att-1"><one:InsertedFile preferredName="a.pdf"/></one:OE>'
    '<one:OE objectID="p-2"><one:T>b</one:T><one:OEChildren>'
    '<one:OE objectID="ink-2"><one:InkParagraph/></one:OE></one:OEChildren></one:OE>'
    '</one:OEChildren></one:Outline></one:Page>'
)


class _FakeProtectedApp(_FakeReplaceApp):
    def GetPageContent(self, page_id: str, page_info: int) -> str:
        return _PROTECTED_PAGE


def test_replace_refuses_page_with_attachment(monkeypatch, clean_pagestate, _backup_dir):
    fake = _FakeProtectedApp()
    _use(monkeypatch, fake)
    with pytest.raises(OneNoteError) as exc:
        com._replace_page_impl("{PAGE}", _NEW_PAGE, frozenset(), pagestate.revision(_PROTECTED_PAGE))
    assert exc.value.code == "bad_request"
    message = str(exc.value)
    assert "file att-1" in message and "ink ink-2" in message
    assert "remove_ids" in message
    assert fake.calls == []
    assert backups.list_backups(_backup_dir) == []


def test_remove_ids_must_cover_every_protected_item(monkeypatch, clean_pagestate):
    fake = _FakeProtectedApp()
    _use(monkeypatch, fake)
    with pytest.raises(OneNoteError) as exc:
        com._replace_page_impl("{PAGE}", _NEW_PAGE, frozenset(), pagestate.revision(_PROTECTED_PAGE),
                               remove_ids=frozenset({"att-1"}))
    assert "ink ink-2" in str(exc.value) and "att-1" not in str(exc.value)


def test_remove_ids_allows_removal(monkeypatch, clean_pagestate):
    fake = _FakeProtectedApp()
    _use(monkeypatch, fake)
    com._replace_page_impl("{PAGE}", _NEW_PAGE, frozenset(), pagestate.revision(_PROTECTED_PAGE),
                           remove_ids=frozenset({"att-1", "ink-2"}))
    assert _deleted(fake) == ["outline-1"]


def test_protected_items_lists_nested_unsupported_oes():
    from onenote_mcp.builders import protected_items
    assert protected_items(ET.fromstring(_PROTECTED_PAGE)) == [("att-1", "file"), ("ink-2", "ink")]


def test_restore_over_page_with_protected_items_replaces_them(monkeypatch, clean_pagestate, _backup_dir):
    fake = _FakeProtectedApp()
    _use(monkeypatch, fake)
    backup_id = backups.save(fake, "{PAGE}", "replace_page", root=_backup_dir, retention_days=30)
    com.restore_page(backup_id)
    assert _deleted(fake) == ["outline-1"]


# ---------------------------------------------------------------------------
# writes register as in-flight for their page
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("call", [
    lambda: com.replace_page("{PAGE}", _NEW_PAGE, expected_revision="r"),
    lambda: com.append_page("{PAGE}", f'<one:Outline {_XMLNS}><one:OEChildren/></one:Outline>'),
])
def test_page_writes_pass_their_page_as_write_page_id(monkeypatch, call):
    seen = []
    monkeypatch.setattr(com, "_run_com", lambda fn, *args, **kwargs: seen.append(kwargs.get("write_page_id")))
    call()
    assert seen == ["{PAGE}"]


# ---------------------------------------------------------------------------
# file paths in writes must lie inside ONENOTE_FILE_ROOTS
# ---------------------------------------------------------------------------

def _with_file(tag_xml: str) -> str:
    return f'<one:Outline {_XMLNS}><one:OEChildren><one:OE>{tag_xml}</one:OE></one:OEChildren></one:Outline>'


@pytest.fixture
def allowed_file(tmp_path, monkeypatch):
    folder = tmp_path / "roots"
    folder.mkdir()
    file = folder / "pic.png"
    file.write_bytes(b"x")
    monkeypatch.setenv("ONENOTE_FILE_ROOTS", str(folder))
    return file


@pytest.mark.parametrize("element", [
    '<one:Image><one:File path="{path}"/></one:Image>',
    '<one:InsertedFile pathSource="{path}" preferredName="x"/>',
])
def test_append_with_path_outside_roots_is_refused_before_com(monkeypatch, allowed_file, tmp_path, element):
    outside = tmp_path / "outside.png"
    outside.write_bytes(b"x")
    fake = _FakeReplaceApp()
    _use(monkeypatch, fake)
    with pytest.raises(OneNoteError) as exc:
        com.append_page("{PAGE}", _with_file(element.format(path=outside)))
    assert exc.value.code == "bad_request" and "outside ONENOTE_FILE_ROOTS" in str(exc.value)
    assert fake.calls == []


def test_append_with_path_inside_roots_is_written(monkeypatch, allowed_file):
    fake = _FakeReplaceApp()
    _use(monkeypatch, fake)
    com.append_page("{PAGE}", _with_file(f'<one:Image><one:File path="{allowed_file}"/></one:Image>'))
    assert [call for call, _ in fake.calls] == ["update"]


def test_replace_with_path_and_no_roots_is_refused(monkeypatch, allowed_file, clean_pagestate):
    monkeypatch.delenv("ONENOTE_FILE_ROOTS")
    fake = _FakeReplaceApp()
    _use(monkeypatch, fake)
    page = (
        f'<one:Page {_XMLNS}><one:Title><one:OE><one:T>t</one:T></one:OE></one:Title>'
        f'<one:Outline><one:OEChildren><one:OE><one:InsertedFile pathSource="{allowed_file}" preferredName="x"/>'
        '</one:OE></one:OEChildren></one:Outline></one:Page>'
    )
    with pytest.raises(OneNoteError) as exc:
        com.replace_page("{PAGE}", page, expected_revision="r")
    assert "Inserting files is disabled" in str(exc.value)
    assert fake.calls == []
