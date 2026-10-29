"""Unit tests for structure.py: creating and organising notebooks, section
groups, sections and pages.  A fake app stands in for OneNote.Application and
records the hierarchy calls it receives.
"""
import xml.etree.ElementTree as ET

import pytest

try:
    from onenote_mcp import com, structure
    from onenote_mcp.errors import OneNoteError
    _COM_AVAILABLE = True
except Exception:
    _COM_AVAILABLE = False

pytestmark = pytest.mark.skipif(not _COM_AVAILABLE, reason="comtypes/winreg unavailable")

NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
X = f'xmlns:one="{NS}"'

_TREE = (
    f'<one:Notebooks {X}>'
    '<one:Notebook ID="{NB}" name="Spike" nickname="Spike" path="D:\\nb\\Spike">'
    '<one:Section ID="{S1}" name="Top"><one:Page ID="{P1}" name="One" pageLevel="1"/>'
    '<one:Page ID="{P2}" name="Two" pageLevel="1"/><one:Page ID="{P3}" name="Three" pageLevel="2"/></one:Section>'
    '<one:SectionGroup ID="{G}" name="Group"><one:Section ID="{S2}" name="Inside"/></one:SectionGroup>'
    '</one:Notebook>'
    '<one:Notebook ID="{OTHER}" name="Private" path="D:\\nb\\Private"><one:Section ID="{SX}" name="Secret"/></one:Notebook>'
    '</one:Notebooks>'
)


class _FakeApp:
    def __init__(self) -> None:
        self.calls: list[tuple] = []
        self.tree = _TREE
        self.appears_on_open = ""  # notebook XML that OpenHierarchy adds to the tree

    def GetHierarchy(self, start: str, scope: int) -> str:
        if start:
            root = ET.fromstring(self.tree)
            el = next(e for e in root.iter() if e.get("ID") == start)
            return ET.tostring(el, encoding="unicode")
        return self.tree

    def OpenHierarchy(self, path: str, relative_to: str, create_type: int) -> str:
        self.calls.append(("open", path, relative_to, create_type))
        self.tree = self.tree.replace("</one:Notebooks>", self.appears_on_open + "</one:Notebooks>")
        return "{NEW}"

    def UpdateHierarchy(self, xml: str) -> None:
        self.calls.append(("update", xml))

    def CloseNotebook(self, notebook_id: str, force: bool) -> None:
        self.calls.append(("close", notebook_id, force))

    def DeleteHierarchy(self, object_id: str, date: float, permanently: bool) -> None:
        self.calls.append(("delete", object_id, permanently))

    def GetPageContent(self, page_id: str, page_info: int) -> str:
        return f'<one:Page {X} ID="{page_id}"><one:Title><one:OE><one:T>x</one:T></one:OE></one:Title></one:Page>'


@pytest.fixture
def app(monkeypatch):
    monkeypatch.delenv("ONENOTE_ALLOWED_NOTEBOOKS", raising=False)
    fake = _FakeApp()
    monkeypatch.setattr(com, "_app", lambda: fake)
    return fake


def _bad_request(call) -> str:
    with pytest.raises(OneNoteError) as exc:
        call()
    assert exc.value.code == "bad_request"
    return str(exc.value)


# ---- create ---------------------------------------------------------------------

def test_create_section_path_and_type(app):
    assert structure.create_section("{G}", "Notes") == {"id": "{NEW}", "name": "Notes"}
    assert app.calls == [("open", "Notes.one", "{G}", 3)]


def test_create_section_strips_dot_one(app):
    structure.create_section("{NB}", "Notes.one")
    assert app.calls[0][1] == "Notes.one"


def test_create_section_with_colour_updates_hierarchy(app):
    structure.create_section("{G}", "Coloured", color="#FF8800")
    (_, xml) = app.calls[1]
    parent = ET.fromstring(xml)
    assert parent.tag == f"{{{NS}}}SectionGroup" and parent.get("ID") == "{G}"
    assert parent[0].attrib == {"ID": "{NEW}", "color": "#FF8800"}


def test_bad_colour_rejected(app):
    assert "color must be #RRGGBB" in _bad_request(lambda: structure.create_section("{NB}", "C", color="red"))


def test_create_section_group_type(app):
    assert structure.create_section_group("{NB}", "05_Team") == {"id": "{NEW}", "name": "05_Team"}
    assert app.calls == [("open", "05_Team", "{NB}", 2)]


@pytest.mark.parametrize("name", ["", "  ", "a/b", "a\\b", "x:y", "what?", "a*b", 'q"', "<x>", "a|b", "..", "."])
def test_bad_name_rejected(app, name):
    assert "name" in _bad_request(lambda: structure.create_section("{NB}", name))
    assert app.calls == []


def test_duplicate_section_rejected(app):
    assert "already has a section named 'top'" in _bad_request(lambda: structure.create_section("{NB}", "top"))


def test_duplicate_group_rejected(app):
    assert "already has a section group named 'Group'" in _bad_request(lambda: structure.create_section_group("{NB}", "Group"))


def test_section_parent_must_be_notebook_or_group(app):
    assert "notebook or section group" in _bad_request(lambda: structure.create_section("{S1}", "x"))


def test_create_in_disallowed_notebook_refused(app, monkeypatch):
    monkeypatch.setenv("ONENOTE_ALLOWED_NOTEBOOKS", "Spike")
    assert "not in an allowed notebook" in _bad_request(lambda: structure.create_section("{OTHER}", "x"))


def test_create_notebook_joins_url_with_slash(app):
    app.appears_on_open = '<one:Notebook ID="{NEW}" name="Glasto" path="https://d.docs.live.net/abc/test/Glasto/"/>'
    result = structure.create_notebook("Glasto", "https://d.docs.live.net/abc/test/")
    assert app.calls == [("open", "https://d.docs.live.net/abc/test/Glasto", "", 1)]
    assert result == {"id": "{NEW}", "name": "Glasto", "path": "https://d.docs.live.net/abc/test/Glasto/"}


def test_create_notebook_joins_local_path_with_backslash(app, tmp_path):
    structure.create_notebook("Local", str(tmp_path))
    assert app.calls[0][1] == str(tmp_path / "Local")


def test_create_notebook_refuses_existing_local_folder(app, tmp_path):
    (tmp_path / "Exists").mkdir()
    assert "already exists" in _bad_request(lambda: structure.create_notebook("Exists", str(tmp_path)))


def test_create_notebook_refuses_relative_location(app):
    assert "absolute folder path or a URL" in _bad_request(lambda: structure.create_notebook("N", "relative\\dir"))


def test_create_notebook_refused_by_allowlist(app, monkeypatch, tmp_path):
    monkeypatch.setenv("ONENOTE_ALLOWED_NOTEBOOKS", "Spike")
    assert "ONENOTE_ALLOWED_NOTEBOOKS" in _bad_request(lambda: structure.create_notebook("New", str(tmp_path)))
    monkeypatch.setenv("ONENOTE_ALLOWED_NOTEBOOKS", "Spike, New")
    structure.create_notebook("New", str(tmp_path))


# ---- open / close ------------------------------------------------------------------

def test_open_notebook(app):
    app.appears_on_open = '<one:Notebook ID="{NEW}" name="Opened" path="D:\\nb\\Opened"/>'
    assert structure.open_notebook("D:\\nb\\Opened") == {"id": "{NEW}", "name": "Opened", "path": "D:\\nb\\Opened"}
    assert app.calls == [("open", "D:\\nb\\Opened", "", 0)]


def test_open_disallowed_notebook_is_closed_again(app, monkeypatch):
    monkeypatch.setenv("ONENOTE_ALLOWED_NOTEBOOKS", "Spike")
    app.appears_on_open = '<one:Notebook ID="{NEW}" name="Opened" path="D:\\nb\\Opened"/>'
    assert "ONENOTE_ALLOWED_NOTEBOOKS" in _bad_request(lambda: structure.open_notebook("D:\\nb\\Opened"))
    assert app.calls[-1] == ("close", "{NEW}", False)


def test_close_notebook_calls_close(app):
    assert structure.close_notebook("{NB}") == "ok"
    assert app.calls == [("close", "{NB}", False)]


def test_close_needs_a_notebook_id(app):
    assert "is not a notebook" in _bad_request(lambda: structure.close_notebook("{S1}"))
