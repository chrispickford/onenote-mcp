"""Unit tests for the title write in _create_page_impl.

A fake app stands in for OneNote.Application and records the XML it receives.
"""
import xml.etree.ElementTree as ET

import pytest

try:
    from onenote_mcp import com, pagestate
    from onenote_mcp.builders import parse_page
    _COM_AVAILABLE = True
except Exception:
    _COM_AVAILABLE = False

pytestmark = pytest.mark.skipif(
    not _COM_AVAILABLE,
    reason="com.py module-level imports (comtypes/winreg) unavailable in this environment",
)

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"


class _FakeApp:
    """Stand-in for OneNote.Application providing CreateNewPage and UpdatePageContent."""

    def __init__(self) -> None:
        self.updates: list[str] = []

    def CreateNewPage(self, section_id: str) -> str:
        return "{NEW-PAGE}"

    def UpdatePageContent(self, xml: str, date_expected_last_modified: float) -> None:
        self.updates.append(xml)

    def GetPageContent(self, page_id: str, page_info: int) -> str:
        return self.updates[-1] if self.updates else f'<one:Page xmlns:one="{_ONE_NS}" ID="{page_id}"/>'


def test_create_page_escapes_title(monkeypatch):
    monkeypatch.delenv("ONENOTE_ALLOWED_NOTEBOOKS", raising=False)
    app = _FakeApp()
    monkeypatch.setattr(com, "_app", lambda: app)
    title = "<b>x</b> & ]]> \"q\" äöü ß"

    assert com._create_page_impl("{SECTION}", title) == "{NEW-PAGE}"

    (xml,) = app.updates
    t_el = ET.fromstring(xml).find(f".//{{{_ONE_NS}}}T")
    assert t_el.text == "&lt;b&gt;x&lt;/b&gt; &amp; ]]&gt; \"q\" äöü ß"
    assert parse_page(xml).title == title


def test_create_page_counts_as_a_read(monkeypatch):
    # The client knows what a page it just created holds, so a replace may
    # follow without a get_page.
    monkeypatch.delenv("ONENOTE_ALLOWED_NOTEBOOKS", raising=False)
    pagestate.forget_all()
    app = _FakeApp()
    monkeypatch.setattr(com, "_app", lambda: app)

    page_id = com._create_page_impl("{SECTION}", "Fresh")

    assert pagestate.remembered(page_id) == pagestate.revision(app.GetPageContent(page_id, 0))
    pagestate.forget_all()
