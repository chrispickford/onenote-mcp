"""Unit tests for the schema pin in com.py: every XML-exchanging COM call asks
OneNote for the 2013 schema, which is the namespace the parsers understand.

Newer OneNote builds answer in a newer namespace by default (e.g. .../2026/onenote),
which the 2013-namespace parsers would read as an empty notebook list.
"""
import pytest

try:
    from onenote_mcp import com
    _COM_AVAILABLE = True
except Exception:
    _COM_AVAILABLE = False

pytestmark = pytest.mark.skipif(
    not _COM_AVAILABLE,
    reason="com.py module-level imports (comtypes/winreg) unavailable in this environment",
)


class _RawApp:
    """Records the positional arguments each COM method receives."""

    def __init__(self) -> None:
        self.calls: list[tuple] = []

    def __getattr__(self, name):
        def method(*args):
            self.calls.append((name, args))
            return "ok"
        return method


def test_get_hierarchy_pins_2013_schema():
    raw = _RawApp()
    com._SchemaPinnedApp(raw).GetHierarchy("", 4)
    assert raw.calls == [("GetHierarchy", ("", 4, com.XS_2013))]


def test_get_page_content_pins_2013_schema():
    raw = _RawApp()
    com._SchemaPinnedApp(raw).GetPageContent("{P}", 1)
    assert raw.calls == [("GetPageContent", ("{P}", 1, com.XS_2013))]


def test_update_page_content_pins_2013_schema():
    raw = _RawApp()
    com._SchemaPinnedApp(raw).UpdatePageContent("<x/>", 0.0)
    assert raw.calls == [("UpdatePageContent", ("<x/>", 0.0, com.XS_2013))]


def test_update_hierarchy_pins_2013_schema():
    raw = _RawApp()
    com._SchemaPinnedApp(raw).UpdateHierarchy("<x/>")
    assert raw.calls == [("UpdateHierarchy", ("<x/>", com.XS_2013))]


def test_other_methods_pass_through():
    raw = _RawApp()
    com._SchemaPinnedApp(raw).DeletePageContent("{P}", "{O}", 0.0)
    assert raw.calls == [("DeletePageContent", ("{P}", "{O}", 0.0))]


def test_xs_2013_value():
    # XMLSchema enum: xs2007 = 0, xs2010 = 1, xs2013 = 2.
    assert com.XS_2013 == 2
