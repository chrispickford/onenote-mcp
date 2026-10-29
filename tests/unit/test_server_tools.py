"""Unit tests for the registered tools: which tools the raw-XML settings
leave, their generated schemas, and how the structured tools handle content
they cannot write.

Builds the server in-process the way main() does; no OneNote and no stdio
transport involved (COM entry points are stubbed where a tool would reach them).
"""
import asyncio
import json

import pytest

try:
    from mcp.server.fastmcp.exceptions import ToolError
    from onenote_mcp import config, server
    _COM_AVAILABLE = True
except Exception:
    _COM_AVAILABLE = False

pytestmark = pytest.mark.skipif(
    not _COM_AVAILABLE,
    reason="com.py module-level imports (comtypes/winreg) unavailable in this environment",
)

_RAW_XML_TOOLS = {"list_hierarchy_xml", "get_page_xml", "replace_page_xml", "append_page_xml"}
_STRUCTURED_TOOLS = {
    "ping", "validate_handles", "get_image_data", "get_notebooks", "list_pages",
    "get_page", "create_page", "replace_page", "append_page", "list_backups", "restore_page",
}


def _tool_names(environ: dict[str, str]) -> set[str]:
    tools = asyncio.run(server.create_server(raw_xml_tools=config.raw_xml_enabled(environ)).list_tools())
    return {tool.name for tool in tools}


@pytest.mark.parametrize("environ", [
    {},
    {"ONENOTE_ENABLE_RAW_XML": ""},
    {"ONENOTE_ENABLE_RAW_XML": "0"},
    {"ONENOTE_DISABLE_RAW_XML": "1"},
    {"ONENOTE_ENABLE_RAW_XML": "1", "ONENOTE_DISABLE_RAW_XML": "true"},
])
def test_raw_xml_tools_off_by_default_and_when_disabled(environ):
    names = _tool_names(environ)
    assert _STRUCTURED_TOOLS <= names
    assert not names & _RAW_XML_TOOLS


@pytest.mark.parametrize("value", ["1", "true", "TRUE", "True", " true "])
def test_raw_xml_tools_registered_when_enabled(value):
    names = _tool_names({"ONENOTE_ENABLE_RAW_XML": value})
    assert _STRUCTURED_TOOLS | _RAW_XML_TOOLS <= names


def test_create_server_default_has_no_raw_xml_tools():
    tools = asyncio.run(server.create_server().list_tools())
    assert not {tool.name for tool in tools} & _RAW_XML_TOOLS


@pytest.mark.parametrize("name", ["ONENOTE_DISABLE_RAW_XML", "ONENOTE_ENABLE_RAW_XML"])
@pytest.mark.parametrize("value", ["yes", "2", "off", "disabled"])
def test_invalid_value_refuses_to_start(monkeypatch, name, value):
    monkeypatch.setenv(name, value)
    with pytest.raises(SystemExit) as exc:
        server.main()
    message = str(exc.value.code)
    assert name in message
    assert repr(value) in message


# ---------------------------------------------------------------------------
# Nested and unsupported content through the structured tools
# ---------------------------------------------------------------------------

def _input_schema(tool_name: str) -> dict:
    tools = asyncio.run(server.create_server().list_tools())
    return next(tool.inputSchema for tool in tools if tool.name == tool_name)


@pytest.mark.parametrize("tool_name", ["replace_page", "append_page"])
def test_write_tool_schemas_describe_nested_and_unsupported_items(tool_name):
    defs = _input_schema(tool_name)["$defs"]
    for model in ("Paragraph", "InlineImage", "ListItem", "UnsupportedItem"):
        assert "children" in defs[model]["properties"], model
    assert defs["UnsupportedItem"]["properties"]["kind"]["enum"] == ["table", "ink", "file", "media", "unknown"]


_UNSUPPORTED_OUTLINE = {"items": [
    {"type": "paragraph", "text": "kept", "children": [{"type": "unsupported", "kind": "table", "text": "a | b"}]},
]}


@pytest.mark.parametrize("tool_name, com_name, arguments", [
    ("replace_page", "_replace_page_com", {"page_id": "p", "title": "t", "outlines": [_UNSUPPORTED_OUTLINE]}),
    ("append_page", "_append_page_com", {"page_id": "p", "outline": _UNSUPPORTED_OUTLINE}),
])
def test_write_tools_reject_unsupported_items_before_touching_onenote(monkeypatch, tool_name, com_name, arguments):
    calls = []
    monkeypatch.setattr(server, com_name, lambda *args: calls.append(args))
    with pytest.raises(ToolError, match=r"bad_request: Cannot write an item of type 'unsupported' \(kind 'table'\)"):
        asyncio.run(server.create_server().call_tool(tool_name, arguments))
    assert calls == []


def test_both_replace_tools_keep_unsupported_page_objects_and_pass_remove_ids(monkeypatch):
    calls = []
    monkeypatch.setattr(server, "_replace_page_com", lambda *args, **kwargs: calls.append((args, kwargs)))
    page_xml = '<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote"/>'
    asyncio.run(server.create_server().call_tool(
        "replace_page", {"page_id": "p", "title": "t", "outlines": [], "remove_ids": ["a"]},
    ))
    asyncio.run(server.create_server(raw_xml_tools=True).call_tool(
        "replace_page_xml", {"page_id": "p", "page_xml": page_xml, "remove_ids": ["b"]},
    ))
    (structured, structured_kw), (raw, raw_kw) = calls
    assert structured[0] == "p" and structured[2] == server.UNSUPPORTED_PAGE_OBJECT_TAGS
    assert raw == ("p", page_xml, server.UNSUPPORTED_PAGE_OBJECT_TAGS)
    assert structured_kw["remove_ids"] == frozenset({"a"}) and raw_kw["remove_ids"] == frozenset({"b"})


def test_replace_page_ignores_the_read_only_outline_height(monkeypatch):
    calls = []
    monkeypatch.setattr(server, "_replace_page_com", lambda *args, **kwargs: calls.append(args))
    read_outline = {"height": 134.27, "items": [{"type": "paragraph", "text": "x"}]}
    asyncio.run(server.create_server().call_tool(
        "replace_page", {"page_id": "p", "title": "t", "outlines": [read_outline]},
    ))
    ((_, page_xml, _),) = [args for args in calls]
    assert "Size" not in page_xml and "134.27" not in page_xml


def test_write_tool_schemas_have_no_outline_height():
    for tool_name in ("replace_page", "append_page"):
        assert "height" not in _input_schema(tool_name)["$defs"]["Outline"]["properties"]


def test_get_page_reports_nested_and_unsupported_content_compactly(monkeypatch):
    xml = (
        '<one:Page xmlns:one="http://schemas.microsoft.com/office/onenote/2013/onenote">'
        "<one:Title><one:OE><one:T>Notes</one:T></one:OE></one:Title>"
        "<one:Outline><one:OEChildren>"
        '<one:OE style="font-size:20.0pt"><one:T>parent</one:T>'
        "<one:OEChildren><one:OE><one:T>child</one:T></one:OE></one:OEChildren></one:OE>"
        '<one:OE><one:InkWord recognizedText="scribble"><one:CallbackID callbackID="{1}"/></one:InkWord></one:OE>'
        "</one:OEChildren></one:Outline>"
        '<one:InkDrawing><one:Position x="1.0" y="2.0"/><one:CallbackID callbackID="{2}"/></one:InkDrawing>'
        "</one:Page>"
    )
    monkeypatch.setattr(server, "_get_page_com", lambda page_id, include_binary: xml)
    assert asyncio.run(server.get_page("p")) == {
        "title": "Notes",
        "outlines": [{"items": [
            {"type": "paragraph", "text": "parent", "font_size": 20.0,
             "children": [{"type": "paragraph", "text": "child"}]},
            {"type": "unsupported", "kind": "ink", "text": "scribble"},
        ]}],
        "unsupported": [{"kind": "ink", "position": {"x": 1.0, "y": 2.0}}],
    }


# ---------------------------------------------------------------------------
# revisions
# ---------------------------------------------------------------------------

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"


def _call(name, arguments, raw_xml_tools=False):
    result = asyncio.run(server.create_server(raw_xml_tools=raw_xml_tools).call_tool(name, arguments))
    return result


def test_get_page_returns_revision_and_last_modified(monkeypatch):
    page_xml = (f'<one:Page xmlns:one="{_ONE_NS}" ID="p" lastModifiedTime="2026-10-29T12:00:00.000Z">'
                '<one:Title><one:OE><one:T>T</one:T></one:OE></one:Title></one:Page>')

    def fake_get_page(page_id, include_binary=False):
        server.pagestate.remember(page_id, "abcdef0123456789")
        return page_xml

    monkeypatch.setattr(server, "_get_page_com", fake_get_page)
    (content,) = _call("get_page", {"page_id": "p"})
    structured = json.loads(content.text)
    assert structured["revision"] == "abcdef0123456789"
    assert structured["last_modified"] == "2026-10-29T12:00:00.000Z"


@pytest.mark.parametrize("tool,arguments,raw", [
    ("replace_page", {"page_id": "p", "title": "t", "outlines": []}, False),
    ("replace_page_xml", {"page_id": "p", "page_xml": f'<one:Page xmlns:one="{_ONE_NS}"/>'}, True),
])
def test_replace_tools_pass_expected_revision(monkeypatch, tool, arguments, raw):
    calls = []
    monkeypatch.setattr(server, "_replace_page_com", lambda *args, **kwargs: calls.append(kwargs))
    _call(tool, {**arguments, "expected_revision": "r1"}, raw_xml_tools=raw)
    assert calls[0]["expected_revision"] == "r1"


@pytest.mark.parametrize("tool,arguments,raw", [
    ("append_page", {"page_id": "p", "outline": {"items": []}}, False),
    ("append_page_xml", {"page_id": "p", "content_xml": f'<one:Outline xmlns:one="{_ONE_NS}"/>'}, True),
])
def test_append_tools_pass_expected_revision(monkeypatch, tool, arguments, raw):
    calls = []
    monkeypatch.setattr(server, "_append_page_com", lambda *args, **kwargs: calls.append(kwargs))
    _call(tool, {**arguments, "expected_revision": "r2"}, raw_xml_tools=raw)
    assert calls[0]["expected_revision"] == "r2"


# ---------------------------------------------------------------------------
# backups
# ---------------------------------------------------------------------------

def test_list_backups_tool_passes_filters(monkeypatch):
    calls = []
    monkeypatch.setattr(server, "_list_backups_com", lambda **kwargs: calls.append(kwargs) or [{"backup_id": "b"}])
    _, structured = _call("list_backups", {"page_id": "p", "since": "2026-10-29T00:00:00Z"})
    assert calls == [{"page_id": "p", "section_id": None, "since": "2026-10-29T00:00:00Z"}]
    assert structured == {"result": [{"backup_id": "b"}]}


def test_restore_page_tool(monkeypatch):
    monkeypatch.setattr(server, "_restore_page_com", lambda backup_id, target_section_id=None: {
        "page_id": "p", "restored_over": target_section_id is None, "missing_attachments": [],
    })
    (content,) = _call("restore_page", {"backup_id": "b"})
    assert json.loads(content.text)["restored_over"] is True
