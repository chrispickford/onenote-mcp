"""E2E tests: tables and hyperlinks written by replace_page read back unchanged.

Run with:  pytest -m e2e tests/e2e/test_tables_links_via_mcp.py -v
"""
import pytest

from tests.e2e.helpers import without_ids

_TABLE = {
    "type": "table",
    "header_row": True,
    "column_widths": [120.0, 220.0],
    "rows": [
        [{"items": [{"type": "paragraph", "text": "Item", "bold": True}]},
         {"items": [{"type": "paragraph", "text": "Notes", "bold": True}]}],
        [{"items": [{"type": "paragraph", "text": "Radios"}]},
         {"items": [
             {"type": "list", "style": "bullet", "items": [{"text": "10 handsets"}, {"text": "2 chargers"}]},
             {"type": "table", "rows": [[{"items": [{"type": "paragraph", "text": "nested"}]}]]},
         ]}],
        [{"items": [{"type": "paragraph", "text": ""}]},
         {"items": [{"type": "paragraph", "text": "empty cell on the left"}]}],
    ],
}

_LINKED = {
    "type": "paragraph",
    "segments": [
        {"text": "See "},
        {"text": "the site", "link": "https://example.com/a?b=1&c=2"},
        {"text": " or "},
        {"text": "mail us", "link": "mailto:team@example.com", "bold": True},
    ],
}


@pytest.mark.e2e
def test_table_and_links_round_trip(mcp_session, claudespike_section_id):
    page_id = mcp_session.call_tool("create_page", {"section_id": claudespike_section_id, "title": "Tables and links"})
    mcp_session.call_tool("replace_page", {
        "page_id": page_id, "title": "Tables and links", "outlines": [{"items": [_LINKED, _TABLE]}],
    })

    (outline,) = mcp_session.call_tool("get_page", {"page_id": page_id})["outlines"]
    assert without_ids(outline["items"]) == [_LINKED, _TABLE]

    # Writing the read content back unchanged keeps it unchanged.
    mcp_session.call_tool("replace_page", {
        "page_id": page_id, "title": "Tables and links", "outlines": [{"items": outline["items"]}],
    })
    (again,) = mcp_session.call_tool("get_page", {"page_id": page_id})["outlines"]
    assert without_ids(again["items"]) == [_LINKED, _TABLE]
