"""E2E tests: to-do, Important and Question tags round-trip through OneNote.

Run with:  pytest -m e2e tests/e2e/test_tags_via_mcp.py -v
"""
import pytest

from tests.e2e.helpers import without_ids

_ITEMS = [
    {"type": "paragraph", "text": "Book the compound", "tag": {"kind": "todo"}},
    {"type": "paragraph", "text": "Radios collected", "tag": {"kind": "todo", "checked": True}},
    {"type": "paragraph", "text": "Insurance", "style": "h2", "tag": {"kind": "important"}},
    {"type": "list", "style": "bullet", "items": [{"text": "Which gate?", "tag": {"kind": "question"}}]},
]


@pytest.mark.e2e
def test_tags_round_trip(mcp_session, claudespike_section_id):
    page_id = mcp_session.call_tool("create_page", {"section_id": claudespike_section_id, "title": "Tags"})
    mcp_session.call_tool("replace_page", {"page_id": page_id, "title": "Tags", "outlines": [{"items": _ITEMS}]})

    (outline,) = mcp_session.call_tool("get_page", {"page_id": page_id})["outlines"]
    assert without_ids(outline["items"]) == _ITEMS
