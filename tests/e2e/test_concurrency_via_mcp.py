"""E2E tests for page revisions: writes refuse pages changed since they were read.

The "user edit" is made in-process through com.py, i.e. outside the MCP
server, so the server's remembered revision is stale exactly as it would be
after an edit in the OneNote window.

Run with:  pytest -m e2e tests/e2e/test_concurrency_via_mcp.py -v
"""
import pytest

from onenote_mcp import com

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
_OUTLINE = (
    f'<one:Outline xmlns:one="{_ONE_NS}"><one:OEChildren>'
    "<one:OE><one:T>edited outside the server</one:T></one:OE>"
    "</one:OEChildren></one:Outline>"
)


def _paragraphs(content: dict) -> list[str]:
    return [item.get("text") for outline in content["outlines"] for item in outline["items"]]


@pytest.mark.e2e
def test_replace_after_outside_edit_is_conflict(mcp_session, claudespike_section_id):
    page_id = mcp_session.call_tool("create_page", {"section_id": claudespike_section_id, "title": "Revision conflict"})
    mcp_session.call_tool("get_page", {"page_id": page_id})

    com.append_page(page_id, _OUTLINE)

    with pytest.raises(RuntimeError, match="conflict:"):
        mcp_session.call_tool("replace_page", {
            "page_id": page_id, "title": "Revision conflict",
            "outlines": [{"items": [{"type": "paragraph", "text": "would overwrite"}]}],
        })
    assert "edited outside the server" in _paragraphs(mcp_session.call_tool("get_page", {"page_id": page_id}))


@pytest.mark.e2e
def test_replace_after_fresh_read_succeeds(mcp_session, claudespike_section_id):
    page_id = mcp_session.call_tool("create_page", {"section_id": claudespike_section_id, "title": "Revision fresh"})
    com.append_page(page_id, _OUTLINE)
    first = mcp_session.call_tool("get_page", {"page_id": page_id})
    assert first["revision"]

    mcp_session.call_tool("replace_page", {
        "page_id": page_id, "title": "Revision fresh",
        "outlines": [{"items": [{"type": "paragraph", "text": "replaced"}]}],
    })
    assert _paragraphs(mcp_session.call_tool("get_page", {"page_id": page_id})) == ["replaced"]


@pytest.mark.e2e
def test_replace_of_unread_page_is_bad_request(mcp_session, claudespike_section_id):
    # Created outside the server, so the server has never seen its content.
    page_id = com.create_page(claudespike_section_id, "Revision unread")
    with pytest.raises(RuntimeError, match="bad_request: Read the page with get_page first"):
        mcp_session.call_tool("replace_page", {"page_id": page_id, "title": "x", "outlines": []})


@pytest.mark.e2e
def test_two_replaces_in_a_row_need_one_read(mcp_session, claudespike_section_id):
    page_id = mcp_session.call_tool("create_page", {"section_id": claudespike_section_id, "title": "Revision twice"})
    mcp_session.call_tool("get_page", {"page_id": page_id})
    for text in ("one", "two"):
        mcp_session.call_tool("replace_page", {
            "page_id": page_id, "title": "Revision twice",
            "outlines": [{"items": [{"type": "paragraph", "text": text}]}],
        })
    assert _paragraphs(mcp_session.call_tool("get_page", {"page_id": page_id})) == ["two"]
