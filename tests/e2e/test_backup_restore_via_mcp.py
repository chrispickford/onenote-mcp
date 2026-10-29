"""E2E tests for backups and restore_page through the MCP server.

Run with:  pytest -m e2e tests/e2e/test_backup_restore_via_mcp.py -v
"""
import pytest


def _texts(mcp_session, page_id: str) -> list[str]:
    content = mcp_session.call_tool("get_page", {"page_id": page_id})
    return [item.get("text") for outline in content["outlines"] for item in outline["items"]]


def _replace(mcp_session, page_id: str, text: str) -> None:
    mcp_session.call_tool("replace_page", {
        "page_id": page_id, "title": "Backup restore",
        "outlines": [{"items": [{"type": "paragraph", "text": text}]}],
    })


@pytest.mark.e2e
def test_restore_after_replace(mcp_session, claudespike_section_id):
    page_id = mcp_session.call_tool("create_page", {"section_id": claudespike_section_id, "title": "Backup restore"})
    _replace(mcp_session, page_id, "before")
    _replace(mcp_session, page_id, "after")
    assert _texts(mcp_session, page_id) == ["after"]

    latest = mcp_session.call_tool("list_backups", {"page_id": page_id})[0]
    assert latest["tool"] == "replace_page"
    result = mcp_session.call_tool("restore_page", {"backup_id": latest["backup_id"]})

    assert result["restored_over"] is True
    assert _texts(mcp_session, page_id) == ["before"]
    assert mcp_session.call_tool("list_backups", {"page_id": page_id})[0]["tool"] == "restore_page"
