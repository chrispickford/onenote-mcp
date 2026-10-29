"""Smoke test: MCP stdio transport wire-up through server.py.

Calls ``get_notebooks`` through the real stdio transport and asserts the
basic shape of the response.  This proves the MCP wire works end-to-end
without any mocking.
"""

import xml.etree.ElementTree as ET

import pytest

from tests.e2e.mcp_client import mcp_session  # noqa: F401 — re-export fixture


@pytest.mark.e2e
def test_get_notebooks_via_mcp_transport(mcp_session):
    """get_notebooks over real stdio MCP returns a non-empty notebook skeleton."""
    result = mcp_session.call_tool("get_notebooks")

    assert isinstance(result, list), (
        f"Expected list from get_notebooks, got {type(result).__name__}: {result!r}"
    )
    assert len(result) > 0, (
        "Expected at least one notebook — open a notebook in OneNote desktop"
    )
    first = result[0]
    assert "id" in first and "name" in first
    assert "sections" in first and isinstance(first["sections"], list), (
        f"Each notebook must carry a 'sections' list, got keys: {list(first.keys())}"
    )


@pytest.mark.e2e
def test_get_notebooks_lists_sections_inside_section_groups(mcp_session, claudespike_section_id):
    """A section inside a section group appears under that group, not at the top."""
    from onenote_mcp import com

    def make():
        app = com._app()
        notebooks = ET.fromstring(app.GetHierarchy("", 2))
        nb_id = next(nb.get("ID") for nb in notebooks if nb.get("name") == "ClaudeSpike")
        group_id = app.OpenHierarchy("Grouped", nb_id, 2)  # cftFolder: opens it if it exists
        section_id = app.OpenHierarchy("Inside.one", group_id, 3)  # cftSection
        return group_id, section_id

    group_id, section_id = com._run_com(make)
    (spike,) = mcp_session.call_tool("get_notebooks")
    group = next(g for g in spike["section_groups"] if g["id"] == group_id)
    assert section_id in [s["id"] for s in group["sections"]]
    assert section_id not in [s["id"] for s in spike["sections"]]
