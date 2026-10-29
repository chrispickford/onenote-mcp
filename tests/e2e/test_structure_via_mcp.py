"""E2E tests for the structure tools: section groups, sections, notebooks.

Section and section-group tests work inside ClaudeSpike.  Creating a notebook
leaves it behind (OneNote cannot delete notebooks), so that test only runs
when ONENOTE_E2E_NOTEBOOK_LOCATION names where it may create one; it prints
the notebook's name and location so it can be deleted by hand.

Run with:  pytest -m e2e tests/e2e/test_structure_via_mcp.py -v -s
"""
import os
import time

import pytest

from onenote_mcp import structure
from tests.e2e.conftest import TEST_NOTEBOOK


def _spike(mcp_session) -> dict:
    return next(nb for nb in mcp_session.call_tool("get_notebooks") if nb["name"] == TEST_NOTEBOOK)


def _group(container: dict, group_id: str) -> dict:
    return next(g for g in container.get("section_groups", []) if g["id"] == group_id)


@pytest.mark.e2e
def test_create_section_group_and_sections(mcp_session, claudespike_section_id):
    stamp = int(time.time())
    spike = _spike(mcp_session)
    group = mcp_session.call_tool("create_section_group", {"parent_id": spike["id"], "name": f"Group {stamp}"})
    nested = mcp_session.call_tool("create_section_group", {"parent_id": group["id"], "name": "Nested"})
    general = mcp_session.call_tool("create_section", {"parent_id": group["id"], "name": "General", "color": "#3C8DBC"})
    deep = mcp_session.call_tool("create_section", {"parent_id": nested["id"], "name": "Deep.one"})

    tree = _group(_spike(mcp_session), group["id"])
    assert tree["name"] == f"Group {stamp}"
    assert tree["sections"] == [{"id": general["id"], "name": "General", "color": "#3C8DBC"}]
    assert _group(tree, nested["id"])["sections"][0]["id"] == deep["id"]
    assert _group(tree, nested["id"])["sections"][0]["name"] == "Deep"

    page_id = mcp_session.call_tool("create_page", {"section_id": general["id"], "title": "In a group"})
    assert [p["id"] for p in mcp_session.call_tool("list_pages", {"section_id": general["id"]})] == [page_id]

    with pytest.raises(RuntimeError, match="already has a section named 'general'"):
        mcp_session.call_tool("create_section", {"parent_id": group["id"], "name": "general"})


@pytest.mark.e2e
def test_structure_outside_the_allowlist_is_refused(mcp_session, claudespike_section_id):
    with pytest.raises(RuntimeError, match="not found in any open notebook|not in an allowed notebook"):
        mcp_session.call_tool("create_section", {"parent_id": "{00000000-0000-0000-0000-000000000000}{1}{B0}", "name": "x"})


@pytest.mark.e2e
def test_create_notebook(claudespike_section_id, monkeypatch):
    location = os.environ.get("ONENOTE_E2E_NOTEBOOK_LOCATION")
    if not location:
        pytest.skip("set ONENOTE_E2E_NOTEBOOK_LOCATION to a folder or URL where a test notebook may be created")
    name = f"ClaudeSpike E2E {int(time.time())}"
    # The suite allowlists ClaudeSpike only; allow the new notebook for this call.
    monkeypatch.setenv("ONENOTE_ALLOWED_NOTEBOOKS", f"{TEST_NOTEBOOK}, {name}")
    created = structure.create_notebook(name, location)
    print(f"\nLeft behind for deletion by hand: notebook {name!r} at {created['path']}")
    assert created["name"] == name
    section = structure.create_section(created["id"], "First")
    assert section["name"] == "First"
    assert structure.close_notebook(created["id"]) == "ok"
