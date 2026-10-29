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


@pytest.mark.e2e
def test_rename_move_and_reorder(mcp_session, claudespike_section_id):
    stamp = int(time.time())
    spike = _spike(mcp_session)
    group = mcp_session.call_tool("create_section_group", {"parent_id": spike["id"], "name": f"Moves {stamp}"})
    first = mcp_session.call_tool("create_section", {"parent_id": spike["id"], "name": f"First {stamp}"})
    second = mcp_session.call_tool("create_section", {"parent_id": spike["id"], "name": f"Second {stamp}"})
    pages = [mcp_session.call_tool("create_page", {"section_id": first["id"], "title": title}) for title in ("A", "B", "C")]

    # reorder, then make B a subpage of C
    mcp_session.call_tool("reorder_pages", {"section_id": first["id"], "page_ids": [pages[2], pages[0], pages[1]]})
    mcp_session.call_tool("set_page_level", {"page_id": pages[1], "level": 2})
    listed = mcp_session.call_tool("list_pages", {"section_id": first["id"]})
    assert [(p["name"], p["level"]) for p in listed] == [("C", 1), ("A", 1), ("B", 2)]

    # move page A into the second section: new id, content kept
    moved = mcp_session.call_tool("move_page", {"page_id": pages[0], "target_section_id": second["id"]})
    assert moved["id_changed"] is True
    assert [p["id"] for p in mcp_session.call_tool("list_pages", {"section_id": second["id"]})] == [moved["page_id"]]
    assert mcp_session.call_tool("get_page", {"page_id": moved["page_id"]})["title"] == "A"

    # rename and move the first section into the group
    mcp_session.call_tool("rename", {"object_id": first["id"], "name": f"Renamed {stamp}"})
    mcp_session.call_tool("move_section", {"section_id": first["id"], "new_parent_id": group["id"]})
    tree = _group(_spike(mcp_session), group["id"])
    (inside,) = tree["sections"]
    assert inside["name"] == f"Renamed {stamp}"
    after_move = mcp_session.call_tool("list_pages", {"section_id": inside["id"]})
    assert [p["name"] for p in after_move] == ["C", "B"]
    # Moving a section keeps its id and its pages' ids (verified against OneNote 16.0.20527).
    assert inside["id"] == first["id"]
    assert [p["id"] for p in after_move] == [pages[2], pages[1]]

    with pytest.raises(RuntimeError, match="into itself"):
        mcp_session.call_tool("move_section_group", {"group_id": group["id"], "new_parent_id": group["id"]})


@pytest.mark.e2e
def test_delete_page_section_and_group(mcp_session, claudespike_section_id):
    stamp = int(time.time())
    spike = _spike(mcp_session)
    group = mcp_session.call_tool("create_section_group", {"parent_id": spike["id"], "name": f"Delete {stamp}"})
    section = mcp_session.call_tool("create_section", {"parent_id": group["id"], "name": "Doomed"})
    keep = mcp_session.call_tool("create_page", {"section_id": section["id"], "title": "Keep me"})
    gone = mcp_session.call_tool("create_page", {"section_id": section["id"], "title": "Delete me"})

    result = mcp_session.call_tool("delete_page", {"page_id": gone})
    assert result["backup_id"] and result["recycle_bin"] is True
    assert [p["id"] for p in mcp_session.call_tool("list_pages", {"section_id": section["id"]})] == [keep]

    with pytest.raises(RuntimeError, match="Section has 1 pages; pass include_pages=true"):
        mcp_session.call_tool("delete_section", {"section_id": section["id"]})
    with pytest.raises(RuntimeError, match="include_contents=true"):
        mcp_session.call_tool("delete_section_group", {"group_id": group["id"]})

    deleted = mcp_session.call_tool("delete_section_group", {"group_id": group["id"], "include_contents": True})
    assert deleted["pages_backed_up"] == 1
    assert group["id"] not in [g["id"] for g in _spike(mcp_session).get("section_groups", [])]

    # The deleted page comes back from its backup into another section.
    restored = mcp_session.call_tool("restore_page", {"backup_id": result["backup_id"], "target_section_id": claudespike_section_id})
    assert restored["restored_over"] is False
    assert mcp_session.call_tool("get_page", {"page_id": restored["page_id"]})["title"] == "Delete me"
