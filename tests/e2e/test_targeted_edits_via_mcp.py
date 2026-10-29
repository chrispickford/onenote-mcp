"""E2E tests for the targeted edit tools: update_item, insert_items,
delete_items and move_item.  Each edit must leave untouched items (and their
ids) as they were, including an image and an attachment in the same outline.

Run with:  pytest -m e2e tests/e2e/test_targeted_edits_via_mcp.py -v
"""
import base64
import struct
import zlib

import pytest

from tests.e2e.helpers import without_ids


def _png() -> str:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)
    raw = b"".join(b"\x00" + b"\x10\x60\xa0" * 30 for _ in range(15))
    png = (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 30, 15, 8, 2, 0, 0, 0))
           + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))
    return base64.b64encode(png).decode()


def _items(mcp_session, page_id):
    (outline,) = mcp_session.call_tool("get_page", {"page_id": page_id})["outlines"]
    return outline["items"]


def _ids(items) -> dict[str, str]:
    """Map each paragraph text / attachment name / item type to its id (lists have none; their items do)."""
    return {item.get("text") or item.get("name") or item["type"]: item["id"] for item in items if item["type"] != "list"}


@pytest.mark.e2e
def test_targeted_edits(mcp_session, claudespike_section_id, tmp_path):
    notes = tmp_path / "notes.txt"
    notes.write_text("kept through edits")
    page_id = mcp_session.call_tool("create_page", {"section_id": claudespike_section_id, "title": "Targeted edits"})
    mcp_session.call_tool("replace_page", {"page_id": page_id, "title": "Targeted edits", "outlines": [{"items": [
        {"type": "paragraph", "text": "one"},
        {"type": "paragraph", "text": "two"},
        {"type": "inline_image", "data": _png()},
        {"type": "attachment", "path": str(notes), "name": "notes.txt"},
        {"type": "paragraph", "text": "three"},
        {"type": "list", "style": "bullet", "items": [{"text": "a"}, {"text": "b"}]},
    ]}]})
    items = _items(mcp_session, page_id)
    ids = _ids(items)
    list_ids = [entry["id"] for entry in items[-1]["items"]]

    # update the middle paragraph: same id, new text, everything else untouched
    mcp_session.call_tool("update_item", {"page_id": page_id, "item_id": ids["two"],
                                          "item": {"type": "paragraph", "text": "TWO", "bold": True}})
    # insert after the first paragraph
    result = mcp_session.call_tool("insert_items", {"page_id": page_id, "after_id": ids["one"],
                                                    "items": [{"type": "paragraph", "text": "inserted"}]})
    assert len(result["ids"]) == 1
    # delete one list item
    mcp_session.call_tool("delete_items", {"page_id": page_id, "item_ids": [list_ids[0]]})
    # move "three" under "one"
    mcp_session.call_tool("move_item", {"page_id": page_id, "item_id": ids["three"], "under_id": ids["one"]})

    items = _items(mcp_session, page_id)
    assert [item["type"] for item in items] == ["paragraph", "paragraph", "paragraph", "inline_image", "attachment", "list"]
    assert without_ids(items[0]) == {"type": "paragraph", "text": "one", "children": [{"type": "paragraph", "text": "three"}]}
    assert items[0]["id"] == ids["one"] and items[0]["children"][0]["id"] == ids["three"]
    assert items[1]["text"] == "inserted" and items[1]["id"] == result["ids"][0]
    assert (items[2]["text"], items[2].get("bold"), items[2]["id"]) == ("TWO", True, ids["two"])
    assert items[3]["id"] == ids["inline_image"] and items[3]["handle"].startswith("mcpref:")
    assert items[4]["id"] == ids["notes.txt"]
    assert [entry["text"] for entry in items[5]["items"]] == ["b"] and items[5]["items"][0]["id"] == list_ids[1]


@pytest.mark.e2e
def test_list_item_update_and_heading_insert(mcp_session, claudespike_section_id):
    page_id = mcp_session.call_tool("create_page", {"section_id": claudespike_section_id, "title": "List edits"})
    mcp_session.call_tool("replace_page", {"page_id": page_id, "title": "List edits", "outlines": [{"items": [
        {"type": "paragraph", "text": "Intro", "style": "h2"},
        {"type": "list", "style": "numbered", "items": [{"text": "first"}, {"text": "second"}]},
    ]}]})
    items = _items(mcp_session, page_id)
    second = items[1]["items"][1]["id"]

    mcp_session.call_tool("update_item", {"page_id": page_id, "item_id": second,
                                          "item": {"text": "second, done", "tag": {"kind": "todo", "checked": True}}})
    mcp_session.call_tool("insert_items", {"page_id": page_id, "after_id": items[1]["items"][1]["id"],
                                           "items": [{"type": "paragraph", "text": "Next steps", "style": "h2"},
                                                     {"type": "paragraph", "text": "Summary", "style": "h3"}]})

    assert without_ids(_items(mcp_session, page_id)) == [
        {"type": "paragraph", "text": "Intro", "style": "h2"},
        {"type": "list", "style": "numbered", "items": [
            {"text": "first"}, {"text": "second, done", "tag": {"kind": "todo", "checked": True}},
        ]},
        {"type": "paragraph", "text": "Next steps", "style": "h2"},
        {"type": "paragraph", "text": "Summary", "style": "h3"},
    ]


@pytest.mark.e2e
def test_replace_text_rename_and_set_tag(mcp_session, claudespike_section_id):
    page_id = mcp_session.call_tool("create_page", {"section_id": claudespike_section_id, "title": "Find replace"})
    mcp_session.call_tool("replace_page", {"page_id": page_id, "title": "Find replace", "outlines": [{"items": [
        {"type": "paragraph", "segments": [{"text": "Meet at "}, {"text": "Gate A", "bold": True}, {"text": " by gate b"}]},
        {"type": "table", "rows": [["gate", "time"], ["Gate C", "09:00"]]},
        {"type": "paragraph", "text": "Check radios"},
    ]}]})

    result = mcp_session.call_tool("replace_text", {"page_id": page_id, "find": "gate", "replace": "Door"})
    assert result["replacements"] == 4
    assert mcp_session.call_tool("replace_text", {"page_id": page_id, "find": "nowhere", "replace": "x"})["replacements"] == 0

    mcp_session.call_tool("rename_page", {"page_id": page_id, "title": "Find & replace <done>"})
    items = _items(mcp_session, page_id)
    mcp_session.call_tool("set_tag", {"page_id": page_id, "item_id": items[2]["id"], "tag": {"kind": "todo", "checked": True}})

    content = mcp_session.call_tool("get_page", {"page_id": page_id})
    assert content["title"] == "Find & replace <done>"
    first, table, last = without_ids(content["outlines"][0]["items"])
    assert first == {"type": "paragraph", "segments": [{"text": "Meet at "}, {"text": "Door A", "bold": True}, {"text": " by Door b"}]}
    assert [[cell["items"][0]["text"] for cell in row] for row in table["rows"]] == [["Door", "time"], ["Door C", "09:00"]]
    assert last == {"type": "paragraph", "text": "Check radios", "tag": {"kind": "todo", "checked": True}}
