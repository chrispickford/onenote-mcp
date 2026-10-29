"""E2E tests: content the structured tools cannot re-create is reported, not dropped.

A page seeded via replace_page_xml with a file attached inside its outline and
an indented paragraph must read back through get_page as an attachment item and
the nested child, and replace_page must refuse to write the attachment back or to
drop it silently.  A file attached directly to the page canvas must survive
replace_page.  Requires OneNote desktop with the ClaudeSpike notebook open.
The test pages are left in place after the test (no cleanup); the unique
titles avoid collisions on re-runs.
"""

from __future__ import annotations

import time

import pytest

from tests.e2e.helpers import without_ids

from tests.e2e.mcp_client import mcp_session  # noqa: F401 (re-export fixture)

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"


def _seed_page(mcp_session, section_id: str, title: str, attachment) -> str:
    """Create a page holding an attached file and a paragraph with an indented child; return its id."""
    attachment.write_text("seeded by the e2e suite")
    page_id = mcp_session.call_tool("create_page", {"section_id": section_id, "title": title})
    page_xml = (
        f'<one:Page xmlns:one="{_ONE_NS}" ID="{page_id}">'
        f"<one:Title><one:OE><one:T>{title}</one:T></one:OE></one:Title>"
        "<one:Outline><one:OEChildren>"
        f'<one:OE><one:InsertedFile pathSource="{attachment}" preferredName="seed.txt"/></one:OE>'
        "<one:OE><one:T>Parent paragraph</one:T>"
        "<one:OEChildren><one:OE><one:T>Indented child</one:T></one:OE></one:OEChildren></one:OE>"
        "</one:OEChildren></one:Outline>"
        "</one:Page>"
    )
    result = mcp_session.call_tool("replace_page_xml", {"page_id": page_id, "page_xml": page_xml})
    assert result == "ok", f"replace_page_xml returned {result!r}"
    return page_id


_PARENT = {
    "type": "paragraph",
    "text": "Parent paragraph",
    "children": [{"type": "paragraph", "text": "Indented child"}],
}


@pytest.mark.e2e
def test_get_page_reports_file_and_indented_paragraph(mcp_session, claudespike_section_id, tmp_path):
    page_id = _seed_page(mcp_session, claudespike_section_id, f"test-unsupported-{int(time.time())}", tmp_path / "seed.txt")

    attached, parent = mcp_session.call_tool("get_page", {"page_id": page_id})["outlines"][0]["items"]
    assert attached.pop("id")
    assert attached == {"type": "attachment", "name": "seed.txt"}
    assert without_ids(parent) == _PARENT


@pytest.mark.e2e
def test_replace_page_refuses_to_drop_the_file(mcp_session, claudespike_section_id, tmp_path):
    title = f"test-unsupported-rewrite-{int(time.time())}"
    page_id = _seed_page(mcp_session, claudespike_section_id, title, tmp_path / "seed.txt")
    content = mcp_session.call_tool("get_page", {"page_id": page_id})

    with pytest.raises(RuntimeError, match=r"bad_request: An attachment read from a page cannot be written back"):
        mcp_session.call_tool("replace_page", {
            "page_id": page_id, "title": title, "outlines": content["outlines"],
        })
    assert mcp_session.call_tool("get_page", {"page_id": page_id})["outlines"] == content["outlines"]

    # Leaving the marker out is not enough: the page still holds the file.
    with pytest.raises(RuntimeError, match=r"bad_request: The page holds content a rewrite cannot re-create: file"):
        mcp_session.call_tool("replace_page", {
            "page_id": page_id, "title": title, "outlines": [{"items": [_PARENT]}],
        })

    # Listing its ID in remove_ids deletes the file deliberately; the indented child stays.
    file_id = content["outlines"][0]["items"][0]["id"]
    result = mcp_session.call_tool("replace_page", {
        "page_id": page_id, "title": title, "outlines": [{"items": [_PARENT]}], "remove_ids": [file_id],
    })
    assert result == "ok", f"replace_page returned {result!r}"
    assert without_ids(mcp_session.call_tool("get_page", {"page_id": page_id})["outlines"][0]["items"]) == [_PARENT]


def _without_z(page_object: dict) -> dict:
    """Return a get_page `unsupported` entry without the stacking order of its position."""
    position = {key: value for key, value in page_object["position"].items() if key != "z"}
    return {**page_object, "position": position}


def _page_with_attached_file(title: str, attachment: str) -> str:
    """Return page XML with one paragraph and a file attached directly to the canvas."""
    return (
        f'<one:Page xmlns:one="{_ONE_NS}">'
        f"<one:Title><one:OE><one:T>{title}</one:T></one:OE></one:Title>"
        "<one:Outline><one:OEChildren><one:OE><one:T>old text</one:T></one:OE></one:OEChildren></one:Outline>"
        f'<one:InsertedFile pathSource="{attachment}" preferredName="attachment.txt">'
        '<one:Position x="400.0" y="300.0"/></one:InsertedFile>'
        "</one:Page>"
    )


@pytest.mark.e2e
def test_replace_page_keeps_a_file_attached_to_the_canvas(mcp_session, claudespike_section_id, tmp_path):
    attachment = tmp_path / "attachment.txt"
    attachment.write_text("attached by the e2e suite")
    title = f"test-unsupported-attached-file-{int(time.time())}"
    page_id = mcp_session.call_tool("create_page", {"section_id": claudespike_section_id, "title": title})
    seed = mcp_session.call_tool("replace_page_xml", {
        "page_id": page_id, "page_xml": _page_with_attached_file(title, str(attachment)),
    })
    assert seed == "ok", f"replace_page_xml returned {seed!r}"
    (attached,) = mcp_session.call_tool("get_page", {"page_id": page_id})["unsupported"]
    assert attached["kind"] == "file" and attached["text"] == "attachment.txt"

    new_items = [{"type": "paragraph", "text": "new text"}]
    result = mcp_session.call_tool("replace_page", {
        "page_id": page_id, "title": title, "outlines": [{"items": new_items}],
    })
    assert result == "ok", f"replace_page returned {result!r}"
    content = mcp_session.call_tool("get_page", {"page_id": page_id})
    assert [without_ids(outline["items"]) for outline in content["outlines"]] == [new_items]
    # Same place and size; OneNote renumbers the stacking order (z) once the
    # outline below the file is deleted.
    (kept,) = content["unsupported"]
    assert _without_z(kept) == _without_z(attached)

    # The raw tool has the same protection: the canvas file stays.
    page_xml = (
        f'<one:Page xmlns:one="{_ONE_NS}"><one:Title><one:OE><one:T>{title}</one:T></one:OE></one:Title></one:Page>'
    )
    assert mcp_session.call_tool("replace_page_xml", {"page_id": page_id, "page_xml": page_xml}) == "ok"
    (still_there,) = mcp_session.call_tool("get_page", {"page_id": page_id})["unsupported"]
    assert still_there["text"] == "attachment.txt"
