"""E2E tests: images inserted from a file or from base64 data, and attachments.

The server never opens these files; OneNote reads them from ONENOTE_FILE_ROOTS
(set by conftest.py to pytest's temp folder and the test files folder).

Run with:  pytest -m e2e tests/e2e/test_files_via_mcp.py -v
"""
import base64
import struct
import zlib

import pytest


def _png(width: int, height: int) -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)
    raw = b"".join(b"\x00" + b"\x20\x80\xc0" * width for _ in range(height))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


@pytest.mark.e2e
def test_images_and_attachment_from_files(mcp_session, claudespike_section_id, tmp_path):
    picture = tmp_path / "picture.png"
    picture.write_bytes(_png(60, 30))
    notes = tmp_path / "notes.txt"
    notes.write_text("attached by the e2e suite")
    page_id = mcp_session.call_tool("create_page", {"section_id": claudespike_section_id, "title": "Files"})

    mcp_session.call_tool("replace_page", {
        "page_id": page_id, "title": "Files",
        "outlines": [{"items": [
            {"type": "paragraph", "text": "from a path"},
            {"type": "inline_image", "path": str(picture)},
            {"type": "inline_image", "data": base64.b64encode(_png(20, 10)).decode(), "width": 40.0},
            {"type": "attachment", "path": str(notes), "name": "Meeting notes.txt"},
        ]}],
        "images": [{"path": str(picture), "width": 120.0, "height": 60.0, "position": {"x": 400.0, "y": 100.0}}],
    })

    content = mcp_session.call_tool("get_page", {"page_id": page_id})
    items = content["outlines"][0]["items"]
    assert [item["type"] for item in items] == ["paragraph", "inline_image", "inline_image", "attachment"]
    assert items[1]["handle"].startswith("mcpref:")
    assert items[2]["width"] == pytest.approx(40.0) and items[2]["height"] == pytest.approx(20.0)
    assert items[3]["name"] == "Meeting notes.txt" and items[3]["id"]
    (floating,) = content["images"]
    assert (floating["width"], floating["height"]) == (pytest.approx(120.0), pytest.approx(60.0))


@pytest.mark.e2e
def test_file_outside_roots_is_refused(mcp_session, claudespike_section_id):
    page_id = mcp_session.call_tool("create_page", {"section_id": claudespike_section_id, "title": "Files refused"})
    with pytest.raises(RuntimeError, match="outside ONENOTE_FILE_ROOTS"):
        mcp_session.call_tool("append_page", {
            "page_id": page_id,
            "outline": {"items": [{"type": "attachment", "path": "C:\\Windows\\win.ini"}]},
        })
