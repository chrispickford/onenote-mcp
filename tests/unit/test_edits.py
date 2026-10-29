"""Unit tests for edits.py: targeted item edits on a parsed page tree.

OneNote replaces an outline's whole child list when the outline is sent with
its objectID (verified against 16.0.20527), so every edit returns the
complete modified outlines; OEs keep their objectIDs, new ones have none.
"""
import xml.etree.ElementTree as ET

import pytest

from onenote_mcp import edits
from onenote_mcp.errors import OneNoteError
from onenote_mcp.models import ListItem, Paragraph, Tag

NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
X = f'xmlns:one="{NS}"'
q = lambda tag: f"{{{NS}}}{tag}"  # noqa: E731


def _page(body: str, defs: str = "") -> ET.Element:
    return ET.fromstring(
        f'<one:Page {X} ID="{{P}}">{defs}<one:Title><one:OE objectID="title"><one:T>T</one:T></one:OE></one:Title>'
        f"{body}</one:Page>"
    )


_BODY = (
    '<one:Outline objectID="o1"><one:OEChildren>'
    '<one:OE objectID="a"><one:T>alpha</one:T><one:OEChildren>'
    '<one:OE objectID="a1"><one:T>alpha child</one:T></one:OE></one:OEChildren></one:OE>'
    '<one:OE objectID="b"><one:T>beta</one:T></one:OE>'
    '<one:OE objectID="l1"><one:List><one:Bullet bullet="2"/></one:List><one:T>item 1</one:T></one:OE>'
    '<one:OE objectID="l2"><one:List><one:Bullet bullet="2"/></one:List><one:T>item 2</one:T></one:OE>'
    '</one:OEChildren></one:Outline>'
    '<one:Outline objectID="o2"><one:OEChildren><one:OE objectID="c"><one:T>gamma</one:T></one:OE>'
    '</one:OEChildren></one:Outline>'
)


def _texts(outline: ET.Element) -> list[tuple[str | None, str]]:
    """(objectID, text) of the outline's top-level OEs."""
    return [
        (oe.get("objectID"), "".join(t.text or "" for t in oe.findall(q("T"))))
        for oe in outline.find(q("OEChildren")).findall(q("OE"))
    ]


def _xml(page_root, edit) -> ET.Element:
    return ET.fromstring(edits.edit_xml("{P}", page_root, edit))


def test_find_oe_nested():
    oe, parent, outline = edits.find_oe(_page(_BODY), "a1")
    assert oe.find(q("T")).text == "alpha child"
    assert outline.get("objectID") == "o1"
    assert oe in list(parent)


@pytest.mark.parametrize("item_id", ["missing", "title", "o1"])
def test_find_oe_missing_is_bad_request(item_id):
    with pytest.raises(OneNoteError) as exc:
        edits.find_oe(_page(_BODY), item_id)
    assert exc.value.code == "bad_request"
    assert "not found on the page; re-read it with get_page" in str(exc.value)


def test_update_replaces_only_target_and_keeps_its_id():
    page = _page(_BODY)
    edit = edits.update_item(page, "b", Paragraph(type="paragraph", text="beta 2"))
    (outline,) = _xml(page, edit).findall(q("Outline"))
    assert outline.get("objectID") == "o1"
    assert _texts(outline) == [("a", "alpha"), ("b", "beta 2"), ("l1", "item 1"), ("l2", "item 2")]
    assert outline.find(f".//{q('OE')}[@objectID='a1']") is not None, "untouched children are re-sent"


def test_update_list_item_with_list_item_keeps_list_style():
    page = _page(_BODY)
    edit = edits.update_item(page, "l2", ListItem(text="item two", tag=Tag(kind="todo")))
    oe = _xml(page, edit).find(f".//{q('OE')}[@objectID='l2']")
    assert oe.find(f"{q('List')}/{q('Bullet')}") is not None
    assert oe.find(q("T")).text == "item two"
    assert oe.find(q("Tag")) is not None


def test_update_to_several_oes_gives_new_ids():
    from onenote_mcp.models import List
    page = _page(_BODY)
    edit = edits.update_item(page, "b", List(type="list", style="numbered", items=[ListItem(text="x"), ListItem(text="y")]))
    assert _texts(_xml(page, edit).find(q("Outline")))[1:3] == [(None, "x"), (None, "y")]


def test_insert_after_position():
    page = _page(_BODY)
    edit = edits.insert_items(page, [Paragraph(type="paragraph", text="new")], after_id="a")
    assert _texts(_xml(page, edit).find(q("Outline")))[:3] == [("a", "alpha"), (None, "new"), ("b", "beta")]


def test_insert_before_nested_item():
    page = _page(_BODY)
    edit = edits.insert_items(page, [Paragraph(type="paragraph", text="first child")], before_id="a1")
    outline = _xml(page, edit).find(q("Outline"))
    children = outline.find(f".//{q('OE')}[@objectID='a']/{q('OEChildren')}")
    assert [t.text for t in children.iter(q("T"))] == ["first child", "alpha child"]


@pytest.mark.parametrize("at_end, expected", [(True, ["gamma", "end"]), (False, ["end", "gamma"])])
def test_insert_into_outline(at_end, expected):
    page = _page(_BODY)
    edit = edits.insert_items(page, [Paragraph(type="paragraph", text="end")], outline_id="o2", at_end=at_end)
    (outline,) = _xml(page, edit).findall(q("Outline"))
    assert [text for _, text in _texts(outline)] == expected


def test_insert_needs_exactly_one_anchor():
    with pytest.raises(OneNoteError, match="exactly one of before_id, after_id or outline_id"):
        edits.insert_items(_page(_BODY), [Paragraph(type="paragraph", text="x")], before_id="a", after_id="b")


def test_insert_list_item_after_list_item_uses_its_style():
    page = _page(_BODY)
    edit = edits.insert_items(page, [ListItem(text="item 3")], after_id="l2")
    oe = [oe for oe in _xml(page, edit).find(q("Outline")).find(q("OEChildren")) if oe.get("objectID") is None][0]
    assert oe.find(f"{q('List')}/{q('Bullet')}") is not None


def test_delete_items():
    page = _page(_BODY)
    edit = edits.delete_items(page, ["b", "a1"])
    outline = _xml(page, edit).find(q("Outline"))
    assert [oid for oid, _ in _texts(outline)] == ["a", "l1", "l2"]
    assert outline.find(f".//{q('OE')}[@objectID='a1']") is None
    assert edit.delete_outline_ids == []


def test_deleting_the_last_item_deletes_the_outline():
    page = _page(_BODY)
    edit = edits.delete_items(page, ["c"])
    assert edit.delete_outline_ids == ["o2"]
    assert _xml(page, edit).findall(q("Outline")) == []


def test_move_under_item():
    page = _page(_BODY)
    edit = edits.move_item(page, "b", under_id="a")
    outline = _xml(page, edit).find(q("Outline"))
    assert [oid for oid, _ in _texts(outline)] == ["a", "l1", "l2"]
    children = outline.find(f".//{q('OE')}[@objectID='a']/{q('OEChildren')}")
    assert [oe.get("objectID") for oe in children.findall(q("OE"))] == ["a1", "b"]


def test_move_before_within_outline_keeps_id():
    page = _page(_BODY)
    edit = edits.move_item(page, "l2", before_id="a")
    assert [oid for oid, _ in _texts(_xml(page, edit).find(q("Outline")))] == ["l2", "a", "b", "l1"]


def test_move_to_other_outline_sends_both_without_reusing_the_id():
    page = _page(_BODY)
    edit = edits.move_item(page, "b", after_id="c")
    o1, o2 = _xml(page, edit).findall(q("Outline"))
    assert "b" not in [oid for oid, _ in _texts(o1)]
    assert _texts(o2) == [("c", "gamma"), (None, "beta")]


def test_move_into_own_subtree_refused():
    with pytest.raises(OneNoteError, match="into itself"):
        edits.move_item(_page(_BODY), "a", under_id="a1")


def test_new_heading_reuses_page_style_index():
    defs = '<one:QuickStyleDef index="7" name="h2" fontSize="14.0"/><one:QuickStyleDef index="0" name="p"/>'
    page = _page(_BODY, defs)
    edit = edits.insert_items(page, [Paragraph(type="paragraph", text="H", style="h2")], after_id="b")
    xml = _xml(page, edit)
    new = [oe for oe in xml.iter(q("OE")) if oe.get("objectID") is None][0]
    assert new.get("quickStyleIndex") == "7"
    assert sorted(d.get("index") for d in xml.findall(q("QuickStyleDef"))) == ["0", "7"]


def test_new_heading_gets_free_index():
    defs = "".join(f'<one:QuickStyleDef index="{i}" name="s{i}"/>' for i in range(4))
    page = _page(_BODY, defs)
    edit = edits.insert_items(page, [Paragraph(type="paragraph", text="H", style="h2")], after_id="b")
    xml = _xml(page, edit)
    new = [oe for oe in xml.iter(q("OE")) if oe.get("objectID") is None][0]
    assert new.get("quickStyleIndex") == "4"
    assert [d.get("name") for d in xml.findall(q("QuickStyleDef"))] == ["s0", "s1", "s2", "s3", "h2"]


def test_existing_oes_keep_their_style_indexes():
    defs = '<one:QuickStyleDef index="0" name="h2"/><one:QuickStyleDef index="1" name="p"/>'
    body = _BODY.replace('<one:OE objectID="b">', '<one:OE objectID="b" quickStyleIndex="1">')
    page = _page(body, defs)
    edit = edits.insert_items(page, [Paragraph(type="paragraph", text="H", style="h1")], after_id="a")
    xml = _xml(page, edit)
    assert xml.find(f".//{q('OE')}[@objectID='b']").get("quickStyleIndex") == "1"
    new = [oe for oe in xml.iter(q("OE")) if oe.get("objectID") is None][0]
    assert new.get("quickStyleIndex") == "2"


def test_new_todo_reuses_tagdef():
    defs = '<one:TagDef index="3" type="0" symbol="3" name="Zu erledigen"/>'
    page = _page(_BODY, defs)
    edit = edits.insert_items(page, [Paragraph(type="paragraph", text="t", tag=Tag(kind="todo"))], after_id="b")
    xml = _xml(page, edit)
    assert [d.get("index") for d in xml.findall(q("TagDef"))] == ["3"]
    assert xml.find(f".//{q('Tag')}").get("index") == "3"


def test_tagdefs_come_before_quick_style_defs():
    page = _page(_BODY, '<one:QuickStyleDef index="0" name="p"/>')
    edit = edits.insert_items(page, [Paragraph(type="paragraph", text="t", tag=Tag(kind="question"))], after_id="b")
    names = [child.tag.split("}")[1] for child in _xml(page, edit)]
    assert names[:2] == ["TagDef", "QuickStyleDef"]


def test_outline_with_ink_cannot_be_resent():
    body = _BODY.replace('<one:OE objectID="b"><one:T>beta</one:T></one:OE>',
                         '<one:OE objectID="b"><one:T>beta</one:T></one:OE><one:OE objectID="ink"><one:InkParagraph/></one:OE>')
    with pytest.raises(OneNoteError) as exc:
        edits.update_item(_page(body), "a", Paragraph(type="paragraph", text="x"))
    assert exc.value.code == "bad_request"
    assert "ink ink" in str(exc.value) and "cannot be re-sent" in str(exc.value)


def test_deleting_the_ink_itself_is_allowed():
    body = _BODY.replace('<one:OE objectID="b"><one:T>beta</one:T></one:OE>',
                         '<one:OE objectID="ink"><one:InkParagraph/></one:OE>')
    edit = edits.delete_items(_page(body), ["ink"])
    assert edit.outlines


def test_update_protected_item_refused():
    body = _BODY.replace('<one:T>beta</one:T>', '<one:InsertedFile preferredName="a.pdf"/>')
    with pytest.raises(OneNoteError, match="cannot be changed"):
        edits.update_item(_page(body), "b", Paragraph(type="paragraph", text="x"))
