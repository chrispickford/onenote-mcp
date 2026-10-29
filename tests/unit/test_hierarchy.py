"""Unit tests for hierarchy.py: the notebook tree get_notebooks returns, and
lookups in hierarchy XML."""
import xml.etree.ElementTree as ET

from onenote_mcp import hierarchy

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"

_TREE = (
    f'<one:Notebooks xmlns:one="{_ONE_NS}">'
    '<one:Notebook ID="{NB}" name="Folder" nickname="Shown">'
    '<one:Section ID="{S1}" name="Top" color="#F6B078"/>'
    '<one:SectionGroup ID="{G1}" name="01_Background">'
    '<one:Section ID="{S2}" name="General"><one:Page ID="{P}" name="Brief"/></one:Section>'
    '<one:SectionGroup ID="{G2}" name="Nested"><one:Section ID="{S3}" name="Deep"/></one:SectionGroup>'
    '</one:SectionGroup>'
    '<one:SectionGroup ID="{RB}" name="OneNote_RecycleBin" isRecycleBin="true">'
    '<one:Section ID="{S9}" name="Deleted Pages" isInRecycleBin="true"/></one:SectionGroup>'
    '<one:Section ID="{S8}" name="Binned" isInRecycleBin="true"/>'
    '</one:Notebook></one:Notebooks>'
)


def test_sections_in_groups_listed_under_the_group():
    (nb,) = hierarchy.parse_notebook_tree(_TREE)
    assert nb["id"] == "{NB}" and nb["name"] == "Shown"
    assert nb["sections"] == [{"id": "{S1}", "name": "Top", "color": "#F6B078"}]
    (group,) = nb["section_groups"]
    assert group["id"] == "{G1}" and group["name"] == "01_Background"
    assert group["sections"] == [{"id": "{S2}", "name": "General"}]


def test_nested_groups():
    (nb,) = hierarchy.parse_notebook_tree(_TREE)
    (nested,) = nb["section_groups"][0]["section_groups"]
    assert nested == {"id": "{G2}", "name": "Nested", "sections": [{"id": "{S3}", "name": "Deep"}]}


def test_recycle_bin_hidden():
    text = str(hierarchy.parse_notebook_tree(_TREE))
    assert "RecycleBin" not in text and "Binned" not in text and "{S9}" not in text


def test_notebook_without_groups_has_no_section_groups_key():
    xml = f'<one:Notebooks xmlns:one="{_ONE_NS}"><one:Notebook ID="{{N}}" name="N"><one:Section ID="{{S}}" name="S"/></one:Notebook></one:Notebooks>'
    assert hierarchy.parse_notebook_tree(xml) == [{"id": "{N}", "name": "N", "sections": [{"id": "{S}", "name": "S"}]}]


def test_find_object_and_notebook_of():
    root = ET.fromstring(_TREE)
    page = hierarchy.find_object(root, "{P}")
    assert page.get("name") == "Brief"
    assert hierarchy.notebook_of(root, page).get("ID") == "{NB}"
    assert hierarchy.find_object(root, "{NOPE}") is None
