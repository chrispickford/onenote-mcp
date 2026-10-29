"""Unit tests for pagestate.py: page revision tokens and the per-page memory
of the revision a client last read.

OneNote's own dateExpectedLastModified check accepts stale writes (verified
against 16.0.20527), so the server compares a revision token of its own.
"""
import pytest

from onenote_mcp import pagestate
from onenote_mcp.com import OneNoteError

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"


def _page(text: str, path_cache: str = "C:\\T\\{A}.bin", callback: str = "{CB-1}") -> str:
    return (
        f'<one:Page xmlns:one="{_ONE_NS}" ID="{{P}}" lastModifiedTime="2026-10-29T12:00:00.000Z">'
        f'<one:Outline><one:OEChildren><one:OE><one:T>{text}</one:T></one:OE>'
        f'<one:OE><one:Image><one:CallbackID callbackID="{callback}"/></one:Image></one:OE>'
        f'<one:OE><one:InsertedFile pathCache="{path_cache}" pathSource="D:\\a.txt" preferredName="a.txt"/></one:OE>'
        '</one:OEChildren></one:Outline></one:Page>'
    )


@pytest.fixture(autouse=True)
def _clean_state():
    pagestate.forget_all()
    yield
    pagestate.forget_all()


def test_revision_ignores_volatile_attributes():
    first = pagestate.revision(_page("a", path_cache="C:\\T\\{A}.bin", callback="{CB-1}"))
    second = pagestate.revision(_page("a", path_cache="C:\\T\\{B}.bin", callback="{CB-2}"))
    assert first == second


def test_revision_changes_with_content():
    assert pagestate.revision(_page("a")) != pagestate.revision(_page("b"))


def test_revision_is_short_hex():
    token = pagestate.revision(_page("a"))
    assert len(token) == 16 and int(token, 16) >= 0


def test_expected_prefers_given():
    pagestate.remember("{P}", "remembered")
    assert pagestate.expected("{P}", "given", required=True) == "given"


def test_expected_uses_remembered():
    pagestate.remember("{P}", "remembered")
    assert pagestate.expected("{P}", None, required=True) == "remembered"


def test_expected_required_without_read_is_bad_request():
    with pytest.raises(OneNoteError) as exc:
        pagestate.expected("{P}", None, required=True)
    assert exc.value.code == "bad_request"
    assert "Read the page with get_page first" in str(exc.value)


def test_expected_not_required_returns_none():
    assert pagestate.expected("{P}", None, required=False) is None


def test_check_raises_conflict_on_mismatch():
    with pytest.raises(OneNoteError) as exc:
        pagestate.check("{P}", _page("now"), pagestate.revision(_page("then")))
    assert exc.value.code == "conflict"
    assert "changed since it was read" in str(exc.value)


def test_check_passes_on_match_and_without_expectation():
    pagestate.check("{P}", _page("a"), pagestate.revision(_page("a")))
    pagestate.check("{P}", _page("a"), None)


def test_after_write_remembers_new_revision():
    pagestate.after_write("{P}", before_xml=_page("a"), after_xml=_page("b"), checked=True)
    assert pagestate.remembered("{P}") == pagestate.revision(_page("b"))


def test_unchecked_write_over_unseen_change_forgets_page():
    # The client read "a"; someone changed it to "x"; an append (no check) then
    # wrote on top.  The client never saw "x", so the next checked write must
    # re-read rather than trust the post-append revision.
    pagestate.remember("{P}", pagestate.revision(_page("a")))
    pagestate.after_write("{P}", before_xml=_page("x"), after_xml=_page("x+appended"), checked=False)
    assert pagestate.remembered("{P}") is None


def test_unchecked_write_to_unread_page_remembers_nothing():
    # Appending never shows the client the page, so a later replace still
    # needs a read first.
    pagestate.after_write("{P}", before_xml=_page("a"), after_xml=_page("a+appended"), checked=False)
    assert pagestate.remembered("{P}") is None


def test_unchecked_write_over_seen_state_remembers():
    pagestate.remember("{P}", pagestate.revision(_page("a")))
    pagestate.after_write("{P}", before_xml=_page("a"), after_xml=_page("a+appended"), checked=False)
    assert pagestate.remembered("{P}") == pagestate.revision(_page("a+appended"))


def test_revision_ignores_selection_state():
    plain = _page("a")
    selected = plain.replace("<one:OE><one:T>a</one:T>", '<one:OE selected="partial"><one:T selected="all">a</one:T>')
    assert pagestate.revision(plain) == pagestate.revision(selected)
