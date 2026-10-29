"""Per-page state the server keeps for the session: page revisions.

A revision is a short token for a page's content.  get_page returns it, the
server remembers the revision each client read last, and writes that change
or remove existing content refuse to run when the page no longer has the
expected revision, so edits the user made in OneNote since the read are not
overwritten.  OneNote's own dateExpectedLastModified check is not used: it
accepts stale writes (verified against OneNote 16.0.20527).
"""

import hashlib
import re
import threading

from onenote_mcp.errors import OneNoteError

# Attributes whose value changes between two reads of an unchanged page:
# InsertedFile@pathCache names a fresh temp copy, and image callback IDs are
# per read.
_VOLATILE_ATTRS = re.compile(r'\s(?:pathCache|callbackID)="[^"]*"')

_REVISIONS: dict[str, str] = {}
_LOCK = threading.Lock()


def revision(page_xml: str) -> str:
    """Return the revision token (16 hex characters) of a basic page read."""
    stable = _VOLATILE_ATTRS.sub("", page_xml)
    return hashlib.sha1(stable.encode("utf-8")).hexdigest()[:16]


def remember(page_id: str, token: str) -> None:
    """Record *token* as the revision the client last saw for *page_id*."""
    with _LOCK:
        _REVISIONS[page_id] = token


def remembered(page_id: str) -> str | None:
    """Return the revision last recorded for *page_id*, or None."""
    with _LOCK:
        return _REVISIONS.get(page_id)


def forget(page_id: str) -> None:
    with _LOCK:
        _REVISIONS.pop(page_id, None)


def forget_all() -> None:
    with _LOCK:
        _REVISIONS.clear()


def expected(page_id: str, given: str | None, *, required: bool) -> str | None:
    """Return the revision a write to *page_id* must find.

    *given* (the caller's expected_revision) wins; otherwise the remembered
    revision.  With neither, a *required* check raises bad_request, and an
    optional one returns None (no check).
    """
    if given is not None:
        return given
    token = remembered(page_id)
    if token is None and required:
        raise OneNoteError(
            f"Read the page with get_page first: page {page_id!r} has not been read in "
            "this session, so the server cannot tell whether it changed since. "
            "Alternatively pass expected_revision from an earlier get_page.",
            code="bad_request",
        )
    return token


def check(page_id: str, current_xml: str, expected_token: str | None) -> None:
    """Raise conflict unless the page (read as *current_xml*) still has *expected_token*.

    No-op when *expected_token* is None.
    """
    if expected_token is None or revision(current_xml) == expected_token:
        return
    raise OneNoteError(
        f"Page {page_id!r} changed since it was read; re-read it with get_page and retry.",
        code="conflict",
    )


def after_write(page_id: str, *, before_xml: str, after_xml: str, checked: bool) -> None:
    """Update the remembered revision after the server wrote to *page_id*.

    A checked write proved the client saw *before_xml*, so the new revision is
    remembered.  An unchecked write (append, insert) remembers it only when
    the page still had the remembered revision before the write; otherwise
    (never read, or changed since the read) the client has not seen some
    content now on the page, so the page is forgotten and the next checked
    write needs a fresh read.
    """
    token = remembered(page_id)
    if checked or (token is not None and token == revision(before_xml)):
        remember(page_id, revision(after_xml))
    else:
        forget(page_id)
