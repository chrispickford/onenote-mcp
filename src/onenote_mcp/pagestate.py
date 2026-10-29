"""Per-page state the server keeps for the session: page revisions.

A revision is a short token for a page's content.  get_page returns it, the
server remembers the revision each client read last, and writes that change
or remove existing content refuse to run when the page no longer has the
expected revision, so edits the user made in OneNote since the read are not
overwritten.  OneNote's own dateExpectedLastModified check is not used: it
accepts stale writes (verified against OneNote 16.0.20527).
"""

import hashlib
import os
import re
import threading
import time

from onenote_mcp import config
from onenote_mcp.errors import OneNoteError

# Attributes whose value changes between two reads of an unchanged page:
# InsertedFile@pathCache names a fresh temp copy, image callback IDs are per
# read, and selected="all|partial" marks where the user's cursor is.
_VOLATILE_ATTRS = re.compile(r'\s(?:pathCache|callbackID|selected)="[^"]*"')

_REVISIONS: dict[str, str] = {}
_LOCK = threading.RLock()


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


# ---- In-flight writes ---------------------------------------------------------
#
# A COM call that times out keeps running on its abandoned worker thread and
# may still complete.  A write to that page until then could interleave with
# it or apply twice, so it is refused with 'busy'.  A stuck write stops
# blocking once its thread ends, or after config.abandon_write_after seconds.

_WRITES: dict[str, tuple[threading.Thread, float]] = {}


def _blocking(page_id: str, abandon_after: float) -> bool:
    entry = _WRITES.get(page_id)
    if entry is None:
        return False
    thread, started = entry
    if thread.is_alive() and time.monotonic() - started < abandon_after:
        return True
    del _WRITES[page_id]
    return False


def begin_write(page_id: str, thread: threading.Thread, abandon_after: float) -> None:
    """Register *thread* as the write to *page_id*; busy if another still runs."""
    with _LOCK:
        if _blocking(page_id, abandon_after):
            raise OneNoteError(
                f"A previous write to this page is still running (page {page_id!r}); it timed "
                "out but may still complete. Wait, re-read the page with get_page, then retry.",
                code="busy",
            )
        _WRITES[page_id] = (thread, time.monotonic())


def end_write(page_id: str, thread: threading.Thread) -> None:
    """Forget the write to *page_id* if *thread* is still the registered one."""
    with _LOCK:
        if _WRITES.get(page_id, (None,))[0] is thread:
            del _WRITES[page_id]


def writes_in_flight(abandon_after: float | None = None) -> list[str]:
    """Return the pages whose timed-out write is still running and blocking."""
    if abandon_after is None:
        abandon_after = config.abandon_write_after(os.environ)
    with _LOCK:
        return sorted(page_id for page_id in list(_WRITES) if _blocking(page_id, abandon_after))


def clear_writes() -> None:
    with _LOCK:
        _WRITES.clear()
