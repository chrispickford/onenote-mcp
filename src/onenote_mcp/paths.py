"""Which local files OneNote may be asked to insert (images and attachments).

The server never opens these files itself: OneNote reads them.  But inserting
a file copies it into a notebook that may sync to the cloud, so paths must lie
inside the folders listed in ONENOTE_FILE_ROOTS; unset, inserting from a path
is disabled.
"""

import os

from onenote_mcp.errors import OneNoteError


def _within(path: str, root: str) -> bool:
    """Whether *path* is *root* or inside it (both absolute, normalised)."""
    try:
        return os.path.commonpath([path, root]) == root
    except ValueError:  # different drives
        return False


def check_insert_path(path: str, roots: list[str] | None, *, extra_roots: list[str] | tuple = ()) -> str:
    """Return *path* resolved to an absolute path if OneNote may insert it.

    *roots* are the configured ONENOTE_FILE_ROOTS (None or empty: disabled);
    *extra_roots* are allowed in addition for internal callers, e.g. the
    backup folder for restore_page.  Symlinks and '..' are resolved before the
    comparison, which ignores case.  Raises bad_request otherwise, and when
    the file does not exist.
    """
    allowed_roots = [*(roots or []), *extra_roots]
    if not allowed_roots:
        raise OneNoteError(
            f"Inserting files is disabled: set ONENOTE_FILE_ROOTS to the folders files may be "
            f"inserted from (refused {path!r}).",
            code="bad_request",
        )
    if not os.path.isabs(path):
        raise OneNoteError(f"File path {path!r} must be an absolute path.", code="bad_request")
    resolved = os.path.realpath(path)
    candidate = os.path.normcase(resolved)
    if not any(_within(candidate, os.path.normcase(os.path.realpath(root))) for root in allowed_roots):
        raise OneNoteError(
            f"File {path!r} is outside ONENOTE_FILE_ROOTS, so OneNote may not insert it.",
            code="bad_request",
        )
    if not os.path.isfile(resolved):
        raise OneNoteError(f"File {path!r} does not exist.", code="bad_request")
    return resolved
