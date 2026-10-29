# CLAUDE.md

## Running commands

This repo uses a local `.venv`. Always call the venv's Python via its
**absolute path** (resolve it from the repo root) — bare `python`, bare
`pytest`, and relative forms like `.venv\Scripts\python.exe` or
`.\.venv\Scripts\python.exe` fail in this environment:

```
<repo-root>\.venv\Scripts\python.exe
```

### Common commands

| Task | Command (prefix with the absolute venv python) |
|------|------------------------------------------------|
| Unit tests | `-m pytest tests/unit/ -q` |
| E2E tests | `-m pytest -m e2e tests/e2e/ -v` |
| Collect tests | `-m pytest tests --collect-only` |
| Run a script | `scripts/<script>.py` |
| Run the server | `-m onenote_mcp.server` |

## Architecture

```
src/onenote_mcp/
  server.py     MCP tool definitions (FastMCP, stdio transport)
  com.py        COM wrapper: worker threads, timeouts, image-handle cache,
                notebook allowlist, page writes and targeted edits
  structure.py  notebooks, section groups, sections: create, rename, move,
                reorder, delete (COM impls on top of com.py)
  edits.py      targeted item edits on a page tree (pure)
  hierarchy.py  hierarchy XML: navigation tree, lookups, UpdateHierarchy XML (pure)
  builders.py   model <-> OneNote XML conversion
  models.py     Pydantic models = the JSON schema of the structured tools
  backups.py    local page backups and delete manifests
  pagestate.py  page revisions and in-flight writes (per server process)
  paths.py      ONENOTE_FILE_ROOTS check for files OneNote inserts
  config.py     ONENOTE_* environment settings
  errors.py     OneNoteError (code-prefixed)
  images.py     pixel size from image header bytes (proportional image scaling)
```

Layering: `server.py` calls `structure.py` and `com.py`; `structure.py` calls `com.py`;
`com.py` uses `edits.py`, `builders.py` and `models.py`, and also `backups.py`, `pagestate.py`, `paths.py`,
`hierarchy.py`, `config.py`. Every XML-exchanging COM call is pinned to the 2013
schema (`com._SchemaPinnedApp`). COM calls run on
fresh daemon threads with hard timeouts (`_run_com`); functions named `_*_impl`
run on the worker thread and must not call CoInitialize themselves.

## Tests

- Unit tests need no OneNote and must stay that way (fake `app` objects stand
  in for COM where needed).
- E2E tests drive the real MCP server against the real OneNote desktop app and
  create pages. They are restricted to the **ClaudeSpike** notebook (a throwaway
  one) and skip when it is not open or OneNote is unreachable; never point them
  at a real notebook.
