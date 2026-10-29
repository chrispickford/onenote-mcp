"""Shared fixtures for E2E tests — all require OneNote desktop running.

Every e2e test, in-process or through the spawned server, is restricted to the
ClaudeSpike notebook by the allowlist set below, so a test can never write to
a real notebook.
"""
import os
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent.parent / "src"))

from onenote_mcp import com
from tests.e2e.mcp_client import mcp_session  # noqa: F401 — re-export for all e2e files

_ONE_NS = "http://schemas.microsoft.com/office/onenote/2013/onenote"
_NS = {"one": _ONE_NS}

TEST_NOTEBOOK = "ClaudeSpike"
os.environ["ONENOTE_ALLOWED_NOTEBOOKS"] = TEST_NOTEBOOK
# The raw XML tools are off by default; their e2e tests need them.
os.environ["ONENOTE_ENABLE_RAW_XML"] = "1"
# Backups of test pages stay next to the test notebook, not in %LOCALAPPDATA%.
os.environ["ONENOTE_BACKUP_DIR"] = r"D:\scratch\onenote-mcp-tests\backups"

# ---------------------------------------------------------------------------
# OneNote availability check (done once at session start)
# ---------------------------------------------------------------------------

def _check_onenote() -> tuple[bool, str]:
    """Return (available, reason_if_not)."""
    try:
        xml = com.list_hierarchy()
        root = ET.fromstring(xml)
        nbs = root.findall(f"{{{_ONE_NS}}}Notebook")
        if not nbs:
            return False, "No notebooks found — open at least one notebook in OneNote desktop."
        return True, ""
    except Exception as exc:
        return False, f"OneNote not reachable: {exc}"

_ONENOTE_AVAILABLE, _ONENOTE_REASON = _check_onenote()

# ---------------------------------------------------------------------------
# Session-scoped fixtures
# ---------------------------------------------------------------------------

@pytest.fixture(scope="session")
def onenote_available():
    """Skip the test if OneNote is not reachable."""
    if not _ONENOTE_AVAILABLE:
        pytest.skip(f"OneNote not available: {_ONENOTE_REASON}")


@pytest.fixture(scope="session")
def claudespike_section_id(onenote_available):
    """Return the ID of the first section in the ClaudeSpike notebook.

    Skips when ClaudeSpike is not open; never falls back to another notebook.
    Creates a "Tests" section when the notebook has none.
    """
    root = ET.fromstring(com.list_hierarchy())
    nb = next((nb for nb in root.findall(f"{{{_ONE_NS}}}Notebook") if nb.get("name") == TEST_NOTEBOOK), None)
    if nb is None:
        pytest.skip(f"{TEST_NOTEBOOK} notebook not open")
    sections = nb.findall(f".//{{{_ONE_NS}}}Section")
    if sections:
        return sections[0].get("ID")
    # OpenHierarchy(path, relative_to_id, cftSection=3) creates the section.
    return com._run_com(lambda: com._app().OpenHierarchy("Tests.one", nb.get("ID"), 3))
