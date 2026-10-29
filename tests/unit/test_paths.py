"""Unit tests for paths.py: files may only be inserted from ONENOTE_FILE_ROOTS."""
import os

import pytest

from onenote_mcp.errors import OneNoteError
from onenote_mcp.paths import check_insert_path


@pytest.fixture
def root(tmp_path):
    folder = tmp_path / "allowed"
    folder.mkdir()
    (folder / "a.png").write_bytes(b"x")
    (tmp_path / "outside.png").write_bytes(b"x")
    return folder


def _refused(path, roots, **kwargs) -> str:
    with pytest.raises(OneNoteError) as exc:
        check_insert_path(str(path), roots, **kwargs)
    assert exc.value.code == "bad_request"
    return str(exc.value)


def test_path_inside_root_is_returned_absolute(root):
    assert check_insert_path(str(root / "a.png"), [str(root)]) == os.path.realpath(root / "a.png")


def test_path_outside_roots_refused(root, tmp_path):
    assert "is outside ONENOTE_FILE_ROOTS" in _refused(tmp_path / "outside.png", [str(root)])


def test_dotdot_escape_refused(root):
    assert "is outside ONENOTE_FILE_ROOTS" in _refused(f"{root}\\..\\outside.png", [str(root)])


def test_sibling_folder_with_same_prefix_refused(root, tmp_path):
    sibling = tmp_path / "allowed-not"
    sibling.mkdir()
    (sibling / "a.png").write_bytes(b"x")
    assert "is outside ONENOTE_FILE_ROOTS" in _refused(sibling / "a.png", [str(root)])


def test_case_insensitive_root_match(root):
    assert check_insert_path(str(root / "a.png"), [str(root).upper()])


def test_other_drive_refused(root):
    other = "Z:\\x.png" if not str(root).upper().startswith("Z:") else "Y:\\x.png"
    assert "is outside ONENOTE_FILE_ROOTS" in _refused(other, [str(root)])


def test_roots_unset_disables(root):
    assert "Inserting files is disabled: set ONENOTE_FILE_ROOTS" in _refused(root / "a.png", None)


def test_empty_roots_disable(root):
    assert "Inserting files is disabled" in _refused(root / "a.png", [])


def test_missing_file_refused(root):
    assert "does not exist" in _refused(root / "nope.png", [str(root)])


def test_extra_roots_allow_without_configured_roots(root):
    assert check_insert_path(str(root / "a.png"), None, extra_roots=[str(root)])


def test_relative_path_refused(root):
    assert "must be an absolute path" in _refused("a.png", [str(root)])
