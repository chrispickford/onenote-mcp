"""Unit-test defaults: backups go to a per-test temporary folder, never to the
real %LOCALAPPDATA%\onenote-mcp\backups."""
import pytest


@pytest.fixture(autouse=True)
def _backup_dir(tmp_path, monkeypatch):
    monkeypatch.setenv("ONENOTE_BACKUP_DIR", str(tmp_path / "backups"))
    return tmp_path / "backups"
