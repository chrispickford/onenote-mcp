"""Environment settings of the server (the ONENOTE_* variables).

Each function takes the environment mapping explicitly, so callers pass
os.environ and tests pass plain dicts.  The allowlist and the timeouts keep
living in com.py, which reads them at call time.
"""

from collections.abc import Mapping
from pathlib import Path

_TRUE = ("1", "true")
_FALSE = ("", "0", "false")


def _flag(environ: Mapping[str, str], name: str) -> bool:
    """Parse a 1/true/0/false/empty flag (case-insensitive); unset is false."""
    raw = environ.get(name, "")
    value = raw.strip().lower()
    if value in _TRUE:
        return True
    if value in _FALSE:
        return False
    raise ValueError(f"{name} must be 1, true, 0, false or empty, got {raw!r}")


def _positive_float(environ: Mapping[str, str], name: str, default: float) -> float:
    """Read a positive float, falling back to *default* when unset or invalid."""
    try:
        value = float(environ[name])
    except (KeyError, ValueError):
        return default
    return value if value > 0 else default


def raw_xml_enabled(environ: Mapping[str, str]) -> bool:
    """Whether the four raw-XML tools are registered.

    Off unless ONENOTE_ENABLE_RAW_XML is 1/true; ONENOTE_DISABLE_RAW_XML=1
    (the upstream setting) forces them off and wins.  Raises ValueError for
    any other value of either variable.
    """
    enabled = _flag(environ, "ONENOTE_ENABLE_RAW_XML")
    disabled = _flag(environ, "ONENOTE_DISABLE_RAW_XML")
    return enabled and not disabled


def file_roots(environ: Mapping[str, str]) -> list[str] | None:
    """Folders files may be inserted from (ONENOTE_FILE_ROOTS, ';'-separated).

    None when unset, which disables inserting files from a path.
    """
    raw = environ.get("ONENOTE_FILE_ROOTS")
    if raw is None:
        return None
    return [part.strip() for part in raw.split(";") if part.strip()]


def backup_dir(environ: Mapping[str, str]) -> Path:
    """Folder for page backups (ONENOTE_BACKUP_DIR, default %LOCALAPPDATA%\\onenote-mcp\\backups)."""
    override = environ.get("ONENOTE_BACKUP_DIR", "").strip()
    if override:
        return Path(override)
    return Path(environ.get("LOCALAPPDATA", str(Path.home()))) / "onenote-mcp" / "backups"


def backup_retention_days(environ: Mapping[str, str]) -> float:
    """Days backups are kept (ONENOTE_BACKUP_RETENTION_DAYS, default 30)."""
    return _positive_float(environ, "ONENOTE_BACKUP_RETENTION_DAYS", 30.0)


def abandon_write_after(environ: Mapping[str, str]) -> float:
    """Seconds after which a stuck write stops blocking its page (ONENOTE_ABANDON_WRITE_AFTER, default 300)."""
    return _positive_float(environ, "ONENOTE_ABANDON_WRITE_AFTER", 300.0)
