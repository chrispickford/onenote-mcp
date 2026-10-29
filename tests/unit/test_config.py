"""Unit tests for config.py: parsing of the ONENOTE_* environment settings."""
import os

import pytest

from onenote_mcp import config


def test_raw_xml_off_by_default():
    assert config.raw_xml_enabled({}) is False


@pytest.mark.parametrize("value", ["1", "true", "TRUE", " true "])
def test_enable_raw_xml(value):
    assert config.raw_xml_enabled({"ONENOTE_ENABLE_RAW_XML": value}) is True


@pytest.mark.parametrize("value", ["", "0", "false"])
def test_enable_raw_xml_off_values(value):
    assert config.raw_xml_enabled({"ONENOTE_ENABLE_RAW_XML": value}) is False


def test_disable_wins():
    environ = {"ONENOTE_ENABLE_RAW_XML": "1", "ONENOTE_DISABLE_RAW_XML": "1"}
    assert config.raw_xml_enabled(environ) is False


@pytest.mark.parametrize("name", ["ONENOTE_ENABLE_RAW_XML", "ONENOTE_DISABLE_RAW_XML"])
def test_bad_value_raises(name):
    with pytest.raises(ValueError) as exc:
        config.raw_xml_enabled({name: "yes"})
    assert name in str(exc.value)
    assert "'yes'" in str(exc.value)


def test_file_roots_split():
    assert config.file_roots({"ONENOTE_FILE_ROOTS": "C:\\a; D:\\b ;"}) == ["C:\\a", "D:\\b"]


def test_file_roots_unset_is_none():
    assert config.file_roots({}) is None


def test_backup_defaults():
    environ = {"LOCALAPPDATA": "C:\\Users\\x\\AppData\\Local"}
    assert config.backup_retention_days(environ) == 30.0
    assert config.abandon_write_after(environ) == 300.0
    assert str(config.backup_dir(environ)) == os.path.join(
        "C:\\Users\\x\\AppData\\Local", "onenote-mcp", "backups"
    )


def test_backup_dir_override():
    assert str(config.backup_dir({"ONENOTE_BACKUP_DIR": "D:\\bk"})) == "D:\\bk"


@pytest.mark.parametrize("value", ["abc", "-1", "0"])
def test_bad_numbers_fall_back_to_default(value):
    assert config.backup_retention_days({"ONENOTE_BACKUP_RETENTION_DAYS": value}) == 30.0
    assert config.abandon_write_after({"ONENOTE_ABANDON_WRITE_AFTER": value}) == 300.0
