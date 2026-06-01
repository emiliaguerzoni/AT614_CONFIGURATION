"""Test per il dialogo di editing del settings.ini."""
from pathlib import Path

import pytest

from at614_editor.domain.settings_ini import (
    FOLDER_KEY_LABELS,
    _FOLDER_KEYS,
    parse_settings_ini,
    write_settings_ini,
)
from at614_editor.ui.dialogs.settings_ini_dialog import SettingsIniEditorDialog


# ------------------------------------------------------------------ helpers

def _make_ini(tmp_path: Path, content: str) -> Path:
    ini = tmp_path / "settings.ini"
    ini.write_text(content, encoding="cp1252")
    return ini


# ------------------------------------------------------------------ domain

def test_write_settings_ini_updates_existing_key(tmp_path: Path) -> None:
    ini = _make_ini(
        tmp_path,
        "' commento\r\nFolderConfigurazioneTest=C:\\vecchio\\TEST\r\n",
    )
    write_settings_ini(ini, {"FolderConfigurazioneTest": "C:\\nuovo\\TEST"})
    result = parse_settings_ini(ini)
    assert str(result["FolderConfigurazioneTest"]) == "C:\\nuovo\\TEST"


def test_write_settings_ini_adds_missing_key(tmp_path: Path) -> None:
    ini = _make_ini(tmp_path, "FolderConfigurazioneBancoCollaudo=C:\\DIST\r\n")
    write_settings_ini(ini, {"FolderGraphSaved": "\\\\server\\GRAPH"})
    result = parse_settings_ini(ini)
    assert str(result["FolderGraphSaved"]).rstrip("\\") == "\\\\server\\GRAPH"
    # La chiave preesistente non deve essere rimossa
    assert "FolderConfigurazioneBancoCollaudo" in result


def test_write_settings_ini_preserves_comments(tmp_path: Path) -> None:
    ini = _make_ini(
        tmp_path,
        "' File dei settaggi\r\nFolderConfigurazioneBancoCollaudo=C:\\DIST\r\n",
    )
    write_settings_ini(ini, {"FolderConfigurazioneBancoCollaudo": "C:\\NUOVA\\DIST"})
    text = ini.read_text(encoding="cp1252")
    assert "' File dei settaggi" in text


def test_write_settings_ini_empty_value_skips_key(tmp_path: Path) -> None:
    ini = _make_ini(tmp_path, "FolderConfigurazioneBancoCollaudo=C:\\DIST\r\n")
    # Passare una stringa vuota non deve sovrascrivere
    write_settings_ini(ini, {})
    result = parse_settings_ini(ini)
    assert "FolderConfigurazioneBancoCollaudo" in result


# ------------------------------------------------------------------ UI

def test_settings_ini_editor_shows_all_keys(tmp_path: Path, qt_app) -> None:
    ini = _make_ini(
        tmp_path,
        "FolderConfigurazioneBancoCollaudo=C:\\DIST\r\nFolderConfigurazioneTest=C:\\TEST\r\n",
    )
    dialog = SettingsIniEditorDialog(ini_path=ini)
    assert set(dialog._fields.keys()) == set(_FOLDER_KEYS.keys())


def test_settings_ini_editor_populates_existing_values(tmp_path: Path, qt_app) -> None:
    ini = _make_ini(
        tmp_path,
        "FolderConfigurazioneBancoCollaudo=C:\\DIST\r\nFolderConfigurazioneTest=C:\\TEST\r\n",
    )
    dialog = SettingsIniEditorDialog(ini_path=ini)
    assert dialog._fields["FolderConfigurazioneBancoCollaudo"].text() == "C:\\DIST"
    assert dialog._fields["FolderConfigurazioneTest"].text() == "C:\\TEST"


def test_settings_ini_editor_save_writes_file(tmp_path: Path, qt_app) -> None:
    ini = _make_ini(tmp_path, "FolderConfigurazioneBancoCollaudo=C:\\OLD\r\n")
    dialog = SettingsIniEditorDialog(ini_path=ini)
    dialog._fields["FolderConfigurazioneBancoCollaudo"].setText("C:\\NEW\\DIST")
    dialog._on_save()
    result = parse_settings_ini(ini)
    assert str(result["FolderConfigurazioneBancoCollaudo"]) == "C:\\NEW\\DIST"


def test_settings_ini_editor_current_values(tmp_path: Path, qt_app) -> None:
    ini = _make_ini(
        tmp_path,
        "FolderGraphSaved=\\\\server\\GRAPH\r\n",
    )
    dialog = SettingsIniEditorDialog(ini_path=ini)
    values = dialog.current_values()
    assert values["FolderGraphSaved"].rstrip("\\") == "\\\\server\\GRAPH"


def test_folder_key_labels_covers_all_keys() -> None:
    assert set(FOLDER_KEY_LABELS.keys()) == set(_FOLDER_KEYS.keys())
