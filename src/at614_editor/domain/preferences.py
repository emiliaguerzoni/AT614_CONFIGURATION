"""Persistenza delle preferenze utente dell'applicazione.

Le preferenze vengono salvate in ``~/.at614-editor/preferences.json``.
"""
from __future__ import annotations

import json
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

_PREFS_FILE = Path.home() / ".at614-editor" / "preferences.json"
_KEY_SETTINGS_INI = "settings_ini_path"


def _load_raw() -> dict:
    try:
        if _PREFS_FILE.is_file():
            return json.loads(_PREFS_FILE.read_text(encoding="utf-8"))
    except Exception as exc:
        logger.warning(f"Preferences: impossibile leggere {_PREFS_FILE}: {exc}")
    return {}


def _save_raw(data: dict) -> None:
    try:
        _PREFS_FILE.parent.mkdir(parents=True, exist_ok=True)
        _PREFS_FILE.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    except Exception as exc:
        logger.warning(f"Preferences: impossibile salvare {_PREFS_FILE}: {exc}")


def load_settings_ini_path() -> Path | None:
    """Restituisce il percorso dell'ultimo settings.ini usato, o None."""
    raw = _load_raw()
    value = raw.get(_KEY_SETTINGS_INI)
    if value:
        p = Path(value)
        if p.is_file():
            logger.debug(f"Preferences: settings.ini salvato trovato in {p}")
            return p
        logger.debug(f"Preferences: settings.ini salvato non trovato più in {p}")
    return None


def save_settings_ini_path(path: Path) -> None:
    """Salva il percorso del settings.ini scelto."""
    raw = _load_raw()
    raw[_KEY_SETTINGS_INI] = str(path.resolve())
    _save_raw(raw)
    logger.info(f"Preferences: settings.ini salvato → {path.resolve()}")
