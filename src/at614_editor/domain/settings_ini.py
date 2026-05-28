"""Parser per il file settings.ini del banco AT614 (software VB6).

Il file usa la forma ``Chiave=Valore`` con righe di commento che iniziano
con ``'``.  Non ha intestazioni di sezione (non è un INI standard).
"""
from __future__ import annotations

from pathlib import Path


# Chiavi del settings.ini → nome cartella atteso sotto la project root
_FOLDER_KEYS: dict[str, str] = {
    "FolderConfigurazioneBancoCollaudo": "DISTRIBUTORE",
    "FolderSettaggiProgramma": "SETTAGGI PROGRAMMA",
    "FolderConfigurazioneTest": "TEST",
    "FolderFileCurveLimite": "CURVE_LIMITE",
    "FolderRampeXY": "RAMPE_XY",
    "FolderFileCurveComando": "CURVE_COMANDO",
    "FolderConfigurazioneModuli": "CONFIGURAZIONE",
    "FolderGraphSaved": "GRAPH",
    "FolderGraphSaved_LastACQ": "RAMPE_XY_LAST",
}

# Le due chiavi obbligatorie per determinare la project root
_KEY_DISTRIBUTORE = "FolderConfigurazioneBancoCollaudo"
_KEY_TEST = "FolderConfigurazioneTest"


def parse_settings_ini(path: Path) -> dict[str, Path]:
    """Legge il settings.ini e restituisce ``{chiave: percorso_assoluto}``."""
    result: dict[str, Path] = {}
    try:
        text = path.read_text(encoding="cp1252")
    except OSError:
        return result

    for line in text.splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("'"):
            continue
        key, sep, value = stripped.partition("=")
        if not sep:
            continue
        key = key.strip()
        value = value.strip()
        if key in _FOLDER_KEYS and value:
            result[key] = Path(value)

    return result


def discover_root_from_settings_ini(settings_path: Path) -> Path | None:
    """Restituisce la project root leggendo un settings.ini AT614.

    La root è il parent comune delle cartelle DISTRIBUTORE e TEST definite
    nel file.  Se le due cartelle non esistono o hanno parent diversi,
    restituisce ``None``.
    """
    mappings = parse_settings_ini(settings_path)
    dist_path = mappings.get(_KEY_DISTRIBUTORE)
    test_path = mappings.get(_KEY_TEST)

    if dist_path is None or test_path is None:
        return None

    if dist_path.parent != test_path.parent:
        return None

    root = dist_path.parent
    if (root / "DISTRIBUTORE").exists() and (root / "TEST").exists():
        return root

    return None
