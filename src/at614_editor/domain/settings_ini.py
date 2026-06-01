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

# Label descrittive per l'editor UI (chiave → (label, tooltip))
FOLDER_KEY_LABELS: dict[str, tuple[str, str]] = {
    "FolderConfigurazioneBancoCollaudo": (
        "Distributori banco (DISTRIBUTORE)",
        "Cartella dei file .cfg dei distributori (es. \\\\server\\...\\AT614_BANCO_0001\\DISTRIBUTORE)",
    ),
    "FolderSettaggiProgramma": (
        "Settaggi programma (SETTAGGI PROGRAMMA)",
        "Cartella profili CE16/MMS2218 per ogni computer di collaudo",
    ),
    "FolderConfigurazioneTest": (
        "Sequenze test (TEST)",
        "Cartella dei file .csv delle sequenze di test",
    ),
    "FolderFileCurveLimite": (
        "Curve limite (CURVE_LIMITE)",
        "Cartella dei file .csv delle curve limite",
    ),
    "FolderRampeXY": (
        "Rampe XY (RAMPE_XY)",
        "Cartella delle rampe XY di riferimento",
    ),
    "FolderFileCurveComando": (
        "Curve comando (CURVE_COMANDO)",
        "Cartella dei file .csv delle curve di comando",
    ),
    "FolderConfigurazioneModuli": (
        "Parametri moduli (CONFIGURAZIONE)",
        "Cartella condivisa dei file .txt con i parametri da scrivere (può essere su altro share)",
    ),
    "FolderGraphSaved": (
        "Archivio grafici output (GRAPH)",
        "Cartella di rete dove il software VB6 salva i CSV acquisizione (può essere su share separato)",
    ),
    "FolderGraphSaved_LastACQ": (
        "Ultima acquisizione (RAMPE_XY_LAST)",
        "Cartella con l'ultima rampa XY acquisita",
    ),
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


def write_settings_ini(path: Path, mappings: dict[str, str]) -> None:
    """Riscrive il settings.ini mantenendo commenti e righe non-chiave esistenti.

    ``mappings`` deve contenere solo chiavi in ``_FOLDER_KEYS``; le righe
    esistenti con quelle chiavi vengono aggiornate; le chiavi non presenti
    come riga vengono aggiunte in coda.
    """
    # Leggi il testo esistente (se presente) per preservare i commenti
    existing_lines: list[str] = []
    try:
        existing_lines = path.read_text(encoding="cp1252").splitlines()
    except OSError:
        pass

    updated: set[str] = set()
    result_lines: list[str] = []

    for line in existing_lines:
        stripped = line.strip()
        if not stripped or stripped.startswith("'"):
            result_lines.append(line)
            continue
        key, sep, _value = stripped.partition("=")
        key = key.strip()
        if sep and key in mappings:
            result_lines.append(f"{key}={mappings[key]}")
            updated.add(key)
        else:
            result_lines.append(line)

    # Aggiungi le chiavi nuove non ancora presenti nel file
    for key, value in mappings.items():
        if key not in updated:
            result_lines.append(f"{key}={value}")

    path.write_text("\r\n".join(result_lines) + "\r\n", encoding="cp1252")


def discover_root_from_settings_ini(settings_path: Path) -> Path | None:
    """Restituisce la project root leggendo un settings.ini AT614.

    La root è il parent comune delle cartelle DISTRIBUTORE e TEST definite
    nel file.  Restituisce ``None`` solo se le due chiavi obbligatorie
    mancano o i loro percorsi hanno parent diversi (banco diverso).
    Non richiede che i percorsi siano fisicamente raggiungibili.
    """
    mappings = parse_settings_ini(settings_path)
    dist_path = mappings.get(_KEY_DISTRIBUTORE)
    test_path = mappings.get(_KEY_TEST)

    if dist_path is None or test_path is None:
        return None

    if dist_path.parent != test_path.parent:
        return None

    return dist_path.parent
