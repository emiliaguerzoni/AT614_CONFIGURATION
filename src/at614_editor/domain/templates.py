from __future__ import annotations

from pathlib import Path


CATEGORY_DIRECTORIES = {
    "distributori": "DISTRIBUTORE",
    "sequenze_test": "TEST",
    "curve_comando": "CURVE_COMANDO",
    "curve_limite": "CURVE_LIMITE",
    "rampe_xy": "RAMPE_XY",
    "file_esterni": "TEST",
}

SETTINGS_PROGRAM_DIR = "SETTAGGI PROGRAMMA"
DEFAULT_SETTINGS_PROFILE = "DEFAULT"

DISTRIBUTORE_TEMPLATE = (
    "// Questo file racchiude la lista dei codici dei moduli da montare\r\n"
    "// sul distributore selezionato\r\n"
    "\r\n"
    "sezione1=\r\n"
    "sezione2=\r\n"
    "sezione3=\r\n"
    "sezione4=\r\n"
    "sezione5=\r\n"
    "\r\n"
    "// File di calibrazione specifici dei moduli CE16 presenti sul distributore.\r\n"
    "calibrazioneCE16_1=\r\n"
    "calibrazioneCE16_2=\r\n"
    "calibrazioneCE16_3=\r\n"
    "calibrazioneCE16_4=\r\n"
    "calibrazioneCE16_5=\r\n"
)

TEST_SEQUENCE_TEMPLATE = (
    "NOME TEST;ID TEST;PARAMETRO 0;PARAMETRO 1;PARAMETRO 2;PARAMETRO 3;"
    "PARAMETRO 4;PARAMETRO 5;PARAMETRO 6;PARAMETRO7\r\n"
)

CURVA_COMANDO_TEMPLATE = "tensione;stato;stato\r\n"
CURVA_LIMITE_TEMPLATE = "COMANDO;POSIZIONE\r\n"
RAMPA_XY_TEMPLATE = "Tempo;Tensione\r\n"
CE16_TEMPLATE = "PUNTI;POSIZIONE\r\n"
MMS2218_TEMPLATE = "PUNTO X;PUNTO Y\r\n"
FILE_ESTERNO_TEMPLATE = ""

CATEGORY_DEFAULT_NAMES = {
    "distributori": ("nuovo_distributore", ".cfg"),
    "sequenze_test": ("nuova_sequenza", ".csv"),
    "curve_comando": ("nuova_curva_comando", ".csv"),
    "curve_limite": ("nuova_curva_limite", ".csv"),
    "rampe_xy": ("nuova_rampa", ".csv"),
    "calibrazioni_ce16": ("nuovo_ce16", ".cfg"),
    "calibrazioni_mms2218": ("nuovo_adc", ".csv"),
    "file_esterni": ("nuovo_file_esterno", ".txt"),
}

CATEGORY_TEMPLATES = {
    "distributori": DISTRIBUTORE_TEMPLATE,
    "sequenze_test": TEST_SEQUENCE_TEMPLATE,
    "curve_comando": CURVA_COMANDO_TEMPLATE,
    "curve_limite": CURVA_LIMITE_TEMPLATE,
    "rampe_xy": RAMPA_XY_TEMPLATE,
    "calibrazioni_ce16": CE16_TEMPLATE,
    "calibrazioni_mms2218": MMS2218_TEMPLATE,
    "file_esterni": FILE_ESTERNO_TEMPLATE,
}


def supports_new_resource(category_key: str) -> bool:
    return category_key in CATEGORY_TEMPLATES


def _next_unique_path(directory: Path, stem: str, suffix: str) -> Path:
    candidate = directory / f"{stem}{suffix}"
    if not candidate.exists():
        return candidate

    for index in range(1, 10_000):
        candidate = directory / f"{stem} {index}{suffix}"
        if not candidate.exists():
            return candidate

    raise RuntimeError(f"Impossibile trovare un nome libero in {directory}")


def _target_directory(category_key: str, project_root: Path) -> Path:
    if category_key in CATEGORY_DIRECTORIES:
        return project_root / CATEGORY_DIRECTORIES[category_key]

    if category_key == "calibrazioni_ce16":
        return project_root / SETTINGS_PROGRAM_DIR / DEFAULT_SETTINGS_PROFILE / "CE16"

    if category_key == "calibrazioni_mms2218":
        return project_root / SETTINGS_PROGRAM_DIR / DEFAULT_SETTINGS_PROFILE / "MMS2218"

    raise ValueError(f"Categoria senza directory associata: {category_key}")


def create_new_resource(category_key: str, project_root: Path) -> Path:
    if not supports_new_resource(category_key):
        raise ValueError(f"Categoria non supportata per la creazione: {category_key}")

    stem, suffix = CATEGORY_DEFAULT_NAMES[category_key]
    target_dir = _target_directory(category_key, project_root)
    target_dir.mkdir(parents=True, exist_ok=True)

    new_path = _next_unique_path(target_dir, stem, suffix)
    new_path.write_bytes(CATEGORY_TEMPLATES[category_key].encode("utf-8"))
    return new_path
