from __future__ import annotations

import logging
from pathlib import Path

from at614_editor.domain.models import DistributoreConfig, DistributoreLine

logger = logging.getLogger(__name__)

SECTION_PREFIX = "sezione"
CE16_PREFIX = "calibrazioneCE16_"


def _detect_encoding(raw: bytes) -> str:
    try:
        raw.decode("utf-8")
    except UnicodeDecodeError:
        return "cp1252"
    return "utf-8"


def _detect_line_ending(text: str) -> str:
    if "\r\n" in text:
        return "\r\n"
    return "\n"


def _parse_index(key: str, prefix: str) -> int | None:
    if not key.startswith(prefix):
        return None

    suffix = key[len(prefix) :]
    if not suffix.isdigit():
        return None
    return int(suffix)


def parse(path: Path) -> DistributoreConfig:
    try:
        logger.info(f"Parsing distributore: {path}")
        raw = path.read_bytes()
        encoding = _detect_encoding(raw)
        logger.debug(f"Encoding rilevato: {encoding}")
        text = raw.decode(encoding)
        line_ending = _detect_line_ending(text)
        logger.debug(f"Line ending rilevato: {repr(line_ending)}")
        lines: list[DistributoreLine] = []
        sezioni: dict[int, str] = {}
        calibrazioni_ce16: dict[int, str] = {}
        extras: dict[str, str] = {}
    except Exception as e:
        logger.exception(f"Errore durante il parsing di {path}: {e}")
        raise

    for raw_line in text.splitlines():
        stripped = raw_line.strip()
        if not stripped or stripped.startswith("//") or "=" not in raw_line:
            lines.append(DistributoreLine(key=None, value=None, raw=raw_line))
            continue

        key, value = raw_line.split("=", 1)
        key = key.strip()
        value = value.strip()
        lines.append(DistributoreLine(key=key, value=value, raw=raw_line))

        section_index = _parse_index(key, SECTION_PREFIX)
        if section_index is not None:
            sezioni[section_index] = value
            continue

        ce16_index = _parse_index(key, CE16_PREFIX)
        if ce16_index is not None:
            calibrazioni_ce16[ce16_index] = value
            continue

        extras[key] = value

    return DistributoreConfig(
        source_path=path,
        sezioni=sezioni,
        calibrazioni_ce16=calibrazioni_ce16,
        extras=extras,
        lines=lines,
        encoding=encoding,
        line_ending=line_ending,
        endswith_newline=text.endswith(("\n", "\r\n")),
    )


def serialize(config: DistributoreConfig) -> bytes:
    serialized_lines: list[str] = []

    for line in config.lines:
        if line.key is None:
            serialized_lines.append(line.raw)
            continue

        section_index = _parse_index(line.key, SECTION_PREFIX)
        if section_index is not None:
            value = config.sezioni.get(section_index, line.value or "")
            serialized_lines.append(f"{line.key}={value}")
            continue

        ce16_index = _parse_index(line.key, CE16_PREFIX)
        if ce16_index is not None:
            value = config.calibrazioni_ce16.get(ce16_index, line.value or "")
            serialized_lines.append(f"{line.key}={value}")
            continue

        value = config.extras.get(line.key, line.value or "")
        serialized_lines.append(f"{line.key}={value}")

    text = config.line_ending.join(serialized_lines)
    if config.endswith_newline:
        text += config.line_ending
    return text.encode(config.encoding)
