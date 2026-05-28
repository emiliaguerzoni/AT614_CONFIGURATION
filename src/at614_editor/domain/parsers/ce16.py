from __future__ import annotations

from pathlib import Path

from at614_editor.domain.models import Ce16Resource
from at614_editor.domain.parsers.point_series import parse as parse_point_series
from at614_editor.domain.parsers.point_series import serialize as serialize_point_series


def _extract_slot_index(path: Path) -> int:
    stem = path.stem
    prefix = "CE16_"
    if not stem.startswith(prefix):
        raise ValueError(f"Invalid CE16 file name: {path.name}")

    suffix = stem[len(prefix) :]
    if not suffix.isdigit():
        raise ValueError(f"Invalid CE16 slot index: {path.name}")

    return int(suffix)


def parse(path: Path) -> Ce16Resource:
    if path.parent.name != "CE16":
        raise ValueError(f"Invalid CE16 parent directory: {path.parent}")

    series = parse_point_series(path)
    return Ce16Resource(
        source_path=path,
        profile_name=path.parent.parent.name,
        slot_index=_extract_slot_index(path),
        series=series,
    )


def serialize(resource: Ce16Resource) -> bytes:
    return serialize_point_series(resource.series)
