from __future__ import annotations

from pathlib import Path

from at614_editor.domain.models import Mms2218Resource
from at614_editor.domain.parsers.point_series import parse as parse_point_series
from at614_editor.domain.parsers.point_series import serialize as serialize_point_series


def _extract_channel_index(path: Path) -> int:
    stem = path.stem
    prefix = "ADC"
    if not stem.startswith(prefix):
        raise ValueError(f"Invalid MMS2218 file name: {path.name}")

    suffix = stem[len(prefix) :]
    if not suffix.isdigit():
        raise ValueError(f"Invalid MMS2218 channel index: {path.name}")

    return int(suffix)


def parse(path: Path) -> Mms2218Resource:
    if path.parent.name != "MMS2218":
        raise ValueError(f"Invalid MMS2218 parent directory: {path.parent}")

    series = parse_point_series(path)
    return Mms2218Resource(
        source_path=path,
        profile_name=path.parent.parent.name,
        channel_index=_extract_channel_index(path),
        series=series,
    )


def serialize(resource: Mms2218Resource) -> bytes:
    return serialize_point_series(resource.series)
