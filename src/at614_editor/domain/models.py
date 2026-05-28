from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(slots=True)
class DistributoreLine:
    key: str | None
    value: str | None
    raw: str


@dataclass(slots=True)
class DistributoreConfig:
    source_path: Path
    sezioni: dict[int, str]
    calibrazioni_ce16: dict[int, str]
    extras: dict[str, str]
    lines: list[DistributoreLine]
    encoding: str
    line_ending: str
    endswith_newline: bool


@dataclass(slots=True)
class TestRow:
    name: str
    test_id: str
    index_raw: str
    parameters: list[str]


@dataclass(slots=True)
class TestSequence:
    source_path: Path
    header_fields: list[str]
    rows: list[TestRow]
    encoding: str
    line_ending: str
    endswith_newline: bool


@dataclass(slots=True)
class TestFileReference:
    test_id: str
    parameter_index: int
    resource_type: str
    raw_value: str


@dataclass(slots=True)
class PointSeriesRow:
    values: list[str]


@dataclass(slots=True)
class PointSeriesResource:
    source_path: Path
    header_fields: list[str]
    rows: list[PointSeriesRow]
    encoding: str
    line_ending: str
    endswith_newline: bool


@dataclass(slots=True)
class Ce16Resource:
    source_path: Path
    profile_name: str
    slot_index: int
    series: PointSeriesResource


@dataclass(slots=True)
class Mms2218Resource:
    source_path: Path
    profile_name: str
    channel_index: int
    series: PointSeriesResource

