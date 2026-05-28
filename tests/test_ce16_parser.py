from pathlib import Path

from at614_editor.domain.parsers.ce16 import parse, serialize


FIXTURES_DIR = Path(__file__).parent / "fixtures" / "SETTAGGI PROGRAMMA"


def test_parse_ce16_extracts_profile_and_slot() -> None:
    fixture = FIXTURES_DIR / "MERLO" / "CE16" / "CE16_3.cfg"

    resource = parse(fixture)

    assert resource.profile_name == "MERLO"
    assert resource.slot_index == 3
    assert resource.series.header_fields == ["PUNTI", "POSIZIONE"]
    assert resource.series.rows[1].values == ["1000", "0"]


def test_ce16_round_trip_bytes() -> None:
    fixture = FIXTURES_DIR / "MERLO" / "CE16" / "CE16_1.cfg"
    raw = fixture.read_bytes()

    resource = parse(fixture)

    assert serialize(resource) == raw
