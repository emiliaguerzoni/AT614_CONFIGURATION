from pathlib import Path

from at614_editor.domain.parsers.mms2218 import parse, serialize


FIXTURES_DIR = Path(__file__).parent / "fixtures" / "SETTAGGI PROGRAMMA"


def test_parse_mms2218_extracts_profile_and_channel() -> None:
    fixture = FIXTURES_DIR / "MERLO" / "MMS2218" / "ADC0.csv"

    resource = parse(fixture)

    assert resource.profile_name == "MERLO"
    assert resource.channel_index == 0
    assert resource.series.header_fields == ["PUNTO X", "PUNTO Y"]
    assert resource.series.rows[-1].values == ["1024", "5"]


def test_mms2218_round_trip_bytes() -> None:
    fixture = FIXTURES_DIR / "DEFAULT" / "MMS2218" / "ADC0.csv"
    raw = fixture.read_bytes()

    resource = parse(fixture)

    assert serialize(resource) == raw
