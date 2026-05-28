from pathlib import Path

from at614_editor.domain.parsers.distributore import parse, serialize


FIXTURES_DIR = Path(__file__).parent / "fixtures" / "DISTRIBUTORE"


def test_parse_distributore_base_fields() -> None:
    fixture = FIXTURES_DIR / "15.1001.310_RC.cfg"

    config = parse(fixture)

    assert config.sezioni[1] == "15.1001.310_RC"
    assert config.sezioni[5] == "15.1001.310_RC"
    assert config.calibrazioni_ce16[1] == "MERLO/CE16_1.cfg"
    assert config.calibrazioni_ce16[5] == "MERLO/CE16_5.cfg"
    assert config.extras == {}


def test_parse_distributore_with_extras() -> None:
    fixture = FIXTURES_DIR / "15.1001.356_C.cfg"

    config = parse(fixture)

    assert config.extras["Temperatura_olio_min"] == "40"
    assert config.extras["Temperatura_olio_max"] == "80"
    assert config.extras["Pressure_min"] == "20"
    assert config.extras["Pressure_max"] == "45"


def test_distributore_round_trip_bytes() -> None:
    fixture = FIXTURES_DIR / "15.1001.356_C.cfg"
    raw = fixture.read_bytes()

    config = parse(fixture)

    assert serialize(config) == raw
