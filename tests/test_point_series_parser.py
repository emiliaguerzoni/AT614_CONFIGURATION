from pathlib import Path

from at614_editor.domain.parsers.point_series import parse, serialize


FIXTURES_ROOT = Path(__file__).parent / "fixtures"


def test_parse_curve_comando_point_series() -> None:
    fixture = FIXTURES_ROOT / "CURVE_COMANDO" / "REAL_TEST_CURVE_COMANDO.csv"

    resource = parse(fixture)

    assert resource.header_fields == ["tensione", "stato", "stato"]
    assert resource.rows[0].values == ["0", "-2", "errore extend"]
    assert resource.rows[-1].values == ["5", "2", "errore retract"]


def test_parse_limite_point_series() -> None:
    fixture = FIXTURES_ROOT / "CURVE_LIMITE" / "REAL_LIMITE_INF.csv"

    resource = parse(fixture)

    assert resource.header_fields == ["COMANDO", "POSIZIONE"]
    assert resource.rows[1].values == ["2165", "1,4"]


def test_point_series_round_trip_real_rampa() -> None:
    fixture = FIXTURES_ROOT / "RAMPE_XY" / "REAL_TEST_RAMPA.csv"
    raw = fixture.read_bytes()

    resource = parse(fixture)

    assert serialize(resource) == raw
