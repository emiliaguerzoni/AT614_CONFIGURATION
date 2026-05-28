from pathlib import Path

from at614_editor.domain.parsers.test_csv import parse, serialize


FIXTURES_DIR = Path(__file__).parent / "fixtures" / "TEST"


def test_parse_test_csv_preserves_header_and_rows() -> None:
    fixture = FIXTURES_DIR / "15.1001.310_RC.csv"

    sequence = parse(fixture)

    assert sequence.header_fields[:3] == ["NOME TEST", "ID TEST", "PARAMETRO 0"]
    assert sequence.rows[0].name == "SCRIVI NODE ID"
    assert sequence.rows[0].test_id == "ASSEGNA_NODEID"
    assert sequence.rows[0].index_raw == "0"
    assert sequence.rows[3].parameters[0] == "15.1001.310_B MODULO MLT FD5 D_C0 TDV100.txt"


def test_test_csv_round_trip_with_trailing_fields() -> None:
    fixture = FIXTURES_DIR / "15.1001.356_C.csv"
    raw = fixture.read_bytes()

    sequence = parse(fixture)

    assert serialize(sequence) == raw


def test_parse_test_csv_preserves_special_index_tokens(tmp_path: Path) -> None:
    fixture = tmp_path / "tokens.csv"
    fixture.write_text(
        "NOME TEST;ID TEST;PARAMETRO 0\n"
        "A;TEST_A;i\n"
        "B;TEST_B;i++\n"
        "C;TEST_C;++i\n",
        encoding="utf-8",
    )

    sequence = parse(fixture)

    assert [row.index_raw for row in sequence.rows] == ["i", "i++", "++i"]
    assert serialize(sequence) == fixture.read_bytes()
