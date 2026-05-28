from pathlib import Path

from at614_editor.domain.models import TestRow as DomainTestRow
from at614_editor.domain.parsers.test_csv import parse
from at614_editor.domain.schemas import get_test_schema_catalog
from at614_editor.domain.test_schema import describe_test_parameters, extract_test_file_references


FIXTURES_DIR = Path(__file__).parent / "fixtures" / "TEST"


def test_extract_file_references_for_acquisizione_can() -> None:
    sequence = parse(FIXTURES_DIR / "15.1001.356_C.csv")
    row = next(item for item in sequence.rows if item.test_id == "ACQUISIZIONE_CAN")

    references = extract_test_file_references(row)

    assert [(ref.parameter_index, ref.resource_type, ref.raw_value) for ref in references] == [
        (1, "curva_comando", "TEST_CURVE_GRADINO.csv"),
        (2, "rampa_xy", "TEST_RAMPA_GRADINO.csv"),
        (6, "limite_inf", "LIMITE_INF_GRADINO.csv"),
        (7, "limite_sup", "LIMITE_SUP_GRADINO.csv"),
    ]


def test_extract_file_references_for_risposte_gradino() -> None:
    sequence = parse(FIXTURES_DIR / "15.1001.356_C.csv")
    row = next(item for item in sequence.rows if item.name == "TEMPI EXTEND")

    references = extract_test_file_references(row)

    assert [(ref.parameter_index, ref.resource_type, ref.raw_value) for ref in references] == [
        (1, "curva_comando", "TEST_CURVE_COMANDO_GR.csv"),
        (2, "rampa_xy", "TEST_RAMPA_EXTEND_GR.csv"),
    ]


def test_extract_file_references_for_scrittura_parametri() -> None:
    sequence = parse(FIXTURES_DIR / "15.1001.310_RC.csv")
    row = next(item for item in sequence.rows if item.test_id == "SCRITTURA_PARAMETRI")

    references = extract_test_file_references(row)

    assert [(ref.parameter_index, ref.resource_type, ref.raw_value) for ref in references] == [
        (1, "external_config", "15.1001.310_B MODULO MLT FD5 D_C0 TDV100.txt")
    ]


def test_describe_common_test_parameters_from_schema() -> None:
    sequence = parse(FIXTURES_DIR / "15.1001.310_RC.csv")

    lettura_cfg_hw = next(item for item in sequence.rows if item.test_id == "LETTURA_CFG_HW")
    test_ciclica = next(item for item in sequence.rows if item.name == "SPURGO VALVOLE")
    calibrazione = next(item for item in sequence.rows if item.test_id == "CALIBRAZIONE")

    assert [descriptor.label for descriptor in describe_test_parameters(lettura_cfg_hw)] == [
        "Sincronia cattura",
        "Versione firmware attesa",
        "Release hardware attesa",
    ]
    assert [descriptor.label for descriptor in describe_test_parameters(test_ciclica)] == [
        "Tempo ciclica",
        "Tempo singolo ciclo",
        "Source address",
        "Riferimento neutro",
        "Riferimento max retract",
        "Riferimento max extend",
        "Tolleranza",
    ]
    assert [descriptor.label for descriptor in describe_test_parameters(calibrazione)] == [
        "Nome file calibrazione",
        "Sincronia calibrazione",
        "Corsa retract",
        "Corsa extend",
        "Tolleranza corsa",
    ]


def test_schema_catalog_loads_auto_and_override_aliases() -> None:
    catalog = get_test_schema_catalog()

    assert "CICLICA_CAN" in catalog
    assert "TEST_CICLICA_CAN" in catalog
    assert "SERIAL_NUMBER" in catalog


def test_describe_test_parameters_uses_auto_schema_for_unrefined_ids() -> None:
    row = DomainTestRow(
        name="",
        test_id="ACQUISIZIONE_PRE",
        index_raw="0",
        parameters=[""] * 17,
    )

    assert [descriptor.label for descriptor in describe_test_parameters(row)[:4]] == [
        "Nome file rampa xy",
        "Nome file curva comando",
        "Scheda indirizzo x1",
        "Var tipo x1",
    ]


def test_describe_test_parameters_skips_pure_trailing_empty_fields() -> None:
    sequence = parse(FIXTURES_DIR / "15.1001.356_C.csv")
    row = next(item for item in sequence.rows if item.test_id == "SERIAL_NUMBER")

    assert describe_test_parameters(row) == []
