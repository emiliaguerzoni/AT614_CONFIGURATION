from pathlib import Path

from at614_editor.domain.project import load_project


FIXTURES_ROOT = Path(__file__).parent / "fixtures"


def test_project_loads_distributori_and_tests() -> None:
    project = load_project(FIXTURES_ROOT)

    assert len(project.distributori) == 2
    assert len(project.test_sequences) == 2
    assert len(project.point_series_resources) == 14
    assert len(project.ce16_resources) == 6
    assert len(project.mms2218_resources) == 2


def test_project_builds_distributore_to_test_references() -> None:
    project = load_project(FIXTURES_ROOT)

    distributore = FIXTURES_ROOT / "DISTRIBUTORE" / "15.1001.310_RC.cfg"
    target_test = FIXTURES_ROOT / "TEST" / "15.1001.310_RC.csv"

    assert target_test in project.get_uses(distributore)
    assert distributore in project.get_used_by(target_test)


def test_project_builds_distributore_to_ce16_references_with_profile_alias() -> None:
    project = load_project(FIXTURES_ROOT)

    distributore = FIXTURES_ROOT / "DISTRIBUTORE" / "15.1001.310_RC.cfg"
    merlo_ce16 = FIXTURES_ROOT / "SETTAGGI PROGRAMMA" / "MERLO" / "CE16" / "CE16_1.cfg"
    default_ce16 = FIXTURES_ROOT / "SETTAGGI PROGRAMMA" / "DEFAULT" / "CE16" / "CE16_1.cfg"

    assert merlo_ce16 in project.get_uses(distributore)
    assert default_ce16 not in project.get_uses(distributore)
    assert distributore in project.get_used_by(merlo_ce16)


def test_project_builds_test_to_resource_references() -> None:
    project = load_project(FIXTURES_ROOT)

    source_test = FIXTURES_ROOT / "TEST" / "15.1001.356_C.csv"
    expected_targets = {
        FIXTURES_ROOT / "CURVE_COMANDO" / "TEST_CURVE_COMANDO.csv",
        FIXTURES_ROOT / "CURVE_COMANDO" / "TEST_CURVE_COMANDO_GR.csv",
        FIXTURES_ROOT / "CURVE_COMANDO" / "TEST_CURVE_GRADINO.csv",
        FIXTURES_ROOT / "RAMPE_XY" / "TEST_RAMPA.csv",
        FIXTURES_ROOT / "RAMPE_XY" / "TEST_RAMPA_EXTEND_GR.csv",
        FIXTURES_ROOT / "RAMPE_XY" / "TEST_RAMPA_GRADINO.csv",
        FIXTURES_ROOT / "RAMPE_XY" / "TEST_RAMPA_RETRACT_GR.csv",
        FIXTURES_ROOT / "CURVE_LIMITE" / "LIMITE_INF.csv",
        FIXTURES_ROOT / "CURVE_LIMITE" / "LIMITE_SUP.csv",
        FIXTURES_ROOT / "CURVE_LIMITE" / "LIMITE_INF_GRADINO.csv",
        FIXTURES_ROOT / "CURVE_LIMITE" / "LIMITE_SUP_GRADINO.csv",
    }

    assert expected_targets.issubset(project.get_uses(source_test))


def test_project_resolves_file_reference_with_leading_description_prefix() -> None:
    from at614_editor.domain.project import _resolve_resource_path

    resource_index = {
        "example.txt": {Path("\\\\server\\share\\PRODUZIONE\\EXAMPLE.txt")},
        "example": {Path("\\\\server\\share\\PRODUZIONE\\EXAMPLE.txt")},
    }
    resolved = _resolve_resource_path(
        resource_index,
        "parametri banco gianluca example.txt",
        Path("\\\\server\\share\\PRODUZIONE"),
    )

    assert resolved == Path("\\\\server\\share\\PRODUZIONE\\EXAMPLE.txt")


def test_project_tracks_unresolved_test_file_references() -> None:
    project = load_project(FIXTURES_ROOT)

    source_test = FIXTURES_ROOT / "TEST" / "15.1001.356_C.csv"

    assert (
        "15.1001.356_C MODULO MLT FD5 MERLO FW 3.0.2.0.txt"
        in project.get_unresolved_references(source_test)
    )


def test_project_parses_point_series_resources() -> None:
    project = load_project(FIXTURES_ROOT)

    target = FIXTURES_ROOT / "CURVE_LIMITE" / "REAL_LIMITE_INF.csv"
    resource = project.point_series_resources[target]

    assert resource.header_fields == ["COMANDO", "POSIZIONE"]
    assert resource.rows[0].values == ["500", "6,8"]


def test_project_parses_settings_program_resources() -> None:
    project = load_project(FIXTURES_ROOT)

    ce16_target = FIXTURES_ROOT / "SETTAGGI PROGRAMMA" / "MERLO" / "CE16" / "CE16_3.cfg"
    mms_target = FIXTURES_ROOT / "SETTAGGI PROGRAMMA" / "MERLO" / "MMS2218" / "ADC0.csv"

    assert project.ce16_resources[ce16_target].slot_index == 3
    assert project.mms2218_resources[mms_target].channel_index == 0


