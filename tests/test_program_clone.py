import shutil
from pathlib import Path

import pytest

from at614_editor.domain.parsers.distributore import parse as parse_distributore
from at614_editor.domain.parsers.test_csv import parse as parse_test_csv
from at614_editor.domain.program_clone import (
    build_clone_plan,
    derive_target_name,
    execute_clone_plan,
    validate_plan,
)
from at614_editor.domain.project import load_project


FIXTURES_ROOT = Path(__file__).parent / "fixtures"


def _build_temp_project(tmp_path: Path) -> Path:
    project_root = tmp_path / "project"
    shutil.copytree(FIXTURES_ROOT, project_root)
    return project_root


def test_derive_target_name_replaces_source_code() -> None:
    assert derive_target_name("15.1001.356_C.csv", "15.1001.500", "15.1001.356_C") == "15.1001.500.csv"


def test_derive_target_name_appends_new_code_when_not_matching() -> None:
    assert derive_target_name("TEST_CURVE_COMANDO.csv", "500", "15.1001.356_C") == "TEST_CURVE_COMANDO_500.csv"


def test_derive_target_name_avoids_redundant_suffix() -> None:
    assert derive_target_name("LIMITE_INF_500.csv", "500", "15.1001.356_C") == "LIMITE_INF_500.csv"


def test_build_clone_plan_collects_tests_and_resources(tmp_path: Path) -> None:
    project_root = _build_temp_project(tmp_path)
    project = load_project(project_root)
    source = project_root / "DISTRIBUTORE" / "15.1001.356_C.cfg"

    plan = build_clone_plan(project, source, "15.1001.999")

    assert plan.new_distributore_path.name == "15.1001.999.cfg"
    assert {action.source_path.name for action in plan.test_actions} == {"15.1001.356_C.csv"}
    resource_names = {action.source_path.name for action in plan.resource_actions}
    assert "TEST_CURVE_COMANDO.csv" in resource_names
    assert "LIMITE_INF.csv" in resource_names


def test_build_clone_plan_rejects_missing_source(tmp_path: Path) -> None:
    project_root = _build_temp_project(tmp_path)
    project = load_project(project_root)

    with pytest.raises(ValueError):
        build_clone_plan(project, project_root / "DISTRIBUTORE" / "non_esiste.cfg", "X")


def test_build_clone_plan_rejects_empty_code(tmp_path: Path) -> None:
    project_root = _build_temp_project(tmp_path)
    project = load_project(project_root)
    source = project_root / "DISTRIBUTORE" / "15.1001.356_C.cfg"

    with pytest.raises(ValueError):
        build_clone_plan(project, source, "  ")


def test_execute_clone_plan_creates_distributore_and_renamed_refs(tmp_path: Path) -> None:
    project_root = _build_temp_project(tmp_path)
    project = load_project(project_root)
    source = project_root / "DISTRIBUTORE" / "15.1001.356_C.cfg"

    plan = build_clone_plan(project, source, "15.1001.999")
    created = execute_clone_plan(project, plan)

    new_dist_path = project_root / "DISTRIBUTORE" / "15.1001.999.cfg"
    assert new_dist_path in created
    assert new_dist_path.exists()

    new_config = parse_distributore(new_dist_path)
    assert new_config.sezioni[1] == "15.1001.999"

    new_test_path = project_root / "TEST" / "15.1001.999.csv"
    assert new_test_path.exists()
    new_seq = parse_test_csv(new_test_path)
    acq_row = next(row for row in new_seq.rows if row.test_id == "ACQUISIZIONE_CAN")
    assert acq_row.parameters[0].endswith("_15.1001.999.csv") or acq_row.parameters[0] == "TEST_CURVE_GRADINO_15.1001.999.csv"


def test_execute_clone_plan_respects_riusa_mode(tmp_path: Path) -> None:
    project_root = _build_temp_project(tmp_path)
    project = load_project(project_root)
    source = project_root / "DISTRIBUTORE" / "15.1001.356_C.cfg"

    plan = build_clone_plan(project, source, "15.1001.999")
    for action in plan.resource_actions:
        action.mode = "riusa"

    execute_clone_plan(project, plan)
    new_test_path = project_root / "TEST" / "15.1001.999.csv"
    new_seq = parse_test_csv(new_test_path)

    acq_row = next(row for row in new_seq.rows if row.test_id == "ACQUISIZIONE_CAN")
    assert acq_row.parameters[0] == "TEST_CURVE_GRADINO.csv"

    duplicated_curve = project_root / "CURVE_COMANDO" / "TEST_CURVE_GRADINO_15.1001.999.csv"
    assert not duplicated_curve.exists()


def test_validate_plan_detects_existing_target(tmp_path: Path) -> None:
    project_root = _build_temp_project(tmp_path)
    project = load_project(project_root)
    source = project_root / "DISTRIBUTORE" / "15.1001.356_C.cfg"

    plan = build_clone_plan(project, source, "15.1001.999")

    new_dist_path = project_root / "DISTRIBUTORE" / "15.1001.999.cfg"
    new_dist_path.write_bytes(b"placeholder")

    errors = validate_plan(plan)
    assert any("15.1001.999.cfg" in error for error in errors)


def test_execute_clone_plan_fails_when_target_exists(tmp_path: Path) -> None:
    project_root = _build_temp_project(tmp_path)
    project = load_project(project_root)
    source = project_root / "DISTRIBUTORE" / "15.1001.356_C.cfg"

    plan = build_clone_plan(project, source, "15.1001.999")
    (project_root / "DISTRIBUTORE" / "15.1001.999.cfg").write_bytes(b"placeholder")

    with pytest.raises(ValueError):
        execute_clone_plan(project, plan)


def test_clone_program_dialog_creates_program_and_returns_path(tmp_path: Path, qt_app) -> None:
    from at614_editor.ui.dialogs.clone_program_dialog import CloneProgramDialog

    project_root = _build_temp_project(tmp_path)
    project = load_project(project_root)
    source = project_root / "DISTRIBUTORE" / "15.1001.356_C.cfg"

    dialog = CloneProgramDialog(project, source_distributore=source)
    dialog.new_code_field.setText("15.1001.999")
    dialog._on_accept()

    assert dialog.new_distributore_path() == project_root / "DISTRIBUTORE" / "15.1001.999.cfg"
    assert (project_root / "DISTRIBUTORE" / "15.1001.999.cfg").exists()
    assert (project_root / "TEST" / "15.1001.999.csv").exists()


def test_clone_program_dialog_riusa_mode_skips_duplication(tmp_path: Path, qt_app) -> None:
    from at614_editor.ui.dialogs.clone_program_dialog import CloneProgramDialog

    project_root = _build_temp_project(tmp_path)
    project = load_project(project_root)
    source = project_root / "DISTRIBUTORE" / "15.1001.356_C.cfg"

    dialog = CloneProgramDialog(project, source_distributore=source)
    dialog.new_code_field.setText("15.1001.999")

    for stored_action, combo, _item in dialog._action_refs:
        if stored_action.kind == "curva_comando":
            combo.setCurrentIndex(1)

    dialog._on_accept()

    assert (project_root / "DISTRIBUTORE" / "15.1001.999.cfg").exists()
    duplicated_curve = project_root / "CURVE_COMANDO" / "TEST_CURVE_GRADINO_15.1001.999.csv"
    assert not duplicated_curve.exists()
