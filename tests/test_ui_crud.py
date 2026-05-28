import shutil
from pathlib import Path

from at614_editor.domain.parsers.distributore import parse as parse_distributore
from at614_editor.domain.parsers.point_series import parse as parse_point_series
from at614_editor.ui.editors.dist_editor import DistEditor
from at614_editor.ui.editors.point_series_editors import RampEditor
from at614_editor.ui.main_window import MainWindow


FIXTURES_ROOT = Path(__file__).parent / "fixtures"


def _build_temp_project(tmp_path: Path) -> Path:
    project_root = tmp_path / "project"
    shutil.copytree(FIXTURES_ROOT, project_root)
    return project_root


def test_dist_editor_save_persists_sections_ce16_and_params(tmp_path: Path, qt_app) -> None:
    project_root = _build_temp_project(tmp_path)
    window = MainWindow(project_root=project_root)
    distributore_path = project_root / "DISTRIBUTORE" / "15.1001.356_C.cfg"

    assert window.editor_shell.resource_tree.activate_path(distributore_path)

    editor = window.editor_shell.current_editor
    assert isinstance(editor, DistEditor)

    editor.section_cards[0].file_label.setText("15.1001.356_C_ALT.csv")
    editor.ce16_widgets[0].name_label.setText("MERLO/CE16_4.cfg")
    editor.parameters_table.item(0, 1).setText("99")

    saved_path = window.editor_shell.save_current_resource()

    assert saved_path == distributore_path

    parsed = parse_distributore(distributore_path)
    assert parsed.sezioni[1] == "15.1001.356_C_ALT"
    assert parsed.calibrazioni_ce16[1] == "MERLO/CE16_4.cfg"
    assert parsed.extras["Pressure_max"] == "99"


def test_ramp_editor_save_persists_point_changes(tmp_path: Path, qt_app) -> None:
    project_root = _build_temp_project(tmp_path)
    window = MainWindow(project_root=project_root)
    ramp_path = project_root / "RAMPE_XY" / "REAL_TEST_RAMPA.csv"

    assert window.editor_shell.resource_tree.activate_path(ramp_path)

    editor = window.editor_shell.current_editor
    assert isinstance(editor, RampEditor)

    editor.points_table.item(0, 1).setText("2600")
    saved_path = window.editor_shell.save_current_resource()

    assert saved_path == ramp_path
    parsed = parse_point_series(ramp_path)
    assert parsed.rows[0].values == ["0", "2600"]


def test_ramp_editor_duplicate_creates_new_resource_and_reselects_it(tmp_path: Path, qt_app) -> None:
    project_root = _build_temp_project(tmp_path)
    window = MainWindow(project_root=project_root)
    ramp_path = project_root / "RAMPE_XY" / "REAL_TEST_RAMPA.csv"

    assert window.editor_shell.resource_tree.activate_path(ramp_path)

    editor = window.editor_shell.current_editor
    assert isinstance(editor, RampEditor)
    editor.points_table.item(0, 1).setText("2600")

    duplicate_path = window.editor_shell.duplicate_current_resource()

    assert duplicate_path is not None
    assert duplicate_path.exists()
    assert duplicate_path.name == "REAL_TEST_RAMPA - copia.csv"
    assert window.editor_shell.current_title() == duplicate_path.name

    parsed = parse_point_series(duplicate_path)
    assert parsed.rows[0].values == ["0", "2600"]
