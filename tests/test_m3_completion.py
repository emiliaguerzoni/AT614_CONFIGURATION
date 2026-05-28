import shutil
from pathlib import Path

from at614_editor.domain.parsers.distributore import parse as parse_distributore
from at614_editor.domain.parsers.point_series import parse as parse_point_series
from at614_editor.domain.templates import (
    CATEGORY_DEFAULT_NAMES,
    CATEGORY_TEMPLATES,
    create_new_resource,
    supports_new_resource,
)
from at614_editor.ui.editors.dist_editor import DistEditor
from at614_editor.ui.editors.external_file_editor import ExternalFileEditor
from at614_editor.ui.editors.point_series_editors import Ce16Editor, MmsEditor
from at614_editor.ui.main_window import MainWindow
from at614_editor.ui.resource_tree import ResourceTreeSelection


FIXTURES_ROOT = Path(__file__).parent / "fixtures"


def _build_temp_project(tmp_path: Path) -> Path:
    project_root = tmp_path / "project"
    shutil.copytree(FIXTURES_ROOT, project_root)
    return project_root


# ---------------------------------------------------------------------------
# CE16Editor / MmsEditor — duplicate now supported
# ---------------------------------------------------------------------------


def test_ce16_editor_duplicate_creates_sibling_file(tmp_path: Path, qt_app) -> None:
    project_root = _build_temp_project(tmp_path)
    window = MainWindow(project_root=project_root)
    ce16_path = project_root / "SETTAGGI PROGRAMMA" / "MERLO" / "CE16" / "CE16_3.cfg"

    assert window.editor_shell.resource_tree.activate_path(ce16_path)

    editor = window.editor_shell.current_editor
    assert isinstance(editor, Ce16Editor)
    assert editor.supports_duplicate is True

    duplicate_path = window.editor_shell.duplicate_current_resource()

    assert duplicate_path is not None
    assert duplicate_path.exists()
    assert duplicate_path.parent == ce16_path.parent
    assert duplicate_path.name == "CE16_6.cfg"

    series = parse_point_series(duplicate_path)
    assert series.header_fields == ["PUNTI", "POSIZIONE"]


def test_mms_editor_duplicate_creates_sibling_file(tmp_path: Path, qt_app) -> None:
    project_root = _build_temp_project(tmp_path)
    window = MainWindow(project_root=project_root)
    mms_path = project_root / "SETTAGGI PROGRAMMA" / "MERLO" / "MMS2218" / "ADC0.csv"

    assert window.editor_shell.resource_tree.activate_path(mms_path)

    editor = window.editor_shell.current_editor
    assert isinstance(editor, MmsEditor)
    assert editor.supports_duplicate is True

    duplicate_path = window.editor_shell.duplicate_current_resource()

    assert duplicate_path is not None
    assert duplicate_path.exists()
    assert duplicate_path.name == "ADC1.csv"

    series = parse_point_series(duplicate_path)
    assert series.header_fields == ["PUNTO X", "PUNTO Y"]


# ---------------------------------------------------------------------------
# ExternalFileEditor — dual mode
# ---------------------------------------------------------------------------


def test_external_file_editor_warning_mode_when_path_missing(qt_app) -> None:
    editor = ExternalFileEditor("missing.txt", source_path=None)

    assert editor.supports_save is False
    assert editor.supports_duplicate is False
    assert "non risolto" in editor.warning_label.text().lower()
    assert editor.preview_editor.isReadOnly() is True


def test_external_file_editor_reads_existing_file(tmp_path: Path, qt_app) -> None:
    file_path = tmp_path / "params.txt"
    file_path.write_text("parameter A = 12\nparameter B = 34\n", encoding="utf-8")

    editor = ExternalFileEditor(file_path.name, source_path=file_path)

    assert editor.supports_save is True
    assert editor.supports_duplicate is True
    assert "modalità testo" in editor.warning_label.text().lower()
    assert editor.preview_editor.isReadOnly() is False
    assert "parameter A = 12" in editor.preview_editor.toPlainText()


def test_external_file_editor_save_writes_back(tmp_path: Path, qt_app) -> None:
    file_path = tmp_path / "params.txt"
    file_path.write_text("originale\n", encoding="utf-8")

    editor = ExternalFileEditor(file_path.name, source_path=file_path)
    editor.preview_editor.setPlainText("modificato\n")

    saved = editor.save_changes()

    assert saved == file_path
    assert file_path.read_text(encoding="utf-8") == "modificato\n"


def test_external_file_editor_duplicate_writes_sibling(tmp_path: Path, qt_app) -> None:
    file_path = tmp_path / "params.txt"
    file_path.write_text("contenuto\n", encoding="utf-8")

    editor = ExternalFileEditor(file_path.name, source_path=file_path)
    duplicate_path = editor.duplicate_resource()

    assert duplicate_path is not None
    assert duplicate_path.exists()
    assert duplicate_path.name == "params - copia.txt"
    assert duplicate_path.read_text(encoding="utf-8") == "contenuto\n"


# ---------------------------------------------------------------------------
# "Nuovo" action — create_new_resource via templates + shell wiring
# ---------------------------------------------------------------------------


def test_supports_new_resource_covers_authoring_categories() -> None:
    expected = {
        "distributori",
        "sequenze_test",
        "curve_comando",
        "curve_limite",
        "rampe_xy",
        "calibrazioni_ce16",
        "calibrazioni_mms2218",
        "file_esterni",
    }
    assert set(CATEGORY_TEMPLATES) == expected
    assert set(CATEGORY_DEFAULT_NAMES) == expected
    for category_key in expected:
        assert supports_new_resource(category_key)
    assert supports_new_resource("output_banco") is False


def test_create_new_distributore_writes_template_with_empty_sections(tmp_path: Path) -> None:
    project_root = _build_temp_project(tmp_path)

    new_path = create_new_resource("distributori", project_root)

    assert new_path.exists()
    assert new_path.parent == project_root / "DISTRIBUTORE"
    config = parse_distributore(new_path)
    assert set(config.sezioni.keys()) == {1, 2, 3, 4, 5}
    assert all(value == "" for value in config.sezioni.values())
    assert set(config.calibrazioni_ce16.keys()) == {1, 2, 3, 4, 5}
    assert all(value == "" for value in config.calibrazioni_ce16.values())


def test_create_new_curva_comando_has_expected_header(tmp_path: Path) -> None:
    project_root = _build_temp_project(tmp_path)

    new_path = create_new_resource("curve_comando", project_root)

    assert new_path.exists()
    assert new_path.parent == project_root / "CURVE_COMANDO"
    assert new_path.read_text(encoding="utf-8").startswith("tensione;stato;stato")


def test_create_new_ce16_targets_default_profile(tmp_path: Path) -> None:
    project_root = _build_temp_project(tmp_path)

    new_path = create_new_resource("calibrazioni_ce16", project_root)

    assert new_path.exists()
    assert new_path.parent == project_root / "SETTAGGI PROGRAMMA" / "DEFAULT" / "CE16"


def test_create_new_resource_avoids_collisions(tmp_path: Path) -> None:
    project_root = _build_temp_project(tmp_path)

    first = create_new_resource("rampe_xy", project_root)
    second = create_new_resource("rampe_xy", project_root)
    third = create_new_resource("rampe_xy", project_root)

    assert {first.name, second.name, third.name} == {
        "nuova_rampa.csv",
        "nuova_rampa 1.csv",
        "nuova_rampa 2.csv",
    }


def test_shell_new_button_is_enabled_only_for_authoring_categories(tmp_path: Path, qt_app) -> None:
    project_root = _build_temp_project(tmp_path)
    window = MainWindow(project_root=project_root)
    shell = window.editor_shell

    shell.show_selection(
        ResourceTreeSelection("output_banco", "Output banco", "Output banco", None)
    )
    assert shell.new_button.isEnabled() is False

    shell.show_selection(
        ResourceTreeSelection("rampe_xy", "Rampe XY", "Rampe XY", None)
    )
    assert shell.new_button.isEnabled() is True


def test_shell_create_new_in_rampe_category_opens_the_new_file(tmp_path: Path, qt_app) -> None:
    project_root = _build_temp_project(tmp_path)
    window = MainWindow(project_root=project_root)
    shell = window.editor_shell

    shell.show_selection(
        ResourceTreeSelection("rampe_xy", "Rampe XY", "Rampe XY", None)
    )

    new_path = shell.create_new_in_current_category()

    assert new_path is not None
    assert new_path.exists()
    assert new_path.parent == project_root / "RAMPE_XY"
    assert shell.current_title() == new_path.name


def test_shell_create_new_external_file_opens_external_editor(tmp_path: Path, qt_app) -> None:
    project_root = _build_temp_project(tmp_path)
    window = MainWindow(project_root=project_root)
    shell = window.editor_shell

    shell.show_selection(
        ResourceTreeSelection("file_esterni", "File esterni", "File esterni", None)
    )

    new_path = shell.create_new_in_current_category()

    assert new_path is not None
    assert new_path.exists()
    assert isinstance(shell.current_editor, ExternalFileEditor)
    assert shell.current_editor.supports_save is True


# ---------------------------------------------------------------------------
# DistEditor — does not mutate original config in-place
# ---------------------------------------------------------------------------


def test_dist_editor_collect_model_does_not_mutate_source_config(tmp_path: Path, qt_app) -> None:
    project_root = _build_temp_project(tmp_path)
    window = MainWindow(project_root=project_root)
    distributore_path = project_root / "DISTRIBUTORE" / "15.1001.356_C.cfg"

    assert window.editor_shell.resource_tree.activate_path(distributore_path)

    editor = window.editor_shell.current_editor
    assert isinstance(editor, DistEditor)

    original_sezioni = dict(editor.config.sezioni)
    original_ce16 = dict(editor.config.calibrazioni_ce16)
    original_extras = dict(editor.config.extras)

    editor.section_cards[0].file_label.setText("15.1001.356_C_ALT.csv")
    editor.ce16_widgets[0].name_label.setText("MERLO/CE16_4.cfg")
    editor.parameters_table.item(0, 1).setText("99")

    model = editor._collect_model(distributore_path)

    assert model.sezioni[1] == "15.1001.356_C_ALT"
    assert editor.config.sezioni == original_sezioni
    assert editor.config.calibrazioni_ce16 == original_ce16
    assert editor.config.extras == original_extras
