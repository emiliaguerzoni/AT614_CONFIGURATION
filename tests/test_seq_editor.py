import shutil
from pathlib import Path

from PySide6.QtWidgets import QAbstractItemView

from at614_editor.domain.parsers.test_csv import parse as parse_test_csv
from at614_editor.ui.editors.seq_editor import SeqEditor
from at614_editor.ui.main_window import MainWindow


FIXTURES_ROOT = Path(__file__).parent / "fixtures"


def _build_temp_project(tmp_path: Path) -> Path:
    project_root = tmp_path / "project"
    shutil.copytree(FIXTURES_ROOT, project_root)
    return project_root


def _create_external_fixture(project_root: Path, file_name: str) -> Path:
    external_dir = project_root / "EXTERNAL"
    external_dir.mkdir(exist_ok=True)
    file_path = external_dir / file_name
    file_path.write_text("placeholder", encoding="utf-8")
    return file_path


def _create_resource_placeholder(project_root: Path, relative_path: str) -> Path:
    file_path = project_root / relative_path
    file_path.parent.mkdir(parents=True, exist_ok=True)
    file_path.write_text("placeholder", encoding="utf-8")
    return file_path


def test_seq_editor_save_persists_common_fields_and_file_refs(tmp_path: Path, qt_app) -> None:
    project_root = _build_temp_project(tmp_path)
    _create_external_fixture(project_root, "15.1001.356_C MODULO MLT FD5 MERLO FW 3.0.2.0.txt")
    window = MainWindow(project_root=project_root)
    test_path = project_root / "TEST" / "15.1001.356_C.csv"

    assert window.editor_shell.resource_tree.activate_path(test_path)

    editor = window.editor_shell.current_editor
    assert isinstance(editor, SeqEditor)
    editor.select_row(8)

    editor.name_field.setText("TEMPI EXTEND REVISIONE")
    editor.index_field.setText("88")
    curve_widget = next(widget for widget in editor.file_reference_widgets if widget.reference_text() == "TEST_CURVE_COMANDO_GR.csv")
    curve_widget.name_label.setText("TEST_CURVE_COMANDO.csv")

    saved_path = window.editor_shell.save_current_resource()
    assert saved_path == test_path

    parsed = parse_test_csv(test_path)
    row = parsed.rows[8]
    assert row.name == "TEMPI EXTEND REVISIONE"
    assert row.index_raw == "88"
    assert row.parameters[0] == "TEST_CURVE_COMANDO.csv"


def test_seq_editor_save_is_blocked_on_validation_error(tmp_path: Path, qt_app) -> None:
    project_root = _build_temp_project(tmp_path)
    _create_external_fixture(project_root, "15.1001.356_C MODULO MLT FD5 MERLO FW 3.0.2.0.txt")
    window = MainWindow(project_root=project_root)
    test_path = project_root / "TEST" / "15.1001.356_C.csv"

    assert window.editor_shell.resource_tree.activate_path(test_path)

    editor = window.editor_shell.current_editor
    assert isinstance(editor, SeqEditor)

    editor.select_row(8)
    reference_widget = next(widget for widget in editor.file_reference_widgets if widget.reference_text() == "TEST_CURVE_COMANDO_GR.csv")
    reference_widget.name_label.setText("CURVA_ASSENTE.csv")
    editor.select_row(0)

    saved_path = window.editor_shell.save_current_resource()

    assert saved_path is None
    assert editor.current_row_index == 8
    assert window.editor_shell.right_panel.tabs.currentIndex() == 1
    assert window.editor_shell.right_panel.validation_list.item(0).text() == "Riga 9 · Curva comando: — riferimento mancante —"

    parsed = parse_test_csv(test_path)
    assert parsed.rows[8].parameters[0] == "TEST_CURVE_COMANDO_GR.csv"


def test_seq_editor_save_allows_warning_only_rows(tmp_path: Path, qt_app) -> None:
    project_root = _build_temp_project(tmp_path)
    _create_external_fixture(project_root, "15.1001.310_B MODULO MLT FD5 D_C0 TDV100.txt")
    _create_resource_placeholder(project_root, "CURVE_LIMITE/LIMITE_INF 15.1001.311_A.csv")
    _create_resource_placeholder(project_root, "CURVE_LIMITE/LIMITE_SUP 15.1001.311_A.csv")
    window = MainWindow(project_root=project_root)
    test_path = project_root / "TEST" / "15.1001.310_RC.csv"

    assert window.editor_shell.resource_tree.activate_path(test_path)

    editor = window.editor_shell.current_editor
    assert isinstance(editor, SeqEditor)

    editor.select_row(0)
    editor.name_field.setText("")

    saved_path = window.editor_shell.save_current_resource()

    assert saved_path == test_path
    parsed = parse_test_csv(test_path)
    assert parsed.rows[0].name == ""


def test_seq_editor_duplicate_creates_new_sequence_and_reselects_it(tmp_path: Path, qt_app) -> None:
    project_root = _build_temp_project(tmp_path)
    window = MainWindow(project_root=project_root)
    test_path = project_root / "TEST" / "15.1001.356_C.csv"

    assert window.editor_shell.resource_tree.activate_path(test_path)

    editor = window.editor_shell.current_editor
    assert isinstance(editor, SeqEditor)
    editor.select_row(11)
    editor.name_field.setText("Acquisizione duplicata")

    duplicate_path = window.editor_shell.duplicate_current_resource()

    assert duplicate_path is not None
    assert duplicate_path.exists()
    assert duplicate_path.name == "15.1001.356_C - copia.csv"
    assert window.editor_shell.current_title() == duplicate_path.name

    parsed = parse_test_csv(duplicate_path)
    assert parsed.rows[11].name == "Acquisizione duplicata"


def test_seq_editor_add_duplicate_delete_and_reindex_rows(qt_app) -> None:
    window = MainWindow(project_root=FIXTURES_ROOT)
    test_path = FIXTURES_ROOT / "TEST" / "15.1001.310_RC.csv"

    assert window.editor_shell.resource_tree.activate_path(test_path)

    editor = window.editor_shell.current_editor
    assert isinstance(editor, SeqEditor)
    assert editor.row_table.rowCount() == 10

    editor.select_row(1)
    editor.add_row()
    assert editor.row_table.rowCount() == 11
    assert editor.current_row_index == 2
    assert editor.index_field.text() == "2"

    editor.name_field.setText("TEST AGGIUNTO")
    editor.test_id_field.setCurrentText("ASSEGNA_NODEID")
    editor._sync_and_refresh_current_row()
    assert editor.parameter_labels() == ["Node ID destinazione"]

    editor.duplicate_selected_row()
    assert editor.row_table.rowCount() == 12
    assert editor.current_row_index == 3
    assert editor.row_table.item(3, 2).text() == "3"

    editor.delete_selected_row()
    assert editor.row_table.rowCount() == 11
    assert [editor.row_table.item(index, 2).text() for index in range(editor.row_table.rowCount())] == [
        str(index) for index in range(editor.row_table.rowCount())
    ]


def test_seq_editor_supports_row_reorder_with_internal_move(qt_app) -> None:
    window = MainWindow(project_root=FIXTURES_ROOT)
    test_path = FIXTURES_ROOT / "TEST" / "15.1001.310_RC.csv"

    assert window.editor_shell.resource_tree.activate_path(test_path)

    editor = window.editor_shell.current_editor
    assert isinstance(editor, SeqEditor)
    assert editor.row_table.dragDropMode() == QAbstractItemView.DragDropMode.InternalMove

    original_names = [row.name for row in editor.working_rows[:4]]

    editor.move_row(0, 3)

    assert [row.name for row in editor.working_rows[:4]] == [
        original_names[1],
        original_names[2],
        original_names[3],
        original_names[0],
    ]
    assert editor.row_table.item(3, 3).text() == editor.current_row_status()
    assert [editor.row_table.item(index, 2).text() for index in range(editor.row_table.rowCount())] == [
        str(index) for index in range(editor.row_table.rowCount())
    ]
