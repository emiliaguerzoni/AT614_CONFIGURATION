from pathlib import Path

from at614_editor.ui.main_window import MainWindow
from at614_editor.ui.components.file_widget import FileWidget
from at614_editor.ui.resource_tree import ResourceTreeSelection
from at614_editor.ui.resource_tree import CATEGORY_DEFINITIONS
from at614_editor.ui.editors.dist_editor import DistEditor
from at614_editor.ui.editors.external_file_editor import ExternalFileEditor
from at614_editor.ui.editors.point_series_editors import Ce16Editor, RampEditor
from at614_editor.ui.editors.seq_editor import SeqEditor


FIXTURES_ROOT = Path(__file__).parent / "fixtures"


def test_main_window_builds_shell_layout(qt_app) -> None:
    window = MainWindow(project_root=FIXTURES_ROOT)

    assert window.windowTitle() == "AT614 Configuration Editor"
    assert window.editor_shell.resource_tree.minimumWidth() == 240
    assert window.editor_shell.resource_tree.maximumWidth() == 240
    assert window.editor_shell.right_panel.minimumWidth() == 260
    assert window.editor_shell.right_panel.maximumWidth() == 260
    assert window.editor_shell.resource_tree.top_level_labels() == [label for _, label in CATEGORY_DEFINITIONS]
    assert window.editor_shell.right_panel.tab_labels() == ["Riferimenti", "Validazione", "Azioni rapide"]
    assert window.editor_shell.right_panel.quick_action_labels() == [
        "Nuovo collegato",
        "Duplica collegato",
        "Sostituisci",
        "Apri",
        "Rinomina",
        "Elimina",
    ]


def test_main_window_opens_first_resource(qt_app) -> None:
    window = MainWindow(project_root=FIXTURES_ROOT)

    window.open_first_available_resource()

    assert window.active_page_name() == "editor"
    assert window.editor_shell.current_title() == "15.1001.310_RC.cfg"


def test_workspace_home_shows_project_summary(qt_app) -> None:
    window = MainWindow(project_root=FIXTURES_ROOT)

    assert "Distributori: 2" in window.home_page.summary_lines()
    assert "Calibrazioni CE16: 6" in window.home_page.summary_lines()
    assert "Calibrazioni MMS2218: 2" in window.home_page.summary_lines()


def test_dist_editor_has_fixed_sections_slots_and_params(qt_app) -> None:
    window = MainWindow(project_root=FIXTURES_ROOT)
    distributore_path = FIXTURES_ROOT / "DISTRIBUTORE" / "15.1001.356_C.cfg"

    assert window.editor_shell.resource_tree.activate_path(distributore_path)

    editor = window.editor_shell.current_editor
    assert isinstance(editor, DistEditor)
    assert len(editor.section_cards) == 5
    assert len(editor.ce16_widgets) == 5
    assert editor.parameters_table.rowCount() == 4
    assert editor.parameters_table.item(0, 0).text() == "Pressure_max"
    assert editor.ce16_widgets[0].uses_button.text() == "Usi (2)"


def test_ramp_editor_renders_table_and_toolbar(qt_app) -> None:
    window = MainWindow(project_root=FIXTURES_ROOT)
    ramp_path = FIXTURES_ROOT / "RAMPE_XY" / "REAL_TEST_RAMPA.csv"

    assert window.editor_shell.resource_tree.activate_path(ramp_path)

    editor = window.editor_shell.current_editor
    assert isinstance(editor, RampEditor)
    assert editor.points_table.rowCount() == 14
    assert editor.chart_stage.toolbar_labels() == ["Snap griglia", "Inserisci punto intermedio"]


def test_ce16_editor_shows_optional_badge(qt_app) -> None:
    window = MainWindow(project_root=FIXTURES_ROOT)
    ce16_path = FIXTURES_ROOT / "SETTAGGI PROGRAMMA" / "MERLO" / "CE16" / "CE16_3.cfg"

    assert window.editor_shell.resource_tree.activate_path(ce16_path)

    editor = window.editor_shell.current_editor
    assert isinstance(editor, Ce16Editor)
    assert editor.badge_label.text() == "opzionale"


def test_external_file_editor_opens_for_unresolved_reference(qt_app) -> None:
    window = MainWindow(project_root=FIXTURES_ROOT)

    window.editor_shell.show_selection(
        ResourceTreeSelection(
            "file_esterni",
            "File esterni",
            "15.1001.356_C MODULO MLT FD5 MERLO FW 3.0.2.0.txt",
            None,
        )
    )

    editor = window.editor_shell.current_editor
    assert isinstance(editor, ExternalFileEditor)
    assert editor.file_widget.name_label.text() == "15.1001.356_C MODULO MLT FD5 MERLO FW 3.0.2.0.txt"
    assert "Riferimento esterno non risolto" in editor.warning_label.text()


def test_dist_editor_open_button_navigates_to_ce16_editor(qt_app) -> None:
    window = MainWindow(project_root=FIXTURES_ROOT)
    distributore_path = FIXTURES_ROOT / "DISTRIBUTORE" / "15.1001.356_C.cfg"

    assert window.editor_shell.resource_tree.activate_path(distributore_path)

    editor = window.editor_shell.current_editor
    assert isinstance(editor, DistEditor)

    editor.ce16_widgets[0].open_button.click()

    assert window.editor_shell.current_title() == "MERLO · CE16_1.cfg"
    assert isinstance(window.editor_shell.current_editor, Ce16Editor)


def test_dist_editor_uses_button_is_enabled_for_referenced_ce16(qt_app) -> None:
    # M5: count <= 10 → popup non bloccante (non aggiorna rpanel)
    window = MainWindow(project_root=FIXTURES_ROOT)
    distributore_path = FIXTURES_ROOT / "DISTRIBUTORE" / "15.1001.356_C.cfg"

    assert window.editor_shell.resource_tree.activate_path(distributore_path)

    editor = window.editor_shell.current_editor
    assert isinstance(editor, DistEditor)

    # Il pulsante è abilitato (il CE16 è referenziato da almeno 1 file)
    assert editor.ce16_widgets[0].uses_button.isEnabled()
    # Il click non deve causare eccezioni (mostra popup contestuale)
    editor.ce16_widgets[0].uses_button.click()


def test_dist_editor_open_button_navigates_to_test_sequence(qt_app) -> None:
    window = MainWindow(project_root=FIXTURES_ROOT)
    distributore_path = FIXTURES_ROOT / "DISTRIBUTORE" / "15.1001.356_C.cfg"

    assert window.editor_shell.resource_tree.activate_path(distributore_path)

    editor = window.editor_shell.current_editor
    assert isinstance(editor, DistEditor)

    editor.section_cards[0].open_button.click()

    assert window.editor_shell.current_title() == "15.1001.356_C.csv"


def test_test_reference_open_button_navigates_to_ramp_editor(qt_app) -> None:
    window = MainWindow(project_root=FIXTURES_ROOT)
    test_path = FIXTURES_ROOT / "TEST" / "15.1001.356_C.csv"

    assert window.editor_shell.resource_tree.activate_path(test_path)

    editor = window.editor_shell.current_editor
    assert isinstance(editor, SeqEditor)
    editor.select_row(8)

    reference_widget = next(widget for widget in editor.file_reference_widgets if widget.reference_text() == "TEST_RAMPA_EXTEND_GR.csv")
    reference_widget.open_button.click()

    assert window.editor_shell.current_title() == "TEST_RAMPA_EXTEND_GR.csv"
    assert isinstance(window.editor_shell.current_editor, RampEditor)


def test_test_reference_uses_button_is_enabled(qt_app) -> None:
    # M5: count <= 10 → popup non bloccante (non aggiorna rpanel)
    window = MainWindow(project_root=FIXTURES_ROOT)
    test_path = FIXTURES_ROOT / "TEST" / "15.1001.356_C.csv"

    assert window.editor_shell.resource_tree.activate_path(test_path)

    editor = window.editor_shell.current_editor
    assert isinstance(editor, SeqEditor)
    editor.select_row(8)

    reference_widget = next(
        widget for widget in editor.file_reference_widgets
        if widget.reference_text() == "TEST_CURVE_COMANDO_GR.csv"
    )
    # Il pulsante è abilitato (il file è referenziato da almeno 1 test)
    assert reference_widget.uses_button.isEnabled()
    # Il click non deve causare eccezioni (mostra popup contestuale)
    reference_widget.uses_button.click()


def test_seq_editor_opens_with_common_fields_and_schema_labels(qt_app) -> None:
    window = MainWindow(project_root=FIXTURES_ROOT)
    test_path = FIXTURES_ROOT / "TEST" / "15.1001.356_C.csv"

    assert window.editor_shell.resource_tree.activate_path(test_path)

    editor = window.editor_shell.current_editor
    assert isinstance(editor, SeqEditor)
    assert editor.row_table.rowCount() == 13
    assert editor.name_field.text() == "Assegna Node ID"

    editor.select_row(8)
    assert editor.test_id_field.currentText() == "RISPOSTE_GRADINO"
    assert editor.parameter_labels()[:4] == ["Curva comando", "Rampa XY", "Parametro 3", "Parametro 4"]


def test_seq_editor_uses_extended_schema_for_common_ids(qt_app) -> None:
    window = MainWindow(project_root=FIXTURES_ROOT)
    test_path = FIXTURES_ROOT / "TEST" / "15.1001.310_RC.csv"

    assert window.editor_shell.resource_tree.activate_path(test_path)

    editor = window.editor_shell.current_editor
    assert isinstance(editor, SeqEditor)

    editor.select_row(0)
    assert editor.parameter_labels() == ["Node ID destinazione"]

    editor.select_row(1)
    assert editor.parameter_labels() == [
        "Sincronia cattura",
        "Versione firmware attesa",
        "Release hardware attesa",
    ]

    editor.select_row(4)
    assert editor.parameter_labels() == [
        "Tempo ciclica",
        "Tempo singolo ciclo",
        "Source address",
        "Riferimento neutro",
        "Riferimento max retract",
        "Riferimento max extend",
        "Tolleranza",
    ]


def test_seq_editor_hides_trailing_empty_parameters_without_schema(qt_app) -> None:
    window = MainWindow(project_root=FIXTURES_ROOT)
    test_path = FIXTURES_ROOT / "TEST" / "15.1001.356_C.csv"

    assert window.editor_shell.resource_tree.activate_path(test_path)

    editor = window.editor_shell.current_editor
    assert isinstance(editor, SeqEditor)

    editor.select_row(2)
    assert editor.test_id_field.currentText() == "SERIAL_NUMBER"
    assert editor.parameter_labels() == []


def test_seq_editor_live_validation_updates_row_badge_and_rpanel(qt_app) -> None:
    window = MainWindow(project_root=FIXTURES_ROOT)
    test_path = FIXTURES_ROOT / "TEST" / "15.1001.356_C.csv"

    assert window.editor_shell.resource_tree.activate_path(test_path)

    editor = window.editor_shell.current_editor
    assert isinstance(editor, SeqEditor)

    editor.select_row(8)
    reference_widget = next(widget for widget in editor.file_reference_widgets if widget.reference_text() == "TEST_CURVE_COMANDO_GR.csv")
    reference_widget.name_label.setText("CURVA_ASSENTE.csv")

    assert editor.current_row_status() == "errore"
    assert editor.row_table.item(8, 3).text() == "errore"
    assert window.editor_shell.validation_badge.text() == "errore"
    assert window.editor_shell.right_panel.validation_list.item(0).text() == "Curva comando: — riferimento mancante —"


def test_seq_editor_live_validation_marks_missing_common_fields_as_warning(qt_app) -> None:
    window = MainWindow(project_root=FIXTURES_ROOT)
    test_path = FIXTURES_ROOT / "TEST" / "15.1001.310_RC.csv"

    assert window.editor_shell.resource_tree.activate_path(test_path)

    editor = window.editor_shell.current_editor
    assert isinstance(editor, SeqEditor)

    editor.select_row(0)
    editor.name_field.setText("")

    assert editor.current_row_status() == "warning"
    assert editor.row_table.item(0, 3).text() == "warning"
    assert window.editor_shell.validation_badge.text() == "warning"
    assert editor.current_validation_messages() == ["Nome test mancante"]

