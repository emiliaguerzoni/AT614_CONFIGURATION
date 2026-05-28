from __future__ import annotations

import logging
from pathlib import Path

from PySide6.QtCore import Signal

logger = logging.getLogger(__name__)

from PySide6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from at614_editor.domain.project import AT614Project, load_project
from at614_editor.domain.refactor import delete_resource, get_impact, rename_resource
from at614_editor.domain.templates import create_new_resource, supports_new_resource
from at614_editor.ui.components.impact_dialog import ImpactDialog
from at614_editor.ui.components.chart_stage import ChartStage
from at614_editor.ui.components.read_only_banner import ReadOnlyBanner
from at614_editor.ui.components.rpanel import RPanel
from at614_editor.ui.editors.dist_editor import DistEditor
from at614_editor.ui.editors.external_file_editor import ExternalFileEditor
from at614_editor.ui.editors.output_viewer import OutputViewer
from at614_editor.ui.editors.point_series_editors import Ce16Editor, CurveEditor, LimitEditor, MmsEditor, RampEditor
from at614_editor.ui.editors.seq_editor import SeqEditor
from at614_editor.ui.resource_tree import ResourceTree, ResourceTreeSelection


CATEGORY_ICONS = {
    "distributori": "□",
    "sequenze_test": "▲",
    "curve_comando": "○",
    "curve_limite": "○",
    "rampe_xy": "◇",
    "calibrazioni_ce16": "□",
    "calibrazioni_mms2218": "□",
    "file_esterni": "□",
    "output_banco": "○",
}


class EditorShell(QWidget):
    history_changed = Signal()

    def __init__(self, project_root: Path | None, project: AT614Project | None, parent=None) -> None:
        super().__init__(parent)
        self.project_root = project_root
        self.project = project
        self.current_editor: QWidget | None = None
        self.content_widgets: list[QWidget] = []
        self._current_reference_context: list[str] = []
        self._current_category_key: str | None = None
        self.history: list[ResourceTreeSelection] = []
        self.history_index: int = -1
        self._is_navigating = False
        self._last_selection: ResourceTreeSelection | None = None

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        self.resource_tree = ResourceTree(project)
        self.resource_tree.setFixedWidth(240)
        self.resource_tree.resource_activated.connect(self.show_selection)
        layout.addWidget(self.resource_tree)

        center_widget = QWidget()
        center_layout = QVBoxLayout(center_widget)
        center_layout.setContentsMargins(0, 0, 0, 0)
        center_layout.setSpacing(12)

        self.title_icon = QLabel("□")
        self.title_icon.setStyleSheet("font-size: 20px; font-weight: 600;")
        self.title_label = QLabel("Dashboard")
        self.title_label.setStyleSheet("font-size: 22px; font-weight: 700;")
        self.subtitle_label = QLabel("Seleziona una risorsa dall'albero a sinistra.")
        self.subtitle_label.setStyleSheet("color: #475467;")
        self.validation_badge = QLabel("ok")
        self._apply_validation_badge_style("ok")

        header_layout = QHBoxLayout()
        header_left = QVBoxLayout()

        title_row = QHBoxLayout()
        title_row.addWidget(self.title_icon)
        title_row.addWidget(self.title_label)
        title_row.addWidget(self.validation_badge)
        title_row.addStretch(1)
        header_left.addLayout(title_row)
        header_left.addWidget(self.subtitle_label)
        header_layout.addLayout(header_left, 1)

        self.new_button = QPushButton("Nuovo")
        self.duplicate_button = QPushButton("Duplica")
        self.rename_button = QPushButton("Rinomina")
        self.delete_button = QPushButton("Elimina")
        self.delete_button.setStyleSheet("color: #B42318;")
        self.save_button = QPushButton("Salva")
        self.new_button.clicked.connect(self.create_new_in_current_category)
        self.duplicate_button.clicked.connect(self.duplicate_current_resource)
        self.rename_button.clicked.connect(self.rename_current_resource)
        self.delete_button.clicked.connect(self.delete_current_resource)
        self.save_button.clicked.connect(self.save_current_resource)
        header_layout.addWidget(self.new_button)
        header_layout.addWidget(self.duplicate_button)
        header_layout.addWidget(self.rename_button)
        header_layout.addWidget(self.delete_button)
        header_layout.addWidget(self.save_button)
        center_layout.addLayout(header_layout)

        self.content_scroll = QScrollArea()
        self.content_scroll.setWidgetResizable(True)
        self.content_host = QWidget()
        self.content_layout = QVBoxLayout(self.content_host)
        self.content_layout.setContentsMargins(0, 0, 0, 0)
        self.content_layout.setSpacing(12)
        self.content_scroll.setWidget(self.content_host)
        center_layout.addWidget(self.content_scroll, 1)

        layout.addWidget(center_widget, 1)

        self.right_panel = RPanel()
        self.right_panel.setFixedWidth(260)
        layout.addWidget(self.right_panel)

        self.right_panel.elimina_button.clicked.connect(self.delete_current_resource)
        self.right_panel.rinomina_button.clicked.connect(self.rename_current_resource)

        self.show_empty_state()

    def _apply_validation_badge_style(self, badge_text: str) -> None:
        styles = {
            "ok": "background: #ECFDF3; border: 1px solid #ABEFC6; color: #067647; border-radius: 10px; padding: 2px 8px;",
            "warning": "background: #FFFAEB; border: 1px solid #FEDF89; color: #B54708; border-radius: 10px; padding: 2px 8px;",
            "errore": "background: #FEF3F2; border: 1px solid #FECDCA; color: #B42318; border-radius: 10px; padding: 2px 8px;",
        }
        self.validation_badge.setText(badge_text)
        self.validation_badge.setStyleSheet(styles.get(badge_text, styles["ok"]))

    def _set_validation_badge(self, badge_text: str) -> None:
        self._apply_validation_badge_style(badge_text)

    def show_empty_state(self) -> None:
        selection = ResourceTreeSelection("distributori", "Distributori", "Dashboard", None)
        self._last_selection = selection
        self._set_title(selection, "Seleziona una risorsa per aprire la shell di editing.")
        self._replace_content([QLabel("La milestone M2 espone layout 3 colonne, albero categorie e componenti base riusabili.")])
        self._current_reference_context = []
        self._current_category_key = None
        self._set_validation_badge("ok")
        self.right_panel.set_context([], [])

    def show_selection(self, selection: ResourceTreeSelection) -> None:
        if selection is None:
            return

        logger.debug(f"EditorShell: selezione risorsa - categoria={selection.category_key}, path={getattr(selection, 'path', None)}")

        # Ignora se è la stessa selezione
        if selection == self._last_selection:
            return

        if not self._prompt_save_if_dirty():
            # Ripristina selezione nell'albero in modo silenzioso
            self._is_navigating = True
            try:
                if self._last_selection:
                    self.resource_tree.activate_selection(self._last_selection)
            finally:
                self._is_navigating = False
            return

        self._last_selection = selection

        if not self._is_navigating:
            self.history = self.history[:self.history_index + 1]
            if not self.history or self.history[-1] != selection:
                self.history.append(selection)
                self.history_index += 1
            self.history_changed.emit()

        self._current_category_key = selection.category_key
        self._set_title(selection, self._subtitle_for(selection))

        if selection.path is None and selection.category_key == "file_esterni" and selection.display_label != selection.category_label:
            self._show_external_file(selection.display_label)
            return

        if selection.path is None:
            self._show_category_overview(selection)
            return

        if self.project is None:
            self.show_empty_state()
            return

        if selection.category_key == "distributori":
            self._show_distributore(selection.path)
            return

        if selection.category_key == "sequenze_test":
            self._show_test_sequence(selection.path)
            return

        if selection.category_key in {"curve_comando", "curve_limite", "rampe_xy"}:
            self._show_point_series(selection.category_key, selection.path)
            return

        if selection.category_key == "calibrazioni_ce16":
            self._show_ce16(selection.path)
            return

        if selection.category_key == "calibrazioni_mms2218":
            self._show_mms2218(selection.path)
            return

        if selection.category_key == "file_esterni":
            self._show_external_file(selection.display_label, source_path=selection.path)
            return

        self._show_output_viewer()

    def current_title(self) -> str:
        return self.title_label.text()

    def _set_title(self, selection: ResourceTreeSelection, subtitle: str) -> None:
        self.title_icon.setText(CATEGORY_ICONS.get(selection.category_key, "□"))
        self.title_label.setText(selection.display_label)
        self.subtitle_label.setText(subtitle)

    def _subtitle_for(self, selection: ResourceTreeSelection) -> str:
        if selection.path is None and selection.display_label == selection.category_label:
            return f"Categoria {selection.category_label}"
        return f"Editor {selection.category_label.lower()}"

    def _replace_content(self, widgets: list[QWidget]) -> None:
        while self.content_layout.count():
            item = self.content_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

        self.content_widgets = widgets[:]
        self.current_editor = widgets[0] if widgets else None

        if self.current_editor and hasattr(self.current_editor, "state_changed"):
            self.current_editor.state_changed.connect(self._update_titlebar_actions)

        for widget in widgets:
            self.content_layout.addWidget(widget)
        self.content_layout.addStretch(1)
        self._update_titlebar_actions()

    def _current_source_path(self) -> Path | None:
        if self.current_editor is None:
            return None
        return getattr(self.current_editor, "source_path", None)

    def _update_titlebar_actions(self) -> None:
        supports_save = bool(self.current_editor) and getattr(self.current_editor, "supports_save", False)
        is_dirty = bool(self.current_editor) and getattr(self.current_editor, "is_dirty", lambda: False)()
        supports_duplicate = bool(self.current_editor) and getattr(self.current_editor, "supports_duplicate", False)
        has_source = self._current_source_path() is not None

        self.save_button.setEnabled(supports_save and is_dirty)
        self.duplicate_button.setEnabled(supports_duplicate)
        self.rename_button.setEnabled(has_source)
        self.delete_button.setEnabled(has_source)
        self.right_panel.rinomina_button.setEnabled(has_source)
        self.right_panel.elimina_button.setEnabled(has_source)

        title = self.current_title().removesuffix(" *")
        if is_dirty:
            title += " *"
        self.title_label.setText(title)

        new_enabled = (
            self.project_root is not None
            and self._current_category_key is not None
            and supports_new_resource(self._current_category_key)
        )
        self.new_button.setEnabled(new_enabled)

    def _reload_project(self, selected_path: Path | None = None) -> None:
        if self.project_root is None:
            return

        self.project = load_project(self.project_root)
        self.resource_tree.set_project(self.project)

        if selected_path is not None and self.resource_tree.activate_path(selected_path):
            return

        self.show_empty_state()

    def save_current_resource(self) -> Path | None:
        if not self.current_editor or not getattr(self.current_editor, "supports_save", False):
            return None

        validate_before_save = getattr(self.current_editor, "validate_before_save", None)
        if validate_before_save is not None:
            blocking_messages = validate_before_save()
            if blocking_messages:
                self.right_panel.update_validation(blocking_messages, True)
                self._set_validation_badge("errore")
                return None

        save_changes = getattr(self.current_editor, "save_changes", None)
        if save_changes is None:
            return None

        saved_path = save_changes()
        if getattr(self.current_editor, "is_dirty", lambda: False)():
            self._reload_project(saved_path)
            # Reimpostare lo stato is_dirty a False se l'editor gestisce un modello interno,
            # ma il reload dell'albero e della UI ri-crea l'editor da zero, quindi va bene.
        return saved_path

    def _prompt_save_if_dirty(self) -> bool:
        if not self.current_editor or not getattr(self.current_editor, "is_dirty", lambda: False)():
            return True

        reply = QMessageBox.warning(
            self,
            "Modifiche non salvate",
            "Il file corrente contiene modifiche non salvate.\nVuoi salvarle prima di continuare?",
            QMessageBox.StandardButton.Save | QMessageBox.StandardButton.Discard | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Save,
        )

        if reply == QMessageBox.StandardButton.Save:
            return self.save_current_resource() is not None
        return reply == QMessageBox.StandardButton.Discard

    def create_new_in_current_category(self) -> Path | None:
        if not self._prompt_save_if_dirty():
            return None
        if self.project_root is None or self._current_category_key is None:
            return None
        if not supports_new_resource(self._current_category_key):
            return None

        new_path = create_new_resource(self._current_category_key, self.project_root)

        if self._current_category_key == "file_esterni":
            self._show_external_file(new_path.name, source_path=new_path)
            return new_path

        self._reload_project(new_path)
        return new_path

    def duplicate_current_resource(self) -> Path | None:
        if not self.current_editor or not getattr(self.current_editor, "supports_duplicate", False):
            return None

        duplicate_resource = getattr(self.current_editor, "duplicate_resource", None)
        if duplicate_resource is None:
            return None

        duplicate_path = duplicate_resource()
        if duplicate_path is not None:
            self._reload_project(duplicate_path)
        return duplicate_path

    def _reference_labels_for(self, path: Path) -> list[str]:
        if self.project is None:
            return []
        return sorted(target.name for target in self.project.get_uses(path))

    def _validation_labels_for(self, path: Path) -> list[str]:
        if self.project is None:
            return []
        return sorted(self.project.get_unresolved_references(path))

    def _show_category_overview(self, selection: ResourceTreeSelection) -> None:
        if selection.category_key == "output_banco":
            self._show_output_viewer()
            return

        summary_label = QLabel(f"Panoramica categoria: {selection.category_label}")
        summary_label.setStyleSheet("font-size: 15px; font-weight: 600;")

        detail_label = QLabel("Seleziona una risorsa figlia dall'albero per aprire il layout di editing a 3 colonne.")
        detail_label.setWordWrap(True)

        self._replace_content([summary_label, detail_label])
        self._current_reference_context = []
        self._set_validation_badge("ok")
        self.right_panel.set_context([], [])

    def _show_distributore(self, path: Path) -> None:
        assert self.project is not None
        self._current_reference_context = self._reference_labels_for(path)
        self._replace_content(
            [DistEditor(self.project, path, open_resource=self.open_resource_path, show_usages=self.show_resource_usages)]
        )
        validation = self._validation_labels_for(path)
        self._set_validation_badge("warning" if validation else "ok")
        self.right_panel.set_context(self._current_reference_context, validation, focus_validation=bool(validation))

    def _show_test_sequence(self, path: Path) -> None:
        assert self.project is not None
        logger.info(f"EditorShell: apertura sequenza test {path.name}")
        try:
            self._current_reference_context = self._reference_labels_for(path)
            self.right_panel.set_context(self._current_reference_context, [])
            self._replace_content(
                [
                    SeqEditor(
                        self.project,
                        path,
                        open_resource=self.open_resource_path,
                        show_usages=self.show_resource_usages,
                        update_validation=self.right_panel.update_validation,
                        update_status=self._set_validation_badge,
                    )
                ]
            )
            logger.info(f"EditorShell: sequenza test {path.name} aperta con successo")
        except Exception as e:
            logger.exception(f"EditorShell: ERRORE apertura sequenza test {path.name}: {e}")

    def _show_point_series(self, category_key: str, path: Path) -> None:
        assert self.project is not None
        resource = self.project.point_series_resources[path]
        editor_widget = {
            "curve_comando": CurveEditor(resource),
            "curve_limite": LimitEditor(resource),
            "rampe_xy": RampEditor(resource),
        }[category_key]
        self._replace_content([editor_widget])
        self._current_reference_context = self._reference_labels_for(path)
        self._set_validation_badge("ok")
        self.right_panel.set_context(self._current_reference_context, [])

    def _show_ce16(self, path: Path) -> None:
        assert self.project is not None
        resource = self.project.ce16_resources[path]
        self._replace_content([Ce16Editor(resource)])
        self._current_reference_context = self._reference_labels_for(path)
        self._set_validation_badge("ok")
        self.right_panel.set_context(self._current_reference_context, [])

    def _show_mms2218(self, path: Path) -> None:
        assert self.project is not None
        resource = self.project.mms2218_resources[path]
        self._replace_content([MmsEditor(resource)])
        self._current_reference_context = self._reference_labels_for(path)
        self._set_validation_badge("ok")
        self.right_panel.set_context(self._current_reference_context, [])

    def _show_external_file(self, display_label: str, source_path: Path | None = None) -> None:
        editor = ExternalFileEditor(display_label, source_path=source_path)
        self._replace_content([editor])
        self._current_reference_context = []
        if editor.supports_save:
            self._set_validation_badge("ok")
            self.right_panel.set_context([], [])
        else:
            self._set_validation_badge("warning")
            self.right_panel.set_context([], ["Riferimento esterno non risolto nel progetto"], focus_validation=True)

    def _show_output_viewer(self) -> None:
        initial_archive = self._default_output_archive_path()
        viewer = OutputViewer(initial_archive_path=initial_archive)
        self._replace_content([viewer])
        self._current_reference_context = []
        self._set_validation_badge("ok")
        self.right_panel.set_context([], [])

    def _default_output_archive_path(self) -> Path | None:
        if self.project_root is None:
            return None
        graph_dir = self.project_root / "GRAPH"
        if graph_dir.exists() and graph_dir.is_dir():
            return graph_dir
        return None

    def open_resource_path(self, path: Path) -> bool:
        if self.resource_tree.activate_path(path):
            return True
        # File risolto ma non nell'albero (es. file .txt parametri in FolderConfigurazioneModuli):
        # aprilo direttamente come ExternalFileEditor senza navigare l'albero.
        if path.exists():
            selection = ResourceTreeSelection("file_esterni", "File esterni", path.name, path)
            self._set_title(selection, "Editor file esterno")
            self._show_external_file(path.name, source_path=path)
            return True
        return False

    def delete_current_resource(self) -> None:
        path = self._current_source_path()
        if path is None or self.project is None:
            return

        impact = get_impact(self.project, path)
        dialog = ImpactDialog(
            title=f"Elimina {path.name}",
            description=(
                f"Stai per eliminare «{path.name}». Questa azione è irreversibile.\n"
                "I file che la referenziano conterranno riferimenti mancanti."
            ),
            impact_items=impact,
            action_label="Elimina",
            parent=self,
        )
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        delete_resource(path)
        self._reload_project()

    def rename_current_resource(self) -> None:
        path = self._current_source_path()
        if path is None or self.project is None:
            return

        impact = get_impact(self.project, path)
        dialog = ImpactDialog(
            title=f"Rinomina {path.name}",
            description=(
                f"Rinomina «{path.name}» e aggiorna automaticamente tutti i riferimenti nel progetto."
            ),
            impact_items=impact,
            action_label="Rinomina",
            rename_mode=True,
            initial_name=path.name,
            parent=self,
        )
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        new_name = dialog.new_name()
        if not new_name or new_name == path.name:
            return

        try:
            new_path, _ = rename_resource(self.project, path, new_name)
            self._reload_project(new_path)
        except ValueError as exc:
            QMessageBox.warning(self, "Rinomina non riuscita", str(exc))

    def show_resource_usages(self, path: Path) -> None:
        if self.project is None:
            self.right_panel.show_references([])
            return

        references = sorted(source.name for source in self.project.get_used_by(path))
        self.right_panel.show_references(references)

    def go_back(self) -> None:
        if self.history_index > 0:
            self.history_index -= 1
            self._navigate_to_history()

    def go_forward(self) -> None:
        if self.history_index < len(self.history) - 1:
            self.history_index += 1
            self._navigate_to_history()

    def _navigate_to_history(self) -> None:
        self._is_navigating = True
        try:
            selection = self.history[self.history_index]
            if not self.resource_tree.activate_selection(selection):
                self.show_selection(selection)
        finally:
            self._is_navigating = False
            self.history_changed.emit()
