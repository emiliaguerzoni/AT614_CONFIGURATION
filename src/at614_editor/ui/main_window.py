from __future__ import annotations

import logging
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction
from PySide6.QtWidgets import QMainWindow, QStackedWidget, QToolBar

from PySide6.QtWidgets import QDialog

from at614_editor.domain.project import load_project
from at614_editor.ui.dialogs.clone_program_dialog import CloneProgramDialog
from at614_editor.ui.editor_shell import EditorShell
from at614_editor.ui.workspace_home import WorkspaceHome

logger = logging.getLogger(__name__)


def discover_project_root(start_path: Path) -> Path | None:
    logger.debug(f"Ricerca root del progetto a partire da: {start_path}")
    # 1) Cerca un settings.ini AT614 nella directory corrente e nei suoi genitori
    from at614_editor.domain.settings_ini import discover_root_from_settings_ini  # noqa: PLC0415

    for candidate_dir in [start_path, *start_path.parents]:
        ini_path = candidate_dir / "settings.ini"
        if ini_path.is_file():
            logger.debug(f"Trovato settings.ini in: {ini_path}")
            root = discover_root_from_settings_ini(ini_path)
            if root is not None:
                logger.info(f"Root del progetto identificata da settings.ini: {root}")
                return root

    # 2) Fallback: cerca DISTRIBUTORE + TEST nella directory corrente e nelle sue figlie
    logger.debug("Fallback: ricerca di DISTRIBUTORE + TEST")
    candidates = [start_path]
    try:
        candidates.extend(item for item in start_path.iterdir() if item.is_dir())
    except OSError:
        candidates.extend([])

    for candidate in candidates:
        if (candidate / "DISTRIBUTORE").exists() and (candidate / "TEST").exists():
            logger.info(f"Root del progetto trovata in: {candidate}")
            return candidate
    logger.warning(f"Root del progetto non trovato a partire da: {start_path}")
    return None


class MainWindow(QMainWindow):
    def __init__(self, project_root: Path | None = None, parent=None) -> None:
        super().__init__(parent)
        logger.info(f"MainWindow: inizializzazione con project_root={project_root}")
        self.project_root = project_root
        try:
            self.project = load_project(project_root) if project_root is not None else None
            if self.project:
                logger.info(f"MainWindow: progetto caricato con {len(self.project.distributori)} distribuitori")
            else:
                logger.warning("MainWindow: nessun progetto caricato")
        except Exception as e:
            logger.exception(f"MainWindow: errore durante il caricamento del progetto: {e}")
            self.project = None

        self.setWindowTitle("AT614 Configuration Editor")
        self.resize(1480, 920)
        self.setMinimumSize(1280, 820)

        self.home_page = WorkspaceHome(project_root, self.project)
        self.editor_shell = EditorShell(project_root, self.project)

        self.pages = QStackedWidget()
        self.pages.addWidget(self.home_page)
        self.pages.addWidget(self.editor_shell)
        self.setCentralWidget(self.pages)

        self._build_toolbar()
        self._connect_signals()
        self._apply_light_styles()
        self.show_home()

    def closeEvent(self, event) -> None:
        if not self.editor_shell._prompt_save_if_dirty():
            event.ignore()
        else:
            event.accept()

    def _build_toolbar(self) -> None:
        toolbar = QToolBar("Navigazione")
        toolbar.setMovable(False)
        self.addToolBar(Qt.ToolBarArea.TopToolBarArea, toolbar)

        home_action = QAction("Dashboard", self)
        home_action.triggered.connect(self.show_home)
        toolbar.addAction(home_action)

        shell_action = QAction("Shell", self)
        shell_action.triggered.connect(self.show_editor_shell)
        toolbar.addAction(shell_action)

        toolbar.addSeparator()

        self.back_action = QAction("← Indietro", self)
        self.back_action.triggered.connect(self.editor_shell.go_back)
        self.back_action.setEnabled(False)
        toolbar.addAction(self.back_action)

        self.forward_action = QAction("→ Avanti", self)
        self.forward_action.triggered.connect(self.editor_shell.go_forward)
        self.forward_action.setEnabled(False)
        toolbar.addAction(self.forward_action)

    def _connect_signals(self) -> None:
        self.home_page.open_shell_requested.connect(self.show_editor_shell)
        self.home_page.open_first_resource_requested.connect(self.open_first_available_resource)
        self.home_page.duplicate_program_requested.connect(self.open_clone_program_dialog)
        self.editor_shell.resource_tree.resource_activated.connect(lambda _selection: self.show_editor_shell())
        self.editor_shell.history_changed.connect(self._update_navigation_actions)

    def _update_navigation_actions(self) -> None:
        can_back = self.editor_shell.history_index > 0
        can_forward = self.editor_shell.history_index < len(self.editor_shell.history) - 1
        self.back_action.setEnabled(can_back)
        self.forward_action.setEnabled(can_forward)

    def _apply_light_styles(self) -> None:
        self.setStyleSheet(
            "QMainWindow { background: #F5F7FA; }"
            "QToolBar { background: #FFFFFF; border-bottom: 1px solid #D0D5DD; spacing: 6px; padding: 6px; }"
            "QWidget { color: #101828; }"
            "QPushButton { background: #FFFFFF; border: 1px solid #D0D5DD; border-radius: 8px; padding: 6px 10px; }"
            "QPushButton:hover { background: #F2F4F7; }"
            "QTreeWidget, QListWidget, QTabWidget::pane, QScrollArea { background: #FFFFFF; border: 1px solid #D0D5DD; border-radius: 8px; }"
        )

    def show_home(self) -> None:
        self.pages.setCurrentWidget(self.home_page)

    def show_editor_shell(self) -> None:
        self.pages.setCurrentWidget(self.editor_shell)

    def active_page_name(self) -> str:
        return "home" if self.pages.currentWidget() is self.home_page else "editor"

    def open_first_available_resource(self) -> None:
        selection = self.editor_shell.resource_tree.activate_first_resource()
        if selection is not None:
            self.show_editor_shell()

    def open_clone_program_dialog(self, source_distributore: Path | None = None) -> Path | None:
        if self.project is None or not self.project.distributori:
            return None

        dialog = CloneProgramDialog(self.project, source_distributore=source_distributore, parent=self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return None

        new_path = dialog.new_distributore_path()
        if new_path is None:
            return None

        self.editor_shell._reload_project(new_path)
        self.show_editor_shell()
        return new_path
