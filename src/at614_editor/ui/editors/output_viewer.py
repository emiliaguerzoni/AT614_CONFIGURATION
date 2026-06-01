from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Callable

from PySide6.QtCore import Qt, QDate, QThread, QObject, Signal
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDateEdit,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from at614_editor.domain.output_archive import (
    OutputArchive,
    OutputFile,
    ensure_archive_index,
    read_csv_columns,
)
from at614_editor.ui.components.chart_stage import ChartStage


MAX_TABLE_ROWS = 500


class ArchiveLoader(QObject):
    finished = Signal(object)
    error = Signal(str)

    def __init__(self, root_path: Path) -> None:
        super().__init__()
        self.root_path = root_path

    def run(self) -> None:
        try:
            archive = ensure_archive_index(self.root_path)
        except Exception as exc:
            self.error.emit(str(exc))
            return
        self.finished.emit(archive)


class CsvLoader(QObject):
    finished = Signal(list, list)
    error = Signal(str)

    def __init__(self, path: Path) -> None:
        super().__init__()
        self.path = path

    def run(self) -> None:
        try:
            headers, rows = read_csv_columns(self.path)
        except Exception as exc:
            self.error.emit(str(exc))
            return
        self.finished.emit(headers, rows)


class OutputViewer(QWidget):
    """Viewer sola lettura per archivi output del banco — lazy load."""

    editor_name = "OutputViewer"
    supports_save = False
    supports_duplicate = False

    def __init__(
        self,
        initial_archive_path: Path | None = None,
        initial_selected_path: Path | None = None,
        open_folder_callback: Callable[[Path], None] | None = None,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.archive: OutputArchive | None = None
        self.filtered_files: list[OutputFile] = []
        self.selected_file: OutputFile | None = None
        self.csv_headers: list[str] = []
        self.csv_rows: list[list[str]] = []
        self.open_folder_callback = open_folder_callback
        self._archive_thread: QThread | None = None
        self._archive_loader: ArchiveLoader | None = None
        self._csv_thread: QThread | None = None
        self._csv_loader: CsvLoader | None = None
        self._initial_archive_path: Path | None = initial_archive_path
        self._initial_selected_path: Path | None = initial_selected_path
        self.destroyed.connect(self._cleanup_threads)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        # ── toolbar archivio ──────────────────────────────────────────────
        archive_toolbar = QHBoxLayout()
        archive_toolbar.setContentsMargins(0, 0, 0, 0)
        archive_label = QLabel("Archivio output:")
        archive_label.setStyleSheet("font-weight: 600;")
        self.archive_path_label = QLabel("— non selezionato —")
        self.archive_path_label.setStyleSheet("color: #475467;")
        self.archive_path_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)

        self.load_button = QPushButton("Carica ▶")
        self.load_button.setToolTip("Avvia la scansione della cartella e carica l'elenco file")
        self.load_button.clicked.connect(self._on_load_clicked)
        self.load_button.setEnabled(False)

        self.configure_archive_button = QPushButton("Sfoglia…")
        self.configure_archive_button.setToolTip("Scegli la cartella dell'archivio output")
        self.configure_archive_button.clicked.connect(self._on_configure_archive)

        self.refresh_button = QPushButton("Aggiorna indice")
        self.refresh_button.setToolTip("Riscansiona la cartella anche se il cache è aggiornato")
        self.refresh_button.clicked.connect(self._on_refresh_archive)
        self.refresh_button.setEnabled(False)

        self.toggle_filters_button = QPushButton("Mostra filtri")
        self.toggle_filters_button.setCheckable(True)
        self.toggle_filters_button.toggled.connect(self._toggle_filter_panel)

        archive_toolbar.addWidget(archive_label)
        archive_toolbar.addWidget(self.archive_path_label, 1)
        archive_toolbar.addWidget(self.load_button)
        archive_toolbar.addWidget(self.configure_archive_button)
        archive_toolbar.addWidget(self.refresh_button)
        archive_toolbar.addWidget(self.toggle_filters_button)
        layout.addLayout(archive_toolbar)

        # ── filtri ────────────────────────────────────────────────────────
        self.filter_panel = QWidget()
        self.filter_panel.setVisible(False)
        filter_layout = QHBoxLayout(self.filter_panel)
        filter_layout.setContentsMargins(0, 0, 0, 0)

        self.search_field = QLineEdit()
        self.search_field.setPlaceholderText("Filtra per nome o cartella…")
        self.search_field.setToolTip("Filtra i file per nome o percorso cartella.")
        self.search_field.textChanged.connect(self._apply_filters)

        self.created_from = QDateEdit()
        self.created_from.setSpecialValueText("da data")
        self.created_from.setCalendarPopup(True)
        self.created_from.setDate(QDate(2000, 1, 1))
        self.created_from.setMinimumDate(QDate(2000, 1, 1))
        self.created_from.setToolTip("Mostra solo i file creati a partire da questa data.")
        self.created_from.dateChanged.connect(self._apply_filters)

        self.created_to = QDateEdit()
        self.created_to.setSpecialValueText("a data")
        self.created_to.setCalendarPopup(True)
        self.created_to.setDate(QDate.currentDate())
        self.created_to.setToolTip("Mostra solo i file creati fino a questa data.")
        self.created_to.dateChanged.connect(self._apply_filters)

        filter_layout.addWidget(QLabel("Filtri:"))
        filter_layout.addWidget(self.search_field, 1)
        filter_layout.addWidget(QLabel("Creato dal"))
        filter_layout.addWidget(self.created_from)
        filter_layout.addWidget(QLabel("al"))
        filter_layout.addWidget(self.created_to)
        layout.addWidget(self.filter_panel)

        # ── status + tabella ──────────────────────────────────────────────
        self.status_label = QLabel("Seleziona una cartella archivio e premi 'Carica ▶' per iniziare.")
        self.status_label.setStyleSheet("color: #475467;")
        layout.addWidget(self.status_label)

        self.results_table = QTableWidget(0, 5)
        self.results_table.setHorizontalHeaderLabels(
            ["Cartella", "File", "Creato", "Modificato", "Dimensione"]
        )
        self.results_table.verticalHeader().setVisible(False)
        self.results_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.results_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.results_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.results_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.results_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.results_table.setColumnWidth(2, 140)
        self.results_table.setColumnWidth(3, 140)
        self.results_table.setColumnWidth(4, 100)
        self.results_table.setAlternatingRowColors(True)
        self.results_table.itemSelectionChanged.connect(self._on_selection_changed)
        layout.addWidget(self.results_table, 1)

        # ── assi grafico ──────────────────────────────────────────────────
        axis_row = QHBoxLayout()
        axis_row.addWidget(QLabel("Asse X:"))
        self.x_axis_combo = QComboBox()
        self.x_axis_combo.setToolTip("Seleziona la colonna come asse X del grafico.")
        self.x_axis_combo.currentIndexChanged.connect(self._refresh_plot_summary)
        axis_row.addWidget(self.x_axis_combo)

        axis_row.addWidget(QLabel("Assi Y:"))
        self.y_axes_list = QListWidget()
        self.y_axes_list.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        self.y_axes_list.setMaximumHeight(80)
        self.y_axes_list.setToolTip("Seleziona una o più colonne come assi Y. Ctrl+click per selezione multipla.")
        self.y_axes_list.itemSelectionChanged.connect(self._refresh_plot_summary)
        axis_row.addWidget(self.y_axes_list, 1)
        layout.addLayout(axis_row)

        self.chart_stage = ChartStage(
            title="Output banco",
            legend_text="X / Y selezionabili",
            toolbar_actions=["Esporta", "Confronta", "Stampa", "Apri cartella"],
        )
        for action_button in self.chart_stage.toolbar_buttons:
            label = action_button.text()
            if label == "Esporta":
                action_button.clicked.connect(self._on_export)
            elif label == "Confronta":
                action_button.clicked.connect(self._on_compare)
            elif label == "Stampa":
                action_button.clicked.connect(self._on_print)
            elif label == "Apri cartella":
                action_button.clicked.connect(self._on_open_folder)
        self._set_actions_enabled(False)
        layout.addWidget(self.chart_stage)

        # ── init con path pre-configurata ─────────────────────────────────
        if initial_archive_path is not None:
            self._set_archive_path(initial_archive_path)

    # ── archive path setup ────────────────────────────────────────────────

    def _set_archive_path(self, path: Path) -> None:
        """Imposta la cartella archivio senza scansionare. Abilita il pulsante Carica."""
        self._initial_archive_path = path
        self.archive_path_label.setText(str(path))
        self.load_button.setEnabled(True)
        self.status_label.setText(f"Archivio: {path.name} — premi 'Carica ▶' per scansionare.")

    # ── archive loading ───────────────────────────────────────────────────

    def _on_load_clicked(self) -> None:
        if self._initial_archive_path is not None:
            self.load_archive(self._initial_archive_path)
        elif self.archive is not None:
            self.load_archive(self.archive.root_path)

    def load_archive(self, root_path: Path) -> None:
        self.archive_path_label.setText(str(root_path))
        self.status_label.setText("Indicizzazione archivio in corso…")
        self.results_table.setRowCount(0)
        self.load_button.setEnabled(False)
        self.refresh_button.setEnabled(False)
        self._set_actions_enabled(False)

        self._stop_archive_thread()

        self._archive_loader = ArchiveLoader(root_path)
        self._archive_thread = QThread()
        self._archive_loader.moveToThread(self._archive_thread)
        self._archive_thread.started.connect(self._archive_loader.run)
        self._archive_loader.finished.connect(self._on_archive_loaded)
        self._archive_loader.error.connect(self._on_archive_load_error)
        self._archive_loader.finished.connect(self._archive_thread.quit)
        self._archive_loader.error.connect(self._archive_thread.quit)
        self._archive_thread.finished.connect(self._archive_loader.deleteLater)
        self._archive_thread.finished.connect(self._archive_thread.deleteLater)
        self._archive_thread.start()

    def _stop_archive_thread(self) -> None:
        if self._archive_thread is None:
            return
        if self._archive_thread.isRunning():
            self._archive_thread.quit()
            if not self._archive_thread.wait(5000):
                self._archive_thread.terminate()
                self._archive_thread.wait(1000)
        self._archive_thread = None
        self._archive_loader = None

    def _stop_csv_thread(self) -> None:
        if self._csv_thread is None:
            return
        if self._csv_thread.isRunning():
            self._csv_thread.quit()
            if not self._csv_thread.wait(3000):
                self._csv_thread.terminate()
                self._csv_thread.wait(500)
        self._csv_thread = None
        self._csv_loader = None

    def _cleanup_threads(self) -> None:
        self._stop_archive_thread()
        self._stop_csv_thread()

    def _on_archive_loaded(self, archive: OutputArchive) -> None:
        self.archive = archive
        self._initial_archive_path = None
        self.load_button.setEnabled(False)
        self.refresh_button.setEnabled(True)
        self._apply_filters()
        if self._initial_selected_path is not None:
            self._select_initial_path(self._initial_selected_path)
            self._initial_selected_path = None

    def _on_archive_load_error(self, message: str) -> None:
        self.status_label.setText(f"Errore: {message}")
        self.archive = None
        self.filtered_files = []
        self._render_results()
        self.load_button.setEnabled(True)
        self.refresh_button.setEnabled(False)
        self._archive_thread = None
        self._archive_loader = None

    def _on_configure_archive(self) -> None:
        chosen = QFileDialog.getExistingDirectory(self, "Seleziona archivio output")
        if not chosen:
            return
        self._set_archive_path(Path(chosen))
        self.archive = None
        self.filtered_files = []
        self._render_results()

    def _on_refresh_archive(self) -> None:
        if self.archive is None and self._initial_archive_path is not None:
            self.load_archive(self._initial_archive_path)
        elif self.archive is not None:
            self.load_archive(self.archive.root_path)

    def _toggle_filter_panel(self, visible: bool) -> None:
        self.filter_panel.setVisible(visible)
        self.toggle_filters_button.setText("Nascondi filtri" if visible else "Mostra filtri")

    # ── filters ───────────────────────────────────────────────────────────

    def _apply_filters(self) -> None:
        if self.archive is None:
            self.filtered_files = []
            self._render_results()
            return

        created_from = self._qdate_to_datetime(self.created_from.date(), end_of_day=False)
        created_to = self._qdate_to_datetime(self.created_to.date(), end_of_day=True)

        self.filtered_files = self.archive.filter(
            text=self.search_field.text(),
            created_from=created_from,
            created_to=created_to,
        )
        self._render_results()

    def _qdate_to_datetime(self, qdate: QDate, end_of_day: bool) -> datetime | None:
        if qdate == QDate(2000, 1, 1) and not end_of_day:
            return None
        py_date = qdate.toPython()
        if end_of_day:
            return datetime.combine(py_date, datetime.max.time())
        return datetime.combine(py_date, datetime.min.time())

    # ── table (virtuale, max MAX_TABLE_ROWS righe) ────────────────────────

    def _render_results(self) -> None:
        visible = self.filtered_files[:MAX_TABLE_ROWS]
        self.results_table.setRowCount(len(visible))

        if self.archive is None:
            self.status_label.setText("Seleziona una cartella archivio e premi 'Carica ▶' per iniziare.")
        elif len(self.filtered_files) > MAX_TABLE_ROWS:
            self.status_label.setText(
                f"Mostrando {MAX_TABLE_ROWS} di {len(self.filtered_files)} file "
                f"(su {len(self.archive.files)} totali) — affina i filtri per altri risultati."
            )
        else:
            self.status_label.setText(
                f"{len(self.filtered_files)} file su {len(self.archive.files)} totali."
            )

        for row_index, entry in enumerate(visible):
            self.results_table.setItem(row_index, 0, QTableWidgetItem(entry.folder or "—"))
            self.results_table.setItem(row_index, 1, QTableWidgetItem(entry.name))
            self.results_table.setItem(row_index, 2, QTableWidgetItem(_format_timestamp(entry.created_at)))
            self.results_table.setItem(row_index, 3, QTableWidgetItem(_format_timestamp(entry.modified_at)))
            self.results_table.setItem(row_index, 4, QTableWidgetItem(_format_size(entry.size_bytes)))

    # ── CSV async ─────────────────────────────────────────────────────────

    def _on_selection_changed(self) -> None:
        rows = self.results_table.selectionModel().selectedRows()
        if not rows or self.archive is None:
            self.selected_file = None
            self.x_axis_combo.clear()
            self.y_axes_list.clear()
            self._set_actions_enabled(False)
            return

        row_index = rows[0].row()
        if row_index < 0 or row_index >= len(self.filtered_files):
            return

        entry = self.filtered_files[row_index]
        self.selected_file = entry
        target_path = self.archive.root_path / entry.relative_path

        self._stop_csv_thread()
        self.status_label.setText(f"Lettura {entry.name}…")
        self.x_axis_combo.clear()
        self.y_axes_list.clear()
        self._set_actions_enabled(False)

        self._csv_loader = CsvLoader(target_path)
        self._csv_thread = QThread()
        self._csv_loader.moveToThread(self._csv_thread)
        self._csv_thread.started.connect(self._csv_loader.run)
        self._csv_loader.finished.connect(self._on_csv_loaded)
        self._csv_loader.error.connect(self._on_csv_load_error)
        self._csv_loader.finished.connect(self._csv_thread.quit)
        self._csv_loader.error.connect(self._csv_thread.quit)
        self._csv_thread.finished.connect(self._csv_loader.deleteLater)
        self._csv_thread.finished.connect(self._csv_thread.deleteLater)
        self._csv_thread.start()

    def _on_csv_loaded(self, headers: list, rows_data: list) -> None:
        self.csv_headers = headers
        self.csv_rows = rows_data
        self._csv_thread = None
        self._csv_loader = None

        self.x_axis_combo.blockSignals(True)
        self.x_axis_combo.clear()
        for header in headers:
            self.x_axis_combo.addItem(header)
        self.x_axis_combo.blockSignals(False)

        self.y_axes_list.blockSignals(True)
        self.y_axes_list.clear()
        for header in headers:
            self.y_axes_list.addItem(header)
        for i in range(1, self.y_axes_list.count()):
            self.y_axes_list.item(i).setSelected(True)
        self.y_axes_list.blockSignals(False)

        self._set_actions_enabled(True)
        self._refresh_plot_summary()

        total = len(self.archive.files) if self.archive else 0
        self.status_label.setText(
            f"{len(self.filtered_files)} file su {total} totali. "
            f"File aperto: {self.selected_file.name if self.selected_file else ''} "
            f"({len(rows_data)} righe, {len(headers)} colonne)."
        )

    def _on_csv_load_error(self, message: str) -> None:
        self.status_label.setText(f"Errore lettura file: {message}")
        self._csv_thread = None
        self._csv_loader = None

    # ── plotting ──────────────────────────────────────────────────────────

    def _refresh_plot_summary(self) -> None:
        if self.selected_file is None:
            self.chart_stage.title_label.setText("Output banco")
            self.chart_stage.legend_label.setText("X / Y selezionabili")
            self.chart_stage.plot_series([], [])
            return

        x_label = self.x_axis_combo.currentText() or ""
        y_selected = [item.text() for item in self.y_axes_list.selectedItems()]
        y_label = " · ".join(y_selected) if y_selected else "—"

        self.chart_stage.title_label.setText(self.selected_file.name)
        self.chart_stage.legend_label.setText(f"X: {x_label} - Y: {y_label}")

        if x_label and y_selected and self.csv_headers:
            x_idx = self.csv_headers.index(x_label) if x_label in self.csv_headers else 0
            y_indices = [self.csv_headers.index(y) for y in y_selected if y in self.csv_headers]

            filtered_rows: list[list[str]] = []
            for row in self.csv_rows:
                if len(row) > x_idx:
                    filtered_row = [row[x_idx]] + [
                        row[yi] if len(row) > yi else "" for yi in y_indices
                    ]
                    filtered_rows.append(filtered_row)

            self.chart_stage.plot_series(filtered_rows, [x_label] + y_selected)
        else:
            self.chart_stage.plot_series([], [])

    # ── actions ───────────────────────────────────────────────────────────

    def _set_actions_enabled(self, enabled: bool) -> None:
        for action_button in self.chart_stage.toolbar_buttons:
            action_button.setEnabled(enabled)

    def selected_file_path(self) -> Path | None:
        if self.archive is None or self.selected_file is None:
            return None
        return self.archive.root_path / self.selected_file.relative_path

    def _on_export(self) -> None:
        path = self.selected_file_path()
        if path is None:
            return
        target, _ = QFileDialog.getSaveFileName(
            self, "Esporta output", path.name, "CSV (*.csv);;Tutti i file (*)"
        )
        if not target:
            return
        Path(target).write_bytes(path.read_bytes())

    def _on_compare(self) -> None:
        path = self.selected_file_path()
        if path is None or self.archive is None:
            return
        target, _ = QFileDialog.getOpenFileName(
            self,
            "Confronta con file output",
            str(self.archive.root_path),
            "CSV (*.csv);;Tutti i file (*)",
        )
        if not target:
            return
        self.status_label.setText(
            f"Confronto tra {path.name} e {Path(target).name} — overlay non ancora disponibile."
        )

    def _on_print(self) -> None:
        if self.selected_file is None:
            return
        self.status_label.setText(f"Stampa di {self.selected_file.name} non ancora implementata.")

    def _on_open_folder(self) -> None:
        path = self.selected_file_path()
        if path is None:
            return
        if self.open_folder_callback is not None:
            self.open_folder_callback(path.parent)

    def _select_initial_path(self, selected_path: Path) -> None:
        if self.archive is None:
            return
        if not selected_path.exists() or selected_path.is_dir():
            return
        rel = selected_path.relative_to(self.archive.root_path).as_posix()
        for index, entry in enumerate(self.filtered_files):
            if entry.relative_path == rel and index < MAX_TABLE_ROWS:
                self.results_table.selectRow(index)
                self.results_table.scrollToItem(self.results_table.item(index, 0))
                return


def _format_size(num_bytes: int) -> str:
    if num_bytes < 1024:
        return f"{num_bytes} B"
    if num_bytes < 1024 * 1024:
        return f"{num_bytes / 1024:.1f} KB"
    return f"{num_bytes / (1024 * 1024):.2f} MB"


def _format_timestamp(timestamp: float) -> str:
    return datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d %H:%M")
