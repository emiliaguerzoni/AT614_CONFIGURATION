from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Callable

from PySide6.QtCore import Qt
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
from PySide6.QtCore import QDate

from at614_editor.domain.output_archive import (
    OutputArchive,
    OutputFile,
    ensure_archive_index,
    read_csv_columns,
)
from at614_editor.ui.components.chart_stage import ChartStage
from at614_editor.ui.components.read_only_banner import ReadOnlyBanner


class OutputViewer(QWidget):
    """Viewer sola lettura per archivi output del banco."""

    editor_name = "OutputViewer"
    supports_save = False
    supports_duplicate = False

    def __init__(
        self,
        initial_archive_path: Path | None = None,
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

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        layout.addWidget(ReadOnlyBanner())

        archive_row = QHBoxLayout()
        archive_label = QLabel("Archivio output:")
        archive_label.setStyleSheet("font-weight: 600;")
        self.archive_path_label = QLabel("— non selezionato —")
        self.archive_path_label.setStyleSheet("color: #475467;")
        self.archive_path_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.configure_archive_button = QPushButton("Configura archivio output")
        self.configure_archive_button.clicked.connect(self._on_configure_archive)
        self.refresh_button = QPushButton("Aggiorna indice")
        self.refresh_button.clicked.connect(self._on_refresh_archive)
        self.refresh_button.setEnabled(False)

        archive_row.addWidget(archive_label)
        archive_row.addWidget(self.archive_path_label, 1)
        archive_row.addWidget(self.configure_archive_button)
        archive_row.addWidget(self.refresh_button)
        layout.addLayout(archive_row)

        filters_row = QHBoxLayout()
        self.search_field = QLineEdit()
        self.search_field.setPlaceholderText("Filtra per nome o cartella…")
        self.search_field.setToolTip("Filtra i file per nome o percorso cartella. La ricerca è case-insensitive.")
        self.search_field.textChanged.connect(self._apply_filters)

        self.created_from = QDateEdit()
        self.created_from.setSpecialValueText("data creazione da")
        self.created_from.setCalendarPopup(True)
        self.created_from.setDate(QDate(2000, 1, 1))
        self.created_from.setMinimumDate(QDate(2000, 1, 1))
        self.created_from.setToolTip("Mostra solo i file creati a partire da questa data.")
        self.created_from.dateChanged.connect(self._apply_filters)

        self.created_to = QDateEdit()
        self.created_to.setSpecialValueText("data creazione a")
        self.created_to.setCalendarPopup(True)
        self.created_to.setDate(QDate.currentDate())
        self.created_to.setToolTip("Mostra solo i file creati fino a questa data.")
        self.created_to.dateChanged.connect(self._apply_filters)

        filters_row.addWidget(QLabel("Filtri:"))
        filters_row.addWidget(self.search_field, 1)
        filters_row.addWidget(QLabel("Creato dal"))
        filters_row.addWidget(self.created_from)
        filters_row.addWidget(QLabel("al"))
        filters_row.addWidget(self.created_to)
        layout.addLayout(filters_row)

        self.status_label = QLabel("Configura un archivio output per iniziare.")
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
        self.results_table.itemSelectionChanged.connect(self._on_selection_changed)
        layout.addWidget(self.results_table, 1)

        axis_row = QHBoxLayout()
        axis_row.addWidget(QLabel("Asse X:"))
        self.x_axis_combo = QComboBox()
        self.x_axis_combo.setToolTip("Seleziona la colonna da usare come asse X (orizzontale) del grafico.")
        self.x_axis_combo.currentIndexChanged.connect(self._refresh_plot_summary)
        axis_row.addWidget(self.x_axis_combo)

        axis_row.addWidget(QLabel("Assi Y:"))
        self.y_axes_list = QListWidget()
        self.y_axes_list.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        self.y_axes_list.setMaximumHeight(80)
        self.y_axes_list.setToolTip("Seleziona una o più colonne da rappresentare come assi Y (verticali). Clicca con Ctrl per selezioni multiple.")
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

        if initial_archive_path is not None:
            self.load_archive(initial_archive_path)

    # ------------------------------------------------------------------ archive

    def load_archive(self, root_path: Path) -> None:
        try:
            self.archive = ensure_archive_index(root_path)
        except ValueError as exc:
            self.status_label.setText(str(exc))
            self.archive = None
            self.filtered_files = []
            self._render_results()
            self.refresh_button.setEnabled(False)
            return

        self.archive_path_label.setText(str(root_path))
        self.refresh_button.setEnabled(True)
        self._apply_filters()

    def _on_configure_archive(self) -> None:
        chosen = QFileDialog.getExistingDirectory(self, "Seleziona archivio output")
        if not chosen:
            return
        self.load_archive(Path(chosen))

    def _on_refresh_archive(self) -> None:
        if self.archive is None:
            return
        self.load_archive(self.archive.root_path)

    # ------------------------------------------------------------------ filters

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

    # ------------------------------------------------------------------ table

    def _render_results(self) -> None:
        self.results_table.setRowCount(len(self.filtered_files))
        if self.archive is None:
            self.status_label.setText("Configura un archivio output per iniziare.")
        else:
            self.status_label.setText(
                f"{len(self.filtered_files)} risultati su {len(self.archive.files)} file indicizzati."
            )

        for row_index, entry in enumerate(self.filtered_files):
            self.results_table.setItem(row_index, 0, QTableWidgetItem(entry.folder or "—"))
            self.results_table.setItem(row_index, 1, QTableWidgetItem(entry.name))
            self.results_table.setItem(row_index, 2, QTableWidgetItem(_format_timestamp(entry.created_at)))
            self.results_table.setItem(row_index, 3, QTableWidgetItem(_format_timestamp(entry.modified_at)))
            self.results_table.setItem(row_index, 4, QTableWidgetItem(_format_size(entry.size_bytes)))

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
        try:
            headers, rows_data = read_csv_columns(target_path)
        except ValueError as exc:
            self.status_label.setText(f"Impossibile leggere {entry.name}: {exc}")
            return

        self.csv_headers = headers
        self.csv_rows = rows_data
        self.x_axis_combo.blockSignals(True)
        self.x_axis_combo.clear()
        for header in headers:
            self.x_axis_combo.addItem(header)
        self.x_axis_combo.blockSignals(False)

        self.y_axes_list.blockSignals(True)
        self.y_axes_list.clear()
        for header in headers:
            self.y_axes_list.addItem(header)
        # Auto-seleziona tutte le colonne tranne la prima (X) come assi Y
        for i in range(1, self.y_axes_list.count()):
            self.y_axes_list.item(i).setSelected(True)
        self.y_axes_list.blockSignals(False)

        self._set_actions_enabled(True)
        self._refresh_plot_summary()

    # ------------------------------------------------------------------ plotting

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

    # ------------------------------------------------------------------ actions

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
        self.status_label.setText(f"Confronto richiesto tra {path.name} e {Path(target).name} — overlay non ancora disponibile.")

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
        else:
            self.status_label.setText(f"Cartella: {path.parent}")


def _format_size(num_bytes: int) -> str:
    if num_bytes < 1024:
        return f"{num_bytes} B"
    if num_bytes < 1024 * 1024:
        return f"{num_bytes / 1024:.1f} KB"
    return f"{num_bytes / (1024 * 1024):.2f} MB"


def _format_timestamp(timestamp: float) -> str:
    return datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d %H:%M")
