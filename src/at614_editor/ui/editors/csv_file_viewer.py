from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QVBoxLayout,
    QWidget,
)

from at614_editor.domain.output_archive import read_csv_columns
from at614_editor.ui.components.chart_stage import ChartStage
from at614_editor.ui.editors.output_viewer import CsvLoader


class CsvFileViewer(QWidget):
    """Visualizzatore diretto di un singolo file CSV di output banco.

    Non esegue scansioni di cartella né indicizzazione archivio:
    legge un solo file in background e mostra subito il grafico.
    """

    editor_name = "CsvFileViewer"
    supports_save = False
    supports_duplicate = False

    def __init__(self, file_path: Path, parent=None) -> None:
        super().__init__(parent)
        self._file_path = file_path
        self._csv_loader: CsvLoader | None = None
        self.csv_headers: list[str] = []
        self.csv_rows: list[list[str]] = []
        self.destroyed.connect(self._cleanup)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        # ── info bar ──────────────────────────────────────────────────────
        info_bar = QHBoxLayout()
        info_bar.setContentsMargins(0, 0, 0, 0)
        name_label = QLabel(file_path.name)
        name_label.setStyleSheet("font-weight: 600;")
        folder_label = QLabel(str(file_path.parent))
        folder_label.setStyleSheet("color: #475467; font-size: 11px;")
        folder_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        info_bar.addWidget(name_label)
        info_bar.addWidget(folder_label, 1)
        layout.addLayout(info_bar)

        # ── status ────────────────────────────────────────────────────────
        self.status_label = QLabel("Lettura file in corso\u2026")
        self.status_label.setStyleSheet("color: #475467;")
        layout.addWidget(self.status_label)

        # ── selezione assi ────────────────────────────────────────────────
        axis_row = QHBoxLayout()
        axis_row.addWidget(QLabel("Asse X:"))
        self.x_axis_combo = QComboBox()
        self.x_axis_combo.setToolTip("Colonna da usare come asse X del grafico.")
        self.x_axis_combo.currentIndexChanged.connect(self._refresh_plot)
        axis_row.addWidget(self.x_axis_combo)

        axis_row.addWidget(QLabel("Assi Y:"))
        self.y_axes_list = QListWidget()
        self.y_axes_list.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        self.y_axes_list.setMaximumHeight(80)
        self.y_axes_list.setToolTip(
            "Seleziona una o pi\u00f9 colonne come assi Y. Ctrl+click per selezione multipla."
        )
        self.y_axes_list.itemSelectionChanged.connect(self._refresh_plot)
        axis_row.addWidget(self.y_axes_list, 1)
        layout.addLayout(axis_row)

        # ── grafico ───────────────────────────────────────────────────────
        self.chart_stage = ChartStage(
            title=file_path.name,
            legend_text="X / Y selezionabili",
            toolbar_actions=["Esporta", "Stampa", "Apri cartella"],
        )
        layout.addWidget(self.chart_stage, 1)

        # Avvia lettura in background
        self._load_file()

    # ── caricamento asincrono ─────────────────────────────────────────────

    def _load_file(self) -> None:
        self._csv_loader = CsvLoader(self._file_path)
        self._csv_loader.csv_ready.connect(self._on_loaded)
        self._csv_loader.load_error.connect(self._on_error)
        self._csv_loader.finished.connect(self._on_csv_thread_done)
        self._csv_loader.start()

    def _on_loaded(self, headers: list, rows_data: list) -> None:
        self.csv_headers = headers
        self.csv_rows = rows_data

        self.x_axis_combo.blockSignals(True)
        self.x_axis_combo.clear()
        for h in headers:
            self.x_axis_combo.addItem(h)
        self.x_axis_combo.blockSignals(False)

        self.y_axes_list.blockSignals(True)
        self.y_axes_list.clear()
        for h in headers:
            self.y_axes_list.addItem(h)
        for i in range(1, self.y_axes_list.count()):
            self.y_axes_list.item(i).setSelected(True)
        self.y_axes_list.blockSignals(False)

        self.status_label.setText(
            f"{len(rows_data)} righe \u00b7 {len(headers)} colonne \u2014 {self._file_path.name}"
        )
        self._refresh_plot()

    def _on_error(self, message: str) -> None:
        self.status_label.setText(f"Errore lettura: {message}")

    def _on_csv_thread_done(self) -> None:
        if self._csv_loader is not None:
            self._csv_loader.deleteLater()
            self._csv_loader = None

    # ── grafico ───────────────────────────────────────────────────────────

    def _refresh_plot(self) -> None:
        x_label = self.x_axis_combo.currentText() or ""
        y_selected = [item.text() for item in self.y_axes_list.selectedItems()]

        self.chart_stage.title_label.setText(self._file_path.name)
        self.chart_stage.legend_label.setText(
            f"X: {x_label} — Y: {', '.join(y_selected) if y_selected else '—'}"
        )

        if not x_label or not y_selected or not self.csv_headers:
            self.chart_stage.plot_series([], [])
            return

        x_idx = self.csv_headers.index(x_label) if x_label in self.csv_headers else 0
        y_indices = [self.csv_headers.index(y) for y in y_selected if y in self.csv_headers]

        rows: list[list[str]] = []
        for row in self.csv_rows:
            if len(row) > x_idx:
                rows.append(
                    [row[x_idx]] + [row[yi] if len(row) > yi else "" for yi in y_indices]
                )
        self.chart_stage.plot_series(rows, [x_label] + y_selected)

    # ── cleanup ───────────────────────────────────────────────────────────

    def _cleanup(self) -> None:
        if self._csv_loader is not None:
            if self._csv_loader.isRunning():
                self._csv_loader.quit()
                if not self._csv_loader.wait(2000):
                    self._csv_loader.terminate()
                    self._csv_loader.wait(500)
            self._csv_loader.deleteLater()
            self._csv_loader = None
