from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QPushButton, QVBoxLayout, QWidget

try:
    import pyqtgraph as pg  # type: ignore[import-untyped]

    pg.setConfigOptions(antialias=True, background="w", foreground="k")
    _HAS_PG = True
except ImportError:
    _HAS_PG = False

_SERIES_COLORS = ["#2E90FA", "#FF6B6B", "#12B76A", "#F79009", "#9E77ED"]


def _parse_float(value: str) -> float | None:
    try:
        return float(value.strip().replace(",", "."))
    except (ValueError, AttributeError):
        return None


class ChartStage(QWidget):
    def __init__(
        self,
        title: str,
        legend_text: str,
        toolbar_actions: list[str],
        parent=None,
    ) -> None:
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        header_row = QHBoxLayout()
        self.title_label = QLabel(title)
        self.title_label.setStyleSheet("font-size: 16px; font-weight: 600;")
        header_row.addWidget(self.title_label)
        header_row.addStretch(1)

        self.legend_label = QLabel(legend_text)
        self.legend_label.setStyleSheet("color: #475467;")
        header_row.addWidget(self.legend_label)
        layout.addLayout(header_row)

        toolbar_row = QHBoxLayout()
        self.toolbar_buttons: list[QPushButton] = []
        for action_text in toolbar_actions:
            button = QPushButton(action_text)
            self.toolbar_buttons.append(button)
            toolbar_row.addWidget(button)
        toolbar_row.addStretch(1)
        layout.addLayout(toolbar_row)

        if _HAS_PG:
            self._plot_widget: pg.PlotWidget | None = pg.PlotWidget()
            self._plot_widget.setMinimumHeight(260)
            self._plot_widget.setBackground("#F8FAFC")
            self._plot_widget.showGrid(x=True, y=True, alpha=0.3)
            self._plot_widget.getAxis("bottom").setTextPen("k")
            self._plot_widget.getAxis("left").setTextPen("k")
            layout.addWidget(self._plot_widget)
            # Keep backward-compat attribute but hidden
            self.placeholder_label = QLabel()
            self.placeholder_label.setVisible(False)
        else:
            self._plot_widget = None
            stage_frame = QFrame()
            stage_frame.setFrameShape(QFrame.Shape.StyledPanel)
            stage_frame.setMinimumHeight(260)
            stage_frame.setStyleSheet(
                "background: #F8FAFC; border: 1px solid #D0D5DD; border-radius: 8px;"
            )
            stage_layout = QVBoxLayout(stage_frame)
            stage_layout.setContentsMargins(16, 16, 16, 16)
            stage_layout.addStretch(1)
            self.placeholder_label = QLabel("Anteprima grafico")
            self.placeholder_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.placeholder_label.setStyleSheet("color: #667085; font-size: 15px;")
            stage_layout.addWidget(self.placeholder_label)
            stage_layout.addStretch(1)
            layout.addWidget(stage_frame)

    def plot_series(self, row_values: list[list[str]], header_fields: list[str]) -> None:
        """Render the point series data.  X = first column, Y = remaining columns."""
        if not _HAS_PG or self._plot_widget is None:
            return

        self._plot_widget.clear()

        if not row_values or not header_fields:
            return

        y_col_count = len(header_fields) - 1

        if y_col_count <= 0:
            # Single column: plot as index vs value
            x_vals: list[float] = []
            y_vals: list[float] = []
            for i, row in enumerate(row_values):
                v = _parse_float(row[0]) if row else None
                if v is not None:
                    x_vals.append(float(i))
                    y_vals.append(v)
            if x_vals:
                color = _SERIES_COLORS[0]
                pen = pg.mkPen(color=color, width=2)
                self._plot_widget.plot(x_vals, y_vals, pen=pen, symbol="o", symbolSize=5, symbolBrush=color)
            return

        for y_idx in range(y_col_count):
            x_vals = []
            y_vals = []
            for row in row_values:
                xv = _parse_float(row[0]) if row else None
                yv = _parse_float(row[y_idx + 1]) if len(row) > y_idx + 1 else None
                if xv is not None and yv is not None:
                    x_vals.append(xv)
                    y_vals.append(yv)
            if x_vals:
                color = _SERIES_COLORS[y_idx % len(_SERIES_COLORS)]
                pen = pg.mkPen(color=color, width=2)
                self._plot_widget.plot(
                    x_vals, y_vals, pen=pen, symbol="o", symbolSize=5, symbolBrush=color
                )

    def toolbar_labels(self) -> list[str]:
        return [button.text() for button in self.toolbar_buttons]
