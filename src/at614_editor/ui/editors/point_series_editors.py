from __future__ import annotations

from PySide6.QtWidgets import QAbstractItemView, QHBoxLayout, QLabel, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget

from at614_editor.domain.models import Ce16Resource, Mms2218Resource, PointSeriesResource, PointSeriesRow
from at614_editor.domain.parsers.ce16 import serialize as serialize_ce16
from at614_editor.domain.parsers.mms2218 import serialize as serialize_mms2218
from at614_editor.domain.parsers.point_series import serialize as serialize_point_series
from at614_editor.ui.components.chart_stage import ChartStage
from at614_editor.ui.editors.resource_actions import build_duplicate_path, next_indexed_path
from at614_editor.ui.tooltip_manager import get_common_tooltip


class _PointSeriesEditorBase(QWidget):
    editor_name = "PointSeriesEditor"
    supports_save = True
    supports_duplicate = False

    def __init__(
        self,
        source_path,
        title: str,
        legend_text: str,
        toolbar_actions: list[str],
        header_fields: list[str],
        row_values: list[list[str]],
        detail_text: str,
        badge_text: str | None = None,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.source_path = source_path
        self.header_fields = list(header_fields)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        header_row = QHBoxLayout()
        summary_label = QLabel(detail_text)
        summary_label.setStyleSheet("font-size: 15px; font-weight: 600;")
        header_row.addWidget(summary_label)
        header_row.addStretch(1)

        self.badge_label = QLabel(badge_text or "")
        self.badge_label.setVisible(bool(badge_text))
        self.badge_label.setStyleSheet(
            "background: #F2F4F7; border: 1px solid #D0D5DD; border-radius: 10px; padding: 2px 8px;"
        )
        header_row.addWidget(self.badge_label)
        layout.addLayout(header_row)

        self.chart_stage = ChartStage(title=title, legend_text=legend_text, toolbar_actions=toolbar_actions)
        layout.addWidget(self.chart_stage)

        self.points_table = QTableWidget(len(row_values), len(header_fields))
        self.points_table.setHorizontalHeaderLabels(header_fields)
        self.points_table.verticalHeader().setVisible(False)
        self.points_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.points_table.setEditTriggers(QTableWidget.EditTrigger.AllEditTriggers)
        self.points_table.horizontalHeader().setStretchLastSection(True)

        # Aggiungi tooltip alle colonne
        column_tooltips = {
            "X": get_common_tooltip("punto_x"),
            "Y": get_common_tooltip("punto_y"),
        }
        for col_index, header in enumerate(header_fields):
            tooltip = column_tooltips.get(header, "")
            if tooltip:
                self.points_table.horizontalHeaderItem(col_index).setToolTip(tooltip)

        for row_index, values in enumerate(row_values):
            for column_index, value in enumerate(values):
                self.points_table.setItem(row_index, column_index, QTableWidgetItem(value))

        layout.addWidget(self.points_table)

        self.chart_stage.plot_series(row_values, header_fields)
        self.points_table.itemChanged.connect(self._on_table_changed)

    def _on_table_changed(self) -> None:
        row_values: list[list[str]] = []
        for row_index in range(self.points_table.rowCount()):
            row: list[str] = []
            for col_index in range(self.points_table.columnCount()):
                item = self.points_table.item(row_index, col_index)
                row.append(item.text() if item is not None else "")
            row_values.append(row)
        self.chart_stage.plot_series(row_values, self.header_fields)

    def _build_point_series_resource(self, prototype: PointSeriesResource) -> PointSeriesResource:
        rows: list[PointSeriesRow] = []
        for row_index in range(self.points_table.rowCount()):
            values: list[str] = []
            is_empty = True
            for column_index in range(self.points_table.columnCount()):
                item = self.points_table.item(row_index, column_index)
                value = item.text().strip() if item is not None else ""
                values.append(value)
                is_empty = is_empty and not value

            if not is_empty:
                rows.append(PointSeriesRow(values=values))

        return PointSeriesResource(
            source_path=prototype.source_path,
            header_fields=list(self.header_fields),
            rows=rows,
            encoding=prototype.encoding,
            line_ending=prototype.line_ending,
            endswith_newline=prototype.endswith_newline,
        )

    def duplicate_resource(self) -> Path | None:
        if not self.supports_duplicate:
            return None

        duplicate_path = build_duplicate_path(self.source_path)
        resource = self._build_serializable_resource(duplicate_path)
        duplicate_path.write_bytes(self._serialize_resource(resource))
        return duplicate_path

    def _build_serializable_resource(self, source_path):
        raise NotImplementedError()

    def _serialize_resource(self, resource) -> bytes:
        raise NotImplementedError()

    def save_changes(self) -> Path:
        resource = self._build_serializable_resource(self.source_path)
        self.source_path.write_bytes(self._serialize_resource(resource))
        return self.source_path


class CurveEditor(_PointSeriesEditorBase):
    editor_name = "CurveEditor"
    supports_duplicate = True

    def __init__(self, resource: PointSeriesResource, parent=None) -> None:
        self.resource = resource
        super().__init__(
            source_path=resource.source_path,
            title=resource.source_path.name,
            legend_text=" · ".join(resource.header_fields),
            toolbar_actions=["Seleziona punto", "Duplica punto"],
            header_fields=resource.header_fields,
            row_values=[row.values for row in resource.rows],
            detail_text=f"{len(resource.rows)} punti curva comando",
            parent=parent,
        )

    def _build_serializable_resource(self, source_path) -> PointSeriesResource:
        return PointSeriesResource(
            source_path=source_path,
            header_fields=list(self.resource.header_fields),
            rows=self._build_point_series_resource(self.resource).rows,
            encoding=self.resource.encoding,
            line_ending=self.resource.line_ending,
            endswith_newline=self.resource.endswith_newline,
        )

    def _serialize_resource(self, resource: PointSeriesResource) -> bytes:
        return serialize_point_series(resource)


class LimitEditor(_PointSeriesEditorBase):
    editor_name = "LimitEditor"
    supports_duplicate = True

    def __init__(self, resource: PointSeriesResource, parent=None) -> None:
        self.resource = resource
        super().__init__(
            source_path=resource.source_path,
            title=resource.source_path.name,
            legend_text=" · ".join(resource.header_fields),
            toolbar_actions=["Mostra opposta"],
            header_fields=resource.header_fields,
            row_values=[row.values for row in resource.rows],
            detail_text=f"{len(resource.rows)} punti curva limite",
            parent=parent,
        )

    def _build_serializable_resource(self, source_path) -> PointSeriesResource:
        return PointSeriesResource(
            source_path=source_path,
            header_fields=list(self.resource.header_fields),
            rows=self._build_point_series_resource(self.resource).rows,
            encoding=self.resource.encoding,
            line_ending=self.resource.line_ending,
            endswith_newline=self.resource.endswith_newline,
        )

    def _serialize_resource(self, resource: PointSeriesResource) -> bytes:
        return serialize_point_series(resource)


class RampEditor(_PointSeriesEditorBase):
    editor_name = "RampEditor"
    supports_duplicate = True

    def __init__(self, resource: PointSeriesResource, parent=None) -> None:
        self.resource = resource
        super().__init__(
            source_path=resource.source_path,
            title=resource.source_path.name,
            legend_text=" · ".join(resource.header_fields),
            toolbar_actions=["Snap griglia", "Inserisci punto intermedio"],
            header_fields=resource.header_fields,
            row_values=[row.values for row in resource.rows],
            detail_text=f"{len(resource.rows)} punti rampa XY",
            parent=parent,
        )

    def _build_serializable_resource(self, source_path) -> PointSeriesResource:
        return PointSeriesResource(
            source_path=source_path,
            header_fields=list(self.resource.header_fields),
            rows=self._build_point_series_resource(self.resource).rows,
            encoding=self.resource.encoding,
            line_ending=self.resource.line_ending,
            endswith_newline=self.resource.endswith_newline,
        )

    def _serialize_resource(self, resource: PointSeriesResource) -> bytes:
        return serialize_point_series(resource)


class Ce16Editor(_PointSeriesEditorBase):
    editor_name = "Ce16Editor"
    supports_duplicate = True

    def __init__(self, resource: Ce16Resource, parent=None) -> None:
        self.resource = resource
        super().__init__(
            source_path=resource.source_path,
            title=f"{resource.profile_name} · slot {resource.slot_index}",
            legend_text=" · ".join(resource.series.header_fields),
            toolbar_actions=["Ordina punti"],
            header_fields=resource.series.header_fields,
            row_values=[row.values for row in resource.series.rows],
            detail_text=f"{len(resource.series.rows)} punti CE16",
            badge_text="opzionale",
            parent=parent,
        )

    def _build_serializable_resource(self, source_path) -> Ce16Resource:
        return Ce16Resource(
            source_path=source_path,
            profile_name=self.resource.profile_name,
            slot_index=self.resource.slot_index,
            series=self._build_point_series_resource(self.resource.series),
        )

    def _serialize_resource(self, resource: Ce16Resource) -> bytes:
        return serialize_ce16(resource)

    def duplicate_resource(self) -> Path | None:
        duplicate_path = next_indexed_path(self.source_path.parent, "CE16_", ".cfg")
        resource = self._build_serializable_resource(duplicate_path)
        duplicate_path.write_bytes(self._serialize_resource(resource))
        return duplicate_path


class MmsEditor(_PointSeriesEditorBase):
    editor_name = "MmsEditor"
    supports_duplicate = True

    def __init__(self, resource: Mms2218Resource, parent=None) -> None:
        self.resource = resource
        super().__init__(
            source_path=resource.source_path,
            title=f"{resource.profile_name} · canale ADC{resource.channel_index}",
            legend_text=" · ".join(resource.series.header_fields),
            toolbar_actions=["Ordina punti", "Range ADC"],
            header_fields=resource.series.header_fields,
            row_values=[row.values for row in resource.series.rows],
            detail_text=f"{len(resource.series.rows)} punti calibrazione MMS2218",
            parent=parent,
        )

    def _build_serializable_resource(self, source_path) -> Mms2218Resource:
        return Mms2218Resource(
            source_path=source_path,
            profile_name=self.resource.profile_name,
            channel_index=self.resource.channel_index,
            series=self._build_point_series_resource(self.resource.series),
        )

    def _serialize_resource(self, resource: Mms2218Resource) -> bytes:
        return serialize_mms2218(resource)

    def duplicate_resource(self) -> Path | None:
        duplicate_path = next_indexed_path(self.source_path.parent, "ADC", ".csv")
        resource = self._build_serializable_resource(duplicate_path)
        duplicate_path.write_bytes(self._serialize_resource(resource))
        return duplicate_path
