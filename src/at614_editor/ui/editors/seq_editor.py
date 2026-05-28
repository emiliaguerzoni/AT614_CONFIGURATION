from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QHeaderView,
    QPushButton,
    QFrame,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from at614_editor.domain.models import TestRow, TestSequence
from at614_editor.domain.parsers.test_csv import serialize as serialize_test_csv
from at614_editor.domain.project import AT614Project
from at614_editor.domain.schemas import get_test_schema_catalog
from at614_editor.domain.test_schema import (
    TestParameterDescriptor,
    describe_test_parameters,
    extract_test_file_references,
    get_schema_parameter_count,
)
from at614_editor.ui.components.file_widget import FileWidget
from at614_editor.ui.editors.resource_actions import build_duplicate_path
from at614_editor.ui.file_picker import FilePicker
from at614_editor.ui.tooltip_manager import get_common_tooltip, get_tooltip_for_descriptor


RESOURCE_TYPE_ICONS = {
    "curva_comando": "○",
    "limite_inf": "○",
    "limite_sup": "○",
    "rampa_xy": "◇",
    "external_config": "□",
    "generic_file": "□",
}


@dataclass(frozen=True, slots=True)
class RowValidationState:
    badge: str
    messages: tuple[str, ...]
    blocking_messages: tuple[str, ...] = ()


class TestRowTable(QTableWidget):
    row_dropped = Signal(int, int)

    def dropEvent(self, event) -> None:  # type: ignore[override]
        source_row = self.currentRow()
        target_row = self.indexAt(event.position().toPoint()).row()
        if source_row < 0:
            event.ignore()
            return
        if target_row < 0:
            target_row = self.rowCount() - 1

        self.row_dropped.emit(source_row, target_row)
        event.acceptProposedAction()


class SeqEditor(QWidget):
    state_changed = Signal()
    editor_name = "SeqEditor"
    supports_save = True
    supports_duplicate = True

    def __init__(
        self,
        project: AT614Project,
        source_path: Path,
        open_resource: Callable[[Path], bool] | None = None,
        show_usages: Callable[[Path], None] | None = None,
        update_validation: Callable[[list[str], bool], None] | None = None,
        update_status: Callable[[str], None] | None = None,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.project = project
        self.source_path = source_path
        self.open_resource = open_resource
        self.show_usages = show_usages
        self.update_validation = update_validation
        self.update_status = update_status
        self.sequence = project.test_sequences[source_path]
        self.file_picker = FilePicker(project)
        self.working_rows = [
            TestRow(name=row.name, test_id=row.test_id, index_raw=row.index_raw, parameters=list(row.parameters))
            for row in self.sequence.rows
        ]
        self.current_row_index: int | None = None
        self.parameter_bindings: list[tuple[TestParameterDescriptor, QWidget]] = []
        self.file_reference_widgets: list[FileWidget] = []
        self.row_validations: list[RowValidationState] = []
        self._loading_row = False

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        summary_label = QLabel(f"{len(self.working_rows)} righe test caricate")
        summary_label.setStyleSheet("font-size: 15px; font-weight: 600;")
        self.summary_label = summary_label
        layout.addWidget(self.summary_label)

        row_actions = QHBoxLayout()
        self.add_row_button = QPushButton("Aggiungi riga")
        self.duplicate_row_button = QPushButton("Duplica riga")
        self.delete_row_button = QPushButton("Elimina riga")
        self.reindex_button = QPushButton("Reindicizza")

        self.add_row_button.clicked.connect(self.add_row)
        self.duplicate_row_button.clicked.connect(self.duplicate_selected_row)
        self.delete_row_button.clicked.connect(self.delete_selected_row)
        self.reindex_button.clicked.connect(self.reindex_numeric_indices)

        for button in (
            self.add_row_button,
            self.duplicate_row_button,
            self.delete_row_button,
            self.reindex_button,
        ):
            row_actions.addWidget(button)
        row_actions.addStretch(1)
        layout.addLayout(row_actions)

        self.row_table = TestRowTable(len(self.working_rows), 4)
        self.row_table.setHorizontalHeaderLabels(["Nome test", "ID test", "Indice", "Stato"])
        self.row_table.verticalHeader().setVisible(False)
        self.row_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.row_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.row_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.row_table.setDragEnabled(True)
        self.row_table.setAcceptDrops(True)
        self.row_table.setDropIndicatorShown(True)
        self.row_table.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.row_table.horizontalHeader().setStretchLastSection(False)
        self.row_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.row_table.setColumnWidth(1, 170)
        self.row_table.setColumnWidth(2, 70)
        self.row_table.setColumnWidth(3, 90)
        self.row_table.currentCellChanged.connect(self._handle_row_changed)
        self.row_table.row_dropped.connect(self.move_row)
        layout.addWidget(self.row_table)

        self._revalidate_all_rows()
        self._populate_table()

        detail_frame = QFrame()
        detail_layout = QVBoxLayout(detail_frame)
        detail_layout.setContentsMargins(12, 12, 12, 12)
        detail_layout.setSpacing(12)

        detail_title_row = QHBoxLayout()
        detail_title = QLabel("Dettaglio dinamico")
        detail_title.setStyleSheet("font-size: 15px; font-weight: 600;")
        self.row_status_badge = QLabel("ok")
        detail_title_row.addWidget(detail_title)
        detail_title_row.addStretch(1)
        detail_title_row.addWidget(self.row_status_badge)
        detail_layout.addLayout(detail_title_row)

        common_fields_widget = QWidget()
        common_fields_layout = QFormLayout(common_fields_widget)
        common_fields_layout.setContentsMargins(0, 0, 0, 0)
        common_fields_layout.setSpacing(8)

        self.name_field = QLineEdit()
        self.name_field.setToolTip(get_common_tooltip("nome_test"))
        self.test_id_field = QComboBox()
        self.test_id_field.setEditable(True)
        self.test_id_field.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.test_id_field.addItems(sorted(get_test_schema_catalog().keys()))
        self.test_id_field.setCurrentIndex(-1)
        self.test_id_field.lineEdit().setPlaceholderText("Seleziona o digita ID test")
        self.test_id_field.setToolTip(get_common_tooltip("test_id"))
        self.test_id_field.lineEdit().setToolTip(get_common_tooltip("test_id"))
        self.index_field = QLineEdit()
        self.index_field.setToolTip(get_common_tooltip("indice"))

        self.name_field.textChanged.connect(self._handle_live_edit)
        self.test_id_field.currentTextChanged.connect(self._handle_live_edit)
        self.index_field.textChanged.connect(self._handle_live_edit)
        self.name_field.editingFinished.connect(self._sync_common_fields)
        self.index_field.editingFinished.connect(self._sync_common_fields)
        self.test_id_field.activated.connect(lambda _idx: self._sync_and_refresh_current_row())
        self.test_id_field.lineEdit().editingFinished.connect(self._sync_and_refresh_current_row)

        common_fields_layout.addRow("Nome test", self.name_field)
        common_fields_layout.addRow("ID test", self.test_id_field)
        common_fields_layout.addRow("Indice", self.index_field)
        detail_layout.addWidget(common_fields_widget)

        self.parameter_container = QWidget()
        self.parameter_layout = QVBoxLayout(self.parameter_container)
        self.parameter_layout.setContentsMargins(0, 0, 0, 0)
        self.parameter_layout.setSpacing(10)
        detail_layout.addWidget(self.parameter_container)

        layout.addWidget(detail_frame)

        if self.working_rows:
            self.row_table.setCurrentCell(0, 0)
        else:
            self._publish_validation_state(RowValidationState("ok", ()))
        self._update_summary()

    def _populate_table(self) -> None:
        for row_index, row in enumerate(self.working_rows):
            self._update_table_row(row_index, row)

    def _update_summary(self) -> None:
        self.summary_label.setText(f"{len(self.working_rows)} righe test caricate")

    def _ensure_row_parameter_capacity(self, row: TestRow) -> None:
        required_slots = max(len(row.parameters), get_schema_parameter_count(row.test_id))
        if required_slots > len(row.parameters):
            row.parameters.extend([""] * (required_slots - len(row.parameters)))

    def _rebuild_table(self, selected_row: int | None = None) -> None:
        self._revalidate_all_rows()
        self.row_table.blockSignals(True)
        self.row_table.setRowCount(len(self.working_rows))
        self._populate_table()
        self.row_table.blockSignals(False)
        self._update_summary()
        self.current_row_index = None
        self.state_changed.emit()

        if not self.working_rows:
            self.name_field.clear()
            self.test_id_field.clearEditText()
            self.index_field.clear()
            self._clear_parameter_details()
            return

        if selected_row is None:
            selected_row = min(self.current_row_index or 0, len(self.working_rows) - 1)
        self.select_row(selected_row)

    def _update_table_row(self, row_index: int, row: TestRow) -> None:
        self.row_table.setItem(row_index, 0, QTableWidgetItem(row.name))
        self.row_table.setItem(row_index, 1, QTableWidgetItem(row.test_id))
        self.row_table.setItem(row_index, 2, QTableWidgetItem(row.index_raw))
        validation_item = QTableWidgetItem(self.row_validations[row_index].badge)
        validation_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        validation_colors = {
            "ok": ("#ECFDF3", "#067647"),
            "warning": ("#FFFAEB", "#B54708"),
            "errore": ("#FEF3F2", "#B42318"),
        }
        background_color, foreground_color = validation_colors.get(self.row_validations[row_index].badge, validation_colors["ok"])
        validation_item.setBackground(QColor(background_color))
        validation_item.setForeground(QColor(foreground_color))
        self.row_table.setItem(row_index, 3, validation_item)

    def _handle_row_changed(
        self,
        current_row: int,
        _current_column: int,
        previous_row: int,
        _previous_column: int,
    ) -> None:
        if previous_row >= 0 and previous_row == self.current_row_index:
            self._commit_row(previous_row)

        if current_row < 0 or current_row >= len(self.working_rows):
            self.current_row_index = None
            return

        self.current_row_index = current_row
        self._load_row(current_row)

    def _load_row(self, row_index: int) -> None:
        row = self.working_rows[row_index]
        self._ensure_row_parameter_capacity(row)
        self._loading_row = True
        self.name_field.setText(row.name)
        self.test_id_field.setCurrentText(row.test_id)
        self.index_field.setText(row.index_raw)
        self._render_parameter_details(row)
        self._loading_row = False
        self._publish_current_validation()

    def _clear_parameter_details(self) -> None:
        while self.parameter_layout.count():
            item = self.parameter_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

        self.parameter_bindings.clear()
        self.file_reference_widgets.clear()

    def _icon_for_resource_type(self, resource_type: str | None) -> str:
        if resource_type is None:
            return "□"
        return RESOURCE_TYPE_ICONS.get(resource_type, "□")

    def _render_parameter_details(self, row: TestRow) -> None:
        self._clear_parameter_details()
        descriptors = describe_test_parameters(row)

        if not descriptors:
            empty_label = QLabel("Nessun parametro specifico")
            empty_label.setStyleSheet("color: #475467;")
            self.parameter_layout.addWidget(empty_label)
            return

        for descriptor in descriptors:
            wrapper = QWidget()
            wrapper_layout = QVBoxLayout(wrapper)
            wrapper_layout.setContentsMargins(0, 0, 0, 0)
            wrapper_layout.setSpacing(6)

            wrapper_layout.addWidget(QLabel(descriptor.label))
            raw_value = row.parameters[descriptor.parameter_index - 1]

            if descriptor.editor_kind == "file_ref":
                def make_picker_callback(resource_type: str | None, file_widget: FileWidget) -> Callable[[], None]:
                    def picker_callback() -> None:
                        selected_path = self.file_picker.pick_file(
                            parent=self,
                            title=f"Seleziona {descriptor.label.lower()}",
                            resource_type=resource_type,
                        )
                        if selected_path:
                            file_widget.name_label.setText(selected_path)
                            self._handle_live_edit()
                    return picker_callback

                field_widget = FileWidget(
                    self._icon_for_resource_type(descriptor.resource_type),
                    editable=True,
                    on_select=make_picker_callback(descriptor.resource_type, None),  # Sarà impostato dopo
                )
                # Ora aggiorna il callback con il widget corretto
                field_widget.select_button.clicked.disconnect()
                field_widget.select_button.clicked.connect(make_picker_callback(descriptor.resource_type, field_widget))

                tooltip = get_tooltip_for_descriptor(descriptor, row.test_id)
                if tooltip:
                    field_widget.name_label.setToolTip(tooltip)
                target_path = self.project.resolve_resource_path(raw_value)
                field_widget.set_reference(raw_value, uses_count=self.project.get_usage_count(target_path))
                field_widget.open_button.setEnabled(target_path is not None)
                field_widget.uses_button.setEnabled(target_path is not None)
                field_widget.name_label.textChanged.connect(
                    lambda _value, current_widget=field_widget: self._refresh_file_widget_actions(current_widget)
                )
                field_widget.name_label.textChanged.connect(self._handle_live_edit)
                field_widget.open_button.clicked.connect(
                    lambda _checked=False, current_widget=field_widget: self._open_file_reference(current_widget)
                )
                field_widget.uses_button.clicked.connect(
                    lambda _checked=False, current_widget=field_widget: self._show_file_usages(current_widget)
                )
                self.file_reference_widgets.append(field_widget)
            else:
                field_widget = QLineEdit(raw_value)
                tooltip = get_tooltip_for_descriptor(descriptor, row.test_id)
                if tooltip:
                    field_widget.setToolTip(tooltip)
                field_widget.textChanged.connect(self._handle_live_edit)

            wrapper_layout.addWidget(field_widget)
            self.parameter_layout.addWidget(wrapper)
            self.parameter_bindings.append((descriptor, field_widget))

        self.parameter_layout.addStretch(1)

    def _refresh_file_widget_actions(self, file_widget: FileWidget) -> None:
        target_path = self.project.resolve_resource_path(file_widget.reference_text())
        count = self.project.get_usage_count(target_path)
        file_widget.open_button.setEnabled(target_path is not None)
        file_widget.uses_button.setEnabled(target_path is not None and count > 0)
        file_widget.uses_button.setText(f"Usi ({count})")

    def _open_file_reference(self, file_widget: FileWidget) -> bool:
        if self.open_resource is None:
            return False

        target_path = self.project.resolve_resource_path(file_widget.reference_text())
        if target_path is None:
            return False
        return self.open_resource(target_path)

    def _show_file_usages(self, file_widget: FileWidget) -> None:
        target_path = self.project.resolve_resource_path(file_widget.reference_text())
        if target_path is None:
            return

        used_by = self.project.get_used_by(target_path)
        count = len(used_by)
        if count == 0:
            return

        if count > 10:
            if self.show_usages is not None:
                self.show_usages(target_path)
            return

        # Popover inline: menu contestuale sotto il pulsante
        from PySide6.QtWidgets import QMenu  # noqa: PLC0415

        menu = QMenu(self)
        for ref_path in sorted(used_by, key=lambda p: p.name):
            action = menu.addAction(ref_path.name)
            if self.open_resource is not None:
                action.triggered.connect(
                    lambda _checked=False, p=ref_path: self.open_resource(p)  # type: ignore[misc]
                )
        button = file_widget.uses_button
        menu.popup(button.mapToGlobal(button.rect().bottomLeft()))

    def _commit_row(self, row_index: int) -> None:
        row = self.working_rows[row_index]
        row.name = self.name_field.text().strip()
        row.test_id = self.test_id_field.currentText().strip()
        row.index_raw = self.index_field.text().strip()
        self._ensure_row_parameter_capacity(row)

        for descriptor, widget in self.parameter_bindings:
            if descriptor.parameter_index - 1 >= len(row.parameters):
                continue

            if isinstance(widget, FileWidget):
                row.parameters[descriptor.parameter_index - 1] = widget.reference_text()
            elif isinstance(widget, QLineEdit):
                row.parameters[descriptor.parameter_index - 1] = widget.text().strip()

        self._update_row_validation(row_index, publish=row_index == self.current_row_index)
        self._update_table_row(row_index, row)
        self.state_changed.emit()

    def _handle_live_edit(self, _value: str) -> None:
        if self.current_row_index is None or self._loading_row:
            return
        self._commit_row(self.current_row_index)

    def _sync_common_fields(self) -> None:
        if self.current_row_index is None:
            return
        self._commit_row(self.current_row_index)

    def _sync_and_refresh_current_row(self) -> None:
        if self.current_row_index is None:
            return

        self._commit_row(self.current_row_index)
        self._load_row(self.current_row_index)

    def _validate_row(self, row_index: int) -> RowValidationState:
        row = self.working_rows[row_index]
        errors: list[str] = []
        warnings: list[str] = []

        if not row.test_id.strip():
            errors.append("ID test mancante")

        if not row.name.strip():
            warnings.append("Nome test mancante")

        if not row.index_raw.strip():
            warnings.append("Indice mancante")
        elif row.index_raw.isdigit():
            duplicate_index = any(
                other_index != row_index and other_row.index_raw == row.index_raw
                for other_index, other_row in enumerate(self.working_rows)
            )
            if duplicate_index:
                errors.append(f"Indice duplicato: {row.index_raw}")

        explicit_file_reference_indices: set[int] = set()
        for descriptor in describe_test_parameters(row):
            if descriptor.editor_kind != "file_ref":
                continue

            explicit_file_reference_indices.add(descriptor.parameter_index)
            parameter_offset = descriptor.parameter_index - 1
            raw_value = row.parameters[parameter_offset].strip() if parameter_offset < len(row.parameters) else ""
            if not raw_value:
                warnings.append(f"{descriptor.label}: — riferimento mancante —")
                continue
            if self.project.resolve_resource_path(raw_value) is None:
                errors.append(f"{descriptor.label}: — riferimento mancante —")

        for reference in extract_test_file_references(row):
            if reference.parameter_index in explicit_file_reference_indices:
                continue
            if self.project.resolve_resource_path(reference.raw_value) is None:
                errors.append(f"Parametro {reference.parameter_index}: — riferimento mancante —")

        if errors:
            return RowValidationState("errore", tuple(errors + warnings), tuple(errors))
        if warnings:
            return RowValidationState("warning", tuple(warnings))
        return RowValidationState("ok", ())

    def _revalidate_all_rows(self) -> None:
        self.row_validations = [self._validate_row(index) for index, _row in enumerate(self.working_rows)]

    def _update_row_validation(self, row_index: int, publish: bool = False) -> None:
        if row_index < 0 or row_index >= len(self.working_rows):
            return
        self.row_validations[row_index] = self._validate_row(row_index)
        if publish:
            self._publish_validation_state(self.row_validations[row_index])

    def _publish_validation_state(self, validation: RowValidationState) -> None:
        badge_styles = {
            "ok": "background: #ECFDF3; border: 1px solid #ABEFC6; color: #067647; border-radius: 10px; padding: 2px 8px;",
            "warning": "background: #FFFAEB; border: 1px solid #FEDF89; color: #B54708; border-radius: 10px; padding: 2px 8px;",
            "errore": "background: #FEF3F2; border: 1px solid #FECDCA; color: #B42318; border-radius: 10px; padding: 2px 8px;",
        }
        self.row_status_badge.setText(validation.badge)
        self.row_status_badge.setStyleSheet(badge_styles.get(validation.badge, badge_styles["ok"]))

        if self.update_validation is not None:
            self.update_validation(list(validation.messages), False)
        if self.update_status is not None:
            self.update_status(validation.badge)

    def _publish_current_validation(self) -> None:
        if self.current_row_index is None:
            self._publish_validation_state(RowValidationState("ok", ()))
            return
        self._publish_validation_state(self.row_validations[self.current_row_index])

    def _refresh_table_rows(self) -> None:
        self.row_table.blockSignals(True)
        for row_index, row in enumerate(self.working_rows):
            self._update_table_row(row_index, row)
        self.row_table.blockSignals(False)

    def validate_before_save(self) -> list[str]:
        if self.current_row_index is not None:
            self._commit_row(self.current_row_index)

        self._revalidate_all_rows()
        self._refresh_table_rows()

        blocking_messages: list[str] = []
        first_error_row: int | None = None
        for row_index, validation in enumerate(self.row_validations):
            if validation.badge != "errore":
                continue

            if first_error_row is None:
                first_error_row = row_index

            blocking_messages.extend(
                f"Riga {row_index + 1} · {message}" for message in validation.blocking_messages
            )

        if first_error_row is not None:
            self.select_row(first_error_row)
        else:
            self._publish_current_validation()

        return blocking_messages

    def _next_numeric_index(self) -> str:
        numeric_indices = [int(row.index_raw) for row in self.working_rows if row.index_raw.isdigit()]
        if not numeric_indices:
            return "0"
        return str(max(numeric_indices) + 1)

    def add_row(self) -> None:
        if self.current_row_index is not None:
            self._commit_row(self.current_row_index)

        insert_at = (self.current_row_index + 1) if self.current_row_index is not None else len(self.working_rows)
        new_row = TestRow(name="", test_id="", index_raw=self._next_numeric_index(), parameters=[])
        self.working_rows.insert(insert_at, new_row)
        self.reindex_numeric_indices(select_row=insert_at)

    def duplicate_selected_row(self) -> None:
        if self.current_row_index is None:
            return

        self._commit_row(self.current_row_index)
        source_row = self.working_rows[self.current_row_index]
        duplicate_row = TestRow(
            name=source_row.name,
            test_id=source_row.test_id,
            index_raw=source_row.index_raw,
            parameters=list(source_row.parameters),
        )
        insert_at = self.current_row_index + 1
        self.working_rows.insert(insert_at, duplicate_row)
        self.reindex_numeric_indices(select_row=insert_at)

    def delete_selected_row(self) -> None:
        if self.current_row_index is None:
            return

        del self.working_rows[self.current_row_index]
        if not self.working_rows:
            self._rebuild_table()
            return

        select_row = min(self.current_row_index, len(self.working_rows) - 1)
        self.reindex_numeric_indices(select_row=select_row)

    def reindex_numeric_indices(self, _checked: bool = False, select_row: int | None = None) -> None:
        next_index = 0
        for row in self.working_rows:
            if row.index_raw.isdigit():
                row.index_raw = str(next_index)
                next_index += 1

        self._rebuild_table(selected_row=select_row)

    def move_row(self, source_row: int, target_row: int) -> None:
        if source_row == target_row:
            return
        if source_row < 0 or target_row < 0:
            return
        if source_row >= len(self.working_rows) or target_row >= len(self.working_rows):
            return

        if self.current_row_index is not None:
            self._commit_row(self.current_row_index)

        moved_row = self.working_rows.pop(source_row)
        self.working_rows.insert(target_row, moved_row)
        self.reindex_numeric_indices(select_row=target_row)

    def _build_sequence(self, source_path: Path) -> TestSequence:
        if self.current_row_index is not None:
            self._commit_row(self.current_row_index)

        rows = [
            TestRow(name=row.name, test_id=row.test_id, index_raw=row.index_raw, parameters=list(row.parameters))
            for row in self.working_rows
        ]
        return TestSequence(
            source_path=source_path,
            header_fields=list(self.sequence.header_fields),
            rows=rows,
            encoding=self.sequence.encoding,
            line_ending=self.sequence.line_ending,
            endswith_newline=self.sequence.endswith_newline,
        )

    def is_dirty(self) -> bool:
        return self.sequence != self._build_sequence(self.source_path)

    def save_changes(self) -> Path:
        sequence = self._build_sequence(self.source_path)
        self.source_path.write_bytes(serialize_test_csv(sequence))
        return self.source_path

    def duplicate_resource(self) -> Path:
        duplicate_path = build_duplicate_path(self.source_path)
        sequence = self._build_sequence(duplicate_path)
        duplicate_path.write_bytes(serialize_test_csv(sequence))
        return duplicate_path

    def select_row(self, row_index: int) -> None:
        self.row_table.setCurrentCell(row_index, 0)

    def parameter_labels(self) -> list[str]:
        return [descriptor.label for descriptor, _widget in self.parameter_bindings]

    def current_row_status(self) -> str:
        if self.current_row_index is None:
            return "ok"
        return self.row_validations[self.current_row_index].badge

    def current_validation_messages(self) -> list[str]:
        if self.current_row_index is None:
            return []
        return list(self.row_validations[self.current_row_index].messages)
