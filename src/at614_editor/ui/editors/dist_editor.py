from __future__ import annotations

from pathlib import Path
from typing import Callable

from PySide6.QtWidgets import QAbstractItemView, QGridLayout, QLabel, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget

from at614_editor.domain.models import DistributoreConfig
from at614_editor.domain.parsers.distributore import serialize as serialize_distributore
from at614_editor.domain.project import AT614Project
from at614_editor.ui.components.file_widget import FileWidget
from at614_editor.ui.components.resource_card import ResourceCard
from at614_editor.ui.editors.resource_actions import build_duplicate_path
from at614_editor.ui.file_picker import FilePicker
from at614_editor.ui.tooltip_manager import get_common_tooltip


class DistEditor(QWidget):
    editor_name = "DistEditor"
    supports_save = True
    supports_duplicate = True

    def __init__(
        self,
        project: AT614Project,
        source_path: Path,
        open_resource: Callable[[Path], bool] | None = None,
        show_usages: Callable[[Path], None] | None = None,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.project = project
        self.source_path = source_path
        self.open_resource = open_resource
        self.show_usages = show_usages
        self.config = project.distributori[source_path]
        self.section_cards: list[ResourceCard] = []
        self.ce16_widgets: list[FileWidget] = []
        self.file_picker = FilePicker(project)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        sections_title = QLabel("Sezioni distributore")
        sections_title.setStyleSheet("font-size: 15px; font-weight: 600;")
        layout.addWidget(sections_title)

        sections_grid = QWidget()
        sections_layout = QGridLayout(sections_grid)
        sections_layout.setContentsMargins(0, 0, 0, 0)
        sections_layout.setHorizontalSpacing(12)
        sections_layout.setVerticalSpacing(12)

        for offset, section_index in enumerate(range(1, 6)):
            section_code = self.config.sezioni.get(section_index)
            card = ResourceCard(
                title=f"Sezione {section_index}",
                file_name=f"{section_code}.csv" if section_code else "— non configurato —",
                meta_text="Sequenza collegata" if section_code else "Sezione non attiva · Attiva sezione",
                state_text="ok" if section_code else "opz.",
                editable=True,
            )
            card.open_button.clicked.connect(
                lambda _checked=False, current_card=card: self._open_reference(current_card.file_name_text())
            )
            card.uses_button.clicked.connect(
                lambda _checked=False, current_card=card: self._show_reference_usages(
                    current_card.file_name_text(), current_card.uses_button
                )
            )
            card.file_label.textChanged.connect(self._refresh_reference_actions)
            self.section_cards.append(card)
            sections_layout.addWidget(card, offset // 2, offset % 2)
        layout.addWidget(sections_grid)

        ce16_title = QLabel("Calibrazioni CE16")
        ce16_title.setStyleSheet("font-size: 15px; font-weight: 600;")
        layout.addWidget(ce16_title)

        ce16_grid = QWidget()
        ce16_layout = QGridLayout(ce16_grid)
        ce16_layout.setContentsMargins(0, 0, 0, 0)
        ce16_layout.setHorizontalSpacing(12)
        ce16_layout.setVerticalSpacing(12)

        for offset, slot_index in enumerate(range(1, 6)):
            def make_picker_callback(fw: FileWidget) -> None:
                def picker_callback() -> None:
                    selected_path = self.file_picker.pick_file(
                        parent=self,
                        title="Seleziona calibrazione CE16",
                        resource_type="ce16",
                    )
                    if selected_path:
                        fw.name_label.setText(selected_path)
                return picker_callback

            file_widget = FileWidget("□", editable=True)
            file_widget.name_label.setToolTip(get_common_tooltip("ce16"))
            file_widget.set_reference(self.config.calibrazioni_ce16.get(slot_index))
            # Connetti il pulsante Seleziona al file picker
            file_widget.select_button.clicked.connect(make_picker_callback(file_widget))
            file_widget.open_button.clicked.connect(
                lambda _checked=False, current_widget=file_widget: self._open_reference(current_widget.reference_text())
            )
            file_widget.uses_button.clicked.connect(
                lambda _checked=False, current_widget=file_widget: self._show_reference_usages(
                    current_widget.reference_text(), current_widget.uses_button
                )
            )
            file_widget.name_label.textChanged.connect(self._refresh_reference_actions)
            self.ce16_widgets.append(file_widget)
            ce16_layout.addWidget(file_widget, offset // 2, offset % 2)
        layout.addWidget(ce16_grid)

        extras_title = QLabel("Parametri aggiuntivi")
        extras_title.setStyleSheet("font-size: 15px; font-weight: 600;")
        layout.addWidget(extras_title)

        self.parameters_table = QTableWidget(max(1, len(self.config.extras)), 2)
        self.parameters_table.setHorizontalHeaderLabels(["Parametro", "Valore"])
        self.parameters_table.verticalHeader().setVisible(False)
        self.parameters_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.parameters_table.setEditTriggers(QTableWidget.EditTrigger.AllEditTriggers)
        self.parameters_table.setToolTip(get_common_tooltip("parametro_aggiuntivo"))

        if self.config.extras:
            for row_index, (key, value) in enumerate(sorted(self.config.extras.items())):
                self.parameters_table.setItem(row_index, 0, QTableWidgetItem(key))
                self.parameters_table.setItem(row_index, 1, QTableWidgetItem(value))
        else:
            self.parameters_table.setItem(0, 0, QTableWidgetItem("Nessun parametro aggiuntivo"))
            self.parameters_table.setItem(0, 1, QTableWidgetItem(""))

        self.parameters_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.parameters_table)
        self._refresh_reference_actions()

    def _refresh_reference_actions(self) -> None:
        for card in self.section_cards:
            target_path = self.project.resolve_resource_path(card.file_name_text())
            count = self.project.get_usage_count(target_path)
            card.open_button.setEnabled(target_path is not None)
            card.uses_button.setEnabled(target_path is not None and count > 0)
            card.uses_button.setText(f"Usi ({count})")

        for file_widget in self.ce16_widgets:
            target_path = self.project.resolve_resource_path(file_widget.reference_text())
            count = self.project.get_usage_count(target_path)
            file_widget.open_button.setEnabled(target_path is not None)
            file_widget.uses_button.setEnabled(target_path is not None and count > 0)
            file_widget.uses_button.setText(f"Usi ({count})")

    def _open_reference(self, raw_value: str) -> bool:
        if self.open_resource is None:
            return False

        target_path = self.project.resolve_resource_path(raw_value)
        if target_path is None:
            return False

        return self.open_resource(target_path)

    def _show_reference_usages(self, raw_value: str, button=None) -> None:
        target_path = self.project.resolve_resource_path(raw_value)
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

        # Popover inline per n <= 10
        from PySide6.QtWidgets import QMenu  # noqa: PLC0415

        menu = QMenu(self)
        for ref_path in sorted(used_by, key=lambda p: p.name):
            action = menu.addAction(ref_path.name)
            if self.open_resource is not None:
                action.triggered.connect(
                    lambda _checked=False, p=ref_path: self.open_resource(p)  # type: ignore[misc]
                )
        if button is not None:
            menu.popup(button.mapToGlobal(button.rect().bottomLeft()))
        else:
            from PySide6.QtGui import QCursor  # noqa: PLC0415
            menu.popup(QCursor.pos())

    def _normalized_section_code(self, raw_text: str) -> str | None:
        normalized = raw_text.strip()
        if not normalized or normalized == "— non configurato —":
            return None
        if normalized.lower().endswith(".csv"):
            return normalized[:-4]
        return normalized

    def _normalized_reference(self, raw_text: str) -> str | None:
        normalized = raw_text.strip()
        if not normalized or normalized == "— riferimento mancante —":
            return None
        return normalized

    def _collect_extras(self) -> dict[str, str]:
        extras: dict[str, str] = {}
        for row_index in range(self.parameters_table.rowCount()):
            key_item = self.parameters_table.item(row_index, 0)
            value_item = self.parameters_table.item(row_index, 1)
            key = key_item.text().strip() if key_item is not None else ""
            value = value_item.text().strip() if value_item is not None else ""

            if not key or key == "Nessun parametro aggiuntivo":
                continue

            extras[key] = value

        return extras

    def _collect_model(self, source_path: Path) -> DistributoreConfig:
        sezioni: dict[int, str] = {}
        for section_index, card in enumerate(self.section_cards, start=1):
            section_code = self._normalized_section_code(card.file_name_text())
            if section_code is not None:
                sezioni[section_index] = section_code

        calibrazioni_ce16: dict[int, str] = {}
        for slot_index, file_widget in enumerate(self.ce16_widgets, start=1):
            reference = self._normalized_reference(file_widget.reference_text())
            if reference is not None:
                calibrazioni_ce16[slot_index] = reference

        return DistributoreConfig(
            source_path=source_path,
            sezioni=sezioni,
            calibrazioni_ce16=calibrazioni_ce16,
            extras=self._collect_extras(),
            lines=list(self.config.lines),
            encoding=self.config.encoding,
            line_ending=self.config.line_ending,
            endswith_newline=self.config.endswith_newline,
        )

    def save_changes(self) -> Path:
        model = self._collect_model(self.source_path)
        self.source_path.write_bytes(serialize_distributore(model))
        return self.source_path

    def duplicate_resource(self) -> Path:
        duplicate_path = build_duplicate_path(self.source_path)
        model = self._collect_model(duplicate_path)
        duplicate_path.write_bytes(serialize_distributore(model))
        return duplicate_path
