from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from at614_editor.domain.program_clone import (
    CloneAction,
    CloneProgramPlan,
    build_clone_plan,
    derive_target_name,
    execute_clone_plan,
    validate_plan,
)
from at614_editor.domain.project import AT614Project
from at614_editor.ui.tooltip_manager import get_common_tooltip


KIND_LABELS = {
    "sequenza_test": "Sequenza test",
    "curva_comando": "Curva comando",
    "curva_limite": "Curva limite",
    "rampa_xy": "Rampa XY",
    "file_esterno": "File esterno",
    "resource": "Risorsa",
}


class CloneProgramDialog(QDialog):
    """Wizard di clonazione programma con scelta riusa/duplica per ogni risorsa."""

    def __init__(self, project: AT614Project, source_distributore: Path | None = None, parent=None) -> None:
        super().__init__(parent)
        self.project = project
        self.plan: CloneProgramPlan | None = None
        self.created_files: list[Path] = []

        self.setWindowTitle("Duplica programma")
        self.setMinimumSize(720, 540)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(14)

        title_label = QLabel("Duplica programma")
        title_label.setStyleSheet("font-size: 18px; font-weight: 700;")
        layout.addWidget(title_label)

        description = QLabel(
            "Crea un nuovo programma di collaudo a partire da un distributore esistente. "
            "Per ogni risorsa collegata puoi scegliere se duplicare o riusare quella originale."
        )
        description.setWordWrap(True)
        description.setStyleSheet("color: #475467;")
        layout.addWidget(description)

        form_widget = QWidget()
        form_layout = QFormLayout(form_widget)
        form_layout.setContentsMargins(0, 0, 0, 0)
        form_layout.setSpacing(8)

        self.source_combo = QComboBox()
        self.source_combo.setToolTip("Seleziona il distributore da cui duplicare il programma di collaudo.")
        for path in sorted(project.distributori):
            self.source_combo.addItem(path.name, userData=path)
        if source_distributore is not None:
            index = self.source_combo.findData(source_distributore)
            if index >= 0:
                self.source_combo.setCurrentIndex(index)
        self.source_combo.currentIndexChanged.connect(self._rebuild_plan)

        self.new_code_field = QLineEdit()
        self.new_code_field.setPlaceholderText("Esempio: 15.1001.500")
        self.new_code_field.setToolTip("Inserisci il codice univoco del nuovo distributore. Deve essere diverso dal sorgente.")
        self.new_code_field.textChanged.connect(self._rebuild_plan)

        form_layout.addRow("Distributore sorgente", self.source_combo)
        form_layout.addRow("Codice nuovo distributore", self.new_code_field)
        layout.addWidget(form_widget)

        resources_label = QLabel("Risorse collegate")
        resources_label.setStyleSheet("font-weight: 600;")
        layout.addWidget(resources_label)

        self.resources_table = QTableWidget(0, 4)
        self.resources_table.setHorizontalHeaderLabels(
            ["Tipo", "File sorgente", "Modalità", "Nome destinazione"]
        )
        self.resources_table.verticalHeader().setVisible(False)
        self.resources_table.horizontalHeader().setStretchLastSection(True)
        self.resources_table.setColumnWidth(0, 140)
        self.resources_table.setColumnWidth(1, 220)
        self.resources_table.setColumnWidth(2, 110)
        self.resources_table.setSelectionMode(QTableWidget.SelectionMode.NoSelection)
        self.resources_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        layout.addWidget(self.resources_table, 1)

        preview_label = QLabel("Riepilogo file che verranno creati")
        preview_label.setStyleSheet("font-weight: 600;")
        layout.addWidget(preview_label)

        self.preview_list = QListWidget()
        self.preview_list.setMaximumHeight(120)
        layout.addWidget(self.preview_list)

        self.error_label = QLabel("")
        self.error_label.setStyleSheet("color: #B42318;")
        self.error_label.setWordWrap(True)
        layout.addWidget(self.error_label)

        buttons = QDialogButtonBox()
        self.confirm_button = buttons.addButton("Duplica", QDialogButtonBox.ButtonRole.AcceptRole)
        buttons.addButton("Annulla", QDialogButtonBox.ButtonRole.RejectRole)
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self._rebuild_plan()

    # ------------------------------------------------------------------ plan

    def _selected_source(self) -> Path | None:
        return self.source_combo.currentData()

    def _rebuild_plan(self) -> None:
        source = self._selected_source()
        new_code = self.new_code_field.text().strip()

        self.resources_table.setRowCount(0)
        self.preview_list.clear()
        self.error_label.setText("")
        self.confirm_button.setEnabled(False)

        if source is None or not new_code:
            self.plan = None
            return

        try:
            self.plan = build_clone_plan(self.project, source, new_code)
        except ValueError as exc:
            self.plan = None
            self.error_label.setText(str(exc))
            return

        self._render_actions()
        self._render_preview()

    def _render_actions(self) -> None:
        assert self.plan is not None
        actions = self.plan.all_actions
        self.resources_table.setRowCount(len(actions))
        self._action_refs: list[tuple[CloneAction, QComboBox, QTableWidgetItem]] = []

        for row_index, action in enumerate(actions):
            kind_item = QTableWidgetItem(KIND_LABELS.get(action.kind, action.kind))
            kind_item.setFlags(kind_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.resources_table.setItem(row_index, 0, kind_item)

            source_item = QTableWidgetItem(action.source_path.name)
            source_item.setFlags(source_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.resources_table.setItem(row_index, 1, source_item)

            mode_combo = QComboBox()
            mode_combo.addItem("Duplica", userData="duplica")
            mode_combo.addItem("Riusa", userData="riusa")
            mode_combo.setCurrentIndex(0 if action.mode == "duplica" else 1)
            mode_combo.currentIndexChanged.connect(
                lambda _index, current_action=action, current_combo=mode_combo: self._on_mode_changed(
                    current_action, current_combo
                )
            )
            self.resources_table.setCellWidget(row_index, 2, mode_combo)

            target_item = QTableWidgetItem(action.target_name)
            target_item.setForeground(self._target_color(action.mode))
            self.resources_table.setItem(row_index, 3, target_item)

            self._action_refs.append((action, mode_combo, target_item))

    def _on_mode_changed(self, action: CloneAction, combo: QComboBox) -> None:
        mode = combo.currentData()
        action.mode = mode
        for stored_action, _combo, target_item in self._action_refs:
            if stored_action is action:
                target_item.setText(action.target_name if mode == "duplica" else action.source_path.name)
                target_item.setForeground(self._target_color(mode))
                break
        self._render_preview()

    def _target_color(self, mode: str):
        from PySide6.QtGui import QColor

        return QColor("#067647") if mode == "duplica" else QColor("#475467")

    def _render_preview(self) -> None:
        if self.plan is None:
            return

        self.preview_list.clear()
        self.preview_list.addItem(
            QListWidgetItem(f"[nuovo] {self.plan.new_distributore_path.name}")
        )
        for action in self.plan.all_actions:
            if action.mode == "duplica":
                self.preview_list.addItem(
                    QListWidgetItem(f"[duplicato da {action.source_path.name}] {action.target_name}")
                )
            else:
                self.preview_list.addItem(
                    QListWidgetItem(f"[riusato] {action.source_path.name}")
                )

        errors = validate_plan(self.plan)
        if errors:
            self.error_label.setText(" · ".join(errors))
            self.confirm_button.setEnabled(False)
        else:
            self.error_label.setText("")
            self.confirm_button.setEnabled(True)

    # ------------------------------------------------------------------ accept

    def _on_accept(self) -> None:
        if self.plan is None:
            return
        try:
            self.created_files = execute_clone_plan(self.project, self.plan)
        except ValueError as exc:
            self.error_label.setText(str(exc))
            return
        self.accept()

    def new_distributore_path(self) -> Path | None:
        if self.plan is None:
            return None
        return self.plan.new_distributore_path
