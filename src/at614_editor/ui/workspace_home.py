from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Signal
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QPushButton, QVBoxLayout, QWidget

from at614_editor.domain.project import AT614Project


class WorkspaceHome(QWidget):
    open_shell_requested = Signal()
    open_first_resource_requested = Signal()
    duplicate_program_requested = Signal()

    def __init__(self, project_root: Path | None, project: AT614Project | None, parent=None) -> None:
        super().__init__(parent)
        self._summary_labels: list[QLabel] = []

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(16)

        title_label = QLabel("AT614 Configuration Editor")
        title_label.setStyleSheet("font-size: 28px; font-weight: 700;")
        layout.addWidget(title_label)

        subtitle_label = QLabel("Shell UI milestone M2 — struttura navigabile e componenti riusabili")
        subtitle_label.setStyleSheet("color: #475467; font-size: 15px;")
        layout.addWidget(subtitle_label)

        project_name = project_root.name if project_root is not None else "Nessun progetto rilevato"
        project_label = QLabel(f"Programma attivo: {project_name}")
        project_label.setStyleSheet("font-size: 16px; font-weight: 600;")
        layout.addWidget(project_label)

        summary_frame = QFrame()
        summary_frame.setFrameShape(QFrame.Shape.StyledPanel)
        summary_layout = QVBoxLayout(summary_frame)
        summary_layout.setContentsMargins(16, 16, 16, 16)
        summary_layout.setSpacing(10)

        for summary_text in self._build_summary_lines(project):
            label = QLabel(summary_text)
            label.setStyleSheet("font-size: 14px;")
            self._summary_labels.append(label)
            summary_layout.addWidget(label)

        layout.addWidget(summary_frame)

        actions_row = QHBoxLayout()
        open_shell_button = QPushButton("Apri shell")
        open_shell_button.clicked.connect(self.open_shell_requested)
        actions_row.addWidget(open_shell_button)

        open_first_button = QPushButton("Apri prima risorsa")
        open_first_button.clicked.connect(self.open_first_resource_requested)
        open_first_button.setEnabled(project is not None and bool(project.distributori))
        actions_row.addWidget(open_first_button)

        self.duplicate_program_button = QPushButton("Duplica programma")
        self.duplicate_program_button.clicked.connect(self.duplicate_program_requested)
        self.duplicate_program_button.setEnabled(project is not None and bool(project.distributori))
        actions_row.addWidget(self.duplicate_program_button)

        actions_row.addStretch(1)
        layout.addLayout(actions_row)
        layout.addStretch(1)

    def _build_summary_lines(self, project: AT614Project | None) -> list[str]:
        if project is None:
            return ["Nessuna struttura AT614 rilevata nella cartella corrente."]

        return [
            f"Distributori: {len(project.distributori)}",
            f"Sequenze test: {len(project.test_sequences)}",
            f"Curve e rampe: {len(project.point_series_resources)}",
            f"Calibrazioni CE16: {len(project.ce16_resources)}",
            f"Calibrazioni MMS2218: {len(project.mms2218_resources)}",
            f"File esterni mancanti: {sum(len(items) for items in project.unresolved_file_references.values())}",
        ]

    def summary_lines(self) -> list[str]:
        return [label.text() for label in self._summary_labels]
