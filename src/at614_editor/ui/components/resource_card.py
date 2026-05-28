from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QLineEdit, QPushButton, QVBoxLayout


class ResourceCard(QFrame):
    def __init__(
        self,
        title: str,
        file_name: str,
        meta_text: str,
        state_text: str,
        editable: bool = False,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setObjectName("resource-card")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        title_row = QHBoxLayout()
        self.title_label = QLabel(title)
        self.title_label.setStyleSheet("font-weight: 600;")
        title_row.addWidget(self.title_label)

        title_row.addStretch(1)
        self.state_badge = QLabel(state_text)
        self.state_badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.state_badge.setStyleSheet(
            "background: #FFF4CC; border: 1px solid #E7C25A; border-radius: 10px; padding: 2px 8px;"
        )
        title_row.addWidget(self.state_badge)
        layout.addLayout(title_row)

        self.file_label = QLineEdit(file_name)
        self.file_label.setFrame(False)
        self.file_label.setReadOnly(not editable)
        self.file_label.setStyleSheet("font-size: 14px; font-weight: 600;")
        self.file_label.setToolTip("Nome della sequenza test associata. Costituisce il collegamento logico con la sequenza.")
        layout.addWidget(self.file_label)

        self.meta_label = None
        if meta_text:
            self.meta_label = QLabel(meta_text)
            self.meta_label.setStyleSheet("color: #475467;")
            layout.addWidget(self.meta_label)

        actions_row = QHBoxLayout()
        self.select_button = QPushButton("Sel")
        self.new_button = QPushButton("Nuovo")
        self.duplicate_button = QPushButton("Dup")
        self.open_button = QPushButton("Apri")
        self.uses_button = QPushButton("Usi")

        for button in (
            self.select_button,
            self.new_button,
            self.duplicate_button,
            self.open_button,
            self.uses_button,
        ):
            actions_row.addWidget(button)

        layout.addLayout(actions_row)

    def file_name_text(self) -> str:
        return self.file_label.text().strip()
