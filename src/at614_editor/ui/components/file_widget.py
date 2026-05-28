from __future__ import annotations

from typing import Callable

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QLineEdit, QPushButton, QSizePolicy


class FileWidget(QFrame):
    def __init__(self, resource_label: str, editable: bool = False, parent=None, on_select: Callable[[], None] | None = None) -> None:
        super().__init__(parent)
        self._resource_label = resource_label
        self._on_select = on_select
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setObjectName("file-widget")

        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(8)

        self.icon_label = QLabel(resource_label)
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.icon_label.setFixedWidth(36)
        layout.addWidget(self.icon_label)

        self.name_label = QLineEdit()
        self.name_label.setFrame(False)
        self.name_label.setReadOnly(not editable)
        self.name_label.setPlaceholderText("— riferimento mancante —")
        self.name_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self.name_label.setToolTip("Immetti il nome del file di riferimento")
        layout.addWidget(self.name_label, 1)

        self.select_button = QPushButton("Seleziona")
        self.new_button = QPushButton("Nuovo")
        self.duplicate_button = QPushButton("Duplica")
        self.open_button = QPushButton("Apri")
        self.uses_button = QPushButton()

        # Connetti il pulsante Seleziona al callback
        if self._on_select is not None:
            self.select_button.clicked.connect(self._on_select)

        for button in (
            self.select_button,
            self.new_button,
            self.duplicate_button,
            self.open_button,
            self.uses_button,
        ):
            layout.addWidget(button)

        self.set_reference(None)

    def set_reference(self, reference_name: str | None, uses_count: int = 0) -> None:
        is_missing = not reference_name
        self.name_label.setText(reference_name or "")
        self.name_label.setStyleSheet("color: #B42318;" if is_missing else "color: #0F172A;")
        self.open_button.setEnabled(not is_missing)
        self.uses_button.setEnabled(not is_missing)
        self.uses_button.setText(f"Usi ({uses_count})")

    def reference_text(self) -> str:
        return self.name_label.text().strip()

    def action_labels(self) -> list[str]:
        return [
            self.select_button.text(),
            self.new_button.text(),
            self.duplicate_button.text(),
            self.open_button.text(),
            self.uses_button.text(),
        ]
