from __future__ import annotations

from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel


class ReadOnlyBanner(QFrame):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setStyleSheet("background: #FFF3D6; border: 1px solid #E7C25A; border-radius: 8px;")

        layout = QHBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(8)

        layout.addWidget(QLabel("🔒"))
        self.text_label = QLabel(
            "Sola lettura — questo file è un output di banco. Non è un file di configurazione."
        )
        layout.addWidget(self.text_label, 1)
