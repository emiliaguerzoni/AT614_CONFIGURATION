"""Dialogo di editing per il file settings.ini del banco AT614."""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from at614_editor.domain.settings_ini import (
    FOLDER_KEY_LABELS,
    _FOLDER_KEYS,
    parse_settings_ini,
    write_settings_ini,
)


class SettingsIniEditorDialog(QDialog):
    """Dialogo che mostra e permette di modificare le chiavi del settings.ini."""

    def __init__(self, ini_path: Path, parent=None) -> None:
        super().__init__(parent)
        self.ini_path = ini_path
        self._fields: dict[str, QLineEdit] = {}

        self.setWindowTitle(f"Modifica settings.ini — {ini_path.name}")
        self.setMinimumWidth(720)
        self.setMinimumHeight(480)
        self.resize(800, 560)

        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(12)

        # --- intestazione ---
        header = QLabel(f"<b>File:</b> {ini_path}")
        header.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        header.setStyleSheet("color: #475467; font-size: 12px;")
        main_layout.addWidget(header)

        info = QLabel(
            "Modifica i percorsi delle cartelle usate dal banco AT614. "
            "Usa il tasto <b>Sfoglia…</b> per scegliere una cartella dal filesystem."
        )
        info.setWordWrap(True)
        info.setStyleSheet("color: #344054;")
        main_layout.addWidget(info)

        # --- form ---
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        form_host = QWidget()
        form = QFormLayout(form_host)
        form.setContentsMargins(8, 8, 8, 8)
        form.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapAllRows)
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(10)

        current_values = parse_settings_ini(ini_path)

        for key in _FOLDER_KEYS:
            label_text, tooltip = FOLDER_KEY_LABELS.get(key, (key, ""))
            current_value = str(current_values.get(key, "")) if key in current_values else ""

            row_widget = QWidget()
            row_layout = QHBoxLayout(row_widget)
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.setSpacing(6)

            field = QLineEdit(current_value)
            field.setPlaceholderText("Percorso cartella…")
            field.setToolTip(tooltip)
            field.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            self._fields[key] = field

            browse_btn = QPushButton("Sfoglia…")
            browse_btn.setFixedWidth(80)
            browse_btn.setToolTip(f"Seleziona la cartella per: {label_text}")
            browse_btn.clicked.connect(lambda _checked, f=field, t=tooltip: self._browse(f, t))

            row_layout.addWidget(field, 1)
            row_layout.addWidget(browse_btn)

            label_widget = QLabel(label_text)
            label_widget.setWordWrap(True)
            label_widget.setToolTip(tooltip)
            label_widget.setStyleSheet("font-weight: 600;")
            form.addRow(label_widget, row_widget)

        scroll.setWidget(form_host)
        main_layout.addWidget(scroll, 1)

        # --- pulsanti ---
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.button(QDialogButtonBox.StandardButton.Save).setText("Salva")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Annulla")
        buttons.accepted.connect(self._on_save)
        buttons.rejected.connect(self.reject)
        main_layout.addWidget(buttons)

    # ------------------------------------------------------------------ slots

    def _browse(self, field: QLineEdit, tooltip: str) -> None:
        start_dir = field.text() if field.text() else ""
        chosen = QFileDialog.getExistingDirectory(
            self, "Seleziona cartella", start_dir
        )
        if chosen:
            field.setText(chosen)

    def _on_save(self) -> None:
        mappings: dict[str, str] = {}
        for key, field in self._fields.items():
            value = field.text().strip()
            if value:
                mappings[key] = value

        try:
            write_settings_ini(self.ini_path, mappings)
        except OSError as exc:
            QMessageBox.critical(
                self,
                "Errore di scrittura",
                f"Impossibile salvare il file:\n{self.ini_path}\n\n{exc}",
            )
            return

        self.accept()

    # ------------------------------------------------------------------ API pubblica

    def current_values(self) -> dict[str, str]:
        """Restituisce i valori correnti dei campi (prima o dopo il salvataggio)."""
        return {key: field.text().strip() for key, field in self._fields.items() if field.text().strip()}
