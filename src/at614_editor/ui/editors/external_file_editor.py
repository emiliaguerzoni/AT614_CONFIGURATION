from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Signal
from PySide6.QtWidgets import QLabel, QPlainTextEdit, QVBoxLayout, QWidget

from at614_editor.ui.components.file_widget import FileWidget
from at614_editor.ui.editors.resource_actions import build_duplicate_path
from at614_editor.ui.tooltip_manager import get_common_tooltip


WARNING_KNOWN_FILE = (
    "Formato esterno non strutturato: modifica in modalità testo assistito. "
    "Il file viene salvato così com'è, senza validazione semantica."
)
WARNING_MISSING_FILE = (
    "Riferimento esterno non risolto: il file non è presente nel progetto. "
    "L'anteprima resta vuota fino a quando la risorsa non viene creata."
)


def _detect_encoding(raw: bytes) -> str:
    try:
        raw.decode("utf-8")
    except UnicodeDecodeError:
        return "cp1252"
    return "utf-8"


class ExternalFileEditor(QWidget):
    state_changed = Signal()
    editor_name = "ExternalFileEditor"

    def __init__(
        self,
        display_label: str,
        source_path: Path | None = None,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.display_label = display_label
        self.source_path = source_path
        self._encoding = "utf-8"
        self._original_text = ""

        file_exists = source_path is not None and source_path.exists() and source_path.is_file()
        self.supports_save = file_exists
        self.supports_duplicate = file_exists

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        self.file_widget = FileWidget("□")
        self.file_widget.set_reference(display_label)
        layout.addWidget(self.file_widget)

        self.warning_label = QLabel(WARNING_KNOWN_FILE if file_exists else WARNING_MISSING_FILE)
        self.warning_label.setWordWrap(True)
        self.warning_label.setStyleSheet(
            "background: #FFF4CC; border: 1px solid #E7C25A; border-radius: 8px;"
            " padding: 8px 12px; color: #6E5400;"
        )
        layout.addWidget(self.warning_label)

        self.preview_editor = QPlainTextEdit()
        self.preview_editor.setReadOnly(not file_exists)
        self.preview_editor.setToolTip("Modifica il contenuto del file. Le modifiche verranno salvate al salvataggio del progetto.")
        if file_exists:
            self._original_text = self._load_file_text(source_path)
            self.preview_editor.setPlainText(self._original_text)
        else:
            self.preview_editor.setPlaceholderText(
                "Anteprima non disponibile: la risorsa non è presente nel progetto."
            )
        self.preview_editor.textChanged.connect(self.state_changed.emit)
        layout.addWidget(self.preview_editor)

    def _load_file_text(self, source_path: Path) -> str:
        raw = source_path.read_bytes()
        self._encoding = _detect_encoding(raw)
        return raw.decode(self._encoding)

    def is_dirty(self) -> bool:
        if not self.supports_save:
            return False
        return self.preview_editor.toPlainText() != self._original_text

    def save_changes(self) -> Path | None:
        if not self.supports_save or self.source_path is None:
            return None

        text = self.preview_editor.toPlainText()
        self.source_path.write_bytes(text.encode(self._encoding))
        return self.source_path

    def duplicate_resource(self) -> Path | None:
        if not self.supports_duplicate or self.source_path is None:
            return None

        duplicate_path = build_duplicate_path(self.source_path)
        text = self.preview_editor.toPlainText()
        duplicate_path.write_bytes(text.encode(self._encoding))
        return duplicate_path
