from __future__ import annotations

from PySide6.QtWidgets import QListWidget, QPushButton, QTabWidget, QVBoxLayout, QWidget


class RPanel(QWidget):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.tabs = QTabWidget()

        self.references_list = QListWidget()
        self.validation_list = QListWidget()

        actions_widget = QWidget()
        actions_layout = QVBoxLayout(actions_widget)
        actions_layout.setContentsMargins(8, 8, 8, 8)
        actions_layout.setSpacing(8)

        self.quick_action_buttons: list[QPushButton] = []
        # Pulsanti con riferimento nominato per il wiring esterno
        self.elimina_button = QPushButton("Elimina")
        self.rinomina_button = QPushButton("Rinomina")
        for action_text in (
            "Nuovo collegato",
            "Duplica collegato",
            "Sostituisci",
            "Apri",
        ):
            button = QPushButton(action_text)
            self.quick_action_buttons.append(button)
            actions_layout.addWidget(button)
        for button in (self.rinomina_button, self.elimina_button):
            self.quick_action_buttons.append(button)
            actions_layout.addWidget(button)
        actions_layout.addStretch(1)

        self.tabs.addTab(self.references_list, "Riferimenti")
        self.tabs.addTab(self.validation_list, "Validazione")
        self.tabs.addTab(actions_widget, "Azioni rapide")

        layout.addWidget(self.tabs)
        self.set_context([], [])

    def set_context(self, references: list[str], validation: list[str], focus_validation: bool = False) -> None:
        self.references_list.clear()
        self.validation_list.clear()

        for item in references or ["Nessun riferimento disponibile"]:
            self.references_list.addItem(item)

        for item in validation or ["Nessun avviso"]:
            self.validation_list.addItem(item)

        self.tabs.setCurrentIndex(1 if focus_validation and validation else 0)

    def show_references(self, references: list[str]) -> None:
        self.references_list.clear()
        for item in references or ["Nessun riferimento disponibile"]:
            self.references_list.addItem(item)
        self.tabs.setCurrentIndex(0)

    def update_validation(self, validation: list[str], focus_validation: bool = False) -> None:
        self.validation_list.clear()
        for item in validation or ["Nessun avviso"]:
            self.validation_list.addItem(item)
        if focus_validation and validation:
            self.tabs.setCurrentIndex(1)

    def tab_labels(self) -> list[str]:
        return [self.tabs.tabText(index) for index in range(self.tabs.count())]

    def quick_action_labels(self) -> list[str]:
        return [button.text() for button in self.quick_action_buttons]
