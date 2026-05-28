from __future__ import annotations

from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QVBoxLayout,
)


class ImpactDialog(QDialog):
    """Modale di conferma per azioni che impattano altri file del progetto.

    Usata sia per *Elimina* che per *Rinomina*. In modalità rename mostra
    un campo di testo per il nuovo nome; in entrambi i casi mostra la lista
    dei file dipendenti (ImpactList).
    """

    def __init__(
        self,
        title: str,
        description: str,
        impact_items: list[str],
        action_label: str,
        rename_mode: bool = False,
        initial_name: str = "",
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumWidth(480)
        self._rename_mode = rename_mode

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(12)

        desc_label = QLabel(description)
        desc_label.setWordWrap(True)
        layout.addWidget(desc_label)

        if rename_mode:
            name_row = QHBoxLayout()
            name_row.addWidget(QLabel("Nuovo nome:"))
            self.name_input = QLineEdit(initial_name)
            name_row.addWidget(self.name_input, 1)
            layout.addLayout(name_row)
        else:
            # placeholder non aggiunto al layout
            self.name_input = QLineEdit()

        if impact_items:
            count = len(impact_items)
            impact_title = QLabel(
                f"Riferimenti che verranno aggiornati ({count}):"
            )
            impact_title.setStyleSheet("font-weight: 600;")
            layout.addWidget(impact_title)

            self.impact_list = QListWidget()
            self.impact_list.setMaximumHeight(200)
            for item in impact_items:
                self.impact_list.addItem(item)
            layout.addWidget(self.impact_list)
        else:
            no_refs = QLabel("Nessun altro file fa riferimento a questa risorsa.")
            no_refs.setStyleSheet("color: #475467;")
            layout.addWidget(no_refs)
            self.impact_list = QListWidget()

        buttons = QDialogButtonBox()
        self._confirm_button = buttons.addButton(
            action_label, QDialogButtonBox.ButtonRole.AcceptRole
        )
        buttons.addButton("Annulla", QDialogButtonBox.ButtonRole.RejectRole)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def new_name(self) -> str:
        """Restituisce il testo del campo nuovo nome (solo in rename_mode)."""
        return self.name_input.text().strip()
