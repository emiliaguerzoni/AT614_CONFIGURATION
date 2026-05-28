"""Modulo per la selezione dei file tramite file dialog, con gestione percorsi relativi."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Callable

from PySide6.QtWidgets import QFileDialog, QWidget

from at614_editor.domain.project import AT614Project

logger = logging.getLogger(__name__)


class FilePicker:
    """Gestisce la selezione dei file tramite dialog, mantenendo percorsi relativi alla project root."""

    def __init__(self, project: AT614Project) -> None:
        """Inizializza il file picker con un progetto AT614.

        Args:
            project: Progetto AT614Project che contiene la root path e l'indice delle risorse
        """
        self.project = project
        self.root_path = project.root_path

    def pick_file(
        self,
        parent: QWidget | None = None,
        title: str = "Seleziona file",
        resource_type: str | None = None,
    ) -> str | None:
        """Apre un dialog per la selezione di un file.

        Args:
            parent: Widget padre per il dialog
            title: Titolo del dialog
            resource_type: Tipo di risorsa ('curva_comando', 'limite_inf', ecc.) - usato per filtrare i file

        Returns:
            Percorso relativo al file selezionato, oppure None se annullato
        """
        try:
            logger.debug(f"FilePicker: apertura dialog per tipo '{resource_type}'")
            # Determina il filtro e la cartella di partenza in base al tipo di risorsa
            file_filter, start_dir = self._get_filter_and_dir(resource_type)
            logger.debug(f"FilePicker: start_dir={start_dir}, filter={file_filter}")

            # Apri il file dialog
            file_path, _ = QFileDialog.getOpenFileName(
                parent,
                title,
                str(start_dir),
                file_filter,
            )

            if not file_path:
                logger.debug("FilePicker: selezione annullata dall'utente")
                return None

            logger.info(f"FilePicker: file selezionato: {file_path}")
            # Converti il percorso assoluto in percorso relativo alla project root
            result = self._to_relative_path(Path(file_path))
            logger.debug(f"FilePicker: percorso relativo: {result}")
            return result
        except Exception as e:
            logger.exception(f"FilePicker: errore durante la selezione del file: {e}")
            return None

    def _get_filter_and_dir(self, resource_type: str | None) -> tuple[str, Path]:
        """Determina il filtro file e la cartella di partenza in base al tipo di risorsa.

        Args:
            resource_type: Tipo di risorsa ('curva_comando', 'limite_inf', 'rampa_xy', 'external_config', ecc.)

        Returns:
            Tupla (file_filter, start_directory)
        """
        # Mappatura tra tipo risorsa e cartella nel progetto
        _folder_moduli = getattr(self.project, "folder_configurazione_moduli", None)
        resource_dirs = {
            "curva_comando": self.root_path / "CURVE_COMANDO",
            "limite_inf": self.root_path / "CURVE_LIMITE",
            "limite_sup": self.root_path / "CURVE_LIMITE",
            "rampa_xy": self.root_path / "RAMPE_XY",
            "sequenze_test": self.root_path / "TEST",
            "ce16": self.root_path / "CONFIGURAZIONE",
            "mms2218": self.root_path / "CONFIGURAZIONE",
            "parametri": self.root_path / "CONFIGURAZIONE",
            # external_config: usa FolderConfigurazioneModuli da settings.ini se disponibile
            "external_config": _folder_moduli if _folder_moduli and _folder_moduli.exists() else self.root_path,
        }

        # Filtri file per tipo
        file_filters = {
            "curva_comando": "File CSV (*.csv);;Tutti i file (*.*)",
            "limite_inf": "File CSV (*.csv);;Tutti i file (*.*)",
            "limite_sup": "File CSV (*.csv);;Tutti i file (*.*)",
            "rampa_xy": "File CSV (*.csv);;Tutti i file (*.*)",
            "sequenze_test": "File CSV (*.csv);;Tutti i file (*.*)",
            "ce16": "File CSV (*.csv);;Tutti i file (*.*)",
            "mms2218": "File CSV (*.csv);;Tutti i file (*.*)",
            "parametri": "File CFG (*.cfg);;Tutti i file (*.*)",
            "external_config": "File CFG (*.cfg);;File CSV (*.csv);;Tutti i file (*.*)",
            "generic_file": "Tutti i file (*.*)",
        }

        # Determina directory di partenza
        start_dir = resource_dirs.get(resource_type or "", self.root_path)
        if not start_dir.exists():
            start_dir = self.root_path

        # Determina filtro
        file_filter = file_filters.get(resource_type or "generic_file", "Tutti i file (*.*)")

        return file_filter, start_dir

    def _to_relative_path(self, file_path: Path) -> str:
        """Restituisce solo il nome del file (compatibile con il formato VB6 che
        aggiunge a runtime il prefisso di cartella da settings.ini).

        Args:
            file_path: Percorso assoluto del file selezionato

        Returns:
            Solo il nome del file (senza directory)
        """
        return file_path.name


class FilePickerButton:
    """Helper per aggiungere la funzionalità di file picker a un QLineEdit."""

    def __init__(self, line_edit, picker: FilePicker, resource_type: str | None = None) -> None:
        """Collega un pulsante di selezione a un QLineEdit.

        Args:
            line_edit: QLineEdit dove inserire il percorso selezionato
            picker: FilePicker da usare per la selezione
            resource_type: Tipo di risorsa per filtrare i file
        """
        self.line_edit = line_edit
        self.picker = picker
        self.resource_type = resource_type

    def pick(self, parent: QWidget | None = None) -> None:
        """Apre il file dialog e inserisce il percorso nel QLineEdit."""
        selected_path = self.picker.pick_file(
            parent=parent,
            title="Seleziona file",
            resource_type=self.resource_type,
        )

        if selected_path:
            self.line_edit.setText(selected_path)
