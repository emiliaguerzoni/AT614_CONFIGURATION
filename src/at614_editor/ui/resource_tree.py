from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from PySide6.QtCore import QObject, QThread, Qt, Signal
from PySide6.QtWidgets import QTreeWidget, QTreeWidgetItem, QVBoxLayout, QWidget


class _DirectoryLoader(QObject):
    """Carica il contenuto di una cartella in un thread separato."""

    finished = Signal(object, list)  # (directory_path, entries)
    error = Signal(object, str)      # (directory_path, error_message)

    def __init__(self, directory_path: Path) -> None:
        super().__init__()
        self._directory_path = directory_path

    def run(self) -> None:
        try:
            entries = sorted(
                self._directory_path.iterdir(),
                key=lambda p: (not p.is_dir(), p.name.lower()),
            )
            self.finished.emit(self._directory_path, entries)
        except OSError as exc:
            self.error.emit(self._directory_path, str(exc))

from at614_editor.domain.project import AT614Project


CATEGORY_DEFINITIONS: list[tuple[str, str]] = [
    ("distributori", "Distributori"),
    ("sequenze_test", "Sequenze test"),
    ("curve_comando", "Curve comando"),
    ("curve_limite", "Curve limite"),
    ("rampe_xy", "Rampe XY"),
    ("calibrazioni_ce16", "Calibrazioni CE16"),
    ("calibrazioni_mms2218", "Calibrazioni MMS2218"),
    ("file_esterni", "File esterni"),
    ("output_banco", "Output banco"),
]


@dataclass(slots=True)
class ResourceTreeSelection:
    category_key: str
    category_label: str
    display_label: str
    path: Path | None


class ResourceTree(QWidget):
    resource_activated = Signal(object)

    def __init__(self, project: AT614Project | None, parent=None) -> None:
        super().__init__(parent)
        self.project = project
        self._path_items: dict[Path, QTreeWidgetItem] = {}
        self._pending_dir_threads: list[tuple] = []  # (loader, thread)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.currentItemChanged.connect(self._emit_current_selection)
        self.tree.itemExpanded.connect(self._on_item_expanded)
        layout.addWidget(self.tree)

        self._build_tree()

    def _build_tree(self) -> None:
        self.tree.clear()
        self._path_items.clear()
        grouped_selections = self._build_grouped_selections()

        for category_key, category_label in CATEGORY_DEFINITIONS:
            top_item = QTreeWidgetItem([category_label])
            top_item.setData(
                0,
                Qt.ItemDataRole.UserRole,
                ResourceTreeSelection(category_key, category_label, category_label, None),
            )
            self.tree.addTopLevelItem(top_item)

            children = grouped_selections.get(category_key, [])
            if not children:
                placeholder_text = "Archivio non configurato in settings.ini" if category_key == "output_banco" else "Nessuna risorsa"
                placeholder_item = QTreeWidgetItem([placeholder_text])
                placeholder_item.setFlags(placeholder_item.flags() & ~Qt.ItemFlag.ItemIsSelectable)
                placeholder_item.setDisabled(True)
                top_item.addChild(placeholder_item)
                continue

            if category_key == "output_banco":
                for selection in children:
                    root_item = QTreeWidgetItem([selection.display_label])
                    root_item.setData(0, Qt.ItemDataRole.UserRole, selection)
                    root_item.setData(0, Qt.ItemDataRole.UserRole + 1, selection.path)
                    top_item.addChild(root_item)
                    if selection.path is not None and selection.path.exists() and selection.path.is_dir():
                        placeholder_item = QTreeWidgetItem(["Carica cartella..."])
                        placeholder_item.setFlags(placeholder_item.flags() & ~Qt.ItemFlag.ItemIsSelectable)
                        root_item.addChild(placeholder_item)
                    if selection.path is not None:
                        self._path_items[selection.path] = root_item
            else:
                for selection in children:
                    child_item = QTreeWidgetItem([selection.display_label])
                    child_item.setData(0, Qt.ItemDataRole.UserRole, selection)
                    top_item.addChild(child_item)
                    if selection.path is not None:
                        self._path_items[selection.path] = child_item

            top_item.setExpanded(True)

    def _build_grouped_selections(self) -> dict[str, list[ResourceTreeSelection]]:
        grouped: dict[str, list[ResourceTreeSelection]] = {key: [] for key, _ in CATEGORY_DEFINITIONS}
        if self.project is None:
            return grouped

        for path in sorted(self.project.distributori):
            grouped["distributori"].append(
                ResourceTreeSelection("distributori", "Distributori", path.name, path)
            )

        for path in sorted(self.project.test_sequences):
            grouped["sequenze_test"].append(
                ResourceTreeSelection("sequenze_test", "Sequenze test", path.name, path)
            )

        for path, resource in sorted(self.project.point_series_resources.items()):
            category_key = {
                "CURVE_COMANDO": "curve_comando",
                "CURVE_LIMITE": "curve_limite",
                "RAMPE_XY": "rampe_xy",
            }.get(path.parent.name)
            if category_key is None:
                continue
            grouped[category_key].append(
                ResourceTreeSelection(category_key, dict(CATEGORY_DEFINITIONS)[category_key], path.name, path)
            )

        for path, resource in sorted(self.project.ce16_resources.items()):
            grouped["calibrazioni_ce16"].append(
                ResourceTreeSelection(
                    "calibrazioni_ce16",
                    "Calibrazioni CE16",
                    f"{resource.profile_name} · {path.name}",
                    path,
                )
            )

        for path, resource in sorted(self.project.mms2218_resources.items()):
            grouped["calibrazioni_mms2218"].append(
                ResourceTreeSelection(
                    "calibrazioni_mms2218",
                    "Calibrazioni MMS2218",
                    f"{resource.profile_name} · {path.name}",
                    path,
                )
            )

        external_labels = sorted(
            {
                reference
                for references in self.project.unresolved_file_references.values()
                for reference in references
                if reference.lower().endswith(".txt")
            }
        )
        for reference in external_labels:
            grouped["file_esterni"].append(
                ResourceTreeSelection("file_esterni", "File esterni", reference, None)
            )

        if self.project.folder_graph_saved is not None:
            grouped["output_banco"].append(
                ResourceTreeSelection(
                    "output_banco", "Output banco",
                    "Archivio grafici (GRAPH)",
                    self.project.folder_graph_saved,
                )
            )
        if self.project.folder_graph_saved_last_acq is not None:
            grouped["output_banco"].append(
                ResourceTreeSelection(
                    "output_banco", "Output banco",
                    "Ultima acquisizione (RAMPE_XY_LAST)",
                    self.project.folder_graph_saved_last_acq,
                )
            )

        return grouped

    def _on_item_expanded(self, item: QTreeWidgetItem) -> None:
        if item.childCount() == 0:
            return
        first_child = item.child(0)
        if first_child.text(0) != "Carica cartella...":
            return
        archive_path = item.data(0, Qt.ItemDataRole.UserRole + 1)
        if not isinstance(archive_path, Path):
            return

        # Sostituisce il placeholder con indicatore di caricamento (no block UI)
        first_child.setText(0, "Caricamento\u2026")
        first_child.setDisabled(True)
        self._start_dir_load(item, archive_path)

    def _start_dir_load(self, item: QTreeWidgetItem, path: Path) -> None:
        """Avvia il caricamento asincrono del contenuto di una cartella."""
        loader = _DirectoryLoader(path)
        thread = QThread()
        loader.moveToThread(thread)
        thread.started.connect(loader.run)
        loader.finished.connect(lambda _p, entries: self._on_dir_loaded(item, entries))
        loader.error.connect(lambda _p, msg: self._on_dir_error(item, msg))
        loader.finished.connect(thread.quit)
        loader.error.connect(thread.quit)
        # Mantiene un riferimento sia al loader che al thread per evitare il GC
        pair = (loader, thread)
        self._pending_dir_threads.append(pair)
        thread.finished.connect(lambda p=pair: self._pending_dir_threads.remove(p)
            if p in self._pending_dir_threads else None)
        thread.finished.connect(loader.deleteLater)
        thread.finished.connect(thread.deleteLater)
        thread.start()

    def _on_dir_loaded(self, item: QTreeWidgetItem, entries: list) -> None:
        # Rimuove il placeholder "Caricamento…"
        if item.childCount() > 0 and item.child(0).text(0) == "Caricamento\u2026":
            item.removeChild(item.child(0))
        self._populate_output_archive_items(item, entries)

    def _on_dir_error(self, item: QTreeWidgetItem, message: str) -> None:
        if item.childCount() > 0 and item.child(0).text(0) == "Caricamento\u2026":
            item.removeChild(item.child(0))
        error_item = QTreeWidgetItem([f"Errore: {message}"])
        error_item.setFlags(error_item.flags() & ~Qt.ItemFlag.ItemIsSelectable)
        error_item.setDisabled(True)
        item.addChild(error_item)

    def _populate_output_archive_items(self, parent_item: QTreeWidgetItem, entries: list) -> None:
        """Popola i figli di un nodo archivio con le entries pre-caricate (già ordinate)."""
        if not entries:
            empty_item = QTreeWidgetItem(["Cartella vuota"])
            empty_item.setFlags(empty_item.flags() & ~Qt.ItemFlag.ItemIsSelectable)
            empty_item.setDisabled(True)
            parent_item.addChild(empty_item)
            return

        for path in entries:
            node_label = path.name
            child_item = QTreeWidgetItem([node_label])
            selection = ResourceTreeSelection(
                "output_banco",
                "Output banco",
                node_label,
                path,
            )
            child_item.setData(0, Qt.ItemDataRole.UserRole, selection)
            child_item.setData(0, Qt.ItemDataRole.UserRole + 1, path)
            parent_item.addChild(child_item)
            self._path_items[path] = child_item
            if path.is_dir():
                placeholder_child = QTreeWidgetItem(["Carica cartella..."])
                placeholder_child.setFlags(placeholder_child.flags() & ~Qt.ItemFlag.ItemIsSelectable)
                child_item.addChild(placeholder_child)

    def _emit_current_selection(self, current: QTreeWidgetItem | None, _previous: QTreeWidgetItem | None) -> None:
        if current is None:
            return

        selection = current.data(0, Qt.ItemDataRole.UserRole)
        if selection is not None:
            self.resource_activated.emit(selection)

    def top_level_labels(self) -> list[str]:
        return [self.tree.topLevelItem(index).text(0) for index in range(self.tree.topLevelItemCount())]

    def set_project(self, project: AT614Project | None) -> None:
        self.project = project
        self._build_tree()

    def activate_first_resource(self) -> ResourceTreeSelection | None:
        for _, _label in CATEGORY_DEFINITIONS:
            pass

        for index in range(self.tree.topLevelItemCount()):
            top_item = self.tree.topLevelItem(index)
            for child_index in range(top_item.childCount()):
                child_item = top_item.child(child_index)
                selection = child_item.data(0, Qt.ItemDataRole.UserRole)
                if isinstance(selection, ResourceTreeSelection) and selection.path is not None:
                    self.tree.setCurrentItem(child_item)
                    return selection
        return None

    def activate_path(self, path: Path) -> bool:
        item = self._path_items.get(path)
        if item is None:
            return False
        self.tree.setCurrentItem(item)
        return True

    def activate_selection(self, target_selection: ResourceTreeSelection) -> bool:
        for index in range(self.tree.topLevelItemCount()):
            top_item = self.tree.topLevelItem(index)
            for child_index in range(top_item.childCount()):
                child_item = top_item.child(child_index)
                selection = child_item.data(0, Qt.ItemDataRole.UserRole)
                if selection == target_selection:
                    self.tree.setCurrentItem(child_item)
                    return True
            selection = top_item.data(0, Qt.ItemDataRole.UserRole)
            if selection == target_selection:
                self.tree.setCurrentItem(top_item)
                return True
        return False
