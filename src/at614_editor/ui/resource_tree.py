from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import QTreeWidget, QTreeWidgetItem, QVBoxLayout, QWidget

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

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.currentItemChanged.connect(self._emit_current_selection)
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
                placeholder_text = "Viewer output in sviluppo" if category_key == "output_banco" else "Nessuna risorsa"
                placeholder_item = QTreeWidgetItem([placeholder_text])
                placeholder_item.setFlags(placeholder_item.flags() & ~Qt.ItemFlag.ItemIsSelectable)
                placeholder_item.setDisabled(True)
                top_item.addChild(placeholder_item)
                continue

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

        return grouped

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
