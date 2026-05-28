"""Test per il modulo domain/refactor.py e l'ImpactDialog."""
from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from at614_editor.domain.project import load_project
from at614_editor.domain.refactor import delete_resource, get_impact, rename_resource


# ---------------------------------------------------------------------------
# Helper per costruire un progetto minimale in tmp_path
# ---------------------------------------------------------------------------


def _build_minimal_project(tmp_path: Path) -> tuple[Path, Path, Path]:
    """Crea un distributore che referenzia una sequenza test.

    Restituisce (project_root, dist_path, test_path).
    """
    dist_dir = tmp_path / "DISTRIBUTORE"
    test_dir = tmp_path / "TEST"
    dist_dir.mkdir()
    test_dir.mkdir()

    dist_path = dist_dir / "DIST_001.cfg"
    dist_path.write_text(
        "sezione1=SEQ_ALPHA\r\n"
        "sezione2=\r\n"
        "sezione3=\r\n"
        "sezione4=\r\n"
        "sezione5=\r\n"
        "calibrazioneCE16_1=\r\n"
        "calibrazioneCE16_2=\r\n"
        "calibrazioneCE16_3=\r\n"
        "calibrazioneCE16_4=\r\n"
        "calibrazioneCE16_5=\r\n",
        encoding="utf-8",
    )

    test_path = test_dir / "SEQ_ALPHA.csv"
    test_path.write_text(
        "NOME TEST;ID TEST;PARAMETRO 0\r\n"
        "Prova;ASSEGNA_NODEID;0\r\n",
        encoding="utf-8",
    )

    return tmp_path, dist_path, test_path


def _build_project_with_file_ref(tmp_path: Path) -> tuple[Path, Path, Path]:
    """Progetto con un test che referenzia un file esterno nel parametro."""
    test_dir = tmp_path / "TEST"
    test_dir.mkdir()

    ref_file = test_dir / "setup_config.txt"
    ref_file.write_text("contenuto setup", encoding="utf-8")

    test_path = test_dir / "SEQ_BETA.csv"
    test_path.write_text(
        "NOME TEST;ID TEST;PARAMETRO 0;PARAMETRO 1\r\n"
        "Scrittura;SCRITTURA_PARAMETRI;0;setup_config.txt\r\n",
        encoding="utf-8",
    )

    return tmp_path, test_path, ref_file


# ---------------------------------------------------------------------------
# Test get_impact
# ---------------------------------------------------------------------------


def test_get_impact_returns_distributor_name(tmp_path: Path) -> None:
    root, dist_path, test_path = _build_minimal_project(tmp_path)
    project = load_project(root)

    result = get_impact(project, test_path)

    assert result == ["DIST_001.cfg"]


def test_get_impact_returns_empty_for_unreferenced(tmp_path: Path) -> None:
    root, _dist_path, test_path = _build_minimal_project(tmp_path)
    # Aggiungi un file non referenziato
    orphan = (tmp_path / "TEST" / "ORPHAN.csv")
    orphan.write_text("NOME TEST;ID TEST;PARAMETRO 0\r\n", encoding="utf-8")
    project = load_project(root)

    result = get_impact(project, orphan)

    assert result == []


# ---------------------------------------------------------------------------
# Test delete_resource
# ---------------------------------------------------------------------------


def test_delete_resource_removes_file(tmp_path: Path) -> None:
    target = tmp_path / "da_eliminare.csv"
    target.write_text("contenuto", encoding="utf-8")

    delete_resource(target)

    assert not target.exists()


# ---------------------------------------------------------------------------
# Test rename_resource — update in distributore
# ---------------------------------------------------------------------------


def test_rename_resource_updates_distributore_sezione(tmp_path: Path) -> None:
    root, dist_path, test_path = _build_minimal_project(tmp_path)
    project = load_project(root)

    new_path, updated = rename_resource(project, test_path, "SEQ_RENAMED.csv")

    # Il file è stato rinominato
    assert new_path.exists()
    assert not test_path.exists()
    # Il distributore è stato aggiornato
    assert dist_path in updated
    dist_content = dist_path.read_text(encoding="utf-8")
    assert "sezione1=SEQ_RENAMED" in dist_content
    assert "SEQ_ALPHA" not in dist_content


def test_rename_resource_preserves_untouched_sections(tmp_path: Path) -> None:
    root, dist_path, test_path = _build_minimal_project(tmp_path)
    project = load_project(root)

    rename_resource(project, test_path, "SEQ_RENAMED.csv")

    dist_content = dist_path.read_text(encoding="utf-8")
    # Le sezioni 2-5 sono vuote e devono rimanere invariate
    assert "sezione2=" in dist_content


# ---------------------------------------------------------------------------
# Test rename_resource — update in sequenza TEST
# ---------------------------------------------------------------------------


def test_rename_resource_updates_test_parameter_value(tmp_path: Path) -> None:
    root, test_path, ref_file = _build_project_with_file_ref(tmp_path)
    project = load_project(root)

    new_path, updated = rename_resource(project, ref_file, "setup_config_v2.txt")

    assert new_path.exists()
    assert not ref_file.exists()
    assert test_path in updated
    test_content = test_path.read_text(encoding="utf-8")
    assert "setup_config_v2.txt" in test_content
    assert "setup_config.txt" not in test_content


# ---------------------------------------------------------------------------
# Test rename_resource — conflitto con file esistente
# ---------------------------------------------------------------------------


def test_rename_resource_raises_on_name_conflict(tmp_path: Path) -> None:
    root, dist_path, test_path = _build_minimal_project(tmp_path)
    # Crea un file con il nome di destinazione
    conflict = tmp_path / "TEST" / "SEQ_EXISTING.csv"
    conflict.write_text("x", encoding="utf-8")
    project = load_project(root)

    with pytest.raises(ValueError, match="SEQ_EXISTING.csv"):
        rename_resource(project, test_path, "SEQ_EXISTING.csv")

    # Il file originale non è stato toccato
    assert test_path.exists()


# ---------------------------------------------------------------------------
# Test ImpactDialog (UI)
# ---------------------------------------------------------------------------


def test_impact_dialog_shows_items(qt_app) -> None:
    from at614_editor.ui.components.impact_dialog import ImpactDialog

    dialog = ImpactDialog(
        title="Elimina risorsa",
        description="Vuoi eliminare questa risorsa?",
        impact_items=["DIST_001.cfg", "DIST_002.cfg"],
        action_label="Elimina",
    )

    items = [dialog.impact_list.item(i).text() for i in range(dialog.impact_list.count())]
    assert items == ["DIST_001.cfg", "DIST_002.cfg"]


def test_impact_dialog_rename_mode_exposes_name_input(qt_app) -> None:
    from at614_editor.ui.components.impact_dialog import ImpactDialog

    dialog = ImpactDialog(
        title="Rinomina risorsa",
        description="Rinomina la risorsa.",
        impact_items=[],
        action_label="Rinomina",
        rename_mode=True,
        initial_name="vecchio_nome.csv",
    )

    assert dialog.new_name() == "vecchio_nome.csv"


def test_impact_dialog_empty_impact_list(qt_app) -> None:
    from at614_editor.ui.components.impact_dialog import ImpactDialog

    dialog = ImpactDialog(
        title="Elimina",
        description="Nessun riferimento.",
        impact_items=[],
        action_label="Elimina",
    )

    assert dialog.impact_list.count() == 0
