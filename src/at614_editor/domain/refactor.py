from __future__ import annotations

from pathlib import Path

from at614_editor.domain.models import DistributoreConfig, TestRow, TestSequence
from at614_editor.domain.parsers.distributore import serialize as serialize_distributore
from at614_editor.domain.parsers.test_csv import serialize as serialize_test_csv
from at614_editor.domain.project import AT614Project


# ---------------------------------------------------------------------------
# Query helpers
# ---------------------------------------------------------------------------


def get_impact(project: AT614Project, path: Path) -> list[str]:
    """Restituisce i nomi (ordinati) dei file che referenziano la risorsa."""
    return sorted(source.name for source in project.get_used_by(path))


# ---------------------------------------------------------------------------
# Delete
# ---------------------------------------------------------------------------


def delete_resource(path: Path) -> None:
    """Elimina il file dal disco.

    La verifica dei riferimenti dipendenti è responsabilità del chiamante;
    questa funzione non aggiorna nessun altro file.
    """
    path.unlink()


# ---------------------------------------------------------------------------
# Rename con aggiornamento cascata
# ---------------------------------------------------------------------------


def rename_resource(
    project: AT614Project,
    old_path: Path,
    new_name: str,
) -> tuple[Path, list[Path]]:
    """Rinomina una risorsa e aggiorna tutti i riferimenti nel progetto.

    Restituisce (new_path, lista_file_aggiornati).
    Lancia ValueError se new_name corrisponde a un file già esistente.
    """
    new_path = old_path.with_name(new_name)
    if new_path.exists():
        raise ValueError(f"Esiste già un file con il nome: {new_name}")

    updated_files: list[Path] = []

    for source_path in project.get_used_by(old_path):
        if source_path in project.distributori:
            updated = _update_distributore_refs(
                project.distributori[source_path], old_path, new_path
            )
            if updated is not None:
                source_path.write_bytes(serialize_distributore(updated))
                updated_files.append(source_path)

        elif source_path in project.test_sequences:
            updated_seq = _update_test_refs(
                project.test_sequences[source_path], old_path.name, new_path.name
            )
            if updated_seq is not None:
                source_path.write_bytes(serialize_test_csv(updated_seq))
                updated_files.append(source_path)

    old_path.rename(new_path)
    return new_path, updated_files


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _update_distributore_refs(
    config: DistributoreConfig, old_path: Path, new_path: Path
) -> DistributoreConfig | None:
    """Restituisce la config aggiornata, oppure None se nulla è cambiato."""
    changed = False

    new_sezioni = dict(config.sezioni)
    for key, val in new_sezioni.items():
        if val == old_path.stem:
            new_sezioni[key] = new_path.stem
            changed = True

    new_cal = dict(config.calibrazioni_ce16)
    for key, val in new_cal.items():
        if val == old_path.name:
            new_cal[key] = new_path.name
            changed = True

    if not changed:
        return None

    return DistributoreConfig(
        source_path=config.source_path,
        sezioni=new_sezioni,
        calibrazioni_ce16=new_cal,
        extras=config.extras,
        lines=config.lines,
        encoding=config.encoding,
        line_ending=config.line_ending,
        endswith_newline=config.endswith_newline,
    )


def _update_test_refs(
    sequence: TestSequence, old_name: str, new_name: str
) -> TestSequence | None:
    """Restituisce la sequenza aggiornata, oppure None se nulla è cambiato."""
    changed = False
    new_rows: list[TestRow] = []

    for row in sequence.rows:
        new_params = list(row.parameters)
        for i, val in enumerate(new_params):
            if val.strip() == old_name:
                new_params[i] = new_name
                changed = True
        new_rows.append(
            TestRow(
                name=row.name,
                test_id=row.test_id,
                index_raw=row.index_raw,
                parameters=new_params,
            )
        )

    if not changed:
        return None

    return TestSequence(
        source_path=sequence.source_path,
        header_fields=sequence.header_fields,
        rows=new_rows,
        encoding=sequence.encoding,
        line_ending=sequence.line_ending,
        endswith_newline=sequence.endswith_newline,
    )
