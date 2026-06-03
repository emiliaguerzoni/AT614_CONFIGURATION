from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from pathlib import Path

from at614_editor.domain.models import Ce16Resource, DistributoreConfig, Mms2218Resource, PointSeriesResource, TestSequence
from at614_editor.domain.parsers.ce16 import parse as parse_ce16
from at614_editor.domain.parsers.distributore import parse as parse_distributore
from at614_editor.domain.parsers.mms2218 import parse as parse_mms2218
from at614_editor.domain.parsers.point_series import parse as parse_point_series
from at614_editor.domain.parsers.test_csv import parse as parse_test_csv
from at614_editor.domain.settings_ini import parse_settings_ini
from at614_editor.domain.test_schema import extract_test_file_references

logger = logging.getLogger(__name__)
_NORMALIZE_SEPARATOR_RE = re.compile(r"[\\/\s_]+")


POINT_SERIES_DIRECTORIES = ("CURVE_COMANDO", "CURVE_LIMITE", "RAMPE_XY")
SETTINGS_PROGRAM_DIR = "SETTAGGI PROGRAMMA"


@dataclass(slots=True)
class AT614Project:
    root_path: Path
    distributori: dict[Path, DistributoreConfig]
    test_sequences: dict[Path, TestSequence]
    point_series_resources: dict[Path, PointSeriesResource]
    ce16_resources: dict[Path, Ce16Resource]
    mms2218_resources: dict[Path, Mms2218Resource]
    resource_index: dict[str, set[Path]]
    uses: dict[Path, set[Path]] = field(default_factory=dict)
    used_by: dict[Path, set[Path]] = field(default_factory=dict)
    unresolved_file_references: dict[Path, set[str]] = field(default_factory=dict)
    # Percorso assoluto di FolderConfigurazioneModuli da settings.ini (file parametri .txt)
    folder_configurazione_moduli: Path | None = None
    # Percorsi archivio output da settings.ini (possono essere share di rete)
    folder_graph_saved: Path | None = None
    folder_graph_saved_last_acq: Path | None = None

    def add_reference(self, source: Path, target: Path) -> None:
        self.uses.setdefault(source, set()).add(target)
        self.used_by.setdefault(target, set()).add(source)

    def get_uses(self, source: Path) -> set[Path]:
        return self.uses.get(source, set())

    def get_used_by(self, target: Path) -> set[Path]:
        return self.used_by.get(target, set())

    def add_unresolved_reference(self, source: Path, raw_value: str) -> None:
        self.unresolved_file_references.setdefault(source, set()).add(raw_value)

    def get_unresolved_references(self, source: Path) -> set[str]:
        return self.unresolved_file_references.get(source, set())

    def resolve_resource_path(self, raw_value: str) -> Path | None:
        resolved_path = _resolve_resource_path(self.resource_index, raw_value, self.folder_configurazione_moduli)
        if resolved_path is not None:
            return resolved_path

        normalized_value = raw_value.strip()
        if not normalized_value:
            return None

        test_stem = Path(normalized_value).stem
        for path in self.test_sequences:
            if path.stem == test_stem:
                return path

        return None

    def get_usage_count(self, target: Path | None) -> int:
        if target is None:
            return 0
        return len(self.get_used_by(target))


def _normalize_resource_key(raw_value: str) -> str:
    normalized = raw_value.strip().lower()
    normalized = normalized.replace("\\", "/")
    normalized = _NORMALIZE_SEPARATOR_RE.sub(" ", normalized)
    return normalized.strip()


def _add_resource_alias(resource_index: dict[str, set[Path]], alias: str, path: Path) -> None:
    normalized_alias = _normalize_resource_key(alias)
    if not normalized_alias:
        return

    resource_index.setdefault(normalized_alias, set()).add(path)


def _build_resource_file_index(root_path: Path) -> dict[str, set[Path]]:
    resource_index: dict[str, set[Path]] = {}
    skipped_dirs = {".git", ".pytest_cache", "__pycache__"}

    for path in sorted(root_path.rglob("*")):
        if not path.is_file():
            continue

        if any(part in skipped_dirs for part in path.parts):
            continue

        relative_parts = path.relative_to(root_path).parts
        _add_resource_alias(resource_index, path.name, path)
        _add_resource_alias(resource_index, path.relative_to(root_path).as_posix(), path)

        if len(relative_parts) >= 4 and relative_parts[0] == SETTINGS_PROGRAM_DIR:
            profile_name = relative_parts[1]
            resource_group = relative_parts[2]
            _add_resource_alias(resource_index, f"{profile_name}/{path.name}", path)
            _add_resource_alias(resource_index, f"{profile_name}/{resource_group}/{path.name}", path)

    return resource_index


def _pick_unique_path(candidates: set[Path] | None) -> Path | None:
    if not candidates or len(candidates) != 1:
        return None

    return next(iter(candidates))


def _pick_preferred_path(candidates: set[Path], preferred_root: Path | None = None) -> Path | None:
    if not candidates:
        return None
    if len(candidates) == 1:
        return next(iter(candidates))

    if preferred_root is not None:
        direct_children = [p for p in candidates if p.parent == preferred_root]
        if len(direct_children) == 1:
            return direct_children[0]
        if direct_children:
            return sorted(direct_children, key=lambda p: (len(p.parts), str(p)))[0]

    return sorted(candidates, key=lambda p: (len(p.parts), str(p)))[0]


def _find_best_suffix_match(resource_index: dict[str, set[Path]], normalized_value: str) -> tuple[str, set[Path]] | None:
    best_key: str | None = None
    best_paths: set[Path] | None = None
    for alias, paths in resource_index.items():
        if alias == normalized_value:
            continue
        if normalized_value.endswith(alias):
            if best_key is None or len(alias) > len(best_key):
                best_key = alias
                best_paths = paths
    if best_key is None:
        return None
    return best_key, best_paths  # type: ignore[return-value]


def _resolve_resource_path(resource_index: dict[str, set[Path]], value: str, preferred_root: Path | None = None) -> Path | None:
    normalized_value = _normalize_resource_key(value)
    if not normalized_value:
        return None

    exact_candidates = resource_index.get(normalized_value)
    if exact_candidates:
        resolved_exact = _pick_unique_path(exact_candidates)
        if resolved_exact is not None:
            logger.debug("resolve_resource_path: exact match %r -> %s", normalized_value, resolved_exact)
            return resolved_exact
        logger.debug(
            "resolve_resource_path: ambiguous exact match %r, candidates=%s",
            normalized_value,
            sorted(str(p) for p in exact_candidates),
        )
        preferred = _pick_preferred_path(exact_candidates, preferred_root)
        if preferred is not None:
            logger.debug("resolve_resource_path: preferred exact candidate %s", preferred)
            return preferred

    file_name = normalized_value.rsplit("/", 1)[-1]
    file_candidates = resource_index.get(file_name)
    if file_candidates:
        candidate = _pick_unique_path(file_candidates)
        if candidate is not None:
            logger.debug("resolve_resource_path: file-name fallback %r -> %s", file_name, candidate)
            return candidate
        logger.debug(
            "resolve_resource_path: ambiguous file-name fallback %r, candidates=%s",
            file_name,
            sorted(str(p) for p in file_candidates),
        )
        preferred = _pick_preferred_path(file_candidates, preferred_root)
        if preferred is not None:
            logger.debug("resolve_resource_path: preferred file-name candidate %s", preferred)
            return preferred

    suffix_match = _find_best_suffix_match(resource_index, normalized_value)
    if suffix_match is not None:
        suffix_key, suffix_candidates = suffix_match
        candidate = _pick_unique_path(suffix_candidates)
        if candidate is not None:
            logger.debug("resolve_resource_path: suffix match %r -> %s", suffix_key, candidate)
            return candidate
        preferred = _pick_preferred_path(suffix_candidates, preferred_root)
        if preferred is not None:
            logger.debug("resolve_resource_path: preferred suffix candidate %s", preferred)
            return preferred
        logger.debug(
            "resolve_resource_path: ambiguous suffix match %r, candidates=%s",
            suffix_key,
            sorted(str(p) for p in suffix_candidates),
        )

    logger.debug("resolve_resource_path: no fallback match for %r", file_name)
    return None


def _index_distributore_references(project: AT614Project) -> None:
    tests_by_stem = {path.stem: path for path in project.test_sequences}

    for distributore_path, distributore in project.distributori.items():
        for section_code in distributore.sezioni.values():
            target = tests_by_stem.get(section_code)
            if target is not None:
                project.add_reference(distributore_path, target)

        for ce16_reference in distributore.calibrazioni_ce16.values():
            target = _resolve_resource_path(project.resource_index, ce16_reference)
            if target is None:
                project.add_unresolved_reference(distributore_path, ce16_reference)
                continue

            if target != distributore_path:
                project.add_reference(distributore_path, target)


def _index_test_references(project: AT614Project) -> None:
    for test_path, sequence in project.test_sequences.items():
        for row in sequence.rows:
            for reference in extract_test_file_references(row):
                target = _resolve_resource_path(project.resource_index, reference.raw_value)
                if target is None:
                    project.add_unresolved_reference(test_path, reference.raw_value)
                    continue

                if target != test_path:
                    project.add_reference(test_path, target)



def load_project(root_path: Path, settings_ini_path: Path | None = None) -> AT614Project:
    # Leggi TUTTI i percorsi direttamente da settings.ini (fonte di verità).
    # Fallback su root_path/subdir solo se settings.ini non è disponibile.
    _ini: dict[str, Path] = {}
    if settings_ini_path is not None and settings_ini_path.exists():
        _ini = parse_settings_ini(settings_ini_path)

    distributore_dir      = _ini.get("FolderConfigurazioneBancoCollaudo") or root_path / "DISTRIBUTORE"
    test_dir              = _ini.get("FolderConfigurazioneTest")           or root_path / "TEST"
    settings_program_dir  = _ini.get("FolderSettaggiProgramma")            or root_path / SETTINGS_PROGRAM_DIR
    point_series_dirs: dict[str, Path] = {
        "CURVE_COMANDO": _ini.get("FolderFileCurveComando") or root_path / "CURVE_COMANDO",
        "CURVE_LIMITE":  _ini.get("FolderFileCurveLimite")  or root_path / "CURVE_LIMITE",
        "RAMPE_XY":      _ini.get("FolderRampeXY")          or root_path / "RAMPE_XY",
    }

    # Percorsi su share diversi: DEVONO venire da settings.ini, non ricostruibili da root_path
    _raw_moduli = _ini.get("FolderConfigurazioneModuli")
    folder_configurazione_moduli: Path | None = _raw_moduli if (_raw_moduli is not None and _raw_moduli.exists()) else None
    folder_graph_saved:           Path | None = _ini.get("FolderGraphSaved")
    folder_graph_saved_last_acq:  Path | None = _ini.get("FolderGraphSaved_LastACQ")

    distributori: dict[Path, DistributoreConfig] = {}
    test_sequences: dict[Path, TestSequence] = {}
    point_series_resources: dict[Path, PointSeriesResource] = {}
    ce16_resources: dict[Path, Ce16Resource] = {}
    mms2218_resources: dict[Path, Mms2218Resource] = {}

    if distributore_dir.exists():
        for path in sorted(distributore_dir.glob("*.cfg")):
            distributori[path] = parse_distributore(path)

    if test_dir.exists():
        for path in sorted(test_dir.glob("*.csv")):
            test_sequences[path] = parse_test_csv(path)

    for category_name, directory_path in point_series_dirs.items():
        if not directory_path.exists():
            continue
        for path in sorted(directory_path.glob("*.csv")):
            point_series_resources[path] = parse_point_series(path)

    if settings_program_dir.exists():
        for profile_dir in sorted(item for item in settings_program_dir.iterdir() if item.is_dir()):
            ce16_dir = profile_dir / "CE16"
            if ce16_dir.exists():
                for path in sorted(ce16_dir.glob("*.cfg")):
                    ce16_resources[path] = parse_ce16(path)

            mms2218_dir = profile_dir / "MMS2218"
            if mms2218_dir.exists():
                for path in sorted(mms2218_dir.glob("*.csv")):
                    mms2218_resources[path] = parse_mms2218(path)

    # Costruisce l'indice base dai file locali sotto root_path
    resource_index = _build_resource_file_index(root_path)

    # Estende l'indice con le risorse già caricate da percorsi di rete:
    # in questo modo resolve_resource_path() funziona anche per file su share.
    for _path in (
        list(distributori)
        + list(test_sequences)
        + list(point_series_resources)
        + list(ce16_resources)
        + list(mms2218_resources)
    ):
        _add_resource_alias(resource_index, _path.name, _path)
        _add_resource_alias(resource_index, _path.stem, _path)

    # Indicizza i file .txt di FolderConfigurazioneModuli (parametri SCRITTURA_PARAMETRI)
    if folder_configurazione_moduli is not None and folder_configurazione_moduli.exists():
        try:
            for _txt_path in sorted(folder_configurazione_moduli.rglob("*.txt")):
                _add_resource_alias(resource_index, _txt_path.name, _txt_path)
                _add_resource_alias(resource_index, _txt_path.stem, _txt_path)
        except OSError:
            pass  # percorso di rete temporaneamente irraggiungibile → nessun indice parziale

    project = AT614Project(
        root_path=root_path,
        distributori=distributori,
        test_sequences=test_sequences,
        point_series_resources=point_series_resources,
        ce16_resources=ce16_resources,
        mms2218_resources=mms2218_resources,
        resource_index=resource_index,
        folder_configurazione_moduli=folder_configurazione_moduli,
        folder_graph_saved=folder_graph_saved,
        folder_graph_saved_last_acq=folder_graph_saved_last_acq,
    )

    _index_distributore_references(project)
    _index_test_references(project)

    return project
