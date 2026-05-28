from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal

from at614_editor.domain.models import DistributoreConfig, TestRow, TestSequence
from at614_editor.domain.parsers.distributore import serialize as serialize_distributore
from at614_editor.domain.parsers.test_csv import serialize as serialize_test_csv
from at614_editor.domain.project import AT614Project
from at614_editor.domain.test_schema import extract_test_file_references


CloneMode = Literal["duplica", "riusa"]


@dataclass(slots=True)
class CloneAction:
    source_path: Path
    target_name: str
    mode: CloneMode = "duplica"
    kind: str = "resource"

    @property
    def target_path(self) -> Path:
        return self.source_path.parent / self.target_name


@dataclass(slots=True)
class CloneProgramPlan:
    source_distributore: Path
    new_distributore_path: Path
    test_actions: list[CloneAction] = field(default_factory=list)
    resource_actions: list[CloneAction] = field(default_factory=list)
    external_actions: list[CloneAction] = field(default_factory=list)

    @property
    def all_actions(self) -> list[CloneAction]:
        return [*self.test_actions, *self.resource_actions, *self.external_actions]


def derive_target_name(original_name: str, new_code: str, source_code: str | None) -> str:
    stem, dot, suffix = original_name.rpartition(".")
    if not dot:
        stem, suffix = original_name, ""
    else:
        suffix = f".{suffix}"

    if source_code and stem == source_code:
        return f"{new_code}{suffix}"

    if new_code in stem:
        return f"{stem}{suffix}"

    return f"{stem}_{new_code}{suffix}"


def build_clone_plan(
    project: AT614Project,
    source_distributore: Path,
    new_code: str,
) -> CloneProgramPlan:
    if source_distributore not in project.distributori:
        raise ValueError(f"Distributore sorgente non presente nel progetto: {source_distributore}")

    new_code = new_code.strip()
    if not new_code:
        raise ValueError("Il codice del nuovo distributore non può essere vuoto")

    source_code = source_distributore.stem
    new_distributore_path = source_distributore.with_name(f"{new_code}{source_distributore.suffix}")

    plan = CloneProgramPlan(
        source_distributore=source_distributore,
        new_distributore_path=new_distributore_path,
    )

    referenced_tests = {
        target
        for target in project.get_uses(source_distributore)
        if target in project.test_sequences
    }

    referenced_resources: set[Path] = set()
    referenced_externals: set[Path] = set()

    for test_path in sorted(referenced_tests):
        plan.test_actions.append(
            CloneAction(
                source_path=test_path,
                target_name=derive_target_name(test_path.name, new_code, source_code),
                kind="sequenza_test",
            )
        )
        for target in project.get_uses(test_path):
            if target in project.point_series_resources:
                referenced_resources.add(target)
            elif target.suffix.lower() == ".txt":
                referenced_externals.add(target)

    for resource_path in sorted(referenced_resources):
        plan.resource_actions.append(
            CloneAction(
                source_path=resource_path,
                target_name=derive_target_name(resource_path.name, new_code, source_code),
                kind=_resource_kind(resource_path),
            )
        )

    for external_path in sorted(referenced_externals):
        plan.external_actions.append(
            CloneAction(
                source_path=external_path,
                target_name=derive_target_name(external_path.name, new_code, source_code),
                kind="file_esterno",
            )
        )

    return plan


def _resource_kind(path: Path) -> str:
    parent = path.parent.name
    return {
        "CURVE_COMANDO": "curva_comando",
        "CURVE_LIMITE": "curva_limite",
        "RAMPE_XY": "rampa_xy",
    }.get(parent, "resource")


def validate_plan(plan: CloneProgramPlan) -> list[str]:
    errors: list[str] = []

    if plan.new_distributore_path.exists():
        errors.append(f"Il file {plan.new_distributore_path.name} esiste già")

    seen_targets: set[Path] = set()
    for action in plan.all_actions:
        if action.mode != "duplica":
            continue
        target = action.target_path
        if target in seen_targets:
            errors.append(f"Collisione: due risorse duplicano in {target.name}")
        seen_targets.add(target)
        if target.exists():
            errors.append(f"Il file {target.name} esiste già")

    return errors


def _rename_map_from_plan(plan: CloneProgramPlan) -> dict[str, str]:
    rename_map: dict[str, str] = {}
    for action in (*plan.test_actions, *plan.resource_actions, *plan.external_actions):
        if action.mode != "duplica":
            continue
        rename_map[action.source_path.name] = action.target_name
    return rename_map


def _apply_rename_to_distributore(
    config: DistributoreConfig, plan: CloneProgramPlan
) -> DistributoreConfig:
    sezioni_rename = {
        action.source_path.stem: action.target_path.stem
        for action in plan.test_actions
        if action.mode == "duplica"
    }
    new_sezioni = {
        index: sezioni_rename.get(code, code) for index, code in config.sezioni.items()
    }
    return DistributoreConfig(
        source_path=plan.new_distributore_path,
        sezioni=new_sezioni,
        calibrazioni_ce16=dict(config.calibrazioni_ce16),
        extras=dict(config.extras),
        lines=list(config.lines),
        encoding=config.encoding,
        line_ending=config.line_ending,
        endswith_newline=config.endswith_newline,
    )


def _apply_rename_to_sequence(
    sequence: TestSequence, source_path: Path, target_path: Path, rename_map: dict[str, str]
) -> TestSequence:
    new_rows: list[TestRow] = []
    for row in sequence.rows:
        new_params = list(row.parameters)
        references = extract_test_file_references(row)
        for reference in references:
            offset = reference.parameter_index - 1
            if offset >= len(new_params):
                continue
            renamed = rename_map.get(new_params[offset].strip())
            if renamed is not None:
                new_params[offset] = renamed
        new_rows.append(
            TestRow(
                name=row.name,
                test_id=row.test_id,
                index_raw=row.index_raw,
                parameters=new_params,
            )
        )

    return TestSequence(
        source_path=target_path,
        header_fields=list(sequence.header_fields),
        rows=new_rows,
        encoding=sequence.encoding,
        line_ending=sequence.line_ending,
        endswith_newline=sequence.endswith_newline,
    )


def execute_clone_plan(project: AT614Project, plan: CloneProgramPlan) -> list[Path]:
    errors = validate_plan(plan)
    if errors:
        raise ValueError("; ".join(errors))

    created_files: list[Path] = []
    rename_map = _rename_map_from_plan(plan)

    for action in plan.resource_actions:
        if action.mode != "duplica":
            continue
        action.target_path.write_bytes(action.source_path.read_bytes())
        created_files.append(action.target_path)

    for action in plan.external_actions:
        if action.mode != "duplica":
            continue
        action.target_path.write_bytes(action.source_path.read_bytes())
        created_files.append(action.target_path)

    for action in plan.test_actions:
        if action.mode != "duplica":
            continue
        sequence = project.test_sequences[action.source_path]
        new_sequence = _apply_rename_to_sequence(
            sequence, action.source_path, action.target_path, rename_map
        )
        action.target_path.write_bytes(serialize_test_csv(new_sequence))
        created_files.append(action.target_path)

    source_config = project.distributori[plan.source_distributore]
    new_config = _apply_rename_to_distributore(source_config, plan)
    plan.new_distributore_path.write_bytes(serialize_distributore(new_config))
    created_files.append(plan.new_distributore_path)

    return created_files
