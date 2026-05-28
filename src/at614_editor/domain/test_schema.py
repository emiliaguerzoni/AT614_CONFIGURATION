from __future__ import annotations

import re
from dataclasses import dataclass
from functools import lru_cache

from at614_editor.domain.models import TestFileReference, TestRow
from at614_editor.domain.schemas import get_test_schema_catalog


FILE_REFERENCE_SUFFIXES = (".csv", ".cfg", ".txt")


@dataclass(frozen=True, slots=True)
class TestParameterDescriptor:
    parameter_index: int
    label: str
    editor_kind: str
    resource_type: str | None = None


SCHEMA_PARAMETER_DESCRIPTORS: dict[str, dict[int, TestParameterDescriptor]] = {
    # Runtime catalog is loaded from schemas/auto + schemas/overrides.
}


_CAMEL_CASE_BOUNDARY_RE = re.compile(r"(?<=[a-z0-9])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])")


def _humanize_identifier(value: str) -> str:
    normalized = _CAMEL_CASE_BOUNDARY_RE.sub(" ", value.replace("_", " ").replace("-", " "))
    normalized = " ".join(part for part in normalized.split() if part)
    if not normalized:
        return value
    return normalized[0].upper() + normalized[1:].lower()


def _schema_exists(test_id: str) -> bool:
    return test_id in get_test_schema_catalog()


@lru_cache(maxsize=None)
def _get_schema_descriptors(test_id: str) -> dict[int, TestParameterDescriptor]:
    schema = get_test_schema_catalog().get(test_id)
    if schema is None:
        return {}

    descriptors: dict[int, TestParameterDescriptor] = {}
    for parameter in schema.parameters:
        label = parameter.label or _humanize_identifier(parameter.name or parameter.raw_name or f"Parametro {parameter.index}")
        descriptors[parameter.index] = TestParameterDescriptor(
            parameter_index=parameter.index,
            label=label,
            editor_kind="file_ref" if parameter.parameter_type == "file_ref" else "text",
            resource_type=parameter.resource_type,
        )

    return descriptors


def has_test_schema(test_id: str) -> bool:
    return _schema_exists(test_id)


def get_schema_parameter_count(test_id: str) -> int:
    descriptors = _get_schema_descriptors(test_id)
    return max(descriptors, default=0)


def describe_test_parameters(row: TestRow) -> list[TestParameterDescriptor]:
    schema_descriptors = _get_schema_descriptors(row.test_id)
    descriptors: list[TestParameterDescriptor] = []
    highest_schema_index = max(schema_descriptors, default=0)
    highest_filled_index = max(
        (parameter_index for parameter_index, value in enumerate(row.parameters, start=1) if value.strip()),
        default=0,
    )
    visible_parameter_count = max(highest_schema_index, highest_filled_index)

    for parameter_index in range(1, visible_parameter_count + 1):
        descriptor = schema_descriptors.get(parameter_index)
        if descriptor is not None:
            descriptors.append(descriptor)
            continue

        descriptors.append(
            TestParameterDescriptor(
                parameter_index=parameter_index,
                label=f"Parametro {parameter_index}",
                editor_kind="text",
            )
        )

    return descriptors


def _looks_like_file_reference(value: str) -> bool:
    normalized = value.strip()
    if not normalized:
        return False
    return normalized.lower().endswith(FILE_REFERENCE_SUFFIXES)


def extract_test_file_references(row: TestRow) -> list[TestFileReference]:
    references: list[TestFileReference] = []
    schema_descriptors = _get_schema_descriptors(row.test_id)

    if _schema_exists(row.test_id):
        for parameter_index, descriptor in schema_descriptors.items():
            if descriptor.editor_kind != "file_ref":
                continue

            parameter_offset = parameter_index - 1
            if parameter_offset >= len(row.parameters):
                continue

            raw_value = row.parameters[parameter_offset].strip()
            if not raw_value:
                continue

            references.append(
                TestFileReference(
                    test_id=row.test_id,
                    parameter_index=parameter_index,
                    resource_type=descriptor.resource_type or "generic_file",
                    raw_value=raw_value,
                )
            )

        return references

    for parameter_offset, raw_value in enumerate(row.parameters, start=1):
        normalized = raw_value.strip()
        if not _looks_like_file_reference(normalized):
            continue

        references.append(
            TestFileReference(
                test_id=row.test_id,
                parameter_index=parameter_offset,
                resource_type="generic_file",
                raw_value=normalized,
            )
        )

    return references
