from __future__ import annotations

import re
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path


@dataclass(frozen=True, slots=True)
class SchemaParameter:
    index: int
    name: str | None = None
    raw_name: str | None = None
    label: str | None = None
    description: str | None = None
    parameter_type: str | None = None
    resource_type: str | None = None
    editor: str | None = None
    hidden: bool = False


@dataclass(frozen=True, slots=True)
class TestSchemaDefinition:
    test_id: str
    display_name: str | None = None
    source_module: str | None = None
    parameters: tuple[SchemaParameter, ...] = ()


@dataclass(frozen=True, slots=True)
class _SchemaDocument:
    test_id: str
    display_name: str | None = None
    source_module: str | None = None
    extends: str | None = None
    replace_parameters: bool = False
    parameters: tuple[SchemaParameter, ...] = ()


_INTEGER_RE = re.compile(r"^-?\d+$")


def _find_schema_root() -> Path | None:
    candidates: list[Path] = []
    cwd = Path.cwd().resolve()
    source_dir = Path(__file__).resolve().parent
    candidates.extend([cwd, *cwd.parents])
    candidates.extend([source_dir, *source_dir.parents])

    seen: set[Path] = set()
    for candidate in candidates:
        if candidate in seen:
            continue
        seen.add(candidate)
        schema_root = candidate / "schemas"
        if (schema_root / "auto").is_dir():
            return schema_root
    return None


def _split_mapping_entry(line: str) -> tuple[str, str]:
    key, separator, value = line.partition(":")
    if separator == "":
        raise ValueError(f"Invalid schema line: {line!r}")
    return key.strip(), value.strip()


def _unescape_double_quoted(value: str) -> str:
    inner = value[1:-1]
    buffer: list[str] = []
    index = 0
    while index < len(inner):
        char = inner[index]
        if char != "\\" or index + 1 >= len(inner):
            buffer.append(char)
            index += 1
            continue

        escaped = inner[index + 1]
        replacements = {
            "\\": "\\",
            '"': '"',
            "n": "\n",
            "r": "\r",
            "t": "\t",
        }
        buffer.append(replacements.get(escaped, escaped))
        index += 2

    return "".join(buffer)


def _parse_scalar(raw_value: str) -> object:
    if raw_value == "":
        return ""
    if raw_value == "[]":
        return []
    if raw_value.startswith('"') and raw_value.endswith('"'):
        return _unescape_double_quoted(raw_value)
    if raw_value.startswith("'") and raw_value.endswith("'"):
        return raw_value[1:-1].replace("''", "'")

    lowered = raw_value.lower()
    if lowered == "true":
        return True
    if lowered == "false":
        return False
    if _INTEGER_RE.match(raw_value):
        return int(raw_value)
    return raw_value


def _coerce_string(value: object | None) -> str | None:
    if value is None:
        return None
    if isinstance(value, str):
        normalized = value.strip()
        return normalized or None
    return str(value)


def _coerce_bool(value: object | None) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() == "true"
    return False


def _coerce_parameter(entry: dict[str, object]) -> SchemaParameter:
    index_value = entry.get("index")
    if isinstance(index_value, int):
        index = index_value
    elif isinstance(index_value, str) and _INTEGER_RE.match(index_value):
        index = int(index_value)
    else:
        raise ValueError(f"Parameter entry without numeric index: {entry!r}")

    return SchemaParameter(
        index=index,
        name=_coerce_string(entry.get("name")),
        raw_name=_coerce_string(entry.get("raw_name")),
        label=_coerce_string(entry.get("label")),
        description=_coerce_string(entry.get("description")),
        parameter_type=_coerce_string(entry.get("type")),
        resource_type=_coerce_string(entry.get("resource_type")),
        editor=_coerce_string(entry.get("editor")),
        hidden=_coerce_bool(entry.get("hidden")),
    )


def _parse_schema_document(path: Path) -> _SchemaDocument:
    text = path.read_text(encoding="utf-8")
    root: dict[str, object] = {}
    parameters: list[dict[str, object]] = []
    in_parameters = False
    current_parameter: dict[str, object] | None = None

    for raw_line in text.splitlines():
        line = raw_line.rstrip()
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue

        indent = len(line) - len(line.lstrip(" "))
        if indent == 0:
            in_parameters = False
            current_parameter = None
            key, value = _split_mapping_entry(stripped)
            if key == "parameters" and value == "":
                root[key] = parameters
                in_parameters = True
                continue
            root[key] = _parse_scalar(value)
            continue

        if not in_parameters:
            continue

        if stripped.startswith("- "):
            current_parameter = {}
            parameters.append(current_parameter)
            remainder = stripped[2:].strip()
            if remainder:
                key, value = _split_mapping_entry(remainder)
                current_parameter[key] = _parse_scalar(value)
            continue

        if current_parameter is None:
            continue

        key, value = _split_mapping_entry(stripped)
        current_parameter[key] = _parse_scalar(value)

    test_id = _coerce_string(root.get("id"))
    if test_id is None:
        raise ValueError(f"Schema without id: {path}")

    parameter_entries = tuple(sorted((_coerce_parameter(entry) for entry in parameters), key=lambda item: item.index))

    return _SchemaDocument(
        test_id=test_id,
        display_name=_coerce_string(root.get("display_name")),
        source_module=_coerce_string(root.get("module")),
        extends=_coerce_string(root.get("extends")),
        replace_parameters=_coerce_bool(root.get("replace_parameters")),
        parameters=parameter_entries,
    )


def _load_schema_documents(directory: Path) -> dict[str, _SchemaDocument]:
    if not directory.is_dir():
        return {}

    documents: dict[str, _SchemaDocument] = {}
    for path in sorted(directory.glob("*.yaml")):
        document = _parse_schema_document(path)
        documents[document.test_id] = document
    return documents


def _merge_parameter(base: SchemaParameter | None, override: SchemaParameter) -> SchemaParameter:
    if base is None:
        return SchemaParameter(
            index=override.index,
            name=override.name,
            raw_name=override.raw_name,
            label=override.label,
            description=override.description,
            parameter_type=override.parameter_type,
            resource_type=override.resource_type,
            editor=override.editor,
            hidden=False,
        )

    return SchemaParameter(
        index=override.index,
        name=override.name or base.name,
        raw_name=override.raw_name or base.raw_name,
        label=override.label or base.label,
        description=override.description or base.description,
        parameter_type=override.parameter_type or base.parameter_type,
        resource_type=override.resource_type or base.resource_type,
        editor=override.editor or base.editor,
        hidden=False,
    )


def _merge_document(base: TestSchemaDefinition | None, override: _SchemaDocument) -> TestSchemaDefinition:
    parameter_map: dict[int, SchemaParameter]
    if base is not None and not override.replace_parameters:
        parameter_map = {parameter.index: parameter for parameter in base.parameters}
    else:
        parameter_map = {}

    for parameter in override.parameters:
        if parameter.hidden:
            parameter_map.pop(parameter.index, None)
            continue
        parameter_map[parameter.index] = _merge_parameter(parameter_map.get(parameter.index), parameter)

    return TestSchemaDefinition(
        test_id=override.test_id,
        display_name=override.display_name or (base.display_name if base is not None else None),
        source_module=(base.source_module if base is not None else None) or override.source_module,
        parameters=tuple(sorted(parameter_map.values(), key=lambda item: item.index)),
    )


@lru_cache(maxsize=1)
def get_test_schema_catalog() -> dict[str, TestSchemaDefinition]:
    schema_root = _find_schema_root()
    if schema_root is None:
        return {}

    auto_documents = _load_schema_documents(schema_root / "auto")
    catalog: dict[str, TestSchemaDefinition] = {
        document.test_id: TestSchemaDefinition(
            test_id=document.test_id,
            display_name=document.display_name,
            source_module=document.source_module,
            parameters=document.parameters,
        )
        for document in auto_documents.values()
    }

    override_documents = _load_schema_documents(schema_root / "overrides")
    for document in override_documents.values():
        base = catalog.get(document.extends or document.test_id)
        catalog[document.test_id] = _merge_document(base, document)

    return catalog