from __future__ import annotations

import argparse
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Sequence


MODULE_PREFIX = "Module_TEST_"
MODULE_PATTERN = f"{MODULE_PREFIX}*.bas"

ENUM_BLOCK_RE = re.compile(
    r"Private\s+Enum\s+eTestParameter\b(.*?)End\s+Enum",
    re.DOTALL | re.IGNORECASE,
)

ENTRY_RE = re.compile(
    r"^(?P<name>[A-Za-z_][A-Za-z0-9_]*)"
    r"(?:\s*=\s*(?P<value>-?\d+))?"
    r"(?:\s*'\s*(?P<comment>.*?))?"
    r"\s*$"
)

SENTINEL_NAMES = {"emax", "enumeromassimodiparametri"}


@dataclass(slots=True)
class ParameterEntry:
    index: int
    name: str
    raw_name: str
    description: str | None


@dataclass(slots=True)
class TestSchema:
    test_id: str
    module: str
    parameters: list[ParameterEntry] = field(default_factory=list)


def _detect_encoding(raw: bytes) -> str:
    try:
        raw.decode("utf-8")
    except UnicodeDecodeError:
        return "cp1252"
    return "utf-8"


def _strip_parameter_prefix(raw_name: str) -> str:
    if len(raw_name) > 1 and raw_name[0] == "e" and raw_name[1].isupper():
        return raw_name[1:]
    return raw_name


def _parse_enum_lines(block_text: str) -> list[ParameterEntry]:
    entries: list[ParameterEntry] = []
    last_index = 0

    for raw_line in block_text.splitlines():
        stripped = raw_line.strip()
        if not stripped or stripped.startswith("'"):
            continue

        match = ENTRY_RE.match(stripped)
        if match is None:
            continue

        raw_name = match.group("name")
        if raw_name.lower() in SENTINEL_NAMES:
            continue

        explicit_value = match.group("value")
        if explicit_value is not None:
            current_index = int(explicit_value)
        else:
            current_index = last_index + 1
        last_index = current_index

        comment = match.group("comment")
        description = comment.strip() if comment else None
        if description == "":
            description = None

        entries.append(
            ParameterEntry(
                index=current_index,
                name=_strip_parameter_prefix(raw_name),
                raw_name=raw_name,
                description=description,
            )
        )

    return entries


def has_test_parameter_enum(path: Path) -> bool:
    raw = path.read_bytes()
    text = raw.decode(_detect_encoding(raw))
    return ENUM_BLOCK_RE.search(text) is not None


def _normalize_test_id(raw_test_id: str, sibling_test_ids: set[str] | None = None) -> str:
    if sibling_test_ids is None:
        return raw_test_id

    if "_" in raw_test_id:
        candidate = raw_test_id.rpartition("_")[0]
        if candidate in sibling_test_ids:
            return candidate

    if raw_test_id in sibling_test_ids:
        return raw_test_id

    return raw_test_id


def extract_schema(path: Path, sibling_test_ids: set[str] | None = None) -> TestSchema:
    stem = path.stem
    if not stem.startswith(MODULE_PREFIX):
        raise ValueError(f"Module name does not start with {MODULE_PREFIX!r}: {path.name}")

    raw_test_id = stem[len(MODULE_PREFIX) :]
    test_id = _normalize_test_id(raw_test_id, sibling_test_ids)
    raw = path.read_bytes()
    encoding = _detect_encoding(raw)
    text = raw.decode(encoding)

    match = ENUM_BLOCK_RE.search(text)
    if match is None:
        return TestSchema(test_id=test_id, module=path.name, parameters=[])

    parameters = _parse_enum_lines(match.group(1))
    return TestSchema(test_id=test_id, module=path.name, parameters=parameters)


def _yaml_scalar(value: str) -> str:
    if value == "":
        return '""'

    forbidden_chars = set(":#\"'[]{},\n\r&*!|>%@`")
    forbidden_starts = ("-", "?", " ", "\t")
    needs_quotes = (
        any(ch in forbidden_chars for ch in value)
        or value.startswith(forbidden_starts)
        or value.endswith((" ", "\t"))
        or value.strip().lower() in {"true", "false", "null", "yes", "no", "on", "off", "~"}
    )

    if needs_quotes:
        escaped = value.replace("\\", "\\\\").replace('"', '\\"')
        return f'"{escaped}"'
    return value


def dump_schema_yaml(schema: TestSchema) -> str:
    lines: list[str] = [
        f"id: {_yaml_scalar(schema.test_id)}",
        f"module: {_yaml_scalar(schema.module)}",
        "parameters:",
    ]

    if not schema.parameters:
        lines.append("  []")
        lines.append("")
        return "\n".join(lines)

    for entry in schema.parameters:
        lines.append(f"  - index: {entry.index}")
        lines.append(f"    name: {_yaml_scalar(entry.name)}")
        lines.append(f"    raw_name: {_yaml_scalar(entry.raw_name)}")
        if entry.description is not None:
            lines.append(f"    description: {_yaml_scalar(entry.description)}")
        lines.append(f"    type: string")

    lines.append("")
    return "\n".join(lines)


def generate_schemas(bas_dir: Path, output_dir: Path) -> list[Path]:
    if not bas_dir.is_dir():
        raise ValueError(f"BAS directory not found: {bas_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    generated: list[Path] = []

    sibling_test_ids = {
        path.stem[len(MODULE_PREFIX) :]
        for path in bas_dir.glob(MODULE_PATTERN)
        if path.stem.startswith(MODULE_PREFIX)
    }

    for path in sorted(bas_dir.glob(MODULE_PATTERN)):
        if not has_test_parameter_enum(path):
            continue
        schema = extract_schema(path, sibling_test_ids=sibling_test_ids)
        output_path = output_dir / f"{schema.test_id}.yaml"
        output_path.write_text(dump_schema_yaml(schema), encoding="utf-8")
        generated.append(output_path)

    return generated


def _main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="at614-bas-extractor",
        description=(
            "Estrae lo schema parametri dei moduli Module_TEST_*.bas in file "
            "YAML usati come catalogo autoritativo dello schema dei test."
        ),
    )
    parser.add_argument(
        "bas_dir",
        type=Path,
        help="cartella contenente i file Module_TEST_*.bas",
    )
    parser.add_argument(
        "output_dir",
        type=Path,
        nargs="?",
        default=Path("schemas/auto"),
        help="cartella di destinazione dei file YAML (default: schemas/auto)",
    )
    args = parser.parse_args(argv)

    generated = generate_schemas(args.bas_dir, args.output_dir)
    for path in generated:
        print(path)
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
