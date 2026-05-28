from __future__ import annotations

from pathlib import Path

from at614_editor.domain.models import TestRow, TestSequence


def _detect_encoding(raw: bytes) -> str:
    try:
        raw.decode("utf-8")
    except UnicodeDecodeError:
        return "cp1252"
    return "utf-8"


def _detect_line_ending(text: str) -> str:
    if "\r\n" in text:
        return "\r\n"
    return "\n"


def parse(path: Path) -> TestSequence:
    raw = path.read_bytes()
    encoding = _detect_encoding(raw)
    text = raw.decode(encoding)
    line_ending = _detect_line_ending(text)
    logical_lines = text.splitlines()

    if not logical_lines:
        raise ValueError("Empty TEST csv file")

    header_fields = logical_lines[0].split(";")
    rows: list[TestRow] = []

    for raw_line in logical_lines[1:]:
        if not raw_line.strip():
            continue

        fields = raw_line.split(";")
        if len(fields) < 3:
            raise ValueError(f"Invalid TEST row: {raw_line}")

        rows.append(
            TestRow(
                name=fields[0],
                test_id=fields[1],
                index_raw=fields[2],
                parameters=fields[3:],
            )
        )

    return TestSequence(
        source_path=path,
        header_fields=header_fields,
        rows=rows,
        encoding=encoding,
        line_ending=line_ending,
        endswith_newline=text.endswith(("\n", "\r\n")),
    )


def serialize(sequence: TestSequence) -> bytes:
    lines = [";".join(sequence.header_fields)]
    for row in sequence.rows:
        lines.append(
            ";".join([row.name, row.test_id, row.index_raw, *row.parameters])
        )

    text = sequence.line_ending.join(lines)
    if sequence.endswith_newline:
        text += sequence.line_ending
    return text.encode(sequence.encoding)
