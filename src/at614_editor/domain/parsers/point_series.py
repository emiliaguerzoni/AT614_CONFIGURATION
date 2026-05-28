from __future__ import annotations

from pathlib import Path

from at614_editor.domain.models import PointSeriesResource, PointSeriesRow


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


def parse(path: Path) -> PointSeriesResource:
    raw = path.read_bytes()
    encoding = _detect_encoding(raw)
    text = raw.decode(encoding)
    line_ending = _detect_line_ending(text)
    logical_lines = text.splitlines()

    if not logical_lines:
        raise ValueError("Empty point series file")

    header_fields = logical_lines[0].split(";")
    rows = [PointSeriesRow(values=line.split(";")) for line in logical_lines[1:] if line.strip()]

    return PointSeriesResource(
        source_path=path,
        header_fields=header_fields,
        rows=rows,
        encoding=encoding,
        line_ending=line_ending,
        endswith_newline=text.endswith(("\n", "\r\n")),
    )


def serialize(resource: PointSeriesResource) -> bytes:
    lines = [";".join(resource.header_fields)]
    lines.extend(";".join(row.values) for row in resource.rows)

    text = resource.line_ending.join(lines)
    if resource.endswith_newline:
        text += resource.line_ending
    return text.encode(resource.encoding)
