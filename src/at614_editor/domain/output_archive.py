from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path


CSV_PATTERN = "*.csv"
CACHE_FILE_NAME = ".at614_output_index.json"


@dataclass(slots=True)
class OutputFile:
    relative_path: str
    folder: str
    name: str
    created_at: float
    modified_at: float
    size_bytes: int


@dataclass(slots=True)
class OutputArchive:
    root_path: Path
    files: list[OutputFile] = field(default_factory=list)
    indexed_at: float = 0.0

    def filter(
        self,
        text: str | None = None,
        created_from: datetime | None = None,
        created_to: datetime | None = None,
        modified_from: datetime | None = None,
        modified_to: datetime | None = None,
    ) -> list[OutputFile]:
        normalized_text = (text or "").strip().lower()
        results: list[OutputFile] = []
        for entry in self.files:
            if normalized_text and normalized_text not in entry.name.lower() and normalized_text not in entry.folder.lower():
                continue
            if created_from is not None and entry.created_at < created_from.timestamp():
                continue
            if created_to is not None and entry.created_at > created_to.timestamp():
                continue
            if modified_from is not None and entry.modified_at < modified_from.timestamp():
                continue
            if modified_to is not None and entry.modified_at > modified_to.timestamp():
                continue
            results.append(entry)
        return results


def _scan_csv_files(root_path: Path) -> list[OutputFile]:
    entries: list[OutputFile] = []
    for path in sorted(root_path.rglob(CSV_PATTERN)):
        if not path.is_file():
            continue
        if any(part.startswith(".") for part in path.relative_to(root_path).parts):
            continue
        stat = path.stat()
        rel = path.relative_to(root_path)
        entries.append(
            OutputFile(
                relative_path=rel.as_posix(),
                folder=rel.parent.as_posix() if rel.parent != Path(".") else "",
                name=path.name,
                created_at=stat.st_ctime,
                modified_at=stat.st_mtime,
                size_bytes=stat.st_size,
            )
        )
    return entries


def scan_archive_metadata(root_path: Path) -> OutputArchive:
    if not root_path.exists() or not root_path.is_dir():
        raise ValueError(f"Archivio output non trovato: {root_path}")

    files = _scan_csv_files(root_path)
    return OutputArchive(
        root_path=root_path,
        files=files,
        indexed_at=datetime.now(tz=timezone.utc).timestamp(),
    )


def load_cached_index(root_path: Path) -> OutputArchive | None:
    cache_path = root_path / CACHE_FILE_NAME
    if not cache_path.exists():
        return None
    try:
        payload = json.loads(cache_path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None

    files = [OutputFile(**item) for item in payload.get("files", [])]
    return OutputArchive(
        root_path=root_path,
        files=files,
        indexed_at=payload.get("indexed_at", 0.0),
    )


def save_cached_index(archive: OutputArchive) -> Path:
    cache_path = archive.root_path / CACHE_FILE_NAME
    payload = {
        "indexed_at": archive.indexed_at,
        "files": [asdict(item) for item in archive.files],
    }
    cache_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return cache_path


def is_cache_stale(archive: OutputArchive) -> bool:
    """Verifica se il cache è obsoleto con una sola scansione invece di due."""
    if archive.indexed_at <= 0:
        return True

    cached_paths = {item.relative_path for item in archive.files}
    current_paths: set[str] = set()

    for path in archive.root_path.rglob(CSV_PATTERN):
        if not path.is_file():
            continue
        rel = path.relative_to(archive.root_path)
        if any(part.startswith(".") for part in rel.parts):
            continue
        if path.stat().st_mtime > archive.indexed_at:
            return True
        current_paths.add(rel.as_posix())

    return cached_paths != current_paths


def ensure_archive_index(root_path: Path, use_cache: bool = True) -> OutputArchive:
    if use_cache:
        cached = load_cached_index(root_path)
        if cached is not None and not is_cache_stale(cached):
            return cached

    archive = scan_archive_metadata(root_path)
    save_cached_index(archive)
    return archive


def read_csv_columns(path: Path, separator: str = ";", max_rows: int = 200_000) -> tuple[list[str], list[list[str]]]:
    if not path.is_file():
        raise ValueError(f"File output non trovato: {path}")

    raw = path.read_bytes()
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError:
        text = raw.decode("cp1252")

    lines = text.splitlines()
    if not lines:
        return [], []

    headers = lines[0].split(separator)
    rows: list[list[str]] = []
    for raw_line in lines[1 : max_rows + 1]:
        if not raw_line.strip():
            continue
        rows.append(raw_line.split(separator))
    return headers, rows
