from __future__ import annotations

from pathlib import Path


def build_duplicate_path(source_path: Path) -> Path:
    for index in range(1, 10_000):
        suffix = " - copia" if index == 1 else f" - copia {index}"
        candidate = source_path.with_name(f"{source_path.stem}{suffix}{source_path.suffix}")
        if not candidate.exists():
            return candidate

    raise RuntimeError(f"Impossibile generare un nome duplicato per {source_path.name}")


def next_indexed_path(parent: Path, prefix: str, suffix: str) -> Path:
    used: set[int] = set()
    for sibling in parent.glob(f"{prefix}*{suffix}"):
        stem = sibling.stem
        if not stem.startswith(prefix):
            continue
        index_part = stem[len(prefix) :]
        if index_part.isdigit():
            used.add(int(index_part))

    next_index = max(used) + 1 if used else 1
    return parent / f"{prefix}{next_index}{suffix}"
