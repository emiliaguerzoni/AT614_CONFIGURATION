from datetime import datetime, timedelta
from pathlib import Path

import pytest

from at614_editor.domain.output_archive import (
    CACHE_FILE_NAME,
    ensure_archive_index,
    is_cache_stale,
    load_cached_index,
    read_csv_columns,
    save_cached_index,
    scan_archive_metadata,
)


def _make_archive(tmp_path: Path) -> Path:
    archive_root = tmp_path / "archive"
    archive_root.mkdir()
    (archive_root / "prodotto_A").mkdir()
    (archive_root / "prodotto_A" / "run_001.csv").write_text(
        "tempo;tensione;pressione\n0;100;5\n1;200;7\n", encoding="utf-8"
    )
    (archive_root / "prodotto_A" / "run_002.csv").write_text(
        "tempo;tensione\n0;100\n", encoding="utf-8"
    )
    (archive_root / "prodotto_B").mkdir()
    (archive_root / "prodotto_B" / "test_001.csv").write_text(
        "x;y\n1;2\n", encoding="utf-8"
    )
    return archive_root


def test_scan_archive_collects_csv_files_only(tmp_path: Path) -> None:
    archive_root = _make_archive(tmp_path)
    (archive_root / "prodotto_A" / "note.txt").write_text("ignored", encoding="utf-8")

    archive = scan_archive_metadata(archive_root)

    names = {entry.name for entry in archive.files}
    assert names == {"run_001.csv", "run_002.csv", "test_001.csv"}
    folders = {entry.folder for entry in archive.files}
    assert folders == {"prodotto_A", "prodotto_B"}


def test_scan_archive_rejects_missing_root(tmp_path: Path) -> None:
    with pytest.raises(ValueError):
        scan_archive_metadata(tmp_path / "non_esiste")


def test_filter_archive_by_text(tmp_path: Path) -> None:
    archive_root = _make_archive(tmp_path)
    archive = scan_archive_metadata(archive_root)

    results = archive.filter(text="run")
    assert {entry.name for entry in results} == {"run_001.csv", "run_002.csv"}

    folder_results = archive.filter(text="prodotto_b")
    assert {entry.name for entry in folder_results} == {"test_001.csv"}


def test_filter_archive_by_date(tmp_path: Path) -> None:
    archive_root = _make_archive(tmp_path)
    archive = scan_archive_metadata(archive_root)

    past = datetime.now() - timedelta(days=365)
    results_past = archive.filter(created_from=past)
    assert len(results_past) == 3

    future = datetime.now() + timedelta(days=1)
    results_future = archive.filter(created_from=future)
    assert results_future == []


def test_cache_round_trip(tmp_path: Path) -> None:
    archive_root = _make_archive(tmp_path)
    archive = scan_archive_metadata(archive_root)

    cache_path = save_cached_index(archive)
    assert cache_path.name == CACHE_FILE_NAME

    loaded = load_cached_index(archive_root)
    assert loaded is not None
    assert {entry.name for entry in loaded.files} == {"run_001.csv", "run_002.csv", "test_001.csv"}


def test_ensure_archive_uses_cache_when_fresh(tmp_path: Path) -> None:
    archive_root = _make_archive(tmp_path)
    first = ensure_archive_index(archive_root)
    indexed_at_first = first.indexed_at

    second = ensure_archive_index(archive_root)
    assert second.indexed_at == indexed_at_first


def test_is_cache_stale_when_new_file_added(tmp_path: Path) -> None:
    archive_root = _make_archive(tmp_path)
    archive = scan_archive_metadata(archive_root)
    save_cached_index(archive)

    (archive_root / "prodotto_A" / "run_999.csv").write_text("a;b\n", encoding="utf-8")

    cached = load_cached_index(archive_root)
    assert cached is not None
    assert is_cache_stale(cached) is True


def test_read_csv_columns_returns_header_and_rows(tmp_path: Path) -> None:
    archive_root = _make_archive(tmp_path)
    target = archive_root / "prodotto_A" / "run_001.csv"

    headers, rows = read_csv_columns(target)

    assert headers == ["tempo", "tensione", "pressione"]
    assert rows == [["0", "100", "5"], ["1", "200", "7"]]


def test_cache_file_is_skipped_in_scan(tmp_path: Path) -> None:
    archive_root = _make_archive(tmp_path)
    (archive_root / CACHE_FILE_NAME).write_text("{}", encoding="utf-8")

    archive = scan_archive_metadata(archive_root)
    names = {entry.name for entry in archive.files}
    assert CACHE_FILE_NAME not in names
