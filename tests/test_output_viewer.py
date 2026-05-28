from pathlib import Path

from at614_editor.ui.editors.output_viewer import OutputViewer


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
    return archive_root


def test_output_viewer_loads_archive_and_populates_results(tmp_path: Path, qt_app) -> None:
    archive_root = _make_archive(tmp_path)
    viewer = OutputViewer(initial_archive_path=archive_root)

    assert viewer.archive is not None
    assert viewer.results_table.rowCount() == 2
    assert viewer.refresh_button.isEnabled()


def test_output_viewer_filter_by_text(tmp_path: Path, qt_app) -> None:
    archive_root = _make_archive(tmp_path)
    viewer = OutputViewer(initial_archive_path=archive_root)

    viewer.search_field.setText("run_002")

    assert viewer.results_table.rowCount() == 1


def test_output_viewer_does_not_allow_save_or_duplicate() -> None:
    assert OutputViewer.supports_save is False
    assert OutputViewer.supports_duplicate is False


def test_output_viewer_select_row_populates_axes(tmp_path: Path, qt_app) -> None:
    archive_root = _make_archive(tmp_path)
    viewer = OutputViewer(initial_archive_path=archive_root)

    viewer.results_table.selectRow(0)

    assert viewer.x_axis_combo.count() >= 2
    assert viewer.y_axes_list.count() >= 2


def test_output_viewer_actions_disabled_until_row_selected(tmp_path: Path, qt_app) -> None:
    archive_root = _make_archive(tmp_path)
    viewer = OutputViewer(initial_archive_path=archive_root)

    for button in viewer.chart_stage.toolbar_buttons:
        assert button.isEnabled() is False

    viewer.results_table.selectRow(0)

    for button in viewer.chart_stage.toolbar_buttons:
        assert button.isEnabled() is True


def test_output_viewer_export_writes_target_file(tmp_path: Path, qt_app, monkeypatch) -> None:
    archive_root = _make_archive(tmp_path)
    viewer = OutputViewer(initial_archive_path=archive_root)
    viewer.results_table.selectRow(0)

    target = tmp_path / "esportato.csv"
    monkeypatch.setattr(
        "at614_editor.ui.editors.output_viewer.QFileDialog.getSaveFileName",
        lambda *args, **kwargs: (str(target), ""),
    )
    viewer._on_export()

    assert target.exists()
    assert "tempo;tensione" in target.read_text(encoding="utf-8")


def test_output_viewer_handles_missing_archive_path(qt_app) -> None:
    viewer = OutputViewer(initial_archive_path=Path("/path/inesistente"))

    assert viewer.archive is None
    assert viewer.results_table.rowCount() == 0
    assert viewer.refresh_button.isEnabled() is False
