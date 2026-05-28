from __future__ import annotations

from pathlib import Path

from at614_editor.__main__ import normalize_project_root_arg
from at614_editor.ui.main_window import discover_project_root


def test_normalize_project_root_arg_strips_surrounding_quotes() -> None:
    raw = '"C:\\PROGETTI\\SOFTWARE\\VB6\\AT614\\Rel x.x.x.x"'
    normalized = normalize_project_root_arg(raw)
    assert normalized == Path(r'C:\PROGETTI\SOFTWARE\VB6\AT614\Rel x.x.x.x')


def test_discover_project_root_invalid_path_returns_none() -> None:
    invalid_path = Path(r'C:\PROGETTI\SOFTWARE\VB6\AT614\Rel x.x.x.x"')
    assert discover_project_root(invalid_path) is None
