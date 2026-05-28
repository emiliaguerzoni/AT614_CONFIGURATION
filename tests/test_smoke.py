from pathlib import Path

from at614_editor import __version__
from at614_editor.__main__ import main


FIXTURES_ROOT = Path(__file__).parent / "fixtures"


def test_main_smoke_ui() -> None:
    result = main(["--smoke-ui", "--project-root", str(FIXTURES_ROOT)])

    assert result == 0


def test_main_version(capsys) -> None:
    result = main(["--version"])

    captured = capsys.readouterr()

    assert result == 0
    assert captured.out.strip() == f"AT614 Configuration Editor {__version__}"

