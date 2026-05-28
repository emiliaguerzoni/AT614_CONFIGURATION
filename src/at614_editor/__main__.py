from __future__ import annotations

import argparse
import logging
import logging.handlers
from pathlib import Path

from PySide6.QtWidgets import QApplication

from at614_editor import __version__
from at614_editor.ui.main_window import MainWindow, discover_project_root


def normalize_project_root_arg(project_root: Path | str) -> Path:
    raw = str(project_root).strip()
    if raw.startswith('"') and raw.endswith('"'):
        raw = raw[1:-1]
    raw = raw.strip('"')
    return Path(raw)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="at614-editor",
        description="AT614 Configuration Editor",
    )
    parser.add_argument(
        "--version",
        action="store_true",
        help="mostra la versione ed esci",
    )
    parser.add_argument(
        "--project-root",
        type=Path,
        help="cartella del progetto AT614 da aprire",
    )
    parser.add_argument(
        "--smoke-ui",
        action="store_true",
        help=argparse.SUPPRESS,
    )
    return parser


def setup_logging() -> None:
    log_dir = Path.home() / ".at614-editor" / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "at614-editor.log"

    handler = logging.handlers.RotatingFileHandler(
        log_file,
        maxBytes=5 * 1024 * 1024,
        backupCount=3,
        encoding="utf-8",
    )
    formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    handler.setFormatter(formatter)

    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    root_logger.addHandler(handler)


def main(argv: list[str] | None = None) -> int:
    setup_logging()
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.version:
        print(f"AT614 Configuration Editor {__version__}")
        return 0

    if args.project_root:
        explicit_root = normalize_project_root_arg(args.project_root).resolve()
        discovered = discover_project_root(explicit_root)
        project_root = discovered if discovered is not None else explicit_root
    else:
        project_root = discover_project_root(Path.cwd())
    app = QApplication.instance() or QApplication([])
    window = MainWindow(project_root=project_root)
    window.show()

    if args.smoke_ui:
        app.processEvents()
        window.close()
        return 0

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
