from __future__ import annotations

import argparse
import logging
import logging.handlers
import sys
import traceback
from pathlib import Path

from PySide6.QtCore import qInstallMessageHandler, QtMsgType
from PySide6.QtGui import QColor, QPalette
from PySide6.QtWidgets import QApplication

from at614_editor import __version__
from at614_editor.domain.preferences import load_settings_ini_path
from at614_editor.domain.settings_ini import discover_root_from_settings_ini
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


def _qt_message_handler(msg_type: QtMsgType, context, message: str) -> None:
    qt_logger = logging.getLogger("Qt")
    level_map = {
        QtMsgType.QtDebugMsg: logging.DEBUG,
        QtMsgType.QtInfoMsg: logging.INFO,
        QtMsgType.QtWarningMsg: logging.WARNING,
        QtMsgType.QtCriticalMsg: logging.ERROR,
        QtMsgType.QtFatalMsg: logging.CRITICAL,
    }
    level = level_map.get(msg_type, logging.WARNING)
    qt_logger.log(level, message)


def _unhandled_exception_hook(exc_type, exc_value, exc_tb) -> None:
    root_logger = logging.getLogger()
    root_logger.critical(
        "ECCEZIONE NON GESTITA:\n" + "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
    )


def setup_logging() -> None:
    log_dir = Path.home() / ".at614-editor" / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "at614-editor.log"

    formatter = logging.Formatter(
        "%(asctime)s | %(name)s | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = logging.handlers.RotatingFileHandler(
        log_file,
        maxBytes=5 * 1024 * 1024,
        backupCount=3,
        encoding="utf-8",
    )
    file_handler.setFormatter(formatter)
    file_handler.setLevel(logging.DEBUG)

    # stdout per vedere i log nel terminale
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    console_handler.setLevel(logging.DEBUG)

    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)
    root_logger.addHandler(file_handler)
    root_logger.addHandler(console_handler)

    # Cattura eccezioni Python non gestite
    sys.excepthook = _unhandled_exception_hook

    # Reindirizza stderr al logger (cattura print di eccezioni PySide6)
    class _StderrToLogger:
        def write(self, message: str) -> None:
            stripped = message.strip()
            if stripped:
                logging.getLogger("stderr").error(stripped)

        def flush(self) -> None:
            pass

    sys.stderr = _StderrToLogger()  # type: ignore[assignment]

    # Cattura messaggi Qt (warning, critical, etc.)
    qInstallMessageHandler(_qt_message_handler)

    root_logger.info("=" * 80)
    root_logger.info("AT614 Configuration Editor - Avvio applicazione")
    root_logger.info("=" * 80)
    root_logger.info(f"Log file: {log_file}")


def apply_light_palette(app: QApplication) -> None:
    app.setStyle("Fusion")
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor("#F5F7FA"))
    palette.setColor(QPalette.WindowText, QColor("#101828"))
    palette.setColor(QPalette.Base, QColor("#FFFFFF"))
    palette.setColor(QPalette.AlternateBase, QColor("#F0F3F8"))
    palette.setColor(QPalette.ToolTipBase, QColor("#FFFFFF"))
    palette.setColor(QPalette.ToolTipText, QColor("#101828"))
    palette.setColor(QPalette.Text, QColor("#101828"))
    palette.setColor(QPalette.Button, QColor("#FFFFFF"))
    palette.setColor(QPalette.ButtonText, QColor("#101828"))
    palette.setColor(QPalette.Link, QColor("#2563EB"))
    palette.setColor(QPalette.Highlight, QColor("#2563EB"))
    palette.setColor(QPalette.HighlightedText, QColor("#FFFFFF"))
    app.setPalette(palette)
    app.setStyleSheet(
        "QToolTip { color: #101828; background-color: #FFFFFF; border: 1px solid #D0D5DD; }"
    )


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
        saved_ini = None
    else:
        # 1) Prova dal settings.ini salvato nelle preferenze
        saved_ini = load_settings_ini_path()
        if saved_ini is not None:
            project_root = discover_root_from_settings_ini(saved_ini)
            if project_root is None:
                logger.warning(f"Preferences: settings.ini salvato non ha prodotto una root valida: {saved_ini}")
                project_root = discover_project_root(Path.cwd())
                saved_ini = None
        else:
            # 2) Fallback: ricerca automatica dalla cwd
            project_root = discover_project_root(Path.cwd())
    app = QApplication.instance() or QApplication([])
    apply_light_palette(app)
    window = MainWindow(project_root=project_root, settings_ini_path=saved_ini)
    window.show()

    if args.smoke_ui:
        app.processEvents()
        window.close()
        return 0

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
