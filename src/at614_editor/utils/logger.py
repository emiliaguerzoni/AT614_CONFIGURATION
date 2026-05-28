"""Sistema centralizzato di logging per il debuggin dell'applicazione AT614."""
from __future__ import annotations

import logging
import sys
from pathlib import Path


def setup_logging(name: str = "at614_editor", level: int = logging.DEBUG) -> logging.Logger:
    """Configura il logging centralizzato con output nel terminale.

    Args:
        name: Nome del logger (default: at614_editor)
        level: Livello di logging (default: DEBUG)

    Returns:
        Logger configurato pronto all'uso
    """
    logger = logging.getLogger(name)
    logger.setLevel(level)

    if logger.handlers:
        return logger

    formatter = logging.Formatter(
        fmt="%(asctime)s | %(name)s | %(levelname)-8s | %(message)s",
        datefmt="%H:%M:%S",
    )

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(level)
    console_handler.setFormatter(formatter)

    logger.addHandler(console_handler)
    logger.propagate = False

    return logger


def get_logger(name: str) -> logging.Logger:
    """Ottiene un logger configurato per il modulo specifico.

    Args:
        name: Nome del modulo (__name__)

    Returns:
        Logger configurato per il modulo
    """
    return logging.getLogger("at614_editor." + name)
