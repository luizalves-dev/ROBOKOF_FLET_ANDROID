from __future__ import annotations

import logging
import sys


LOGGER_ROOT_NAME = "robokof_terminal"


def configure_terminal_logging(level: int = logging.INFO) -> logging.Logger:
    logger = logging.getLogger(LOGGER_ROOT_NAME)
    logger.setLevel(level)

    if not any(getattr(handler, "_robokof_terminal", False) for handler in logger.handlers):
        handler = logging.StreamHandler(sys.stdout)
        handler.setLevel(level)
        handler.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
        handler._robokof_terminal = True
        logger.addHandler(handler)

    logger.propagate = False
    return logger


def get_terminal_logger(name: str = "") -> logging.Logger:
    configure_terminal_logging()
    if not name:
        return logging.getLogger(LOGGER_ROOT_NAME)
    child = logging.getLogger(f"{LOGGER_ROOT_NAME}.{name}")
    child.setLevel(logging.INFO)
    return child
