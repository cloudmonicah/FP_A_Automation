"""Logging for the app: one rotating log file per Windows user.

The automations report progress with print(). configure() routes stdout and
stderr into the log as well, which also keeps print() working in the
windowed .exe, where sys.stdout is None.
"""

from __future__ import annotations

import logging
import sys
import threading
from logging.handlers import RotatingFileHandler
from pathlib import Path

from shared import settings

LOG_FILE_NAME = "fpa_automation.log"
_FORMAT = "%(asctime)s %(levelname)-7s %(name)s: %(message)s"


def log_dir() -> Path:
    return settings.user_dir() / "logs"


def log_file() -> Path:
    return log_dir() / LOG_FILE_NAME


class _StreamToLogger:
    """File-like object that writes complete lines to a logger."""

    def __init__(self, logger: logging.Logger, level: int, passthrough=None) -> None:
        self.logger = logger
        self.level = level
        self.passthrough = passthrough
        self._buffer = ""
        self._lock = threading.Lock()

    def write(self, text: str) -> int:
        if self.passthrough is not None:
            try:
                self.passthrough.write(text)
            except Exception:
                pass
        with self._lock:
            self._buffer += text
            *lines, self._buffer = self._buffer.split("\n")
        for line in lines:
            if line.strip():
                self.logger.log(self.level, line.rstrip())
        return len(text)

    def flush(self) -> None:
        if self.passthrough is not None:
            try:
                self.passthrough.flush()
            except Exception:
                pass

    def isatty(self) -> bool:
        return False


_configured = False


def configure() -> Path:
    """Set up logging once. Returns the log file path."""
    global _configured
    path = log_file()
    if _configured:
        return path
    _configured = True

    path.parent.mkdir(parents=True, exist_ok=True)
    level = getattr(logging, str(settings.get("log_level")).upper(), logging.INFO)

    root = logging.getLogger()
    root.setLevel(level)
    file_handler = RotatingFileHandler(
        path, maxBytes=5 * 1024 * 1024, backupCount=5, encoding="utf-8"
    )
    file_handler.setFormatter(logging.Formatter(_FORMAT))
    root.addHandler(file_handler)

    original_out, original_err = sys.stdout, sys.stderr
    if original_err is not None:
        # Warnings and errors also show in the terminal when there is one.
        console = logging.StreamHandler(original_err)
        console.setLevel(logging.WARNING)
        console.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
        # print() output already reaches the terminal directly.
        console.addFilter(lambda record: record.name not in ("output", "errors"))
        root.addHandler(console)
    sys.stdout = _StreamToLogger(logging.getLogger("output"), logging.INFO, original_out)
    sys.stderr = _StreamToLogger(logging.getLogger("errors"), logging.ERROR, original_err)

    def log_uncaught(exc_type, exc, tb):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc, tb)
            return
        logging.getLogger("uncaught").critical("Unhandled error", exc_info=(exc_type, exc, tb))

    def log_thread(args):
        logging.getLogger("uncaught").critical(
            "Unhandled error in thread %s",
            getattr(args.thread, "name", "?"),
            exc_info=(args.exc_type, args.exc_value, args.exc_traceback),
        )

    sys.excepthook = log_uncaught
    threading.excepthook = log_thread

    logging.getLogger(__name__).info("Logging to %s", path)
    return path
