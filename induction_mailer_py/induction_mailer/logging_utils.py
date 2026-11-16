from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional


def log_event(log_path: str | Path, message: str) -> None:
    """Append a timestamped line to a log file (VBS `LogEvent`)."""
    path = Path(log_path)
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("a", encoding="utf-8", errors="replace") as f:
            f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S}  {message}\n")
    except Exception:
        # Logging must never break the workflow.
        pass


def _iso_week_key(dt: datetime) -> tuple[int, int]:
    iso = dt.isocalendar()
    return int(iso.year), int(iso.week)


def should_reset_log_for_new_week(log_path: str | Path, now: Optional[datetime] = None) -> bool:
    """Port of VBS ShouldResetLogForNewWeek.

    Resets when the ISO week number of the file's last modification differs from now.
    """
    path = Path(log_path)
    if not path.exists():
        return False
    if now is None:
        now = datetime.now()
    try:
        mtime = datetime.fromtimestamp(path.stat().st_mtime)
    except Exception:
        return False
    return _iso_week_key(mtime) != _iso_week_key(now)


def reset_log_if_new_week(log_path: str | Path, now: Optional[datetime] = None) -> bool:
    """Truncate the log file if it belongs to a previous week."""
    path = Path(log_path)
    if not path.exists():
        return False
    if not should_reset_log_for_new_week(path, now=now):
        return False
    try:
        path.write_text("", encoding="utf-8")
        return True
    except Exception:
        return False


@dataclass
class WeeklyFileLogger:
    """Minimal weekly-reset logger compatible with the VBS scripts.

    - Resets log file when week changes
    - Adds a variant prefix (e.g. "[Main] ")
    - Supports verbose toggle
    """

    log_path: str | Path
    variant: str = ""
    extended_logging: bool = True

    def initialize(self) -> None:
        reset = reset_log_if_new_week(self.log_path)
        if reset:
            self.log(f"Weekly log reset detected. Starting new log file: {self.log_path}")
        else:
            if Path(self.log_path).exists():
                self.log(f"Continuing existing weekly log: {self.log_path}")
            else:
                self.log(f"Creating new weekly log file: {self.log_path}")
        self.verbose("InitializeLogging invoked")
        self.log("=== Script execution started ===")
        self.log(f"Logging to file: {self.log_path}")

    def log(self, message: str) -> None:
        log_event(self.log_path, f"{self.variant}{message}")

    def verbose(self, message: str) -> None:
        if self.extended_logging:
            self.log(f"[VERBOSE] {message}")

    def error(self, context: str, exc: Optional[BaseException] = None) -> None:
        if exc is None:
            self.log(f"{context}")
        else:
            self.log(f"{context} | Err: {type(exc).__name__}: {exc}")


def make_log_handler(default_log_path: str | Path) -> Callable[[str, str], None]:
    """Return a (log_path, message) -> None callback, similar to VBS `LogHandler`."""

    def _handler(log_path: str, message: str) -> None:
        used = log_path or str(default_log_path)
        log_event(used, message)

    return _handler
