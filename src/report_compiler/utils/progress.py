"""
Live terminal progress indicator that coexists with the scrolling logs.

The compiler emits detailed INFO logs as it works. This module pins a small,
continuously refreshing status region to the bottom of the terminal that shows:

  * a spinner (so the user can see work is ongoing during long, opaque steps
    like the Word conversion or the final PDF save),
  * the recursion breadcrumb and depth (which nested document is compiling),
  * the current pipeline stage (e.g. ``Stage 8/12 - PDF Conversion``),
  * elapsed time.

Log lines are routed through the same rich console so they print *above* the
live region instead of corrupting it. File logging (``--log-file``) is left
completely untouched.

The reporter degrades to a no-op when disabled (e.g. output is not a TTY) or if
``rich`` is unavailable, so callers can use it unconditionally.
"""

import logging
import sys
import threading
import time
from typing import List, Optional

from .logging_config import get_logger

try:
    from rich.console import Console, Group
    from rich.live import Live
    from rich.rule import Rule
    from rich.spinner import Spinner
    from rich.table import Table
    from rich.text import Text
    _RICH_AVAILABLE = True
except Exception:  # pragma: no cover - rich ships with typer, but stay safe
    _RICH_AVAILABLE = False


class _Frame:
    """One document on the recursion stack and its current stage."""

    __slots__ = ("name", "stage_num", "stage_total", "stage_name")

    def __init__(self, name: str):
        self.name = name
        self.stage_num = 0
        self.stage_total = 0
        self.stage_name = ""


class _LiveLogHandler(logging.Handler):
    """Routes log records through the rich console.

    While a :class:`rich.live.Live` is active on a console, printing through
    that same console makes the text appear above the live region. This handler
    re-uses the original console handler's level and formatter (which already
    embeds ANSI colors via ``ColoredFormatter``); ``Text.from_ansi`` turns those
    codes back into rich styling so colors survive.
    """

    def __init__(self, console: "Console", level: int, formatter: Optional[logging.Formatter]):
        super().__init__(level)
        self._console = console
        if formatter is not None:
            self.setFormatter(formatter)

    def emit(self, record: logging.LogRecord) -> None:
        try:
            msg = self.format(record)
            self._console.print(
                Text.from_ansi(msg), markup=False, highlight=False, soft_wrap=False
            )
        except Exception:
            self.handleError(record)


class ProgressReporter:
    """Tracks compilation progress and renders it as a live status region."""

    def __init__(self, enabled: bool = True):
        self.enabled = bool(enabled) and _RICH_AVAILABLE
        self._lock = threading.RLock()
        self._stack: List[_Frame] = []
        self._start = time.monotonic()
        self._console: Optional["Console"] = None
        self._live: Optional["Live"] = None
        self._spinner = None
        self._logger: Optional[logging.Logger] = None
        self._saved_handlers: List[logging.Handler] = []
        self._added_handlers: List[logging.Handler] = []

    # -- context management ------------------------------------------------
    def __enter__(self) -> "ProgressReporter":
        if self.enabled:
            try:
                self._start_live()
            except Exception:
                # Never let a display problem break the actual compilation.
                self.enabled = False
        return self

    def __exit__(self, exc_type, exc, tb) -> bool:
        self.stop()
        return False

    def _start_live(self) -> None:
        self._console = Console()
        self._spinner = Spinner("dots", style="cyan")
        self._start = time.monotonic()
        self._logger = get_logger()
        self._install_log_handler()
        self._live = Live(
            self,  # self is the renderable (see __rich__); recomputed each frame
            console=self._console,
            refresh_per_second=10,
            transient=True,
        )
        self._live.start()

    def stop(self) -> None:
        if not self.enabled:
            return
        if self._live is not None:
            try:
                self._live.stop()
            except Exception:
                pass
            self._live = None
        self._restore_log_handler()
        self.enabled = False

    # -- logging integration ----------------------------------------------
    def _install_log_handler(self) -> None:
        """Replace the stdout stream handler with one that prints via rich."""
        for handler in list(self._logger.handlers):
            if isinstance(handler, logging.StreamHandler) and getattr(handler, "stream", None) is sys.stdout:
                self._saved_handlers.append(handler)
                self._logger.removeHandler(handler)
                replacement = _LiveLogHandler(self._console, handler.level, handler.formatter)
                self._added_handlers.append(replacement)
                self._logger.addHandler(replacement)

    def _restore_log_handler(self) -> None:
        if self._logger is None:
            return
        for handler in self._added_handlers:
            self._logger.removeHandler(handler)
        for handler in self._saved_handlers:
            self._logger.addHandler(handler)
        self._added_handlers.clear()
        self._saved_handlers.clear()

    # -- progress updates --------------------------------------------------
    def enter_document(self, name: str) -> None:
        """Push a document onto the recursion breadcrumb."""
        if not self.enabled:
            return
        with self._lock:
            self._stack.append(_Frame(name))

    def exit_document(self) -> None:
        """Pop the current document off the breadcrumb."""
        if not self.enabled:
            return
        with self._lock:
            if self._stack:
                self._stack.pop()

    def set_stage(self, num: int, total: int, name: str) -> None:
        """Update the stage of the current (deepest) document."""
        if not self.enabled:
            return
        with self._lock:
            if self._stack:
                frame = self._stack[-1]
                frame.stage_num = num
                frame.stage_total = total
                frame.stage_name = name

    # -- rendering ---------------------------------------------------------
    def _format_elapsed(self) -> str:
        secs = int(time.monotonic() - self._start)
        return f"{secs // 60:02d}:{secs % 60:02d}"

    def __rich__(self):
        # Called by Live on every refresh, so elapsed time and the spinner stay
        # live even while the main thread is blocked in a long operation.
        with self._lock:
            if not self._stack:
                grid = Table.grid(padding=(0, 1))
                grid.add_row(self._spinner, Text("Starting...", style="dim"))
                return Group(Rule(style="dim"), grid)

            top = self._stack[-1]
            breadcrumb = Text(" ▸ ".join(f.name for f in self._stack), style="bold cyan")
            depth = len(self._stack) - 1
            if depth > 0:
                breadcrumb.append(f"   (depth {depth})", style="dim")

            stage = Text("   ")
            if top.stage_total:
                stage.append(f"Stage {top.stage_num}/{top.stage_total}", style="green")
                stage.append(f" · {top.stage_name}", style="white")
            stage.append(f"    {self._format_elapsed()}", style="dim")

        grid = Table.grid(padding=(0, 1))
        grid.add_row(self._spinner, breadcrumb)
        return Group(Rule(style="dim"), grid, stage)
