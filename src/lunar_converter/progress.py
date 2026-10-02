"""Leichte Fortschrittsanzeige für dateibasierte CLI-Läufe."""

from __future__ import annotations

import sys
import time
from collections.abc import Callable
from typing import TextIO


class TerminalProgress:
    def __init__(
        self,
        label: str,
        total: int,
        *,
        stream: TextIO | None = None,
        clock: Callable[[], float] = time.monotonic,
    ) -> None:
        self.label = label
        self.total = total
        self._stream = stream or sys.stdout
        self._clock = clock
        self._started = clock()
        self._last_width = 0
        self._finished = False

    def update(self, completed: int, *, current: str = "", status: str = "") -> None:
        if self.total <= 0:
            return
        completed = min(max(completed, 0), self.total)
        elapsed = max(0.0, self._clock() - self._started)
        ratio = completed / self.total
        filled = round(24 * ratio)
        bar = "#" * filled + "-" * (24 - filled)
        if completed:
            estimated_total = elapsed / completed * self.total
            remaining = max(0.0, estimated_total - elapsed)
            estimates = f"Gesamt ~{_format_duration(estimated_total)} | Rest ~{_format_duration(remaining)}"
        else:
            estimates = "Gesamt --:-- | Rest --:--"
        current_text = _shorten(current, 56)
        details = " | ".join(part for part in (status, current_text) if part)
        line = (
            f"{self.label} [{bar}] {completed}/{self.total} ({ratio:.0%}) | "
            f"Laufzeit {_format_duration(elapsed)} | {estimates}"
        )
        if details:
            line += f" | {details}"
        padding = " " * max(0, self._last_width - len(line))
        self._stream.write(f"\r{line}{padding}")
        self._stream.flush()
        self._last_width = len(line)
        if completed == self.total:
            self._stream.write("\n")
            self._stream.flush()
            self._last_width = 0
            self._finished = True

    def finish(self) -> None:
        if self.total == 0:
            self._stream.write(f"{self.label}: keine Dateien gefunden.\n")
            self._stream.flush()
        elif self._last_width and not self._finished:
            self._stream.write("\n")
            self._stream.flush()
            self._last_width = 0


def _format_duration(seconds: float) -> str:
    total_seconds = max(0, round(seconds))
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    if hours:
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
    return f"{minutes:02d}:{seconds:02d}"


def _shorten(value: str, limit: int) -> str:
    if len(value) <= limit:
        return value
    return f"…{value[-(limit - 1):]}"