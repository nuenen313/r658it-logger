"""
Background worker thread for continuous measurement acquisition.

Supports single-terminal or dual-terminal (front+rear) alternating reads,
each with independent measurement configuration.
"""

from __future__ import annotations

import threading
import time
from dataclasses import dataclass
from typing import Callable, Optional

from instrument.r6581t import (
    R6581T, ReadResult, Terminal, MeasureMode,
    Guard, ResistancePower, OcompState, RTDConfig,
)


@dataclass
class TerminalConfig:
    """Measurement configuration for one terminal."""
    enabled: bool = False
    mode: MeasureMode = MeasureMode.DCV
    range_value: float = 10.0
    nplc: float = 100.0
    guard: Guard = Guard.FLOAT
    resistance_power: ResistancePower = ResistancePower.HIGH
    ocomp: OcompState = OcompState.OFF
    rtd: Optional[RTDConfig] = None  # None = disabled; only valid with RES4W mode


@dataclass
class TerminalReadResult:
    """A read result tagged with which terminal it came from."""
    terminal: Terminal
    result: ReadResult


class MeasurementWorker:
    """
    Periodically triggers readings on the R6581T in a daemon thread.

    For each enabled terminal:
      switch terminal → configure mode → read → emit

    Then waits the remainder of the interval before repeating.
    """

    def __init__(
        self,
        meter: R6581T,
        interval_s: float = 5.0,
        front_config: Optional[TerminalConfig] = None,
        rear_config: Optional[TerminalConfig] = None,
        on_reading: Optional[Callable[[TerminalReadResult], None]] = None,
    ):
        self.meter = meter
        self.interval_s = interval_s
        self.front_config = front_config or TerminalConfig(enabled=True)
        self.rear_config = rear_config or TerminalConfig(enabled=False)
        self.on_reading = on_reading
        self._stop_event = threading.Event()
        self._thread: Optional[threading.Thread] = None
        # Track what's currently configured to avoid redundant reconfiguration.
        # Terminal is tracked separately because set_terminal() causes the meter
        # to re-settle its sense lines, which on resistance modes adds a delay
        # comparable to an entire measurement cycle (doubling the interval).
        self._last_terminal: Optional[Terminal] = None
        self._last_config_key: Optional[tuple] = None

    def start(self) -> None:
        self._stop_event.clear()
        self._last_terminal = None
        self._last_config_key = None
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop_event.set()
        self._thread = None

    @property
    def is_running(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def _run(self) -> None:
        while not self._stop_event.is_set():
            t_start = time.monotonic()

            if self.front_config.enabled:
                self._read_terminal(Terminal.FRONT, self.front_config)
                if self._stop_event.is_set():
                    break

            if self.rear_config.enabled:
                self._read_terminal(Terminal.REAR, self.rear_config)
                if self._stop_event.is_set():
                    break

            read_duration = time.monotonic() - t_start
            remaining = self.interval_s - read_duration
            if remaining > 0:
                elapsed = 0.0
                while elapsed < remaining and not self._stop_event.is_set():
                    time.sleep(0.1)
                    elapsed += 0.1

    def _read_terminal(self, terminal: Terminal, cfg: TerminalConfig) -> None:
        """Switch terminal (if needed), reconfigure (if needed), read, and emit."""
        # Only send :INP:TERM when actually switching terminals.
        # On the R6581T, :INP:TERM causes the meter to re-settle its sense
        # lines even if the terminal hasn't changed — this adds significant
        # delay in resistance modes (high NPLC), effectively doubling the
        # measurement interval.
        if self._last_terminal != terminal:
            try:
                self.meter.set_terminal(terminal)
                self._last_terminal = terminal
                # Terminal changed — force reconfigure since the new terminal
                # may not have the same measurement setup loaded.
                self._last_config_key = None
            except Exception:
                pass

        # Reconfigure only if settings differ from what's currently loaded.
        # The key includes all parameters that affect SCPI state.  Dataclass
        # equality on RTDConfig compares all fields, so any coefficient
        # change triggers reconfiguration.
        config_key = (cfg.mode, cfg.range_value, cfg.nplc, cfg.guard,
                      cfg.resistance_power, cfg.ocomp, cfg.rtd)
        if self._last_config_key != config_key:
            try:
                self.meter.configure(
                    mode=cfg.mode,
                    range_value=cfg.range_value,
                    nplc=cfg.nplc,
                    guard=cfg.guard,
                    resistance_power=cfg.resistance_power,
                    ocomp=cfg.ocomp,
                    rtd=cfg.rtd,
                )
                self._last_config_key = config_key
            except Exception:
                pass

        result = self.meter.read()

        if self._stop_event.is_set():
            return

        if self.on_reading is not None:
            self.on_reading(TerminalReadResult(terminal=terminal, result=result))