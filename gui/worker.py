"""
Background worker thread for continuous measurement acquisition.

Uses stdlib threading so tkinter's main loop can stay responsive
while GPIB I/O blocks in the background.
"""

from __future__ import annotations

import threading
import time
from typing import Callable, Optional

from instrument.r6581t import R6581T, ReadResult


class MeasurementWorker:
    """
    Periodically triggers a reading on the R6581T in a daemon thread
    and delivers results via a callback.

    The callback is called from the worker thread — the GUI must use
    ``root.after()`` or a queue to marshal results onto the main thread.
    """

    def __init__(
        self,
        meter: R6581T,
        interval_s: float = 5.0,
        on_reading: Optional[Callable[[ReadResult], None]] = None,
    ):
        self.meter = meter
        self.interval_s = interval_s
        self.on_reading = on_reading
        self._stop_event = threading.Event()
        self._thread: Optional[threading.Thread] = None

    def start(self) -> None:
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop_event.set()
        # Don't join — the thread is a daemon and will exit on its own.
        # Joining would block the GUI while waiting for a GPIB read to finish.
        self._thread = None

    @property
    def is_running(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def _run(self) -> None:
        while not self._stop_event.is_set():
            t_start = time.monotonic()
            result = self.meter.read()
            read_duration = time.monotonic() - t_start

            # Check again after read — could have been stopped while waiting
            if self._stop_event.is_set():
                break
            if self.on_reading is not None:
                self.on_reading(result)

            # Sleep for the remainder of the interval, accounting for read time
            remaining = self.interval_s - read_duration
            if remaining > 0:
                elapsed = 0.0
                while elapsed < remaining and not self._stop_event.is_set():
                    time.sleep(0.1)
                    elapsed += 0.1