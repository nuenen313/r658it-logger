"""
CSV data logger for measurement results.
"""

from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path
from typing import Optional


class CSVWriter:
    """
    Appends timestamped measurement readings to a CSV file.

    The file is created with a header on first write.  Subsequent calls
    to :meth:`write_row` append a single row and flush immediately so
    data is never lost, even on a crash or Ctrl-C.
    """

    HEADER = ["timestamp", "elapsed_s", "mode", "range", "nplc", "value", "unit", "overload"]

    def __init__(self) -> None:
        self._file = None
        self._writer: Optional[csv.writer] = None
        self._path: Optional[Path] = None
        self._start_time: Optional[datetime] = None

    def open(self, path: str | Path) -> None:
        """Open (or create) the CSV file and write the header if new."""
        # Close any previously open file first
        self.close()

        self._path = Path(path)
        file_exists = self._path.exists() and self._path.stat().st_size > 0

        self._file = open(self._path, "a", newline="", encoding="utf-8")
        self._writer = csv.writer(self._file)

        if not file_exists:
            self._writer.writerow(self.HEADER)
            self._file.flush()

        self._start_time = datetime.now()

    def write_row(
        self,
        value: float,
        mode: str,
        range_label: str,
        nplc: float,
        unit: str,
        overload: bool = False,
    ) -> None:
        """Append one measurement row and flush to disk."""
        if self._writer is None:
            raise RuntimeError("CSV file not open — call open() first")

        now = datetime.now()
        elapsed = (now - self._start_time).total_seconds() if self._start_time else 0.0

        self._writer.writerow([
            now.strftime("%Y-%m-%d %H:%M:%S.%f")[:-3],
            f"{elapsed:.3f}",
            mode,
            range_label,
            nplc,
            "OVERLOAD" if overload else f"{value:.10E}",
            unit,
            overload,
        ])
        self._file.flush()

    def close(self) -> None:
        """Close the CSV file. Safe to call multiple times."""
        if self._file is not None:
            try:
                if not self._file.closed:
                    self._file.close()
            except Exception:
                pass
            self._file = None
            self._writer = None

    @property
    def is_open(self) -> bool:
        return self._file is not None and not self._file.closed

    @property
    def path(self) -> Optional[Path]:
        return self._path