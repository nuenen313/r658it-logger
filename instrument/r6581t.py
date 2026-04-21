"""
Advantest R6581T 8.5-digit DMM driver.

Modernized from xDevs.com teckit project (https://github.com/tin-/teckit).
Original code: devices/r6581t.py, MIT License.

Supported modes:
  - DC Voltage       (DCV)  — via :CONF:VOLT:DC
  - DC Current       (DCI)  — via :CONF:CURR:DC
  - 2-wire Resistance (RES) — via :CONF:RES
  - 4-wire Resistance (FRES)— via :CONF:FRES (from teckit init_inst_fres)

All mode switching uses :CONF: commands (not :SENS:FUNC) as the R6581T
requires :CONF: to fully reconfigure the measurement subsystem.
"""

from __future__ import annotations

import logging
import time
from dataclasses import dataclass
from enum import Enum
from typing import Optional

import pyvisa

logger = logging.getLogger(__name__)


# ======================================================================
# Enumerations
# ======================================================================

class MeasureMode(Enum):
    """Measurement functions available on the R6581T."""
    DCV = "DCV"
    DCI = "DCI"
    RES2W = "RES"
    RES4W = "FRES"


class Guard(Enum):
    """Input guard setting."""
    FLOAT = "FLO"
    CABLE = "CAB"


class ResistancePower(Enum):
    """Ohms source current power level."""
    HIGH = "HI"
    LOW = "LO"


class OcompState(Enum):
    """Offset-compensated ohms."""
    ON = "ON"
    OFF = "OFF"


# ======================================================================
# Range / NPLC definitions
# ======================================================================

RANGES: dict[MeasureMode, list[tuple[str, float]]] = {
    MeasureMode.DCV: [
        ("100 mV", 0.1),
        ("1 V", 1),
        ("10 V", 10),
        ("100 V", 100),
        ("1000 V", 1000),
    ],
    MeasureMode.DCI: [
        ("100 nA", 100e-9),
        ("1 uA", 1e-6),
        ("10 uA", 10e-6),
        ("100 uA", 100e-6),
        ("1 mA", 1e-3),
        ("10 mA", 10e-3),
        ("100 mA", 100e-3),
        ("1 A", 1),
    ],
    MeasureMode.RES2W: [
        ("10 Ohm", 10),
        ("100 Ohm", 100),
        ("1 kOhm", 1e3),
        ("10 kOhm", 10e3),
        ("100 kOhm", 100e3),
        ("1 MOhm", 1e6),
        ("10 MOhm", 10e6),
        ("100 MOhm", 100e6),
        ("1 GOhm", 1e9),
    ],
    MeasureMode.RES4W: [
        ("10 Ohm", 10),
        ("100 Ohm", 100),
        ("1 kOhm", 1e3),
        ("10 kOhm", 10e3),
        ("100 kOhm", 100e3),
        ("1 MOhm", 1e6),
        ("10 MOhm", 10e6),
        ("100 MOhm", 100e6),
        ("1 GOhm", 1e9),
    ],
}

NPLC_VALUES: list[float] = [0.02, 0.1, 1, 10, 50, 100]

UNITS: dict[MeasureMode, str] = {
    MeasureMode.DCV: "V",
    MeasureMode.DCI: "A",
    MeasureMode.RES2W: "Ohm",
    MeasureMode.RES4W: "Ohm",
}


# ======================================================================
# Data classes
# ======================================================================

@dataclass
class MeasurementConfig:
    """Complete snapshot of the current measurement setup."""
    mode: MeasureMode = MeasureMode.DCV
    range_value: float = 10.0
    nplc: float = 100.0
    guard: Guard = Guard.FLOAT
    resistance_power: ResistancePower = ResistancePower.HIGH
    ocomp: OcompState = OcompState.OFF


@dataclass
class ReadResult:
    """Result of a single measurement reading."""
    success: bool
    value: float = 0.0
    raw: str = ""
    error: str = ""


# ======================================================================
# Driver
# ======================================================================

class R6581T:
    """
    Driver for the Advantest R6581T 8.5-digit DMM.

    Usage::

        meter = R6581T("GPIB0::24::INSTR")
        meter.connect()
        meter.configure(MeasureMode.DCV, range_value=10, nplc=100)
        result = meter.read()
        print(result.value)
        meter.disconnect()
    """

    def __init__(self, resource_string: str, timeout_ms: int = 60_000):
        self.resource_string = resource_string
        self.timeout_ms = timeout_ms
        self.config = MeasurementConfig()
        self._inst: Optional[pyvisa.resources.Resource] = None
        self._rm: Optional[pyvisa.ResourceManager] = None
        self._connected = False

    # ------------------------------------------------------------------
    # Connection management
    # ------------------------------------------------------------------

    def connect(self) -> None:
        """Open VISA connection and reset the instrument."""
        self._rm = pyvisa.ResourceManager()
        self._inst = self._rm.open_resource(self.resource_string)
        self._inst.timeout = self.timeout_ms
        self._inst.clear()
        self._inst.write("*RST")
        time.sleep(2)
        self._connected = True
        logger.info("Connected to %s", self.resource_string)

    def disconnect(self) -> None:
        """Close the VISA session."""
        if self._inst is not None:
            try:
                self._inst.close()
            except Exception:
                pass
        self._inst = None
        self._connected = False
        logger.info("Disconnected from %s", self.resource_string)

    @property
    def is_connected(self) -> bool:
        return self._connected

    # ------------------------------------------------------------------
    # Configuration
    # ------------------------------------------------------------------

    def configure(
        self,
        mode: MeasureMode,
        range_value: float = 10.0,
        nplc: float = 100.0,
        guard: Guard = Guard.FLOAT,
        resistance_power: ResistancePower = ResistancePower.HIGH,
        ocomp: OcompState = OcompState.OFF,
    ) -> None:
        """Configure the measurement function and all associated parameters."""
        if not self._connected:
            raise RuntimeError("Not connected to instrument")

        if mode == MeasureMode.DCV:
            self._configure_dcv(range_value, nplc)
        elif mode == MeasureMode.DCI:
            self._configure_dci(range_value, nplc)
        elif mode == MeasureMode.RES2W:
            self._configure_res(range_value, nplc)
        elif mode == MeasureMode.RES4W:
            self._configure_fres(range_value, nplc, guard, resistance_power, ocomp)

        self.config = MeasurementConfig(
            mode=mode,
            range_value=range_value,
            nplc=nplc,
            guard=guard,
            resistance_power=resistance_power,
            ocomp=ocomp,
        )

        logger.info(
            "Configured: mode=%s range=%s nplc=%s",
            mode.name, range_value, nplc,
        )

    def _configure_dcv(self, range_value: float, nplc: float) -> None:
        """
        Configure DC Voltage measurement.
        Uses :CONF:VOLT:DC for full mode switch.
        """
        self._inst.write(":CONF:VOLT:DC")
        self._inst.write(f":SENS:VOLT:DC:RANG {range_value:.6E}")
        self._inst.write(":SENS:VOLT:DC:DIG MAX")
        self._inst.write(f":SENS:VOLT:DC:NPLC {nplc}")
        self._inst.write(":FORM:ELEM NONE")

    def _configure_dci(self, range_value: float, nplc: float) -> None:
        """
        Configure DC Current measurement.
        Uses :CONF:CURR:DC for full mode switch.
        """
        self._inst.write(":CONF:CURR:DC")
        self._inst.write(f":SENS:CURR:DC:RANG {range_value:.6E}")
        self._inst.write(":SENS:CURR:DC:DIG MAX")
        self._inst.write(f":SENS:CURR:DC:NPLC {nplc}")
        self._inst.write(":FORM:ELEM NONE")

    def _configure_res(self, range_value: float, nplc: float) -> None:
        """
        Configure 2-wire resistance measurement.
        Uses :CONF:RES for full mode switch.
        """
        self._inst.write(":CONF:RES")
        self._inst.write(f":SENS:RES:RANG {range_value:.6E}")
        self._inst.write(":SENS:RES:DIG MAX")
        self._inst.write(f":SENS:RES:NPLC {nplc}")
        self._inst.write(":FORM:ELEM NONE")

    def _configure_fres(
        self, range_value: float, nplc: float,
        guard: Guard, power: ResistancePower, ocomp: OcompState,
    ) -> None:
        """
        Configure 4-wire resistance measurement.
        Directly from teckit r6581t.py init_inst_fres() + set_ohmf_range().
        Uses :CONF:FRES to fully switch the measurement subsystem.
        """
        # Full mode switch via :CONF: (as in teckit init_inst_fres)
        self._inst.write(":CONF:FRES")
        self._inst.write(f":INP:GUAR {guard.value}")
        self._inst.write(f":SENS:FRES:POW {power.value}")
        self._inst.write(f":SENS:FRES:RANG {range_value:.6E}")
        self._inst.write(":SENS:FRES:DIG MAX")
        self._inst.write(f":SENS:FRES:NPLC {nplc}")
        self._inst.write(f":SENS:FRES:OCOM {ocomp.value}")
        self._inst.write(":SENS:FRES:SOUR:STAT ON")
        self._inst.write(":FORM:ELEM NONE")

    # ------------------------------------------------------------------
    # Reading
    # ------------------------------------------------------------------

    def read(self) -> ReadResult:
        """Trigger a single measurement and return the result."""
        if not self._connected:
            return ReadResult(success=False, error="Not connected")

        try:
            self._inst.write("READ?")
            raw = self._inst.read().strip()
        except pyvisa.errors.VisaIOError as exc:
            logger.error("VISA read error: %s", exc)
            return ReadResult(success=False, error=str(exc))
        except Exception as exc:
            logger.error("Unexpected read error: %s", exc)
            return ReadResult(success=False, error=str(exc))

        try:
            value = float(raw)
        except ValueError:
            logger.warning("Could not parse reading: %r", raw)
            return ReadResult(success=False, raw=raw, error=f"Parse error: {raw}")

        return ReadResult(success=True, value=value, raw=raw)

    # ------------------------------------------------------------------
    # Utility
    # ------------------------------------------------------------------

    @staticmethod
    def list_resources() -> tuple[str, ...]:
        """Return a tuple of available VISA resource strings."""
        rm = pyvisa.ResourceManager()
        resources = rm.list_resources("?*")
        rm.close()
        return resources

    def identify(self) -> str:
        """Send *IDN? and return the response (may not be supported)."""
        if not self._connected:
            raise RuntimeError("Not connected")
        try:
            return self._inst.query("*IDN?").strip()
        except pyvisa.errors.VisaIOError:
            return "(no response — *IDN? may not be supported)"