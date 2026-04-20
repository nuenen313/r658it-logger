"""
Advantest R6581T 8.5-digit DMM driver.

Modernized from xDevs.com teckit project (https://github.com/tin-/teckit).
Original code: devices/r6581t.py, MIT License.

Supports:
  - DC Voltage  (DCV)       — proven working
  - DC Current  (DCI)       — standard SCPI, untested on this unit
  - 2-wire Resistance (RES) — standard SCPI, untested on this unit
  - 4-wire Resistance (FRES)— proven working (from teckit init_inst_fres)
  - Temperature (TEMP)      — measured as 4-wire RTD resistance, converted
                               to °C in software via Callendar-Van Dusen

Temperature note:
  The R6581T does NOT have a native SCPI TEMP function.  The original teckit
  code's commented-out `:SENS:FUNC 'TEMP'` commands were experimental and
  non-functional.  Instead we measure the RTD probe's resistance in 4-wire
  mode and convert to temperature using the Callendar-Van Dusen equation.
"""

from __future__ import annotations

import logging
import math
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
    DCV = "VOLT:DC"
    DCI = "CURR:DC"
    RES2W = "RES"
    RES4W = "FRES"
    TEMP = "TEMP"       # Software mode: 4W resistance + CVD conversion


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


class RTDType(Enum):
    """Standard RTD types with predefined CVD coefficients."""
    PT100 = "PT100"
    PT1000 = "PT1000"
    USER = "USER"


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
    # TEMP: range is auto-selected based on RTD R0 value
    MeasureMode.TEMP: [],
}

NPLC_VALUES: list[float] = [0.02, 0.1, 1, 10, 50, 100]

UNITS: dict[MeasureMode, str] = {
    MeasureMode.DCV: "V",
    MeasureMode.DCI: "A",
    MeasureMode.RES2W: "Ohm",
    MeasureMode.RES4W: "Ohm",
    MeasureMode.TEMP: "C",
}


# ======================================================================
# Callendar-Van Dusen conversion
# ======================================================================

# IEC 60751 standard coefficients for platinum RTDs (alpha = 0.003850)
_CVD_A = 3.9083e-3
_CVD_B = -5.775e-7
_CVD_C = -4.183e-12   # only used below 0 °C


def rtd_resistance_to_temperature(resistance: float, r0: float = 100.0) -> float:
    """
    Convert RTD resistance to temperature in °C using the
    Callendar-Van Dusen equation (IEC 60751).

    For T >= 0 °C:  R(T) = R0 * (1 + A*T + B*T²)
      -> solve quadratic for T

    For T < 0 °C the full 4th-order equation is used, but since the
    quadratic solution gives a good starting point, we use iterative
    refinement (Newton's method) for the negative range.

    Args:
        resistance: measured resistance in ohms
        r0: RTD resistance at 0 °C (100.0 for PT100, 1000.0 for PT1000)

    Returns:
        temperature in °C
    """
    # Quadratic solution (valid for T >= 0 °C, good approximation for T < 0 °C)
    # R = R0 * (1 + A*T + B*T²)
    # R0*B*T² + R0*A*T + (R0 - R) = 0
    discriminant = (_CVD_A * _CVD_A) - 4.0 * _CVD_B * (1.0 - resistance / r0)
    if discriminant < 0:
        return float("nan")

    temp = (-_CVD_A + math.sqrt(discriminant)) / (2.0 * _CVD_B)

    # For negative temperatures, refine with Newton's method using full equation
    if temp < 0:
        for _ in range(10):
            t = temp
            r_calc = r0 * (1.0 + _CVD_A * t + _CVD_B * t * t
                           + _CVD_C * (t - 100.0) * t * t * t)
            dr_dt = r0 * (_CVD_A + 2.0 * _CVD_B * t
                          + _CVD_C * (4.0 * t * t * t - 300.0 * t * t))
            if abs(dr_dt) < 1e-15:
                break
            temp = t - (r_calc - resistance) / dr_dt

    return temp


def _auto_range_for_rtd(r0: float) -> float:
    """Pick a suitable 4-wire resistance range for the given RTD R0."""
    # RTD resistance at ~200°C is roughly 1.8 * R0
    # Pick the smallest range that covers this
    for _, range_val in RANGES[MeasureMode.RES4W]:
        if range_val >= r0 * 2.0:
            return range_val
    return 10e3  # fallback


# ======================================================================
# Data classes
# ======================================================================

@dataclass
class RTDConfig:
    """RTD probe parameters for temperature measurement."""
    type: RTDType = RTDType.PT100
    r_zero: float = 100.0       # resistance at 0 °C (ohms)


@dataclass
class MeasurementConfig:
    """Complete snapshot of the current measurement setup."""
    mode: MeasureMode = MeasureMode.DCV
    range_value: float = 10.0
    nplc: float = 100.0
    guard: Guard = Guard.FLOAT
    resistance_power: ResistancePower = ResistancePower.HIGH
    ocomp: OcompState = OcompState.OFF
    rtd: RTDConfig = None

    def __post_init__(self):
        if self.rtd is None:
            self.rtd = RTDConfig()


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
        rtd: Optional[RTDConfig] = None,
    ) -> None:
        """Configure the measurement function and all associated parameters."""
        if not self._connected:
            raise RuntimeError("Not connected to instrument")

        if rtd is None:
            rtd = RTDConfig()

        # TEMP mode is implemented as 4-wire resistance + software conversion
        if mode == MeasureMode.TEMP:
            actual_range = _auto_range_for_rtd(rtd.r_zero)
            self._configure_fres(actual_range, nplc, guard, resistance_power, ocomp)
        elif mode == MeasureMode.RES4W:
            self._configure_fres(range_value, nplc, guard, resistance_power, ocomp)
        elif mode == MeasureMode.RES2W:
            self._configure_res(range_value, nplc, guard)
        else:
            # DCV or DCI — straightforward
            func = mode.value
            self._inst.write(f":SENS:FUNC '{func}'")
            self._inst.write(f":SENS:{func}:RANG {range_value:.6E}")
            self._inst.write(f":SENS:{func}:DIG MAX")
            self._inst.write(f":SENS:{func}:NPLC {nplc}")
            self._inst.write(":FORM:ELEM NONE")
            try:
                self._inst.write(f":INP:GUAR {guard.value}")
            except pyvisa.errors.VisaIOError:
                pass

        self.config = MeasurementConfig(
            mode=mode,
            range_value=range_value,
            nplc=nplc,
            guard=guard,
            resistance_power=resistance_power,
            ocomp=ocomp,
            rtd=rtd,
        )

        logger.info(
            "Configured: mode=%s range=%s nplc=%s guard=%s",
            mode.name, range_value, nplc, guard.name,
        )

    def _configure_fres(
        self, range_value: float, nplc: float,
        guard: Guard, power: ResistancePower, ocomp: OcompState,
    ) -> None:
        """Configure 4-wire resistance — matches teckit init_inst_fres()."""
        self._inst.write(":SENS:FUNC 'FRES'")
        self._inst.write(f":SENS:FRES:RANG {range_value:.6E}")
        self._inst.write(":SENS:FRES:DIG MAX")
        self._inst.write(f":SENS:FRES:NPLC {nplc}")
        self._inst.write(f":SENS:FRES:POW {power.value}")
        self._inst.write(f":SENS:FRES:OCOM {ocomp.value}")
        self._inst.write(":SENS:FRES:SOUR:STAT ON")
        self._inst.write(":FORM:ELEM NONE")
        try:
            self._inst.write(f":INP:GUAR {guard.value}")
        except pyvisa.errors.VisaIOError:
            pass

    def _configure_res(
        self, range_value: float, nplc: float, guard: Guard,
    ) -> None:
        """Configure 2-wire resistance."""
        self._inst.write(":SENS:FUNC 'RES'")
        self._inst.write(f":SENS:RES:RANG {range_value:.6E}")
        self._inst.write(":SENS:RES:DIG MAX")
        self._inst.write(f":SENS:RES:NPLC {nplc}")
        self._inst.write(":FORM:ELEM NONE")
        try:
            self._inst.write(f":INP:GUAR {guard.value}")
        except pyvisa.errors.VisaIOError:
            pass

    # ------------------------------------------------------------------
    # Reading
    # ------------------------------------------------------------------

    def read(self) -> ReadResult:
        """
        Trigger a single measurement and return the result.

        In TEMP mode, the raw 4-wire resistance is read and converted to
        temperature using the Callendar-Van Dusen equation.
        """
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

        # Convert resistance to temperature if in TEMP mode
        if self.config.mode == MeasureMode.TEMP:
            # Check for overload before converting
            if abs(value) > 9e+37:
                return ReadResult(success=True, value=value, raw=raw)
            r0 = self.config.rtd.r_zero
            value = rtd_resistance_to_temperature(value, r0)

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