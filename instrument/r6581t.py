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


class Terminal(Enum):
    """Input terminal selection."""
    FRONT = "FRONT"
    REAR = "REAR"


class RTDType(Enum):
    """
    RTD probe type for the meter's internal temperature calculation.

    Pt100 / Pt1000
    ---------------
    Platinum RTDs conforming to IEC 60751.  The meter has built-in
    Callendar-Van Dusen coefficients for both; no user configuration is
    needed.  Choose based on your probe's nominal resistance at 0 °C:
      • Pt100  — 100 Ω at 0 °C.  Most common industrial probe.
      • Pt1000 — 1000 Ω at 0 °C.  Higher impedance; better for long
                 cable runs (less lead-resistance error in 2-wire use).

    Both types have identical α (0.003850 per IEC 60751) and the same
    Callendar-Van Dusen shape; Pt1000 just scales R₀ by ×10.

    USER
    ----
    Lets you enter custom Callendar-Van Dusen coefficients (α, β, δ)
    and R₀.  Use this when your probe deviates from the IEC 60751
    standard, e.g. older DIN 43760 probes (α ≈ 0.003850 but slightly
    different β/δ), legacy Pt100 probes, or non-standard RTD materials.
    The teckit project used: α=0.00375, β=0.160, δ=1.605, R₀=1000 Ω
    (a Pt1000-class probe with older DIN coefficients).

    When to use which
    -----------------
    Use Pt100 or Pt1000 for any modern probe purchased to IEC 60751.
    Use USER only if your probe's datasheet gives explicit α/β/δ values
    that differ from IEC 60751 defaults, or when replicating a prior
    calibration setup (e.g. the teckit/xDevs reference measurement).
    """
    PT100  = "PT100"
    PT1000 = "PT1000"
    USER   = "USER"


class RTDScale(Enum):
    """
    Temperature scale used internally by the R6581T for RTD conversion.

    ITS90   — International Temperature Scale of 1990.  Current
              international standard.  Use for all new measurements.
    IPTS68  — International Practical Temperature Scale of 1968.
              Legacy standard, still relevant when comparing against
              historical calibration data taken before ~1990.  The
              difference from ITS-90 is ≤0.04 °C over 0–100 °C but
              grows at higher temperatures.
    """
    ITS90  = "ITS90"
    IPTS68 = "IPTS68"


class RTDUnit(Enum):
    """Output unit for RTD temperature readings."""
    C = "C"   # Celsius
    F = "F"   # Fahrenheit
    K = "K"   # Kelvin

    def symbol(self) -> str:
        return {"C": "°C", "F": "°F", "K": "K"}[self.value]


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

# When RTD is active the unit comes from the RTDUnit choice, not from UNITS.
# Use RTDUnit.symbol() at runtime — this dict is kept for non-RTD modes only.


# ======================================================================
# Data classes
# ======================================================================

@dataclass
class RTDConfig:
    """
    Configuration for the R6581T's built-in RTD temperature calculation.

    The meter performs resistance measurement in FRES mode and then
    applies a Callendar-Van Dusen conversion internally via the CALC
    subsystem, returning temperature directly in the chosen unit.

    Fields
    ------
    rtd_type : RTDType
        Probe type.  PT100/PT1000 use IEC 60751 built-in coefficients.
        USER requires explicit α, β, δ, R₀ values below.
    scale : RTDScale
        Temperature scale (ITS-90 recommended for new work).
    unit : RTDUnit
        Output unit (°C / °F / K).
    alpha : float
        Temperature coefficient α [Ω/Ω/°C].  IEC 60751 default: 0.003850.
        Used only when rtd_type == USER.
    beta : float
        Callendar-Van Dusen β coefficient.  Used only when rtd_type == USER.
    delta : float
        Callendar-Van Dusen δ coefficient.  Used only when rtd_type == USER.
    r_zero : float
        Nominal resistance at 0 °C in Ω.  100 for Pt100, 1000 for Pt1000.
        Used only when rtd_type == USER.
    """
    rtd_type: RTDType  = RTDType.PT100
    scale:    RTDScale = RTDScale.ITS90
    unit:     RTDUnit  = RTDUnit.C
    # USER-mode Callendar-Van Dusen coefficients
    # Defaults match the teckit/xDevs Pt1000-class probe (DIN 43760 era)
    alpha:  float = 0.00375
    beta:   float = 0.160
    delta:  float = 1.605
    r_zero: float = 1000.0


@dataclass
class MeasurementConfig:
    """Complete snapshot of the current measurement setup."""
    mode: MeasureMode = MeasureMode.DCV
    range_value: float = 10.0
    nplc: float = 100.0
    guard: Guard = Guard.FLOAT
    resistance_power: ResistancePower = ResistancePower.HIGH
    ocomp: OcompState = OcompState.OFF
    rtd: Optional[RTDConfig] = None  # None = pure resistance, not RTD


@dataclass
class ReadResult:
    """Result of a single measurement reading."""
    success: bool
    value: float = 0.0
    raw: str = ""
    error: str = ""
    unit: str = ""  # populated by R6581T.read() based on current config


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

        # Always tear down the CALC subsystem before reconfiguring.
        # If we were previously in RTD mode, the CALC layer stays armed
        # and interferes with subsequent :CONF: mode switches.  Disabling
        # it unconditionally is cheap (one SCPI write) and prevents the
        # meter from getting stuck.
        if self.config.rtd is not None:
            self._inst.write(":CALC:FORM:STAT OFF")

        if mode == MeasureMode.DCV:
            self._configure_dcv(range_value, nplc)
        elif mode == MeasureMode.DCI:
            self._configure_dci(range_value, nplc)
        elif mode == MeasureMode.RES2W:
            self._configure_res(range_value, nplc)
        elif mode == MeasureMode.RES4W:
            self._configure_fres(range_value, nplc, guard, resistance_power, ocomp, rtd)

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
            "Configured: mode=%s range=%s nplc=%s rtd=%s",
            mode.name, range_value, nplc,
            f"{rtd.rtd_type.value}/{rtd.unit.value}" if rtd else "off",
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
        rtd: Optional[RTDConfig] = None,
    ) -> None:
        """
        Configure 4-wire resistance measurement, with optional RTD conversion.
        Directly from teckit r6581t.py init_inst_fres() + set_ohmf_range().
        Uses :CONF:FRES to fully switch the measurement subsystem.

        When rtd is not None, the CALC subsystem is armed to convert the
        resistance reading to temperature before returning the value.  The
        meter then returns temperature directly — no post-processing needed.

        NOTE: The CALC:FORM RTD command block has not been verified on hardware
        yet (see project summary).  If readings come back as raw resistance
        values, check that :CALC:FORM:STAT ON takes effect and that the meter
        firmware supports this feature.
        """
        # Full mode switch via :CONF:
        self._inst.write(":CONF:FRES")
        self._inst.write(f":INP:GUAR {guard.value}")
        self._inst.write(f":SENS:FRES:POW {power.value}")
        self._inst.write(f":SENS:FRES:RANG {range_value:.6E}")
        self._inst.write(":SENS:FRES:DIG MAX")
        self._inst.write(f":SENS:FRES:NPLC {nplc}")
        self._inst.write(f":SENS:FRES:OCOM {ocomp.value}")
        self._inst.write(":SENS:FRES:SOUR:STAT ON")
        self._inst.write(":FORM:ELEM NONE")

        if rtd is not None:
            self._configure_rtd(rtd)

    def _configure_rtd(self, rtd: RTDConfig) -> None:
        """
        Enable the meter's internal RTD temperature calculation via CALC subsystem.

        Must be called after _configure_fres() has set up the FRES measurement.
        The meter will return temperature values directly from READ? in the
        chosen unit, rather than raw resistance.

        Standard probe types (Pt100/Pt1000) use the meter's built-in IEC 60751
        coefficients.  USER type sends explicit Callendar-Van Dusen coefficients.
        """
        self._inst.write(":CALC:FORM RTD")
        self._inst.write(f":CALC:FORM:RTD {rtd.scale.value}")
        self._inst.write(f":CALC:FORM:RTD:UNIT {rtd.unit.value}")

        if rtd.rtd_type == RTDType.USER:
            # Send custom Callendar-Van Dusen coefficients and R0
            self._inst.write(":SENS:TEMP:TRAN RTD")
            self._inst.write(":SENS:TEMP:RTD:TYPE USER")
            self._inst.write(f":SENS:TEMP:RTD:ALPH {rtd.alpha:.6f}")
            self._inst.write(f":SENS:TEMP:RTD:BETA {rtd.beta:.6f}")
            self._inst.write(f":SENS:TEMP:RTD:DELT {rtd.delta:.6f}")
            self._inst.write(f":SENS:TEMP:RTD:RZER {rtd.r_zero:.4f}")
        else:
            # Use the meter's built-in coefficients for standard probe types
            self._inst.write(":SENS:TEMP:TRAN RTD")
            self._inst.write(f":SENS:TEMP:RTD:TYPE {rtd.rtd_type.value}")

        # Arm the calculation — this must come last
        self._inst.write(":CALC:FORM:STAT ON")

    # ------------------------------------------------------------------
    # Terminal switching
    # ------------------------------------------------------------------

    def set_terminal(self, terminal: Terminal) -> None:
        """Switch between FRONT and REAR input terminals."""
        if not self._connected:
            raise RuntimeError("Not connected to instrument")
        self._inst.write(f":INP:TERM {terminal.value}")

    # ------------------------------------------------------------------
    # Reading
    # ------------------------------------------------------------------

    def read(self) -> ReadResult:
        """Trigger a single measurement and return the result."""
        if not self._connected:
            return ReadResult(success=False, error="Not connected")

        # Determine the unit for this reading from the current config
        cfg = self.config
        if cfg.mode == MeasureMode.RES4W and cfg.rtd is not None:
            unit = cfg.rtd.unit.symbol()
        else:
            unit = UNITS.get(cfg.mode, "?")

        try:
            self._inst.write("READ?")
            raw = self._inst.read().strip()
        except pyvisa.errors.VisaIOError as exc:
            logger.error("VISA read error: %s", exc)
            return ReadResult(success=False, error=str(exc), unit=unit)
        except Exception as exc:
            logger.error("Unexpected read error: %s", exc)
            return ReadResult(success=False, error=str(exc), unit=unit)

        try:
            value = float(raw)
        except ValueError:
            logger.warning("Could not parse reading: %r", raw)
            return ReadResult(success=False, raw=raw, error=f"Parse error: {raw}", unit=unit)

        return ReadResult(success=True, value=value, raw=raw, unit=unit)

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