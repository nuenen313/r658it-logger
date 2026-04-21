"""
Main application window for the R6581T data logger (tkinter version).
Supports dual front/rear terminal display with independent configuration.
"""

from __future__ import annotations

import queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from pathlib import Path
from typing import Optional

from instrument.r6581t import (
    R6581T, MeasureMode, RANGES, NPLC_VALUES, UNITS,
    Guard, ResistancePower, OcompState, Terminal,
    ReadResult,
)
from data.csv_writer import CSVWriter
from gui.worker import MeasurementWorker, TerminalReadResult, TerminalConfig


# ======================================================================
# Per-terminal config panel widget
# ======================================================================

class TerminalPanel:
    """
    A reusable group of controls for configuring one terminal's
    measurement mode, range, NPLC, guard, and resistance options.
    """

    def __init__(
        self,
        parent: tk.Widget,
        label: str,
        enabled_default: bool,
        on_toggle: callable,
        on_apply: callable = None,
    ):
        self.enabled_var = tk.BooleanVar(value=enabled_default)
        self.mode_var = tk.StringVar(value="DCV")
        self.range_var = tk.StringVar()
        self.nplc_var = tk.StringVar(value="100")
        self.guard_var = tk.StringVar(value="Float")
        self.power_var = tk.StringVar(value="High")
        self.ocomp_var = tk.StringVar(value="OFF")
        self._on_apply = on_apply

        # --- Main frame ---
        self.frame = ttk.LabelFrame(parent, text=label, padding=6)

        # Enable checkbox + reading display
        top = ttk.Frame(self.frame)
        top.pack(fill=tk.X)

        self.check = ttk.Checkbutton(
            top, text=f"Enable {label}",
            variable=self.enabled_var,
            command=on_toggle,
        )
        self.check.pack(side=tk.LEFT)

        self.reading_label = tk.Label(
            self.frame, text="---", font=("Consolas", 28, "bold"),
            fg="#1a1a1a" if enabled_default else "#aaaaaa",
            bg="#ffffff" if enabled_default else "#e8e8e8",
            bd=2, relief="groove", padx=12, pady=10, anchor="center",
            width=20, height=1,
        )
        self.reading_label.pack(fill=tk.X, pady=(4, 6))

        # --- Config controls ---
        self.config_frame = ttk.Frame(self.frame)
        self.config_frame.pack(fill=tk.X)
        row = 0

        ttk.Label(self.config_frame, text="Mode:").grid(row=row, column=0, sticky=tk.W, pady=1)
        self.mode_combo = ttk.Combobox(
            self.config_frame, textvariable=self.mode_var, width=14,
            values=[m.name for m in MeasureMode], state="readonly",
        )
        self.mode_combo.grid(row=row, column=1, sticky=tk.EW, padx=4, pady=1)
        self.mode_combo.bind("<<ComboboxSelected>>", lambda e: self._on_mode_changed())
        row += 1

        self._range_label = ttk.Label(self.config_frame, text="Range:")
        self._range_label.grid(row=row, column=0, sticky=tk.W, pady=1)
        self.range_combo = ttk.Combobox(
            self.config_frame, textvariable=self.range_var, width=14, state="readonly",
        )
        self.range_combo.grid(row=row, column=1, sticky=tk.EW, padx=4, pady=1)
        row += 1

        ttk.Label(self.config_frame, text="NPLC:").grid(row=row, column=0, sticky=tk.W, pady=1)
        self.nplc_combo = ttk.Combobox(
            self.config_frame, textvariable=self.nplc_var, width=14,
            values=[str(n) for n in NPLC_VALUES], state="readonly",
        )
        self.nplc_combo.grid(row=row, column=1, sticky=tk.EW, padx=4, pady=1)
        row += 1

        ttk.Label(self.config_frame, text="Guard:").grid(row=row, column=0, sticky=tk.W, pady=1)
        self.guard_combo = ttk.Combobox(
            self.config_frame, textvariable=self.guard_var, width=14,
            values=[g.name.capitalize() for g in Guard], state="readonly",
        )
        self.guard_combo.grid(row=row, column=1, sticky=tk.EW, padx=4, pady=1)
        row += 1

        # 4W resistance options
        self._res_sep = ttk.Label(self.config_frame, text="4W Resistance", style="Section.TLabel")
        self._res_sep.grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=(4, 1))
        row += 1

        self._power_label = ttk.Label(self.config_frame, text="Src Power:")
        self._power_label.grid(row=row, column=0, sticky=tk.W, pady=1)
        self.power_combo = ttk.Combobox(
            self.config_frame, textvariable=self.power_var, width=14,
            values=[p.name.capitalize() for p in ResistancePower], state="readonly",
        )
        self.power_combo.grid(row=row, column=1, sticky=tk.EW, padx=4, pady=1)
        row += 1

        self._ocomp_label = ttk.Label(self.config_frame, text="Offset Comp:")
        self._ocomp_label.grid(row=row, column=0, sticky=tk.W, pady=1)
        self.ocomp_combo = ttk.Combobox(
            self.config_frame, textvariable=self.ocomp_var, width=14,
            values=[o.value for o in OcompState], state="readonly",
        )
        self.ocomp_combo.grid(row=row, column=1, sticky=tk.EW, padx=4, pady=1)
        row += 1

        # Apply button
        self.apply_btn = ttk.Button(
            self.config_frame, text="Apply",
            command=self._do_apply,
        )
        self.apply_btn.grid(row=row, column=0, columnspan=2, sticky=tk.EW, pady=(6, 0))
        row += 1

        self.config_frame.columnconfigure(1, weight=1)

        # Initial visibility
        self._on_mode_changed()
        if not enabled_default:
            self.config_frame.pack_forget()

    def _on_mode_changed(self) -> None:
        mode = MeasureMode[self.mode_var.get()]
        ranges = RANGES.get(mode, [])
        if ranges:
            labels = [r[0] for r in ranges]
            self.range_combo["values"] = labels
            self.range_combo.set(labels[len(labels) // 2])
            self._range_label.grid()
            self.range_combo.grid()
        else:
            self.range_combo["values"] = []
            self.range_combo.set("")
            self._range_label.grid_remove()
            self.range_combo.grid_remove()

        is_res4w = mode == MeasureMode.RES4W
        for w in (self._res_sep, self._power_label, self.power_combo,
                  self._ocomp_label, self.ocomp_combo):
            w.grid() if is_res4w else w.grid_remove()

    def update_appearance(self) -> None:
        """Update display and config visibility based on enabled state."""
        if self.enabled_var.get():
            self.reading_label.configure(fg="#1a1a1a", bg="#ffffff")
            self.config_frame.pack(fill=tk.X)
        else:
            self.reading_label.configure(text="---", fg="#aaaaaa", bg="#e8e8e8")
            self.config_frame.pack_forget()

    def get_range_value(self) -> float:
        mode = MeasureMode[self.mode_var.get()]
        ranges = RANGES.get(mode, [])
        selected = self.range_var.get()
        for label, val in ranges:
            if label == selected:
                return val
        return 10.0

    def get_config(self) -> TerminalConfig:
        """Build a TerminalConfig from the current widget values."""
        guard_name = self.guard_var.get().upper()
        power_name = self.power_var.get().upper()
        return TerminalConfig(
            enabled=self.enabled_var.get(),
            mode=MeasureMode[self.mode_var.get()],
            range_value=self.get_range_value(),
            nplc=float(self.nplc_var.get()),
            guard=Guard.FLOAT if guard_name == "FLOAT" else Guard.CABLE,
            resistance_power=ResistancePower.HIGH if power_name == "HIGH" else ResistancePower.LOW,
            ocomp=OcompState.ON if self.ocomp_var.get() == "ON" else OcompState.OFF,
        )

    def _do_apply(self) -> None:
        if self._on_apply is not None:
            self._on_apply(self)

    def set_controls_enabled(self, enabled: bool) -> None:
        state = "readonly" if enabled else tk.DISABLED
        for w in (self.mode_combo, self.range_combo, self.nplc_combo,
                  self.guard_combo, self.power_combo, self.ocomp_combo):
            w.configure(state=state)
        self.check.configure(state=tk.NORMAL if enabled else tk.DISABLED)
        self.apply_btn.configure(state=tk.NORMAL if enabled else tk.DISABLED)


# ======================================================================
# Main window
# ======================================================================

class MainWindow(tk.Tk):
    """Top-level tkinter window with dual front/rear reading display."""

    def __init__(self) -> None:
        super().__init__()
        self.title("R6581T Data Logger")
        self.minsize(960, 780)
        self.configure(bg="#f0f0f0")
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # State
        self._meter: Optional[R6581T] = None
        self._worker: Optional[MeasurementWorker] = None
        self._csv = CSVWriter()
        self._sample_count = 0
        self._result_queue: queue.Queue[TerminalReadResult] = queue.Queue()

        # Tkinter variables (connection & logging only — measurement config is per-panel)
        self._resource_var = tk.StringVar()
        self._csv_path_var = tk.StringVar()
        self._interval_var = tk.StringVar(value="5.0")
        self._status_var = tk.StringVar(value="Ready — connect to an instrument to begin.")

        self._configure_styles()
        self._build_ui()
        self._populate_defaults()
        self._poll_results()

    # ==================================================================
    # Styles
    # ==================================================================

    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")

        bg = "#f0f0f0"
        panel = "#ffffff"
        field = "#ffffff"
        border = "#cccccc"
        fg = "#1a1a1a"
        fg_dim = "#555555"
        accent = "#2266aa"
        green = "#2e7d32"
        red = "#c62828"

        style.configure(".", background=bg, foreground=fg, fieldbackground=field,
                         bordercolor=border, troughcolor="#e0e0e0", font=("Segoe UI", 10))
        style.configure("TFrame", background=bg)
        style.configure("TLabel", background=bg, foreground=fg, font=("Segoe UI", 10))
        style.configure("TLabelframe", background=panel, foreground=accent,
                         bordercolor=border, font=("Segoe UI", 10, "bold"))
        style.configure("TLabelframe.Label", background=panel, foreground=accent,
                         font=("Segoe UI", 10, "bold"))
        style.configure("TButton", background="#e0e0e0", foreground=fg, bordercolor=border,
                         font=("Segoe UI", 10, "bold"), padding=(12, 6))
        style.map("TButton",
                   background=[("active", "#d0d0d0"), ("disabled", "#f5f5f5")],
                   foreground=[("disabled", "#aaaaaa")])
        style.configure("Connect.TButton", background=green, foreground="#ffffff")
        style.map("Connect.TButton", background=[("active", "#388e3c")])
        style.configure("Start.TButton", background=red, foreground="#ffffff")
        style.map("Start.TButton", background=[("active", "#d32f2f")])
        style.configure("TCombobox", fieldbackground=field, background=field,
                         foreground=fg, selectbackground=accent, selectforeground="#ffffff")
        style.configure("TEntry", fieldbackground=field, foreground=fg)
        style.configure("TCheckbutton", background=panel, foreground=fg, font=("Segoe UI", 10))
        style.configure("Section.TLabel", background=bg, foreground=red,
                         font=("Segoe UI", 9, "bold"))
        style.configure("Status.TLabel", background="#e8e8e8", foreground=fg_dim,
                         font=("Segoe UI", 9))

        self.option_add("*TCombobox*Listbox.background", field)
        self.option_add("*TCombobox*Listbox.foreground", fg)
        self.option_add("*TCombobox*Listbox.selectBackground", accent)
        self.option_add("*TCombobox*Listbox.selectForeground", "#ffffff")

    # ==================================================================
    # UI construction
    # ==================================================================

    def _build_ui(self) -> None:
        root_frame = ttk.Frame(self, padding=12)
        root_frame.pack(fill=tk.BOTH, expand=True)

        # --- Dual terminal panels (top) ---
        panels_frame = ttk.Frame(root_frame)
        panels_frame.pack(fill=tk.X, pady=(0, 10))
        panels_frame.columnconfigure(0, weight=1, uniform="panel")
        panels_frame.columnconfigure(1, weight=1, uniform="panel")

        self._front_panel = TerminalPanel(
            panels_frame, "FRONT", enabled_default=True,
            on_toggle=self._on_terminal_toggle,
            on_apply=self._on_apply_panel,
        )
        self._front_panel.frame.grid(row=0, column=0, sticky=tk.NSEW, padx=(0, 6))

        self._rear_panel = TerminalPanel(
            panels_frame, "REAR", enabled_default=False,
            on_toggle=self._on_terminal_toggle,
            on_apply=self._on_apply_panel,
        )
        self._rear_panel.frame.grid(row=0, column=1, sticky=tk.NSEW, padx=(6, 0))

        # --- Connection + Logging (middle) ---
        mid_frame = ttk.Frame(root_frame)
        mid_frame.pack(fill=tk.X, pady=(0, 10))

        self._build_connection_group(mid_frame)
        self._build_logging_group(mid_frame)

        # --- Log (bottom) ---
        log_group = ttk.LabelFrame(root_frame, text="Measurement Log", padding=8)
        log_group.pack(fill=tk.BOTH, expand=True)

        self._log_text = tk.Text(
            log_group, height=8, font=("Consolas", 9),
            bg="#ffffff", fg="#1a1a1a", bd=1, relief="solid",
            insertbackground="#1a1a1a", highlightbackground="#cccccc",
            wrap=tk.WORD, state=tk.DISABLED,
        )
        log_scroll = ttk.Scrollbar(log_group, command=self._log_text.yview)
        self._log_text.configure(yscrollcommand=log_scroll.set)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self._log_text.pack(fill=tk.BOTH, expand=True)

        # --- Status bar ---
        status_bar = ttk.Label(
            self, textvariable=self._status_var, style="Status.TLabel", padding=(8, 4),
        )
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

    def _build_connection_group(self, parent: ttk.Frame) -> None:
        group = ttk.LabelFrame(parent, text="Connection", padding=8)
        group.pack(side=tk.LEFT, fill=tk.BOTH, padx=(0, 8))

        ttk.Label(group, text="VISA Resource:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self._resource_combo = ttk.Combobox(
            group, textvariable=self._resource_var, width=28,
        )
        self._resource_combo.grid(row=0, column=1, padx=4, pady=2, sticky=tk.EW)

        self._refresh_btn = ttk.Button(group, text="Refresh", command=self._refresh_resources)
        self._refresh_btn.grid(row=0, column=2, padx=4, pady=2)

        # Spacer to push button to bottom
        spacer = ttk.Frame(group)
        spacer.grid(row=1, column=0, columnspan=3, sticky=tk.NSEW)
        group.rowconfigure(1, weight=1)

        self._connect_btn = ttk.Button(
            group, text="Connect", style="Connect.TButton",
            command=self._toggle_connection,
        )
        self._connect_btn.grid(row=2, column=0, columnspan=3, sticky=tk.EW + tk.S, pady=(6, 0))

        group.columnconfigure(1, weight=1)

    def _build_logging_group(self, parent: ttk.Frame) -> None:
        group = ttk.LabelFrame(parent, text="Data Logging", padding=8)
        group.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        ttk.Label(group, text="CSV File:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self._csv_entry = ttk.Entry(group, textvariable=self._csv_path_var, width=30)
        self._csv_entry.grid(row=0, column=1, padx=4, pady=2, sticky=tk.EW)
        self._browse_btn = ttk.Button(group, text="Browse...", command=self._browse_csv)
        self._browse_btn.grid(row=0, column=2, padx=4, pady=2)

        ttk.Label(group, text="Interval (s):").grid(row=1, column=0, sticky=tk.W, pady=2)
        self._interval_entry = ttk.Entry(group, textvariable=self._interval_var, width=10)
        self._interval_entry.grid(row=1, column=1, padx=4, pady=2, sticky=tk.W)

        # Spacer to push button to bottom
        spacer = ttk.Frame(group)
        spacer.grid(row=2, column=0, columnspan=3, sticky=tk.NSEW)
        group.rowconfigure(2, weight=1)

        self._start_btn = ttk.Button(
            group, text="Start Logging", style="Start.TButton",
            command=self._toggle_logging, state=tk.DISABLED,
        )
        self._start_btn.grid(row=3, column=0, columnspan=3, sticky=tk.EW + tk.S, pady=(6, 0))

        group.columnconfigure(1, weight=1)

    # ==================================================================
    # Terminal toggle
    # ==================================================================

    def _on_terminal_toggle(self) -> None:
        # Must have at least one enabled
        if not self._front_panel.enabled_var.get() and not self._rear_panel.enabled_var.get():
            self._front_panel.enabled_var.set(True)

        self._front_panel.update_appearance()
        self._rear_panel.update_appearance()

    # ==================================================================
    # Apply configuration
    # ==================================================================

    def _on_apply_panel(self, panel: TerminalPanel) -> None:
        """Apply a single panel's config to the meter (switches terminal first)."""
        if not self._meter or not self._meter.is_connected:
            messagebox.showwarning("Not Connected", "Connect to an instrument first.")
            return

        terminal = Terminal.FRONT if panel is self._front_panel else Terminal.REAR
        cfg = panel.get_config()

        try:
            self._meter.set_terminal(terminal)
            self._meter.configure(
                mode=cfg.mode,
                range_value=cfg.range_value,
                nplc=cfg.nplc,
                guard=cfg.guard,
                resistance_power=cfg.resistance_power,
                ocomp=cfg.ocomp,
            )
            range_text = panel.range_var.get() or "auto"
            self._status_var.set(
                f"Applied {terminal.value}: {cfg.mode.name}  range={range_text}  NPLC={cfg.nplc}"
            )
        except Exception as exc:
            messagebox.showerror("Configuration Error", str(exc))

    # ==================================================================
    # Defaults
    # ==================================================================

    def _populate_defaults(self) -> None:
        self._refresh_resources()
        desktop = Path.home() / "Desktop"
        default_csv = desktop / f"r6581t_log_{datetime.now():%Y%m%d_%H%M%S}.csv"
        self._csv_path_var.set(str(default_csv))

    def _refresh_resources(self) -> None:
        try:
            resources = [r for r in R6581T.list_resources() if not r.endswith("::INTFC")]
        except Exception as exc:
            self._status_var.set(f"Could not list VISA resources: {exc}")
            resources = []
        self._resource_combo["values"] = resources
        if resources:
            self._resource_var.set(resources[0])

    # ==================================================================
    # Connection
    # ==================================================================

    def _toggle_connection(self) -> None:
        if self._meter and self._meter.is_connected:
            self._disconnect()
        else:
            self._do_connect()

    def _do_connect(self) -> None:
        resource = self._resource_var.get().strip()
        if not resource:
            messagebox.showwarning("No Resource", "Enter or select a VISA resource string.")
            return

        self._meter = R6581T(resource)
        try:
            self._meter.connect()
        except Exception as exc:
            messagebox.showerror("Connection Error", str(exc))
            self._meter = None
            return

        self._connect_btn.configure(text="Disconnect")
        self._start_btn.configure(state=tk.NORMAL)
        self._resource_combo.configure(state=tk.DISABLED)
        self._refresh_btn.configure(state=tk.DISABLED)
        self._status_var.set(f"Connected to {resource}")

    def _disconnect(self) -> None:
        if self._worker and self._worker.is_running:
            self._stop_logging()
        if self._meter:
            self._meter.disconnect()
            self._meter = None

        self._connect_btn.configure(text="Connect")
        self._start_btn.configure(state=tk.DISABLED)
        self._resource_combo.configure(state="readonly")
        self._refresh_btn.configure(state=tk.NORMAL)
        self._front_panel.reading_label.configure(text="---")
        self._rear_panel.reading_label.configure(text="---")
        self._status_var.set("Disconnected.")

    # ==================================================================
    # Logging
    # ==================================================================

    def _toggle_logging(self) -> None:
        if self._worker and self._worker.is_running:
            self._stop_logging()
        else:
            self._start_logging()

    def _start_logging(self) -> None:
        if not self._meter or not self._meter.is_connected:
            messagebox.showwarning("Not Connected", "Connect to an instrument first.")
            return

        csv_path = self._csv_path_var.get().strip()
        if csv_path:
            try:
                self._csv.open(csv_path)
            except Exception as exc:
                messagebox.showerror("CSV Error", f"Cannot open file:\n{exc}")
                return

        try:
            interval = float(self._interval_var.get())
        except ValueError:
            messagebox.showwarning("Invalid Interval", "Interval must be a number.")
            return

        self._sample_count = 0

        front_cfg = self._front_panel.get_config()
        rear_cfg = self._rear_panel.get_config()

        self._worker = MeasurementWorker(
            meter=self._meter,
            interval_s=interval,
            front_config=front_cfg,
            rear_config=rear_cfg,
            on_reading=self._enqueue_reading,
        )
        self._worker.start()

        self._start_btn.configure(text="Stop Logging")
        self._set_controls_enabled(False)

        terminals = []
        if front_cfg.enabled:
            terminals.append(f"FRONT({front_cfg.mode.name})")
        if rear_cfg.enabled:
            terminals.append(f"REAR({rear_cfg.mode.name})")

        self._append_log(
            f"--- Logging started at {datetime.now():%H:%M:%S} ---\n"
            f"    {' | '.join(terminals)}  Interval: {interval}s"
        )
        self._status_var.set("Logging... waiting for first reading")

    def _stop_logging(self) -> None:
        if self._worker:
            self._worker.stop()
            self._worker = None

        try:
            while True:
                result = self._result_queue.get_nowait()
                self._on_reading(result)
        except queue.Empty:
            pass

        self._csv.close()

        self._start_btn.configure(text="Start Logging")
        self._set_controls_enabled(True)
        self._append_log(f"--- Logging stopped at {datetime.now():%H:%M:%S} ---")
        self._status_var.set(f"Stopped. {self._sample_count} samples collected.")

    def _set_controls_enabled(self, enabled: bool) -> None:
        entry_state = tk.NORMAL if enabled else tk.DISABLED
        self._front_panel.set_controls_enabled(enabled)
        self._rear_panel.set_controls_enabled(enabled)
        self._csv_entry.configure(state=entry_state)
        self._browse_btn.configure(state=entry_state)
        self._interval_entry.configure(state=entry_state)

    # ==================================================================
    # Result queue
    # ==================================================================

    def _enqueue_reading(self, result: TerminalReadResult) -> None:
        self._result_queue.put(result)

    def _poll_results(self) -> None:
        try:
            while True:
                result = self._result_queue.get_nowait()
                self._on_reading(result)
        except queue.Empty:
            pass
        self.after(50, self._poll_results)

    def _on_reading(self, tr: TerminalReadResult) -> None:
        terminal = tr.terminal
        result = tr.result

        # Get the config for this terminal to know its mode/unit
        if terminal == Terminal.FRONT:
            panel = self._front_panel
        else:
            panel = self._rear_panel

        mode = MeasureMode[panel.mode_var.get()]
        unit = UNITS.get(mode, "?")
        label = panel.reading_label

        if result.success:
            self._sample_count += 1
            is_overload = abs(result.value) > 9e+37

            if is_overload:
                display = "OVERLOAD"
                log_fmt = "OVERLOAD"
                label.configure(text=display, fg="#c62828")
            else:
                display = f"{result.value:+.9f} {unit}"
                log_fmt = f"{result.value:+.10E} {unit}"
                label.configure(text=display, fg="#2e7d32")

            ts = datetime.now().strftime("%H:%M:%S")
            tag = "F" if terminal == Terminal.FRONT else "R"
            self._append_log(f"[{self._sample_count:>5}] {ts} [{tag}] {log_fmt}")

            if self._csv.is_open:
                range_text = panel.range_var.get() or "auto"
                try:
                    self._csv.write_row(
                        value=result.value,
                        mode=mode.name,
                        range_label=range_text,
                        nplc=float(panel.nplc_var.get()),
                        unit=unit,
                        terminal=terminal.value,
                        overload=is_overload,
                    )
                except (ValueError, OSError):
                    pass

            self._status_var.set(f"Sample #{self._sample_count} [{tag}] — {display}")
        else:
            label.configure(text="ERR", fg="#c62828")
            ts = datetime.now().strftime("%H:%M:%S")
            tag = "F" if terminal == Terminal.FRONT else "R"
            self._append_log(f"[  ERR] {ts} [{tag}] {result.error}")

    # ==================================================================
    # Helpers
    # ==================================================================

    def _append_log(self, text: str) -> None:
        self._log_text.configure(state=tk.NORMAL)
        self._log_text.insert(tk.END, text + "\n")
        self._log_text.see(tk.END)
        self._log_text.configure(state=tk.DISABLED)

    def _browse_csv(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Select CSV Output File",
            initialdir=str(Path.home() / "Desktop"),
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")],
        )
        if path:
            self._csv_path_var.set(path)

    def _on_close(self) -> None:
        if self._worker and self._worker.is_running:
            self._stop_logging()
        if self._meter and self._meter.is_connected:
            self._meter.disconnect()
        self.destroy()