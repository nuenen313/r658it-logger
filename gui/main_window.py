"""
Main application window for the R6581T data logger (tkinter version).
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
    Guard, ResistancePower, OcompState,
    ReadResult,
)
from data.csv_writer import CSVWriter
from gui.worker import MeasurementWorker


class MainWindow(tk.Tk):
    """Top-level tkinter window containing all controls and the reading log."""

    def __init__(self) -> None:
        super().__init__()
        self.title("R6581T Data Logger")
        self.minsize(920, 720)
        self.configure(bg="#f0f0f0")
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # State
        self._meter: Optional[R6581T] = None
        self._worker: Optional[MeasurementWorker] = None
        self._csv = CSVWriter()
        self._sample_count = 0
        self._result_queue: queue.Queue[ReadResult] = queue.Queue()

        # Tkinter variables
        self._resource_var = tk.StringVar()
        self._mode_var = tk.StringVar(value="DCV")
        self._range_var = tk.StringVar()
        self._nplc_var = tk.StringVar(value="100")
        self._guard_var = tk.StringVar(value="Float")
        self._power_var = tk.StringVar(value="High")
        self._ocomp_var = tk.StringVar(value="OFF")
        self._csv_path_var = tk.StringVar()
        self._interval_var = tk.StringVar(value="5.0")
        self._status_var = tk.StringVar(value="Ready — connect to an instrument to begin.")

        self._configure_styles()
        self._build_ui()
        self._populate_defaults()

        # Start polling the result queue
        self._poll_results()

    # ==================================================================
    # Styles
    # ==================================================================

    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")

        # Light color palette
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
        style.configure("Section.TLabel", background=bg, foreground=red,
                         font=("Segoe UI", 9, "bold"))
        style.configure("Status.TLabel", background="#e8e8e8", foreground=fg_dim,
                         font=("Segoe UI", 9))

        # Combobox dropdown colors
        self.option_add("*TCombobox*Listbox.background", field)
        self.option_add("*TCombobox*Listbox.foreground", fg)
        self.option_add("*TCombobox*Listbox.selectBackground", accent)
        self.option_add("*TCombobox*Listbox.selectForeground", "#ffffff")

    # ==================================================================
    # UI construction
    # ==================================================================

    def _build_ui(self) -> None:
        root_frame = ttk.Frame(self, padding=16)
        root_frame.pack(fill=tk.BOTH, expand=True)

        # --- Reading display ---
        self._reading_label = tk.Label(
            root_frame, text="---", font=("Consolas", 36, "bold"),
            fg="#1a1a1a", bg="#ffffff", bd=2, relief="groove",
            padx=20, pady=18, anchor="center",
        )
        self._reading_label.pack(fill=tk.X, pady=(0, 12))

        # --- Controls row ---
        controls = ttk.Frame(root_frame)
        controls.pack(fill=tk.X, pady=(0, 12))

        # Left column
        left = ttk.Frame(controls)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 12))
        self._build_connection_group(left)
        self._build_logging_group(left)

        # Right column
        right = ttk.Frame(controls)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._build_measurement_group(right)

        # --- Log ---
        log_frame = ttk.LabelFrame(root_frame, text="Measurement Log", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self._log_text = tk.Text(
            log_frame, height=10, font=("Consolas", 9),
            bg="#ffffff", fg="#1a1a1a", bd=1, relief="solid",
            insertbackground="#1a1a1a", highlightbackground="#cccccc",
            wrap=tk.WORD, state=tk.DISABLED,
        )
        log_scroll = ttk.Scrollbar(log_frame, command=self._log_text.yview)
        self._log_text.configure(yscrollcommand=log_scroll.set)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self._log_text.pack(fill=tk.BOTH, expand=True)

        # --- Status bar ---
        status_bar = ttk.Label(
            self, textvariable=self._status_var, style="Status.TLabel",
            padding=(8, 4),
        )
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

    # ------------------------------------------------------------------

    def _build_connection_group(self, parent: ttk.Frame) -> None:
        group = ttk.LabelFrame(parent, text="Connection", padding=8)
        group.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(group, text="VISA Resource:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self._resource_combo = ttk.Combobox(
            group, textvariable=self._resource_var, width=28,
        )
        self._resource_combo.grid(row=0, column=1, padx=4, pady=2, sticky=tk.EW)

        self._refresh_btn = ttk.Button(group, text="Refresh", command=self._refresh_resources)
        self._refresh_btn.grid(row=0, column=2, padx=4, pady=2)

        self._connect_btn = ttk.Button(
            group, text="Connect", style="Connect.TButton",
            command=self._toggle_connection,
        )
        self._connect_btn.grid(row=1, column=0, columnspan=3, sticky=tk.EW, pady=(6, 0))

        group.columnconfigure(1, weight=1)

    # ------------------------------------------------------------------

    def _build_measurement_group(self, parent: ttk.Frame) -> None:
        group = ttk.LabelFrame(parent, text="Measurement Setup", padding=8)
        group.pack(fill=tk.BOTH, expand=True)
        row = 0

        # Mode
        ttk.Label(group, text="Mode:").grid(row=row, column=0, sticky=tk.W, pady=2)
        self._mode_combo = ttk.Combobox(
            group, textvariable=self._mode_var, width=20,
            values=[m.name for m in MeasureMode], state="readonly",
        )
        self._mode_combo.grid(row=row, column=1, sticky=tk.EW, padx=4, pady=2)
        self._mode_combo.bind("<<ComboboxSelected>>", lambda e: self._on_mode_changed())
        row += 1

        # Range
        self._range_label = ttk.Label(group, text="Range:")
        self._range_label.grid(row=row, column=0, sticky=tk.W, pady=2)
        self._range_combo = ttk.Combobox(
            group, textvariable=self._range_var, width=20, state="readonly",
        )
        self._range_combo.grid(row=row, column=1, sticky=tk.EW, padx=4, pady=2)
        row += 1

        # NPLC
        ttk.Label(group, text="NPLC:").grid(row=row, column=0, sticky=tk.W, pady=2)
        self._nplc_combo = ttk.Combobox(
            group, textvariable=self._nplc_var, width=20,
            values=[str(n) for n in NPLC_VALUES], state="readonly",
        )
        self._nplc_combo.grid(row=row, column=1, sticky=tk.EW, padx=4, pady=2)
        row += 1

        # Guard
        ttk.Label(group, text="Guard:").grid(row=row, column=0, sticky=tk.W, pady=2)
        self._guard_combo = ttk.Combobox(
            group, textvariable=self._guard_var, width=20,
            values=[g.name.capitalize() for g in Guard], state="readonly",
        )
        self._guard_combo.grid(row=row, column=1, sticky=tk.EW, padx=4, pady=2)
        row += 1

        # --- Resistance options ---
        self._res_sep = ttk.Label(group, text="Resistance Options", style="Section.TLabel")
        self._res_sep.grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=(8, 2))
        row += 1

        self._power_label = ttk.Label(group, text="Source Power:")
        self._power_label.grid(row=row, column=0, sticky=tk.W, pady=2)
        self._power_combo = ttk.Combobox(
            group, textvariable=self._power_var, width=20,
            values=[p.name.capitalize() for p in ResistancePower], state="readonly",
        )
        self._power_combo.grid(row=row, column=1, sticky=tk.EW, padx=4, pady=2)
        row += 1

        self._ocomp_label = ttk.Label(group, text="Offset Comp:")
        self._ocomp_label.grid(row=row, column=0, sticky=tk.W, pady=2)
        self._ocomp_combo = ttk.Combobox(
            group, textvariable=self._ocomp_var, width=20,
            values=[o.value for o in OcompState], state="readonly",
        )
        self._ocomp_combo.grid(row=row, column=1, sticky=tk.EW, padx=4, pady=2)
        row += 1

        # Apply button
        self._apply_btn = ttk.Button(
            group, text="Apply Configuration", command=self._apply_configuration,
        )
        self._apply_btn.grid(row=row, column=0, columnspan=2, sticky=tk.EW, pady=(8, 0))

        group.columnconfigure(1, weight=1)

    # ------------------------------------------------------------------

    def _build_logging_group(self, parent: ttk.Frame) -> None:
        group = ttk.LabelFrame(parent, text="Data Logging", padding=8)
        group.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(group, text="CSV File:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self._csv_entry = ttk.Entry(group, textvariable=self._csv_path_var, width=24)
        self._csv_entry.grid(row=0, column=1, padx=4, pady=2, sticky=tk.EW)
        self._browse_btn = ttk.Button(group, text="Browse...", command=self._browse_csv)
        self._browse_btn.grid(row=0, column=2, padx=4, pady=2)

        ttk.Label(group, text="Interval (s):").grid(row=1, column=0, sticky=tk.W, pady=2)
        self._interval_entry = ttk.Entry(group, textvariable=self._interval_var, width=10)
        self._interval_entry.grid(row=1, column=1, padx=4, pady=2, sticky=tk.W)

        self._start_btn = ttk.Button(
            group, text="Start Logging", style="Start.TButton",
            command=self._toggle_logging, state=tk.DISABLED,
        )
        self._start_btn.grid(row=2, column=0, columnspan=3, sticky=tk.EW, pady=(6, 0))

        group.columnconfigure(1, weight=1)

    # ==================================================================
    # Defaults & conditional visibility
    # ==================================================================

    def _populate_defaults(self) -> None:
        self._on_mode_changed()
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

    def _on_mode_changed(self) -> None:
        mode_name = self._mode_var.get()
        mode = MeasureMode[mode_name]

        # Range
        ranges = RANGES.get(mode, [])
        if ranges:
            labels = [r[0] for r in ranges]
            self._range_combo["values"] = labels
            self._range_combo.set(labels[len(labels) // 2])
            self._range_label.grid()
            self._range_combo.grid()
        else:
            self._range_combo["values"] = []
            self._range_combo.set("")
            self._range_label.grid_remove()
            self._range_combo.grid_remove()

        # Resistance options (source power, ocomp, guard — for 4W only)
        is_res4w = mode == MeasureMode.RES4W
        for w in (self._res_sep, self._power_label, self._power_combo,
                  self._ocomp_label, self._ocomp_combo):
            w.grid() if is_res4w else w.grid_remove()

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
        self._apply_configuration()

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
        self._reading_label.configure(text="---")
        self._status_var.set("Disconnected.")

    # ==================================================================
    # Configuration
    # ==================================================================

    def _get_selected_mode(self) -> MeasureMode:
        return MeasureMode[self._mode_var.get()]

    def _get_selected_range_value(self) -> float:
        mode = self._get_selected_mode()
        ranges = RANGES.get(mode, [])
        selected_label = self._range_var.get()
        for label, val in ranges:
            if label == selected_label:
                return val
        return 10.0

    def _apply_configuration(self) -> None:
        if not self._meter or not self._meter.is_connected:
            return

        mode = self._get_selected_mode()
        range_val = self._get_selected_range_value()
        nplc = float(self._nplc_var.get())

        guard_name = self._guard_var.get().upper()
        guard = Guard.FLOAT if guard_name == "FLOAT" else Guard.CABLE

        power_name = self._power_var.get().upper()
        power = ResistancePower.HIGH if power_name == "HIGH" else ResistancePower.LOW

        ocomp = OcompState.ON if self._ocomp_var.get() == "ON" else OcompState.OFF

        try:
            self._meter.configure(
                mode=mode, range_value=range_val, nplc=nplc,
                guard=guard, resistance_power=power, ocomp=ocomp,
            )
            range_text = self._range_var.get() or "auto"
            self._status_var.set(
                f"Configured: {mode.name}  range={range_text}  NPLC={nplc}  guard={guard.name}"
            )
        except Exception as exc:
            messagebox.showerror("Configuration Error", str(exc))

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

        # Apply current configuration before starting
        self._apply_configuration()

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

        self._worker = MeasurementWorker(
            meter=self._meter,
            interval_s=interval,
            on_reading=self._enqueue_reading,
        )
        self._worker.start()

        self._start_btn.configure(text="Stop Logging")
        self._set_controls_enabled(False)
        mode = self._get_selected_mode()
        range_text = self._range_var.get() or "auto"
        nplc_text = self._nplc_var.get()
        self._append_log(
            f"--- Logging started at {datetime.now():%H:%M:%S} ---\n"
            f"    Mode: {mode.name}  Range: {range_text}  NPLC: {nplc_text}  "
            f"Interval: {interval}s"
        )
        self._status_var.set("Logging... waiting for first reading")

    def _stop_logging(self) -> None:
        if self._worker:
            self._worker.stop()
            self._worker = None

        # Drain any remaining results before closing CSV
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
        state = "readonly" if enabled else tk.DISABLED
        entry_state = tk.NORMAL if enabled else tk.DISABLED

        for w in (self._mode_combo, self._range_combo, self._nplc_combo,
                  self._guard_combo, self._power_combo, self._ocomp_combo):
            w.configure(state=state)

        self._apply_btn.configure(state=entry_state)
        self._csv_entry.configure(state=entry_state)
        self._browse_btn.configure(state=entry_state)
        self._interval_entry.configure(state=entry_state)

    # ==================================================================
    # Result queue (thread-safe bridge from worker to GUI)
    # ==================================================================

    def _enqueue_reading(self, result: ReadResult) -> None:
        """Called from the worker thread — puts result in queue for the main thread."""
        self._result_queue.put(result)

    def _poll_results(self) -> None:
        """Called periodically on the main thread to process queued results."""
        try:
            while True:
                result = self._result_queue.get_nowait()
                self._on_reading(result)
        except queue.Empty:
            pass
        self.after(50, self._poll_results)

    def _on_reading(self, result: ReadResult) -> None:
        mode = self._get_selected_mode()
        unit = UNITS.get(mode, "?")

        if result.success:
            self._sample_count += 1

            # 9.9E+37 is the R6581T's overload indicator
            is_overload = abs(result.value) > 9e+37

            if is_overload:
                display = "OVERLOAD"
                log_fmt = "OVERLOAD"
                self._reading_label.configure(text=display, fg="#c62828")
            else:
                display = f"{result.value:+.9f} {unit}"
                log_fmt = f"{result.value:+.10E} {unit}"
                self._reading_label.configure(text=display, fg="#2e7d32")

            ts = datetime.now().strftime("%H:%M:%S")
            self._append_log(f"[{self._sample_count:>5}] {ts}   {log_fmt}")

            if self._csv.is_open:
                range_text = self._range_var.get() or "auto"
                try:
                    self._csv.write_row(
                        value=result.value,
                        mode=mode.name,
                        range_label=range_text,
                        nplc=float(self._nplc_var.get()),
                        unit=unit,
                        overload=is_overload,
                    )
                except (ValueError, OSError):
                    pass  # File was closed between check and write

            self._status_var.set(f"Sample #{self._sample_count} — {display}")
        else:
            self._reading_label.configure(text="ERR", fg="#c62828")
            ts = datetime.now().strftime("%H:%M:%S")
            self._append_log(f"[  ERR] {ts}   {result.error}")

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