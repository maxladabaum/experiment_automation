"""
gui/tab_method.py — Method Creation tab.

Responsible for:
  - CV / SWV / Pause parameter forms
  - MethodSCRIPT generation (delegates to core.mscript_parser helpers)
  - MUX16 channel parsing and script wrapping
  - "Add to Queue", "Generate Script", "Run Now" actions
  - Calling back into QueueTab to add items and refresh display
"""

import threading
import time
import math
from pathlib import Path
from tkinter import messagebox
import tkinter as tk
from tkinter import ttk
import serial.tools.list_ports

from core.mscript_parser import to_si_string
from core.runner import SerialMeasurementRunner, format_port_info
from config import DEVICE_KEYWORDS
from core.session import SessionState
from gui.tab_custom_script import CustomScriptPanel

EMSTAT_PICO_LOW_SPEED_FIXED_BA_RANGES = (
    ("100 nA", "100n"),
    ("2 uA", "2u"),
    ("4 uA", "4u"),
    ("8 uA", "8u"),
    ("16 uA", "16u"),
    ("32 uA", "32u"),
    ("63 uA", "63u"),
    ("125 uA", "125u"),
    ("250 uA", "250u"),
    ("500 uA", "500u"),
    ("1 mA", "1m"),
    ("5 mA", "5m"),
)

EMSTAT_PICO_LOW_SPEED_AUTORANGE_BA_RANGES = (
    ("1 nA", "1n"),
    ("100 nA", "100n"),
    ("2 uA", "2u"),
    ("4 uA", "4u"),
    ("8 uA", "8u"),
    ("16 uA", "16u"),
    ("32 uA", "32u"),
    ("63 uA", "63u"),
    ("100 uA", "100u"),
    ("125 uA", "125u"),
    ("250 uA", "250u"),
    ("500 uA", "500u"),
    ("1 mA", "1m"),
    ("5 mA", "5m"),
)

EMSTAT_PICO_HIGH_SPEED_BA_RANGES = (
    ("100 nA", "59n"),
    ("1 uA", "590n"),
    ("6 uA", "3687500p"),
    ("13 uA", "7375n"),
    ("25 uA", "14750n"),
    ("50 uA", "29500n"),
    ("100 uA", "59u"),
    ("200 uA", "118u"),
    ("1 mA", "590u"),
    ("5 mA", "2950u"),
)

class MethodTab:
    """Manages the 'Method Creation' notebook tab.

    Parameters
    ----------
    parent_frame:
        The ``ttk.Frame`` added to the notebook for this tab.
    session:
        Shared :class:`~core.session.SessionState`.
    on_add_to_queue:
        Callable ``(item: dict) → None`` provided by QueueTab so adding
        items doesn't require a direct reference to QueueTab.
    on_refresh_queue:
        Callable ``() → None`` — triggers a queue display refresh.
    on_script_preview:
        Callable ``(script: str) → None`` — pushes generated script to
        the Script Preview tab.
    on_run_now:
        Callable ``(technique, script, mux_channel) → None`` — triggers
        an immediate run (handled by app.py which has access to all tabs).
    """

    def __init__(
        self,
        parent_frame:      ttk.Frame,
        session:           SessionState,
        on_add_to_queue,
        on_refresh_queue,
        on_script_preview,
        on_run_now,
    ):
        self._frame            = parent_frame
        self._session          = session
        self._add_to_queue     = on_add_to_queue
        self._refresh_queue    = on_refresh_queue
        self._script_preview   = on_script_preview
        self._run_now          = on_run_now

        self.current_technique = "CV"
        self.cv_params:  dict  = {}
        self.lsv_params: dict  = {}
        self.swv_params: dict  = {}
        self.alignment_params: dict = {}
        self.pause_params: dict = {}
        self._library_note = tk.StringVar(value="")
        self._cv_range_mode = tk.StringVar(value="auto")
        self._cv_fixed_range = tk.StringVar(value="125 uA")
        self._cv_auto_min = tk.StringVar(value="1 nA")
        self._cv_auto_max = tk.StringVar(value="100 uA")
        self._lsv_range_mode = tk.StringVar(value="fixed")
        self._lsv_fixed_range = tk.StringVar(value="16 uA")
        self._lsv_auto_min = tk.StringVar(value="100 nA")
        self._lsv_auto_max = tk.StringVar(value="16 uA")
        self._swv_range_mode = tk.StringVar(value="auto")
        self._swv_fixed_range = tk.StringVar(value="100 nA")
        self._swv_auto_min = tk.StringVar(value="100 nA")
        self._swv_auto_max = tk.StringVar(value="25 uA")
        self._swv_bandwidth = tk.StringVar(value="4k")
        self._device_port_var = tk.StringVar(value="Auto (detect)")
        self._device_port_choices = []
        self._device_port_info_by_device = {}

        self._build()

    # ── Build ─────────────────────────────────────────────────────────────────

    def _build(self):
        left = ttk.Frame(self._frame)
        left.pack(side="left", fill="both", expand=True, padx=5)

        ttk.Label(left, text="Select Technique:", font=("Arial", 12, "bold")).pack(pady=5)
        tech_frame = ttk.Frame(left)
        tech_frame.pack(pady=10)
        ttk.Button(tech_frame, text="Cyclic Voltammetry (CV)",
                   command=self._show_cv_params, width=28).pack(pady=5)
        ttk.Button(tech_frame, text="Linear Sweep Voltammetry (LSV)",
                   command=self._show_lsv_params, width=28).pack(pady=5)
        ttk.Button(tech_frame, text="Square Wave Voltammetry (SWV)",
           command=self._show_swv_params, width=28).pack(pady=5)
        ttk.Button(tech_frame, text="Custom Script (File)",
                command=self._show_custom_params, width=28).pack(pady=5)
        ttk.Separator(tech_frame, orient="horizontal").pack(fill="x", pady=6)
        ttk.Button(tech_frame, text="PStrace SWV Preset",
                command=self._run_pstrace_preset, width=28).pack(pady=5)
        ttk.Button(tech_frame, text="Alignment Test (Dry)",
                command=self._show_alignment_params, width=28).pack(pady=5)
        ttk.Separator(tech_frame, orient="horizontal").pack(fill="x", pady=6)
        ttk.Button(tech_frame, text="Pause / Alert",
                   command=self._show_pause_params, width=28).pack(pady=5)

        self._device_status = ttk.Label(left, text="", foreground="blue")
        self._device_status.pack(pady=10)
        ttk.Button(left, text="Check Device Connection",
                   command=self._check_device).pack(pady=5)
        device_pick = ttk.Frame(left)
        device_pick.pack(pady=6, fill="x")
        ttk.Label(device_pick, text="Device port:").pack(side="left")
        self._device_port_box = ttk.Combobox(
            device_pick,
            textvariable=self._device_port_var,
            values=[],
            state="readonly",
            width=45,
        )
        self._device_port_box.pack(side="left", padx=6, fill="x", expand=True)
        self._device_port_box.bind("<<ComboboxSelected>>", self._on_device_port_selected)
        ttk.Button(device_pick, text="Refresh",
                   command=self._refresh_device_ports).pack(side="left")

        # Execution options (global)
        exec_frame = ttk.LabelFrame(left, text="Execution Options")
        exec_frame.pack(fill="x", pady=(12, 0), padx=5)
        exec_frame.columnconfigure(1, weight=1)

        self._var_save_raw = tk.BooleanVar(value=self._session.save_raw_packets)
        ttk.Checkbutton(exec_frame, text="Save raw packets",
                        variable=self._var_save_raw,
                        command=self._sync_save_raw).grid(
                            row=0, column=0, columnspan=2, sticky="w", pady=2)

        self._var_sim_meas = tk.BooleanVar(value=self._session.simulate_measurements)
        ttk.Checkbutton(exec_frame, text="Simulate measurements (no device)",
                        variable=self._var_sim_meas,
                        command=self._sync_sim_meas).grid(
                            row=1, column=0, columnspan=2, sticky="w", pady=2)

        ttk.Label(exec_frame, text="Delay between steps (s):").grid(
            row=2, column=0, sticky="w", pady=2)
        self._var_step_delay = tk.StringVar(value=str(self._session.step_delay))
        delay_entry = ttk.Entry(exec_frame, width=10, textvariable=self._var_step_delay)
        delay_entry.grid(row=2, column=1, sticky="w", pady=2)
        delay_entry.bind("<Return>", self._sync_step_delay)
        delay_entry.bind("<FocusOut>", self._sync_step_delay)

        self._params_frame = ttk.LabelFrame(self._frame, text="Parameters", padding=10)
        self._params_frame.pack(side="right", fill="both", expand=True, padx=5)

        self._show_cv_params()
        self._refresh_device_ports(select_existing=True)

    def _sync_save_raw(self):
        self._session.save_raw_packets = bool(self._var_save_raw.get())

    def _sync_sim_meas(self):
        self._session.simulate_measurements = bool(self._var_sim_meas.get())

    def _sync_step_delay(self, _event=None):
        raw = self._var_step_delay.get().strip()
        if not raw:
            self._var_step_delay.set(str(self._session.step_delay))
            return
        try:
            value = float(raw)
        except ValueError:
            self._var_step_delay.set(str(self._session.step_delay))
            return
        if value < 0:
            value = 0.0
        self._session.step_delay = value
        self._var_step_delay.set(str(value))

    # ── Device check ──────────────────────────────────────────────────────────

    def _check_device(self):
        ports = self._refresh_device_ports(select_existing=True)
        if ports:
            selected = self._session.device_port or "Auto"
            if self._session.device_port is None:
                resolved = self._auto_detect_port(ports)
                if resolved:
                    selected = f"Auto -> {resolved}"
                else:
                    selected = "Auto -> (no match)"
            self._device_status.config(
                text=f"Devices found: {len(ports)} | Selected: {selected}", foreground="green"
            )
        else:
            self._device_status.config(text="No devices found", foreground="red")

    def _refresh_device_ports(self, select_existing: bool = False):
        ports = list(serial.tools.list_ports.comports())
        choices = ["Auto (detect)"]
        info_by_device = {}
        for p in ports:
            summary = format_port_info(p)
            choices.append(summary)
            info_by_device[p.device] = summary
        self._device_port_choices = choices
        self._device_port_info_by_device = info_by_device
        self._device_port_box["values"] = choices

        if select_existing:
            current = self._session.device_port
            if not current:
                self._device_port_var.set("Auto (detect)")
            else:
                match = next((c for c in choices if c.startswith(f"{current}:")), None)
                self._device_port_var.set(match or "Auto (detect)")
        else:
            if self._device_port_var.get() not in choices:
                self._device_port_var.set("Auto (detect)")

        self._on_device_port_selected()
        return ports

    def _on_device_port_selected(self, _event=None):
        sel = (self._device_port_var.get() or "").strip()
        if not sel or sel.startswith("Auto"):
            self._session.device_port = None
            return
        device = sel.split(":", 1)[0].strip()
        self._session.device_port = device or None

    @staticmethod
    def _auto_detect_port(ports):
        candidates = []
        for port in ports:
            haystack = " ".join(
                str(s) for s in (
                    getattr(port, "description", None),
                    getattr(port, "manufacturer", None),
                    getattr(port, "product", None),
                    getattr(port, "hwid", None),
                ) if s
            ).lower()
            if any(str(kw).lower() in haystack for kw in DEVICE_KEYWORDS):
                candidates.append(port.device)
        if not candidates:
            return None
        return sorted(candidates)[0]

    # ── Parameter forms ───────────────────────────────────────────────────────

    def _clear_params(self):
        for w in self._params_frame.winfo_children():
            w.destroy()

    def _show_cv_params(self):
        self._clear_params()
        self.current_technique = "CV"
        self.cv_params = {}
        params = [
            ("Begin Potential (V):",                "begin_potential", "0"),
            ("Vertex 1 (V):",                       "vertex1",         "-0.5"),
            ("Vertex 2 (V):",                       "vertex2",         "0.5"),
            ("Step Potential (V):",                 "step_potential",  "0.002"),
            ("Scan Rate (V/s):",                    "scan_rate",       "0.1"),
            ("Number of Scans:",                    "n_scans",         "1"),
            ("Conditioning Potential (V):",         "cond_potential",  "0"),
            ("Conditioning Time (s):",              "cond_time",       "0"),
            ("MUX16 Channels (1-16, 0=off):",       "mux_channel",     "0"),
        ]
        for i, (label, key, default) in enumerate(params):
            ttk.Label(self._params_frame, text=label).grid(
                row=i, column=0, sticky="w", pady=2)
            entry = ttk.Entry(self._params_frame, width=15)
            entry.insert(0, default)
            entry.grid(row=i, column=1, pady=2)
            self.cv_params[key] = entry

        self._build_ba_range_controls(
            row=len(params),
            technique="CV",
            title="Current Range (EmStat Pico low-speed mode / PSTrace labels)",
        )

        ttk.Label(self._params_frame, text="Library note (optional):").grid(
            row=len(params) + 1, column=0, sticky="w", pady=(8, 2))
        ttk.Entry(self._params_frame, width=40, textvariable=self._library_note).grid(
            row=len(params) + 1, column=1, sticky="w", pady=2)

        btn_frame = ttk.Frame(self._params_frame)
        btn_frame.grid(row=len(params) + 2, column=0, columnspan=2, pady=20)
        ttk.Button(btn_frame, text="Generate Script",
                   command=self._generate_cv_script).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Save to Library",
                   command=self._save_cv_to_library).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Run Now",
                   command=self._run_cv_now).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Add to Queue",
                   command=self._add_cv_to_queue).pack(side="left", padx=5)

    def _show_lsv_params(self):
        self._clear_params()
        self.current_technique = "LSV"
        self.lsv_params = {}
        params = [
            ("Begin Potential (V):",                "begin_potential", "-0.7"),
            ("End Potential (V):",                  "end_potential",   "-1.0"),
            ("Step Potential (V):",                 "step_potential",  "0.001"),
            ("Scan Rate (V/s):",                    "scan_rate",       "0.001"),
            ("Conditioning Potential (V):",         "cond_potential",  "-0.7"),
            ("Conditioning Time (s):",              "cond_time",       "0"),
            ("MUX16 Channels (1-16, 0=off):",       "mux_channel",     "0"),
        ]
        for i, (label, key, default) in enumerate(params):
            ttk.Label(self._params_frame, text=label).grid(
                row=i, column=0, sticky="w", pady=2)
            entry = ttk.Entry(self._params_frame, width=15)
            entry.insert(0, default)
            entry.grid(row=i, column=1, pady=2)
            self.lsv_params[key] = entry

        self._build_ba_range_controls(
            row=len(params),
            technique="LSV",
            title="Current Range (EmStat Pico / PSTrace labels)",
        )

        ttk.Label(self._params_frame, text="Library note (optional):").grid(
            row=len(params) + 1, column=0, sticky="w", pady=(8, 2))
        ttk.Entry(self._params_frame, width=40, textvariable=self._library_note).grid(
            row=len(params) + 1, column=1, sticky="w", pady=2)

        btn_frame = ttk.Frame(self._params_frame)
        btn_frame.grid(row=len(params) + 2, column=0, columnspan=2, pady=20)
        ttk.Button(btn_frame, text="Generate Script",
                   command=self._generate_lsv_script).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Save to Library",
                   command=self._save_lsv_to_library).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Run Now",
                   command=self._run_lsv_now).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Add to Queue",
                   command=self._add_lsv_to_queue).pack(side="left", padx=5)

    def _show_swv_params(self):
        self._clear_params()
        self.current_technique = "SWV"
        self.swv_params = {}
        params = [
            ("Begin Potential (V):",                "begin_potential", "-0.6"),
            ("End Potential (V):",                  "end_potential",   "0"),
            ("Step Potential (V):",                 "step_potential",  "0.002"),
            ("Amplitude (V):",                      "amplitude",       "0.36"),
            ("Frequency (Hz):",                     "frequency",       "200"),
            ("Number of Scans:",                    "n_scans",         "1"),
            ("Delay Between Scans (s):",            "cycle_delay",     "0"),
            ("Conditioning Potential (V):",         "cond_potential",  "-0.6"),
            ("Conditioning Time (s):",              "cond_time",       "0.2"),
            ("MUX16 Channels (1-16, 0=off):",       "mux_channel",     "1-10"),
        ]
        for i, (label, key, default) in enumerate(params):
            ttk.Label(self._params_frame, text=label).grid(
                row=i, column=0, sticky="w", pady=2)
            entry = ttk.Entry(self._params_frame, width=15)
            entry.insert(0, default)
            entry.grid(row=i, column=1, pady=2)
            self.swv_params[key] = entry

        bandwidth_row = len(params)
        ttk.Label(self._params_frame, text="Bandwidth (EmStat Pico):").grid(
            row=bandwidth_row, column=0, sticky="w", pady=2
        )
        bw_box = ttk.Combobox(
            self._params_frame,
            textvariable=self._swv_bandwidth,
            state="readonly",
            width=12,
            values=("4k", "8k"),
        )
        bw_box.grid(row=bandwidth_row, column=1, sticky="w", pady=2)

        self._build_ba_range_controls(
            row=bandwidth_row + 1,
            technique="SWV",
            title="Current Range (EmStat Pico / PSTrace labels)",
        )

        ttk.Label(self._params_frame, text="Library note (optional):").grid(
            row=bandwidth_row + 2, column=0, sticky="w", pady=(8, 2))
        ttk.Entry(self._params_frame, width=40, textvariable=self._library_note).grid(
            row=bandwidth_row + 2, column=1, sticky="w", pady=2)

        btn_frame = ttk.Frame(self._params_frame)
        btn_frame.grid(row=bandwidth_row + 3, column=0, columnspan=2, pady=20)
        ttk.Button(btn_frame, text="Generate Script",
                   command=self._generate_swv_script).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Save to Library",
                   command=self._save_swv_to_library).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Run Now",
                   command=self._run_swv_now).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Add to Queue",
                   command=self._add_swv_to_queue).pack(side="left", padx=5)

    def _show_alignment_params(self):
        self._clear_params()
        self.current_technique = "ALIGNMENT"
        self.alignment_params = {}
        params = [
            ("DC Potential (V):",                 "dc_potential",    "0"),
            ("AC Amplitude (V rms):",             "ac_amplitude",    "0.01"),
            ("Start Frequency (Hz):",             "start_frequency", "10"),
            ("End Frequency (Hz):",               "end_frequency",   "10000"),
            ("Points per Decade:",                "points_per_decade", "4"),
            ("MUX16 Channels (1-16, 0=off):",     "mux_channel",     "0"),
        ]
        for i, (label, key, default) in enumerate(params):
            ttk.Label(self._params_frame, text=label).grid(
                row=i, column=0, sticky="w", pady=2)
            entry = ttk.Entry(self._params_frame, width=15)
            entry.insert(0, default)
            entry.grid(row=i, column=1, pady=2)
            self.alignment_params[key] = entry

        ttk.Label(
            self._params_frame,
            text="Dry alignment preset: EIS sweep for impedance/capacitance checks.",
        ).grid(row=len(params), column=0, columnspan=2, sticky="w", pady=(8, 2))

        ttk.Label(self._params_frame, text="Library note (optional):").grid(
            row=len(params) + 1, column=0, sticky="w", pady=(8, 2))
        ttk.Entry(self._params_frame, width=40, textvariable=self._library_note).grid(
            row=len(params) + 1, column=1, sticky="w", pady=2)

        btn_frame = ttk.Frame(self._params_frame)
        btn_frame.grid(row=len(params) + 2, column=0, columnspan=2, pady=20)
        ttk.Button(btn_frame, text="Generate Script",
                   command=self._generate_alignment_script).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Save to Library",
                   command=self._save_alignment_to_library).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Run Now",
                   command=self._run_alignment_now).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Add to Queue",
                   command=self._add_alignment_to_queue).pack(side="left", padx=5)

    def _show_pause_params(self):
        self._clear_params()
        self.current_technique = "PAUSE"
        self.pause_params = {}

        ttk.Label(self._params_frame, text="Pause Time (sec):").grid(
            row=0, column=0, sticky="w", pady=2)
        t_entry = ttk.Entry(self._params_frame, width=15)
        t_entry.insert(0, "10")
        t_entry.grid(row=0, column=1, pady=2)
        self.pause_params["pause_time"] = t_entry

        ttk.Label(self._params_frame, text="Alert Message:").grid(
            row=1, column=0, sticky="w", pady=2)
        a_entry = ttk.Entry(self._params_frame, width=30)
        a_entry.insert(0, "Paused — click OK to continue.")
        a_entry.grid(row=1, column=1, pady=2, sticky="w")
        self.pause_params["alert_message"] = a_entry

        btn_frame = ttk.Frame(self._params_frame)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=20)
        ttk.Button(btn_frame, text="Add Pause to Queue",
                   command=self._add_pause_to_queue).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Add Alert Pause",
                   command=self._add_alert_pause_to_queue).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Run Pause Now",
                   command=self._run_pause_now).pack(side="left", padx=5)

    # ── Script generation ─────────────────────────────────────────────────────

    @staticmethod
    def _ba_range_lines(fixed_range: str, autorange_enabled: bool, auto_min: str, auto_max: str):
        if autorange_enabled:
            return [f"set_range ba {fixed_range}", f"set_autoranging ba {auto_min} {auto_max}"]
        return [f"set_range ba {fixed_range}", f"set_autoranging ba {fixed_range} {fixed_range}"]

    @staticmethod
    def _range_labels(profile):
        return [label for label, _selector in profile]

    @staticmethod
    def _range_label_value(label: str) -> float:
        text = (label or "").strip()
        value_str, unit = text.split()
        scale = {
            "pA": 1e-12,
            "nA": 1e-9,
            "uA": 1e-6,
            "mA": 1e-3,
            "A": 1.0,
        }[unit]
        return float(value_str) * scale

    @classmethod
    def _normalize_range_label(cls, profile, label: str, direction: str) -> str:
        labels = cls._range_labels(profile)
        if label in labels:
            return label

        target = cls._range_label_value(label)
        choices = [(option_label, cls._range_label_value(option_label)) for option_label in labels]
        if direction == "down":
            eligible = [option_label for option_label, value in choices if value <= target]
            return eligible[-1] if eligible else labels[0]
        if direction == "up":
            eligible = [option_label for option_label, value in choices if value >= target]
            return eligible[0] if eligible else labels[-1]

        return min(choices, key=lambda item: abs(item[1] - target))[0]

    @staticmethod
    def _range_selector(profile, label: str) -> str:
        for option_label, selector in profile:
            if option_label == label:
                return selector
        raise ValueError(f"Unsupported current range selection: {label}")

    @staticmethod
    def _range_index(profile, label: str) -> int:
        for idx, (option_label, _selector) in enumerate(profile):
            if option_label == label:
                return idx
        raise ValueError(f"Unsupported current range selection: {label}")

    def _get_ba_range_state(self, technique: str):
        tech = technique.upper()
        if tech == "CV":
            return (
                EMSTAT_PICO_LOW_SPEED_FIXED_BA_RANGES,
                EMSTAT_PICO_LOW_SPEED_AUTORANGE_BA_RANGES,
                self._cv_range_mode,
                self._cv_fixed_range,
                self._cv_auto_min,
                self._cv_auto_max,
            )
        if tech == "LSV":
            return (
                EMSTAT_PICO_HIGH_SPEED_BA_RANGES,
                EMSTAT_PICO_HIGH_SPEED_BA_RANGES,
                self._lsv_range_mode,
                self._lsv_fixed_range,
                self._lsv_auto_min,
                self._lsv_auto_max,
            )
        if tech == "SWV":
            return (
                EMSTAT_PICO_HIGH_SPEED_BA_RANGES,
                EMSTAT_PICO_HIGH_SPEED_BA_RANGES,
                self._swv_range_mode,
                self._swv_fixed_range,
                self._swv_auto_min,
                self._swv_auto_max,
            )
        raise ValueError(f"Unsupported technique for BA range controls: {technique}")

    def _build_ba_range_controls(self, row: int, technique: str, title: str):
        fixed_profile, auto_profile, mode_var, fixed_var, auto_min_var, auto_max_var = self._get_ba_range_state(technique)
        fixed_labels = self._range_labels(fixed_profile)
        auto_labels = self._range_labels(auto_profile)

        frame = ttk.LabelFrame(self._params_frame, text=title, padding=8)
        frame.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(8, 2))
        frame.columnconfigure(1, weight=1)

        def _sync():
            fixed_state = "readonly" if mode_var.get() == "fixed" else "disabled"
            auto_state = "readonly" if mode_var.get() == "auto" else "disabled"
            fixed_box.configure(state=fixed_state)
            min_box.configure(state=auto_state)
            max_box.configure(state=auto_state)

        fixed_radio = ttk.Radiobutton(
            frame, text="Fixed", variable=mode_var, value="fixed", command=_sync
        )
        fixed_radio.grid(row=0, column=0, sticky="w")
        auto_radio = ttk.Radiobutton(
            frame, text="Autorange", variable=mode_var, value="auto", command=_sync
        )
        auto_radio.grid(row=0, column=1, sticky="w")

        ttk.Label(frame, text="Fixed range:").grid(row=1, column=0, sticky="w", pady=(6, 2))
        fixed_box = ttk.Combobox(
            frame, textvariable=fixed_var, values=fixed_labels, state="readonly", width=14
        )
        fixed_box.grid(row=1, column=1, sticky="w", pady=(6, 2))

        ttk.Label(frame, text="Autorange min:").grid(row=2, column=0, sticky="w", pady=2)
        min_box = ttk.Combobox(
            frame, textvariable=auto_min_var, values=auto_labels, state="readonly", width=14
        )
        min_box.grid(row=2, column=1, sticky="w", pady=2)

        ttk.Label(frame, text="Autorange max:").grid(row=3, column=0, sticky="w", pady=2)
        max_box = ttk.Combobox(
            frame, textvariable=auto_max_var, values=auto_labels, state="readonly", width=14
        )
        max_box.grid(row=3, column=1, sticky="w", pady=2)

        ttk.Label(
            frame,
            text="PSTrace-style labels mapped to EmStat Pico BA selectors.",
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(6, 0))

        _sync()

    def _get_ba_range_config(self, technique: str) -> dict:
        fixed_profile, auto_profile, mode_var, fixed_var, auto_min_var, auto_max_var = self._get_ba_range_state(technique)
        mode = (mode_var.get() or "fixed").strip().lower()
        fixed_label = self._normalize_range_label(fixed_profile, fixed_var.get().strip(), "up")
        auto_min_label = self._normalize_range_label(auto_profile, auto_min_var.get().strip(), "down")
        auto_max_label = self._normalize_range_label(auto_profile, auto_max_var.get().strip(), "up")

        fixed_labels = self._range_labels(fixed_profile)
        auto_labels = self._range_labels(auto_profile)
        if fixed_label not in fixed_labels:
            raise ValueError(f"Invalid fixed current range: {fixed_label or '(empty)'}")
        if auto_min_label not in auto_labels:
            raise ValueError(f"Invalid autorange minimum: {auto_min_label or '(empty)'}")
        if auto_max_label not in auto_labels:
            raise ValueError(f"Invalid autorange maximum: {auto_max_label or '(empty)'}")
        if self._range_index(auto_profile, auto_min_label) > self._range_index(auto_profile, auto_max_label):
            raise ValueError("Autorange minimum must be less than or equal to autorange maximum.")

        if mode == "auto":
            set_range_value = self._range_selector(auto_profile, auto_max_label)
            auto_min_value = self._range_selector(auto_profile, auto_min_label)
            auto_max_value = self._range_selector(auto_profile, auto_max_label)
        else:
            mode = "fixed"
            set_range_value = self._range_selector(fixed_profile, fixed_label)
            auto_min_value = set_range_value
            auto_max_value = set_range_value

        return {
            "mode": mode,
            "fixed_label": fixed_label,
            "auto_min_label": auto_min_label,
            "auto_max_label": auto_max_label,
            "set_range_value": set_range_value,
            "auto_min_value": auto_min_value,
            "auto_max_value": auto_max_value,
        }

    def _serialize_ba_range_config(self, technique: str) -> dict:
        cfg = self._get_ba_range_config(technique)
        return {
            "ba_autorange": "1" if cfg["mode"] == "auto" else "0",
            "ba_range_mode": cfg["mode"],
            "ba_fixed_range": cfg["fixed_label"],
            "ba_auto_min": cfg["auto_min_label"],
            "ba_auto_max": cfg["auto_max_label"],
        }

    def _build_cv_script(self) -> str:
        p = self.cv_params
        begin      = to_si_string(p["begin_potential"].get(), "V")
        v1         = to_si_string(p["vertex1"].get(),         "V")
        v2         = to_si_string(p["vertex2"].get(),         "V")
        step       = to_si_string(p["step_potential"].get(),  "V")
        scan_rate  = to_si_string(p["scan_rate"].get(),       "V/s")
        n_scans    = p["n_scans"].get()
        cond_pot   = to_si_string(p["cond_potential"].get(),  "V")
        cond_time  = p["cond_time"].get()
        ba_cfg     = self._get_ba_range_config("CV")

        parts = [
            "e", "var c", "var p",
            "set_pgstat_mode 2", "set_max_bandwidth 40",
        ]
        parts += self._ba_range_lines(
            ba_cfg["set_range_value"],
            ba_cfg["mode"] == "auto",
            ba_cfg["auto_min_value"],
            ba_cfg["auto_max_value"],
        )
        if float(cond_time) > 0:
            parts += [f"set_e {cond_pot}", "cell_on",
                      f"# Condition for {cond_time}s", f"wait {cond_time}"]
        else:
            parts += [f"set_e {begin}", "cell_on"]

        cv_cmd = f"meas_loop_cv p c {begin} {v1} {v2} {step} {scan_rate}"
        if int(n_scans) > 1:
            cv_cmd += f" nscans({n_scans})"
        parts += ["# CV measurement loop", cv_cmd,
                  "\tpck_start", "\tpck_add p", "\tpck_add c", "\tpck_end",
                  "endloop", "on_finished:", "cell_off"]
        return "\n".join(parts)

    def _build_lsv_script(self) -> str:
        p = self.lsv_params
        begin_v    = float(p["begin_potential"].get())
        end_v      = float(p["end_potential"].get())
        begin      = to_si_string(p["begin_potential"].get(), "V")
        end        = to_si_string(p["end_potential"].get(),   "V")
        step       = to_si_string(p["step_potential"].get(),  "V")
        scan_rate  = to_si_string(p["scan_rate"].get(),       "V/s")
        cond_pot   = to_si_string(p["cond_potential"].get(),  "V")
        cond_time  = p["cond_time"].get()
        ba_cfg     = self._get_ba_range_config("LSV")

        min_mv = int(min(begin_v, end_v) * 1000)
        max_mv = int(max(begin_v, end_v) * 1000)

        parts = [
            "e", "var c", "var p",
            "set_pgstat_chan 1",
            "set_pgstat_mode 0",
            "set_pgstat_chan 0",
            "set_pgstat_mode 2",
            "set_max_bandwidth 4",
            f"set_range_minmax da {min_mv}m {max_mv}m",
        ]
        parts += self._ba_range_lines(
            ba_cfg["set_range_value"],
            ba_cfg["mode"] == "auto",
            ba_cfg["auto_min_value"],
            ba_cfg["auto_max_value"],
        )
        if float(cond_time) > 0:
            parts += [f"set_e {cond_pot}", "cell_on",
                      f"# Condition for {cond_time}s", f"wait {cond_time}"]
            parts += [f"set_e {begin}"]
        else:
            parts += [f"set_e {begin}", "cell_on"]
        parts += [
            f"meas_loop_lsv p c {begin} {end} {step} {scan_rate}",
            "\tpck_start", "\tpck_add p", "\tpck_add c", "\tpck_end",
            "endloop", "on_finished:", "cell_off",
        ]
        return "\n".join(parts)

    def _build_swv_script(self) -> str:
        p = self.swv_params
        begin_v  = float(p["begin_potential"].get())
        end_v    = float(p["end_potential"].get())
        amp_v    = float(p["amplitude"].get())
        cond_time_s = float(p["cond_time"].get())
        freq_hz = float(p["frequency"].get())

        begin     = to_si_string(p["begin_potential"].get(), "V")
        end       = to_si_string(p["end_potential"].get(),   "V")
        step      = to_si_string(p["step_potential"].get(),  "V")
        amplitude = to_si_string(p["amplitude"].get(),       "V")
        frequency = to_si_string(p["frequency"].get(),       "Hz")
        cond_pot  = to_si_string(p["cond_potential"].get(),  "V")
        cond_time = p["cond_time"].get()
        ba_cfg    = self._get_ba_range_config("SWV")
        bandwidth = (self._swv_bandwidth.get() or "4k").strip().lower()
        if bandwidth not in ("4k", "8k"):
            raise ValueError(f"Unsupported SWV bandwidth: {bandwidth}")

        min_mv = int((min(begin_v, end_v) - amp_v) * 1000)
        max_mv = int((max(begin_v, end_v) + amp_v) * 1000)

        use_equilibrium_check = cond_time_s > 0
        eq_interval_s = min(0.2, cond_time_s) if use_equilibrium_check else 0.0
        swv_time_step = to_si_string(str(1.0 / freq_hz), "s") if freq_hz > 0 else "0"
        eq_duration = to_si_string(cond_time, "s") if use_equilibrium_check else "0"
        eq_interval = to_si_string(str(eq_interval_s), "s") if use_equilibrium_check else "0"

        parts = [
            "e", "var c", "var p", "var f", "var r",
            "set_pgstat_chan 1",
            "set_pgstat_mode 0",
            "set_pgstat_chan 0",
            "set_pgstat_mode 3",
            f"set_max_bandwidth {bandwidth}",
            f"set_range_minmax da {min_mv}m {max_mv}m",
        ]
        if use_equilibrium_check:
            parts.insert(5, "var t")
        parts += self._ba_range_lines(
            ba_cfg["set_range_value"],
            ba_cfg["mode"] == "auto",
            ba_cfg["auto_min_value"],
            ba_cfg["auto_max_value"],
        )
        parts += [f"set_e {cond_pot if use_equilibrium_check else begin}", "cell_on"]
        if use_equilibrium_check:
            parts += [
                f"# Equilibrium check at {cond_pot} for {cond_time}s",
                "store_var t 0 eb",
                f"meas_loop_ca p c {cond_pot} {eq_interval} {eq_duration}",
                "\tpck_start",
                "\t\tpck_add t",
                "\t\tpck_add p",
                "\t\tpck_add c",
                "\tpck_end",
                f"\tadd_var t {eq_interval}",
                "endloop",
                "store_var t 0 eb",
                f"set_e {begin}",
            ]
        else:
            parts += [f"set_e {begin}"]
        parts += [f"meas_loop_swv p c f r {begin} {end} {step} {amplitude} {frequency}"]
        if use_equilibrium_check:
            parts += [
                "\tpck_start",
                "\t\tpck_add p",
                "\t\tpck_add c",
                "\t\tpck_add f",
                "\t\tpck_add r",
                "\t\tpck_add t",
                "\tpck_end",
                f"\tadd_var t {swv_time_step}",
            ]
        else:
            parts += [
                "\tpck_start", "\t\tpck_add p", "\t\tpck_add c",
                "\t\tpck_add f", "\t\tpck_add r", "\tpck_end",
            ]
        parts += ["endloop", "on_finished:", "cell_off"]
        return "\n".join(parts)

    def _build_alignment_script(self) -> str:
        p = self.alignment_params
        dc_potential_v = float(p["dc_potential"].get())
        ac_amplitude_v = float(p["ac_amplitude"].get())
        start_frequency_hz = float(p["start_frequency"].get())
        end_frequency_hz = float(p["end_frequency"].get())
        points_per_decade = float(p["points_per_decade"].get())

        if start_frequency_hz <= 0 or end_frequency_hz <= 0:
            raise ValueError("Alignment frequencies must be positive.")
        if ac_amplitude_v <= 0:
            raise ValueError("AC amplitude must be positive.")
        if points_per_decade <= 0:
            raise ValueError("Points per decade must be positive.")

        start_frequency = to_si_string(p["start_frequency"].get(), "Hz")
        end_frequency = to_si_string(p["end_frequency"].get(), "Hz")
        ac_amplitude = to_si_string(p["ac_amplitude"].get(), "V")
        dc_potential = to_si_string(p["dc_potential"].get(), "V")

        decades = abs(math.log10(end_frequency_hz) - math.log10(start_frequency_hz))
        n_points = max(1, int(round(decades * points_per_decade)) + 1)

        parts = [
            "e",
            "var f",
            "var z",
            "var i",
            "set_pgstat_chan 1",
            "set_pgstat_mode 0",
            "set_pgstat_chan 0",
            "set_pgstat_mode 3",
            "set_max_bandwidth 40",
            "set_range ba 100u",
            "set_autoranging ba 1n 100u",
            "cell_on",
            f"set_e {dc_potential}",
            "# Dry alignment EIS sweep",
            f"meas_loop_eis f z i {ac_amplitude} {start_frequency} {end_frequency} {n_points}i {dc_potential}",
            "\tpck_start",
            "\t\tpck_add f",
            "\t\tpck_add z",
            "\t\tpck_add i",
            "\tpck_end",
            "endloop",
            "on_finished:",
            "cell_off",
        ]
        return "\n".join(parts)

    # ── MUX helpers ───────────────────────────────────────────────────────────

    def _get_mux_channels(self, params: dict):
        """Parse MUX channel string.  Returns list of ints, [] if disabled,
        None on validation error."""
        entry = params.get("mux_channel")
        if entry is None:
            return []
        raw = entry.get().strip()
        if raw in ("", "0"):
            return []

        channels, seen = [], set()
        for part in [p.strip() for p in raw.split(",") if p.strip()]:
            if "-" in part:
                try:
                    s, e = part.split("-", 1)
                    start, end = int(s), int(e)
                except Exception:
                    messagebox.showerror("Invalid MUX Channel", f"Bad range: '{part}'")
                    return None
                if start > end or not (1 <= start <= 16) or not (1 <= end <= 16):
                    messagebox.showerror("Invalid MUX Channel",
                                         "MUX16 channels must be 1–16.")
                    return None
                for ch in range(start, end + 1):
                    if ch not in seen:
                        seen.add(ch); channels.append(ch)
            else:
                try:
                    ch = int(part)
                except Exception:
                    messagebox.showerror("Invalid MUX Channel", f"Bad channel: '{part}'")
                    return None
                if not (1 <= ch <= 16):
                    messagebox.showerror("Invalid MUX Channel",
                                         "MUX16 channels must be 1–16.")
                    return None
                if ch not in seen:
                    seen.add(ch); channels.append(ch)
        return channels

    @staticmethod
    def _mux_channel_address(channel: int) -> int:
        idx = channel - 1
        return (idx << 4) | idx

    def _wrap_mux(self, base_script: str, channel: int) -> str:
        lines = base_script.splitlines()
        header = lines[0].strip() if lines and lines[0].strip() in ("e", "l") else "e"
        rest   = lines[1:] if lines and lines[0].strip() in ("e", "l") else lines
        addr   = self._mux_channel_address(channel)
        prefix = [
            header, "# MUX16 channel select",
            "set_gpio_cfg 0x3FFi 1", f"set_gpio {addr}i",
        ]
        return "\n".join(prefix + rest)

    def _get_swv_cycles_and_delay(self):
        try:
            n = int(self.swv_params["n_scans"].get())
        except Exception:
            messagebox.showerror("Invalid SWV Scans", "Number of scans must be an integer.")
            return None, None
        try:
            d = float(self.swv_params.get("cycle_delay").get())
        except Exception:
            messagebox.showerror("Invalid SWV Delay", "Delay must be a number.")
            return None, None
        if n < 1:
            messagebox.showerror("Invalid SWV Scans", "Must be at least 1.")
            return None, None
        if d < 0:
            messagebox.showerror("Invalid SWV Delay", "Delay must be non-negative.")
            return None, None
        return n, d

    # ── Generate script (preview) ─────────────────────────────────────────────

    def _generate_cv_script(self):
        try:
            base   = self._build_cv_script()
            mux    = self._get_mux_channels(self.cv_params)
            if mux is None:
                return
            script = base
            if mux:
                script = self._wrap_mux(base, mux[0])
                if len(mux) > 1:
                    script = (f"# NOTE: Multiple channels selected "
                               f"({', '.join(map(str, mux))}). "
                               f"Preview shows ch {mux[0]}.\n") + script
            self._script_preview(script)
        except Exception as exc:
            messagebox.showerror("Error", f"Failed to generate script: {exc}")

    def _generate_lsv_script(self):
        try:
            base   = self._build_lsv_script()
            mux    = self._get_mux_channels(self.lsv_params)
            if mux is None:
                return
            script = base
            if mux:
                script = self._wrap_mux(base, mux[0])
                if len(mux) > 1:
                    script = (f"# NOTE: Multiple channels selected "
                               f"({', '.join(map(str, mux))}). "
                               f"Preview shows ch {mux[0]}.\n") + script
            self._script_preview(script)
        except Exception as exc:
            messagebox.showerror("Error", f"Failed to generate script: {exc}")

    def _generate_swv_script(self):
        try:
            base   = self._build_swv_script()
            mux    = self._get_mux_channels(self.swv_params)
            if mux is None:
                return
            script = base
            if mux:
                script = self._wrap_mux(base, mux[0])
                if len(mux) > 1:
                    script = (f"# NOTE: Multiple channels selected "
                               f"({', '.join(map(str, mux))}). "
                               f"Preview shows ch {mux[0]}.\n") + script
            self._script_preview(script)
        except Exception as exc:
            messagebox.showerror("Error", f"Failed to generate script: {exc}")

    def _generate_alignment_script(self):
        try:
            base = self._build_alignment_script()
            mux = self._get_mux_channels(self.alignment_params)
            if mux is None:
                return
            script = base
            if mux:
                script = self._wrap_mux(base, mux[0])
                if len(mux) > 1:
                    script = (f"# NOTE: Multiple channels selected "
                               f"({', '.join(map(str, mux))}). "
                               f"Preview shows ch {mux[0]}.\n") + script
            self._script_preview(script)
        except Exception as exc:
            messagebox.showerror("Error", f"Failed to generate script: {exc}")

    # ── Add to queue ──────────────────────────────────────────────────────────

    def _add_one(self, technique: str, script: str, params: dict, mux_channel=None, note: str = ""):
        """Save script via registry and enqueue one item."""
        fp, fn = self._session.registry.save_script(
            technique, script, params, mux_channel, note=note
        )
        details = fn if mux_channel is None else f"{fn} (MUX ch {mux_channel})"
        self._add_to_queue({
            "type":        technique,
            "script_path": str(fp),
            "status":      "pending",
            "details":     details,
        })

    def _save_to_library(self, technique: str, base_script: str, params: dict, mux_channels, note: str = ""):
        saved = []
        if mux_channels:
            for ch in mux_channels:
                _fp, fn = self._session.registry.save_script(
                    technique,
                    self._wrap_mux(base_script, ch),
                    params,
                    ch,
                    note=note,
                )
                saved.append(f"{fn} (MUX ch {ch})")
        else:
            _fp, fn = self._session.registry.save_script(
                technique,
                base_script,
                params,
                note=note,
            )
            saved.append(fn)
        return saved

    @staticmethod
    def _show_library_save_result(technique: str, saved: list):
        if not saved:
            return
        if len(saved) == 1:
            messagebox.showinfo("Saved", f"{technique} saved to library.\n{saved[0]}")
            return
        messagebox.showinfo(
            "Saved",
            f"{technique} saved to library for {len(saved)} item(s):\n" + "\n".join(saved),
        )

    def _save_cv_to_library(self):
        try:
            base = self._build_cv_script()
        except Exception as exc:
            messagebox.showerror("Error", str(exc)); return
        mux = self._get_mux_channels(self.cv_params)
        if mux is None:
            return
        raw_params = {k: v.get() for k, v in self.cv_params.items()}
        raw_params.update(self._serialize_ba_range_config("CV"))
        note = (self._library_note.get() or "").strip()
        saved = self._save_to_library("CV", base, raw_params, mux, note=note)
        self._show_library_save_result("CV", saved)

    def _save_lsv_to_library(self):
        try:
            base = self._build_lsv_script()
        except Exception as exc:
            messagebox.showerror("Error", str(exc)); return
        mux = self._get_mux_channels(self.lsv_params)
        if mux is None:
            return
        raw_params = {k: v.get() for k, v in self.lsv_params.items()}
        raw_params.update(self._serialize_ba_range_config("LSV"))
        note = (self._library_note.get() or "").strip()
        saved = self._save_to_library("LSV", base, raw_params, mux, note=note)
        self._show_library_save_result("LSV", saved)

    def _save_swv_to_library(self):
        try:
            base = self._build_swv_script()
        except Exception as exc:
            messagebox.showerror("Error", str(exc)); return
        mux = self._get_mux_channels(self.swv_params)
        if mux is None:
            return
        n_scans, delay = self._get_swv_cycles_and_delay()
        if n_scans is None:
            return
        raw_params = {k: v.get() for k, v in self.swv_params.items()}
        raw_params["bandwidth"] = self._swv_bandwidth.get()
        raw_params["n_scans"] = str(n_scans)
        raw_params["cycle_delay"] = str(delay)
        raw_params.update(self._serialize_ba_range_config("SWV"))
        note = (self._library_note.get() or "").strip()
        saved = self._save_to_library("SWV", base, raw_params, mux, note=note)
        self._show_library_save_result("SWV", saved)

    def _save_alignment_to_library(self):
        try:
            base = self._build_alignment_script()
        except Exception as exc:
            messagebox.showerror("Error", str(exc)); return
        mux = self._get_mux_channels(self.alignment_params)
        if mux is None:
            return
        raw_params = {k: v.get() for k, v in self.alignment_params.items()}
        note = (self._library_note.get() or "").strip()
        saved = self._save_to_library("ALIGNMENT", base, raw_params, mux, note=note)
        self._show_library_save_result("ALIGNMENT", saved)

    def _add_cv_to_queue(self):
        try:
            base = self._build_cv_script()
        except Exception as exc:
            messagebox.showerror("Error", str(exc)); return
        mux    = self._get_mux_channels(self.cv_params)
        if mux is None:
            return
        # Extract raw string values for hashing
        raw_params = {k: v.get() for k, v in self.cv_params.items()}
        raw_params.update(self._serialize_ba_range_config("CV"))
        note = (self._library_note.get() or "").strip()
        if mux:
            for ch in mux:
                self._add_one("CV", self._wrap_mux(base, ch), raw_params, mux_channel=ch, note=note)
            messagebox.showinfo("Success", f"CV added for MUX channels: {', '.join(map(str, mux))}")
        else:
            self._add_one("CV", base, raw_params, note=note)
            messagebox.showinfo("Success", "CV added to queue")
        self._refresh_queue()

    def _add_lsv_to_queue(self):
        try:
            base = self._build_lsv_script()
        except Exception as exc:
            messagebox.showerror("Error", str(exc)); return
        mux = self._get_mux_channels(self.lsv_params)
        if mux is None:
            return
        raw_params = {k: v.get() for k, v in self.lsv_params.items()}
        raw_params.update(self._serialize_ba_range_config("LSV"))
        note = (self._library_note.get() or "").strip()
        if mux:
            for ch in mux:
                self._add_one("LSV", self._wrap_mux(base, ch), raw_params, mux_channel=ch, note=note)
            messagebox.showinfo("Success", f"LSV added for MUX channels: {', '.join(map(str, mux))}")
        else:
            self._add_one("LSV", base, raw_params, note=note)
            messagebox.showinfo("Success", "LSV added to queue")
        self._refresh_queue()

    def _add_swv_to_queue(self):
        try:
            base = self._build_swv_script()
        except Exception as exc:
            messagebox.showerror("Error", str(exc)); return
        mux = self._get_mux_channels(self.swv_params)
        if mux is None:
            return
        n_scans, delay = self._get_swv_cycles_and_delay()
        if n_scans is None:
            return
        raw_params = {k: v.get() for k, v in self.swv_params.items()}
        raw_params["bandwidth"] = self._swv_bandwidth.get()
        raw_params.update(self._serialize_ba_range_config("SWV"))
        note = (self._library_note.get() or "").strip()

        added = []
        for cycle in range(1, n_scans + 1):
            if mux:
                for ch in mux:
                    script = self._wrap_mux(base, ch)
                    fp, fn = self._session.registry.save_script(
                        "SWV", script, raw_params, ch, note=note
                    )
                    self._add_to_queue({
                        "type": "SWV", "script_path": str(fp),
                        "status": "pending", "details": f"{fn} (MUX ch {ch})",
                    })
                    added.append(f"{fn} (ch {ch})")
            else:
                fp, fn = self._session.registry.save_script("SWV", base, raw_params, note=note)
                self._add_to_queue({
                    "type": "SWV", "script_path": str(fp),
                    "status": "pending", "details": fn,
                })
                added.append(fn)

            if delay > 0 and cycle < n_scans:
                self._add_to_queue({
                    "type": "PAUSE", "status": "pending",
                    "details": f"Pause for {delay:.1f} sec",
                    "pause_seconds": delay,
                })

        self._refresh_queue()
        messagebox.showinfo("Success",
            f"SWV added for {n_scans} scan(s)\nSaved: {', '.join(added)}")

    def _add_alignment_to_queue(self):
        try:
            base = self._build_alignment_script()
        except Exception as exc:
            messagebox.showerror("Error", str(exc)); return
        mux = self._get_mux_channels(self.alignment_params)
        if mux is None:
            return
        raw_params = {k: v.get() for k, v in self.alignment_params.items()}
        note = (self._library_note.get() or "").strip()
        if mux:
            for ch in mux:
                self._add_one("ALIGNMENT", self._wrap_mux(base, ch), raw_params, mux_channel=ch, note=note)
            messagebox.showinfo("Success", f"Alignment test added for MUX channels: {', '.join(map(str, mux))}")
        else:
            self._add_one("ALIGNMENT", base, raw_params, note=note)
            messagebox.showinfo("Success", "Alignment test added to queue")
        self._refresh_queue()

    def _add_pause_to_queue(self):
        try:
            secs = float(self.pause_params["pause_time"].get())
            if secs < 0:
                raise ValueError("Pause time must be non-negative")
        except (KeyError, ValueError, tk.TclError) as exc:
            messagebox.showerror("Invalid Pause", str(exc)); return
        self._add_to_queue({
            "type": "PAUSE", "status": "pending",
            "details": f"Pause for {secs:.1f} sec",
            "pause_seconds": secs,
        })
        self._refresh_queue()
        messagebox.showinfo("Success", f"Pause ({secs:.1f} sec) added to queue")

    def _add_alert_pause_to_queue(self):
        msg = (self.pause_params.get("alert_message", tk.StringVar()).get() or "").strip()
        if not msg:
            messagebox.showerror("Invalid Alert", "Alert message cannot be empty.")
            return
        self._add_to_queue({
            "type": "ALERT", "status": "pending",
            "details": "Alert pause", "alert_message": msg,
        })
        self._refresh_queue()
        messagebox.showinfo("Success", "Alert pause added to queue")

    # ── Run now ───────────────────────────────────────────────────────────────

    def _run_cv_now(self):
        try:
            base = self._build_cv_script()
        except Exception as exc:
            messagebox.showerror("Error", str(exc)); return
        mux = self._get_mux_channels(self.cv_params)
        if mux is None:
            return
        raw_params = {k: v.get() for k, v in self.cv_params.items()}
        raw_params.update(self._serialize_ba_range_config("CV"))
        if mux:
            if len(mux) == 1:
                self._run_now("CV", self._wrap_mux(base, mux[0]), {"mux_channel": mux[0], "params": raw_params})
            else:
                # Multi-channel: delegate to app for sequence run
                self._run_now("CV_MUX_SEQ", base, {"channels": mux, "params": raw_params})
        else:
            self._run_now("CV", base, {"mux_channel": None, "params": raw_params})

    def _run_lsv_now(self):
        try:
            base = self._build_lsv_script()
        except Exception as exc:
            messagebox.showerror("Error", str(exc)); return
        mux = self._get_mux_channels(self.lsv_params)
        if mux is None:
            return
        raw_params = {k: v.get() for k, v in self.lsv_params.items()}
        raw_params.update(self._serialize_ba_range_config("LSV"))
        if mux:
            if len(mux) == 1:
                self._run_now("LSV", self._wrap_mux(base, mux[0]), {"mux_channel": mux[0], "params": raw_params})
            else:
                self._run_now("LSV_MUX_SEQ", base, {"channels": mux, "params": raw_params})
        else:
            self._run_now("LSV", base, {"mux_channel": None, "params": raw_params})

    def _run_swv_now(self):
        try:
            base = self._build_swv_script()
        except Exception as exc:
            messagebox.showerror("Error", str(exc)); return
        mux = self._get_mux_channels(self.swv_params)
        if mux is None:
            return
        n_scans, delay = self._get_swv_cycles_and_delay()
        if n_scans is None:
            return
        raw_params = {k: v.get() for k, v in self.swv_params.items()}
        raw_params["bandwidth"] = self._swv_bandwidth.get()
        raw_params.update(self._serialize_ba_range_config("SWV"))
        if mux:
            if len(mux) == 1 and n_scans == 1:
                self._run_now("SWV", self._wrap_mux(base, mux[0]), {"mux_channel": mux[0], "params": raw_params})
            else:
                self._run_now(
                    "SWV_MUX_CYCLES",
                    base,
                    {"channels": mux, "n_scans": n_scans, "delay": delay, "params": raw_params},
                )
        else:
            if n_scans == 1:
                self._run_now("SWV", base, {"mux_channel": None, "params": raw_params})
            else:
                self._run_now("SWV_CYCLES", base, {"n_scans": n_scans, "delay": delay, "params": raw_params})

    def _run_alignment_now(self):
        try:
            base = self._build_alignment_script()
        except Exception as exc:
            messagebox.showerror("Error", str(exc)); return
        mux = self._get_mux_channels(self.alignment_params)
        if mux is None:
            return
        raw_params = {k: v.get() for k, v in self.alignment_params.items()}
        if mux:
            if len(mux) == 1:
                self._run_now("ALIGNMENT", self._wrap_mux(base, mux[0]), {"mux_channel": mux[0], "params": raw_params})
            else:
                self._run_now("ALIGNMENT_MUX_SEQ", base, {"channels": mux, "params": raw_params})
        else:
            self._run_now("ALIGNMENT", base, {"mux_channel": None, "params": raw_params})

    def _run_pause_now(self):
        try:
            secs = float(self.pause_params["pause_time"].get())
            if secs < 0:
                raise ValueError("Pause time must be non-negative")
        except (KeyError, ValueError, tk.TclError) as exc:
            messagebox.showerror("Invalid Pause", str(exc)); return

        def _do():
            time.sleep(secs)
            self._session.log(f"Immediate pause completed ({secs:.1f} sec)")

        threading.Thread(target=_do, daemon=True).start()
        self._session.log(f"Immediate pause started ({secs:.1f} sec)…")
    def _show_custom_params(self):
        self._clear_params()
        self.current_technique = "CUSTOM"
        self._custom_panel = CustomScriptPanel(
            params_frame      = self._params_frame,
            session           = self._session,
            on_run_now        = self._run_now,
            on_add_to_queue   = self._add_to_queue,
            on_script_preview = self._script_preview,
            save_script_fn    = self._session.registry.save_script,
            wrap_mux_fn       = self._wrap_mux,
            parse_mux_fn      = self._get_mux_channels,
        )

    def _run_pstrace_preset(self):
        script = (
            "e\n"
            "set_gpio_cfg 0x3FFi 1\n"
            "set_gpio 119i\n"
            "var c\n"
            "var p\n"
            "var f\n"
            "var g\n"
            "set_pgstat_chan 1\n"
            "set_pgstat_mode 0\n"
            "set_pgstat_chan 0\n"
            "set_pgstat_mode 3\n"
            "set_max_bandwidth 4k\n"
            "set_range_minmax da -536m 36m\n"
            "set_range ba 59n\n"
            "set_autoranging ba 59n 59n\n"
            "set_e -500m\n"
            "cell_on\n"
            "meas_loop_ca p c -500m 200m 1\n"
            "  pck_start\n"
            "    pck_add p\n"
            "    pck_add c\n"
            "  pck_end\n"
            "endloop\n"
            "meas_loop_swv p c f g -500m 0 2m 36m 100\n"
            "  pck_start\n"
            "    pck_add p\n"
            "    pck_add c\n"
            "    pck_add f\n"
            "    pck_add g\n"
            "  pck_end\n"
            "endloop\n"
            "on_finished:\n"
            "  cell_off\n"
        )
        self._script_preview(script)
        self._run_now("SWV", script, None)
