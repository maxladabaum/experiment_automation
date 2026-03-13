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
from pathlib import Path
from tkinter import messagebox
import tkinter as tk
from tkinter import ttk
import serial.tools.list_ports

from core.mscript_parser import to_si_string
from core.runner import SerialMeasurementRunner
from core.session import SessionState
from gui.tab_custom_script import CustomScriptPanel

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
        self.swv_params: dict  = {}
        self.pause_params: dict = {}
        self._library_note = tk.StringVar(value="")

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
        ttk.Button(tech_frame, text="Square Wave Voltammetry (SWV)",
           command=self._show_swv_params, width=28).pack(pady=5)
        ttk.Button(tech_frame, text="Custom Script (File)",
                command=self._show_custom_params, width=28).pack(pady=5)
        ttk.Separator(tech_frame, orient="horizontal").pack(fill="x", pady=6)
        ttk.Button(tech_frame, text="PStrace SWV Preset",
                command=self._run_pstrace_preset, width=28).pack(pady=5)
        ttk.Separator(tech_frame, orient="horizontal").pack(fill="x", pady=6)
        ttk.Button(tech_frame, text="Pause / Alert",
                   command=self._show_pause_params, width=28).pack(pady=5)

        self._device_status = ttk.Label(left, text="", foreground="blue")
        self._device_status.pack(pady=10)
        ttk.Button(left, text="Check Device Connection",
                   command=self._check_device).pack(pady=5)

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
        ports = list(serial.tools.list_ports.comports())
        if ports:
            self._device_status.config(
                text="Devices found (check console)", foreground="green"
            )
            print("Available serial devices:")
            for p in ports:
                print(f"  {p.device}: {p.description}")
        else:
            self._device_status.config(text="No devices found", foreground="red")

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

        ttk.Label(self._params_frame, text="Library note (optional):").grid(
            row=len(params), column=0, sticky="w", pady=2)
        ttk.Entry(self._params_frame, width=40, textvariable=self._library_note).grid(
            row=len(params), column=1, sticky="w", pady=2)

        btn_frame = ttk.Frame(self._params_frame)
        btn_frame.grid(row=len(params) + 1, column=0, columnspan=2, pady=20)
        ttk.Button(btn_frame, text="Generate Script",
                   command=self._generate_cv_script).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Run Now",
                   command=self._run_cv_now).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Add to Queue",
                   command=self._add_cv_to_queue).pack(side="left", padx=5)

    def _show_swv_params(self):
        self._clear_params()
        self.current_technique = "SWV"
        self.swv_params = {}
        params = [
            ("Begin Potential (V):",                "begin_potential", "-0.5"),
            ("End Potential (V):",                  "end_potential",   "0.5"),
            ("Step Potential (V):",                 "step_potential",  "0.002"),
            ("Amplitude (V):",                      "amplitude",       "0.02"),
            ("Frequency (Hz):",                     "frequency",       "15"),
            ("Number of Scans:",                    "n_scans",         "1"),
            ("Delay Between Scans (s):",            "cycle_delay",     "0"),
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
            self.swv_params[key] = entry

        ttk.Label(self._params_frame, text="Library note (optional):").grid(
            row=len(params), column=0, sticky="w", pady=2)
        ttk.Entry(self._params_frame, width=40, textvariable=self._library_note).grid(
            row=len(params), column=1, sticky="w", pady=2)

        btn_frame = ttk.Frame(self._params_frame)
        btn_frame.grid(row=len(params) + 1, column=0, columnspan=2, pady=20)
        ttk.Button(btn_frame, text="Generate Script",
                   command=self._generate_swv_script).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Run Now",
                   command=self._run_swv_now).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Add to Queue",
                   command=self._add_swv_to_queue).pack(side="left", padx=5)

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

        parts = [
            "e", "var c", "var p",
            "set_pgstat_mode 2", "set_max_bandwidth 40",
            "set_range ba 100u", "set_autoranging ba 1n 100u",
        ]
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

    def _build_swv_script(self) -> str:
        p = self.swv_params
        begin_v  = float(p["begin_potential"].get())
        end_v    = float(p["end_potential"].get())
        amp_v    = float(p["amplitude"].get())

        begin     = to_si_string(p["begin_potential"].get(), "V")
        end       = to_si_string(p["end_potential"].get(),   "V")
        step      = to_si_string(p["step_potential"].get(),  "V")
        amplitude = to_si_string(p["amplitude"].get(),       "V")
        frequency = to_si_string(p["frequency"].get(),       "Hz")
        cond_pot  = to_si_string(p["cond_potential"].get(),  "V")
        cond_time = p["cond_time"].get()

        min_mv = int((min(begin_v, end_v) - amp_v) * 1000)
        max_mv = int((max(begin_v, end_v) + amp_v) * 1000)

        parts = [
            "e", "var c", "var p", "var f", "var r",
            "set_pgstat_chan 1",
            "set_pgstat_mode 0",
            "set_pgstat_chan 0",
            "set_pgstat_mode 3",
            "set_max_bandwidth 4k",
            f"set_range_minmax da {min_mv}m {max_mv}m",
            "set_range ba 59n", "set_autoranging ba 59n 59n", "cell_on",
        ]
        if float(cond_time) > 0:
            parts += [f"# Equilibrate at {cond_pot} for {cond_time}s",
                      f"set_e {cond_pot}", f"wait {cond_time}"]
        parts += [f"set_e {begin}"]
        parts += [
            f"meas_loop_swv p c f r {begin} {end} {step} {amplitude} {frequency}",
            "\tpck_start", "\t\tpck_add p", "\t\tpck_add c",
            "\t\tpck_add f", "\t\tpck_add r", "\tpck_end", "endloop",
            "on_finished:", "cell_off",
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
        note = (self._library_note.get() or "").strip()
        if mux:
            for ch in mux:
                self._add_one("CV", self._wrap_mux(base, ch), raw_params, mux_channel=ch, note=note)
            messagebox.showinfo("Success", f"CV added for MUX channels: {', '.join(map(str, mux))}")
        else:
            self._add_one("CV", base, raw_params, note=note)
            messagebox.showinfo("Success", "CV added to queue")
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
        if mux:
            if len(mux) == 1:
                self._run_now("CV", self._wrap_mux(base, mux[0]), mux[0])
            else:
                # Multi-channel: delegate to app for sequence run
                self._run_now("CV_MUX_SEQ", base, mux)
        else:
            self._run_now("CV", base, None)

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
        if mux:
            if len(mux) == 1 and n_scans == 1:
                self._run_now("SWV", self._wrap_mux(base, mux[0]), mux[0])
            else:
                self._run_now("SWV_MUX_CYCLES", base, (mux, n_scans, delay))
        else:
            if n_scans == 1:
                self._run_now("SWV", base, None)
            else:
                self._run_now("SWV_CYCLES", base, (n_scans, delay))

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
