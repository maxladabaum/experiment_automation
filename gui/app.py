"""
gui/app.py — ElectrochemGUI application class.

This is now a **thin orchestrator**.  It:
  1. Creates the shared :class:`~core.session.SessionState`
  2. Creates each tab class and adds it to the notebook
  3. Wires the inter-tab callbacks so no tab imports another

All business logic lives in the ``core/`` modules.
All UI logic lives in the individual ``gui/tab_*.py`` files.
"""

import threading
import time
from pathlib import Path
from tkinter import messagebox
import tkinter as tk
from tkinter import ttk

from config import (
    APP_VERSION, WINDOW_TITLE, WINDOW_GEOMETRY,
    PREFERRED_STEPS_PER_STROKE, PREFERRED_SYRINGE_UL,
)
from core.session  import SessionState
from core.runner   import SerialMeasurementRunner
from core.session_manager import SessionManager
from gui.session_bar import SessionBar
from gui.tab_script  import ScriptTab
from gui.tab_plotter import PlotterTab
from gui.tab_method  import MethodTab
from gui.tab_queue   import QueueTab
from gui.tab_pump    import PumpTab

try:
    from pump_gui import PumpCtrl, HAS_COM as PUMP_HAS_COM
    PUMP_AVAILABLE = True
except ImportError:
    PumpCtrl       = None
    PUMP_HAS_COM   = False
    PUMP_AVAILABLE = False
    print("Warning: pump backend not found — pump features disabled.")


class ElectrochemGUI:
    """Top-level GUI application.

    Instantiate with a ``tk.Tk`` root window, then call ``root.mainloop()``.
    """

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry(WINDOW_GEOMETRY)

        # ── Pump controller (optional) ────────────────────────────────────────
        if PUMP_AVAILABLE and PumpCtrl is not None:
            self._pump_ctrl = PumpCtrl(
                use_sim=(not PUMP_HAS_COM),
                log_cb=lambda m: self._pump_tab_log(m),
            )
            self._pump_ctrl.configure_calibration(
                PREFERRED_STEPS_PER_STROKE, PREFERRED_SYRINGE_UL
            )
        else:
            self._pump_ctrl = None

        # ── Notebook ──────────────────────────────────────────────────────────
        self._nb = ttk.Notebook(root)
        self._nb.pack(fill="both", expand=True, padx=5, pady=5)

        # ── Session state (shared by all tabs) ────────────────────────────────
        # NEW
        self._session = SessionState(
            log_callback    = self._log,
            status_callback = self._set_status,
        )
        self._session_mgr = SessionManager(log_callback=self._log)

        # ── Tab frames ────────────────────────────────────────────────────────
        pump_frame    = ttk.Frame(self._nb)
        method_frame  = ttk.Frame(self._nb)
        script_frame  = ttk.Frame(self._nb)
        queue_frame   = ttk.Frame(self._nb)
        plotter_frame = ttk.Frame(self._nb)

        if PUMP_AVAILABLE:
            self._nb.add(pump_frame,    text="Pump Control")
        self._nb.add(method_frame,  text="Method Creation")
        self._nb.add(script_frame,  text="Script Preview")
        self._nb.add(queue_frame,   text="Queue & Execution")
        self._nb.add(plotter_frame, text="Plotter")

        # ── Instantiate tabs ──────────────────────────────────────────────────
        self._script_tab = ScriptTab(script_frame)

        self._plotter_tab = PlotterTab(
            parent_frame = plotter_frame,
            session      = self._session,
            notebook     = self._nb,
        )

        self._queue_tab = QueueTab(
            parent_frame = queue_frame,
            session      = self._session,
            plotter      = self._plotter_tab,
            pump_ctrl    = self._pump_ctrl,
            root         = self.root,
        )
        # Wire session callbacks now that queue tab (with its log widget) exists
        self._session._log    = self._log
        self._session._status = self._set_status

        self._method_tab = MethodTab(
            parent_frame      = method_frame,
            session           = self._session,
            on_add_to_queue   = self._queue_tab.add_item,
            on_refresh_queue  = self._queue_tab.refresh,
            on_script_preview = self._script_tab.update,
            on_run_now        = self._run_now,
        )

        if PUMP_AVAILABLE:
            self._pump_tab = PumpTab(
                parent_frame   = pump_frame,
                pump_ctrl      = self._pump_ctrl,
                on_add_to_queue= self._queue_tab.add_item,
                root           = self.root,
            )
        else:
            self._pump_tab = None
        # ── Session bar (bottom of window) ───────────────────────────────────────
        self._session_bar = SessionBar(
            root            = root,
            session_manager = self._session_mgr,
        )
        # Give all tabs access to the session manager for require_experiment() guards
        self._session.session_manager = self._session_mgr
    
    # ── Inter-tab wiring helpers ──────────────────────────────────────────────

    def _log(self, msg: str):
        """Route log messages to the queue tab's log panel."""
        try:
            self._queue_tab.log(msg)
        except Exception:
            print(msg)

    def _set_status(self, msg: str):
        try:
            self._queue_tab.set_status(msg)
        except Exception:
            pass

    def _pump_tab_log(self, msg: str):
        if self._pump_tab is not None:
            self._pump_tab.log(msg)

    # ── Immediate run dispatcher ──────────────────────────────────────────────

    def _run_now(self, technique: str, script_or_base, extra=None):
        """Handle all 'Run Now' requests from MethodTab.

        ``technique`` is one of:
          - ``"CV"`` / ``"SWV"``          → single immediate run
          - ``"CV_MUX_SEQ"``              → sequence over multiple MUX channels
          - ``"SWV_CYCLES"``              → repeated SWV scans (no MUX)
          - ``"SWV_MUX_CYCLES"``          → repeated SWV scans over MUX channels
        ``extra`` carries the additional context needed for each variant.
        """
        if self._session.is_running:
            messagebox.showwarning(
                "Busy",
                "A measurement is already running. "
                "Stop it before starting a new one."
            )
            return

        if technique in ("CV", "SWV"):
            mux_channel = extra   # int or None
            self._run_single(technique, script_or_base, mux_channel)

        elif technique in ("CV_MUX_SEQ", "SWV_MUX_SEQ"):
            base_script = script_or_base
            channels    = extra   # list[int]
            tech        = "CV" if technique.startswith("CV") else "SWV"
            self._run_mux_sequence(tech, base_script, channels)

        elif technique == "SWV_CYCLES":
            n_scans, delay = extra
            self._run_swv_cycles(script_or_base, n_scans, delay)

        elif technique == "SWV_MUX_CYCLES":
            channels, n_scans, delay = extra
            self._run_mux_swv_cycles(script_or_base, channels, n_scans, delay)

    # ── Single run ────────────────────────────────────────────────────────────

    def _run_single(self, technique: str, script: str, mux_channel=None):
        try:
            fp, fn = self._session.registry.save_script(technique, script, mux_channel)
        except Exception as exc:
            messagebox.showerror("File Error", f"Failed to save script: {exc}"); return

        self._queue_tab.clear_log()
        self._session.is_running = True
        self._queue_tab.set_status(f"Running: {technique} — {fn}")
        self._plotter_tab.start_live(f"{technique} (live)", label=technique)

        def worker():
            meas_tag = self._session.next_meas_tag()
            self._log(f"[Tag] {meas_tag}")
            self.root.after(0, self._queue_tab.refresh_labels)
            runner = SerialMeasurementRunner(
                fp,
                log_callback  = self._log,
                data_callback = self._plotter_tab.push_live_point,
            )
            self._session.current_runner = runner
            success, csv_path = runner.execute(meas_tag=meas_tag)
            stopped = not runner.is_running
            self._session.current_runner = None

            def finish():
                self._session.is_running = False
                self._plotter_tab.stop_live()
                if csv_path:
                    self._plotter_tab.plot_data(
                        csv_path,
                        self._session.last_live_plot_color,
                        self._session.last_live_plot_label,
                    )
                if stopped:
                    self._queue_tab.set_status("Ready (stopped)")
                    messagebox.showinfo("Stopped", f"{technique} run was stopped.")
                elif success:
                    self._queue_tab.set_status("Ready")
                    messagebox.showinfo("Complete", f"{technique} run completed.\n{csv_path or ''}")
                else:
                    self._queue_tab.set_status("Ready (last run failed)")
                    messagebox.showerror("Failed", f"{technique} run failed. Check log.")
            self.root.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    # ── MUX sequence run ──────────────────────────────────────────────────────

    def _run_mux_sequence(self, technique: str, base_script: str, channels: list):
        self._queue_tab.clear_log()
        self._session.is_running = True
        last_csv = None

        def worker():
            nonlocal last_csv
            stopped = False
            success = True
            for ch in channels:
                if not self._session.is_running:
                    stopped = True; success = False; break
                mux_script = self._method_tab._wrap_mux(base_script, ch)
                fp, fn = self._session.registry.save_script(technique, mux_script, ch)
                color = self._session.next_plot_color()
                label = f"MUX ch {ch}"
                self.root.after(0, self._plotter_tab.start_live,
                                f"{technique} ch {ch} (live)", color, label)
                self.root.after(0, self._queue_tab.set_status,
                                f"Running: {technique} MUX ch {ch}")
                meas_tag = self._session.next_meas_tag()
                self._log(f"[Tag] {meas_tag}")
                self.root.after(0, self._queue_tab.refresh_labels)
                runner = SerialMeasurementRunner(
                    fp, log_callback=self._log,
                    data_callback=self._plotter_tab.push_live_point)
                self._session.current_runner = runner
                ok, csv_path = runner.execute(meas_tag=meas_tag)
                self._session.current_runner = None
                self.root.after(0, self._plotter_tab.stop_live)
                if csv_path:
                    last_csv = csv_path
                    self.root.after(0, self._plotter_tab.plot_data,
                                   csv_path, color, label)
                if not ok:
                    success = False
                    if not runner.is_running:
                        stopped = True
                    break

            def finish():
                self._session.is_running = False
                if stopped:
                    self._queue_tab.set_status("Ready (stopped)")
                    messagebox.showinfo("Stopped", f"{technique} MUX run stopped.")
                elif success:
                    self._queue_tab.set_status("Ready")
                    messagebox.showinfo("Complete", f"{technique} MUX run completed.")
                else:
                    self._queue_tab.set_status("Ready (failed)")
                    messagebox.showerror("Failed", f"{technique} MUX run failed.")
            self.root.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    # ── SWV multi-scan (no MUX) ───────────────────────────────────────────────

    def _run_swv_cycles(self, base_script: str, n_scans: int, delay: float):
        self._queue_tab.clear_log()
        self._session.is_running = True

        def worker():
            stopped = False; success = True; last_csv = None
            for scan in range(1, n_scans + 1):
                if not self._session.is_running:
                    stopped = True; success = False; break
                fp, fn = self._session.registry.save_script("SWV", base_script)
                color = self._session.next_plot_color()
                label = f"SWV scan {scan}"
                self.root.after(0, self._plotter_tab.start_live,
                                f"SWV (scan {scan}/{n_scans} live)", color, label)
                self.root.after(0, self._queue_tab.set_status,
                                f"Running: SWV scan {scan}/{n_scans}")
                meas_tag = self._session.next_meas_tag()
                self._log(f"[Tag] {meas_tag}")
                self.root.after(0, self._queue_tab.refresh_labels)
                runner = SerialMeasurementRunner(
                    fp, log_callback=self._log,
                    data_callback=self._plotter_tab.push_live_point)
                self._session.current_runner = runner
                ok, csv_path = runner.execute(meas_tag=meas_tag)
                self._session.current_runner = None
                self.root.after(0, self._plotter_tab.stop_live)
                if csv_path:
                    last_csv = csv_path
                    self.root.after(0, self._plotter_tab.plot_data,
                                   csv_path, color, label)
                if not ok:
                    success = False
                    if not runner.is_running:
                        stopped = True
                    break
                if delay > 0 and scan < n_scans:
                    waited = 0.0
                    while waited < delay and self._session.is_running:
                        time.sleep(min(0.5, delay - waited))
                        waited += 0.5

            def finish():
                self._session.is_running = False
                if stopped:
                    self._queue_tab.set_status("Ready (stopped)")
                    messagebox.showinfo("Stopped", "SWV run stopped.")
                elif success:
                    self._queue_tab.set_status("Ready")
                    messagebox.showinfo("Complete", f"SWV {n_scans} scan(s) complete.")
                else:
                    self._queue_tab.set_status("Ready (failed)")
                    messagebox.showerror("Failed", "SWV run failed.")
            self.root.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    # ── SWV multi-scan + MUX ─────────────────────────────────────────────────

    def _run_mux_swv_cycles(self, base_script, channels, n_scans, delay):
        self._queue_tab.clear_log()
        self._session.is_running = True

        def worker():
            stopped = False; success = True; last_csv = None
            for scan in range(1, n_scans + 1):
                for ch in channels:
                    if not self._session.is_running:
                        stopped = True; success = False; break
                    mux_script = self._method_tab._wrap_mux(base_script, ch)
                    fp, fn = self._session.registry.save_script("SWV", mux_script, ch)
                    color = self._session.next_plot_color()
                    label = f"MUX ch {ch} scan {scan}"
                    self.root.after(0, self._plotter_tab.start_live,
                                    f"SWV MUX ch {ch} ({scan}/{n_scans})", color, label)
                    self.root.after(0, self._queue_tab.set_status,
                                    f"Running: SWV MUX ch {ch} scan {scan}/{n_scans}")
                    meas_tag = self._session.next_meas_tag()
                    self._log(f"[Tag] {meas_tag}")
                    self.root.after(0, self._queue_tab.refresh_labels)
                    runner = SerialMeasurementRunner(
                        fp, log_callback=self._log,
                        data_callback=self._plotter_tab.push_live_point)
                    self._session.current_runner = runner
                    ok, csv_path = runner.execute(meas_tag=meas_tag)
                    self._session.current_runner = None
                    self.root.after(0, self._plotter_tab.stop_live)
                    if csv_path:
                        last_csv = csv_path
                        self.root.after(0, self._plotter_tab.plot_data,
                                       csv_path, color, label)
                    if not ok:
                        success = False
                        if not runner.is_running: stopped = True
                        break
                if stopped or not success:
                    break
                if delay > 0 and scan < n_scans:
                    waited = 0.0
                    while waited < delay and self._session.is_running:
                        time.sleep(min(0.5, delay - waited))
                        waited += 0.5

            def finish():
                self._session.is_running = False
                if stopped:
                    self._queue_tab.set_status("Ready (stopped)")
                    messagebox.showinfo("Stopped", "SWV MUX run stopped.")
                elif success:
                    self._queue_tab.set_status("Ready")
                    messagebox.showinfo("Complete", f"SWV MUX {n_scans} scan(s) complete.")
                else:
                    self._queue_tab.set_status("Ready (failed)")
                    messagebox.showerror("Failed", "SWV MUX run failed.")
            self.root.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()
