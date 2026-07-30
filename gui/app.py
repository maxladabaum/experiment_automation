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
    SESSION_ARCHIVE_DIR,
    SLACK_ENABLE, SLACK_BOT_TOKEN, SLACK_SIGNING_SECRET, SLACK_PORT,
    SLACK_ONLY_WHEN_EXPERIMENT,
    NGROK_AUTOSTART, NGROK_PATH, NGROK_DOMAIN,
)
from core.session  import SessionState
from core.runner   import SerialMeasurementRunner
from core.session_manager import SessionManager
from core.session_archive import archive_session
from core.slack_bot import SlackBotServer
from gui.session_bar import SessionBar
from gui.tab_script  import ScriptTab
from gui.tab_plotter import PlotterTab
from gui.tab_method  import MethodTab
from gui.tab_queue   import QueueTab
from gui.tab_pump    import PumpTab
from gui.tab_recipe_maker import RecipeMakerTab
from gui.tab_bayesian_optimization import BayesianOptimizationTab
from gui.tab_automated_titration import AutomatedTitrationTab
from gui.widgets import ScrollableFrame

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
        self.root.minsize(980, 650)
        self._ngrok_proc = None

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
        self._layout_root = ttk.Frame(root)
        self._layout_root.pack(fill="both", expand=True)

        self._content_frame = ttk.Frame(self._layout_root)
        self._content_frame.pack(side="top", fill="both", expand=True)
        self._content_frame.pack_propagate(False)

        self._session_bar_min_height = 200
        self._session_bar_collapsed_height = 30
        self._session_bar_minimized = False
        self._session_bar_frame = ttk.Frame(self._layout_root, height=self._session_bar_min_height)
        self._session_bar_frame.pack(side="bottom", fill="x")
        self._session_bar_frame.pack_propagate(False)

        self._session_bar_resize_grip = ttk.Frame(self._session_bar_frame, height=self._session_bar_collapsed_height)
        self._session_bar_resize_grip.pack(side="top", fill="x")
        self._session_bar_resize_grip.configure(cursor="sb_v_double_arrow")
        self._session_bar_resize_grip.bind("<ButtonPress-1>", self._start_session_bar_resize)
        self._session_bar_label = ttk.Label(
            self._session_bar_resize_grip,
            text="Session / Experiment",
            foreground="#555",
        )
        self._session_bar_label.pack(side="left", padx=8)
        self._session_bar_toggle = ttk.Button(
            self._session_bar_resize_grip,
            text="Hide",
            width=7,
            command=self._toggle_session_bar,
        )
        self._session_bar_toggle.pack(side="right", padx=6, pady=2)

        self._session_bar_body = ttk.Frame(self._session_bar_frame)
        self._session_bar_body.pack(side="top", fill="both", expand=True)

        self._nb = ttk.Notebook(self._content_frame)
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
        recipe_frame  = ttk.Frame(self._nb)
        bo_frame      = ttk.Frame(self._nb)
        titration_frame = ttk.Frame(self._nb)
        plotter_frame = ttk.Frame(self._nb)

        if PUMP_AVAILABLE:
            self._nb.add(pump_frame,    text="Pump Control")
        self._nb.add(method_frame,  text="Method Creation")
        self._nb.add(script_frame,  text="Script Preview")
        self._nb.add(queue_frame,   text="Queue & Execution")
        self._nb.add(recipe_frame,  text="Recipe Maker")
        self._nb.add(bo_frame,      text="Bayesian Optimization")
        self._nb.add(titration_frame, text="Automated Titration")
        self._nb.add(plotter_frame, text="Plotter")
        self._session_gated_tabs = [
            method_frame,
            script_frame,
            queue_frame,
            recipe_frame,
            bo_frame,
            titration_frame,
            plotter_frame,
        ]
        if PUMP_AVAILABLE:
            self._session_gated_tabs.insert(0, pump_frame)

        # ── Instantiate tabs ──────────────────────────────────────────────────
        method_content = self._scrollable_tab_content(method_frame, min_width=980)
        recipe_content = self._scrollable_tab_content(recipe_frame, min_width=1080)
        pump_content = self._scrollable_tab_content(pump_frame, min_width=980) if PUMP_AVAILABLE else pump_frame

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

        self._recipe_tab = RecipeMakerTab(
            parent_frame = recipe_content,
            on_send_to_queue = self._queue_tab.add_item,
        )
        # Wire session callbacks now that queue tab (with its log widget) exists
        self._session._log    = self._session_mgr.log
        self._session._status = self._set_status

        self._method_tab = MethodTab(
            parent_frame      = method_content,
            session           = self._session,
            on_add_to_queue   = self._queue_tab.add_item,
            on_refresh_queue  = self._queue_tab.refresh,
            on_script_preview = self._script_tab.update,
            on_run_now        = self._run_now,
        )

        self._bo_tab = BayesianOptimizationTab(
            parent_frame      = bo_frame,
            session           = self._session,
            on_add_to_queue   = self._queue_tab.add_item,
            on_refresh_queue  = self._queue_tab.refresh,
            on_script_preview = self._script_tab.update,
            on_run_queue      = self._queue_tab.run_queue,
            on_configure_auto_titration=self._configure_post_bo_titration,
            is_auto_titration_locked=lambda: (
                self._automated_titration_tab.bo_settings_locked()
            ),
            on_bo_finished=self._run_post_bo_titration,
        )
        self._automated_titration_tab = AutomatedTitrationTab(
            parent_frame=titration_frame,
            session=self._session,
            on_get_best_parameters=self._bo_tab.get_best_parameter_groups,
            on_send_to_queue=self._queue_tab.add_item,
            on_run_queue=self._queue_tab.run_from_index,
            on_lock_for_bo=self._return_to_bo_setup,
        )
        self._bo_frame = bo_frame
        self._titration_frame = titration_frame
        self._queue_tab.add_completion_callback(self._bo_tab.on_queue_complete)
        if SESSION_ARCHIVE_DIR is not None:
            self._queue_tab.add_completion_callback(self._archive_session_after_queue)

        if PUMP_AVAILABLE:
            self._pump_tab = PumpTab(
                parent_frame   = pump_content,
                pump_ctrl      = self._pump_ctrl,
                on_add_to_queue= self._queue_tab.add_item,
                root           = self.root,
            )
        else:
            self._pump_tab = None
        # ── Session bar (bottom of window) ───────────────────────────────────────
        self._session_bar = SessionBar(
            root             = self._session_bar_body,
            session_manager  = self._session_mgr,
            on_start_session = self._on_session_started,
        )
        # Give all tabs access to the session manager for require_experiment() guards
        self._session.session_manager = self._session_mgr
        self._session_mgr.status_var.trace_add("write", self._on_session_state_change)
        self._apply_session_gate()

        # Slack bot listener (optional)
        self._slack_bot = None
        if SLACK_ENABLE and SLACK_BOT_TOKEN and SLACK_SIGNING_SECRET:
            def _status_provider():
                status = self._session.get_queue_status()
                status["session_name"] = (
                    self._session_mgr.current_session_path.name
                    if self._session_mgr.current_session_path is not None
                    else None
                )
                status["experiment_name"] = (
                    self._session_mgr.current_experiment_path.name
                    if self._session_mgr.current_experiment_path is not None
                    else None
                )
                return status

            def _slack_command_handler(command: str):
                normalized = (command or "").strip().lower()
                if normalized in ("continue", "resume", "continue queue", "resume queue", "proceed", "ok"):
                    return self._queue_tab.resume_active_alert(normalized)
                if (
                    normalized.startswith("eta")
                    or normalized.startswith("queue eta")
                    or normalized.startswith("remaining")
                    or normalized.startswith("time remaining")
                ):
                    return self._queue_tab.get_slack_eta_text()
                return None

            self._slack_bot = SlackBotServer(
                host="0.0.0.0",
                port=SLACK_PORT,
                signing_secret=SLACK_SIGNING_SECRET,
                notifier=self._session_mgr._slack,
                status_provider=_status_provider,
                command_handler=_slack_command_handler,
                status_image_provider=self._bo_tab.get_slack_q_trend_image,
                log_callback=self._session_mgr.log,
            )
            if not SLACK_ONLY_WHEN_EXPERIMENT:
                self._slack_bot.start()
            self._session_mgr.set_experiment_callbacks(
                on_start=lambda _p: self._on_experiment_start_slack(),
                on_end=lambda _p: self._on_experiment_end_slack(),
            )
    
    # ── Inter-tab wiring helpers ──────────────────────────────────────────────

    def _configure_post_bo_titration(self, enabled, groups):
        if not enabled:
            self._automated_titration_tab.cancel_bo_autotitration()
            return
        self._automated_titration_tab.prepare_for_bo(groups)
        self._nb.select(self._titration_frame)

    def _return_to_bo_setup(self):
        self._bo_tab.select_setup_tab()
        self._nb.select(self._bo_frame)

    def _run_post_bo_titration(self):
        groups = self._bo_tab.get_best_parameter_groups()
        self._nb.select(self._titration_frame)
        self._automated_titration_tab.run_locked_after_bo(groups)

    @staticmethod
    def _scrollable_tab_content(parent, *, min_width=980):
        scroller = ScrollableFrame(parent, min_width=min_width)
        scroller.pack(fill="both", expand=True)
        return scroller.content

    def _log(self, msg: str):
        """Route log messages to the queue tab's log panel."""
        try:
            self._queue_tab._append_log_gui(msg)
        except Exception:
            print(msg)

    def _set_status(self, msg: str):
        try:
            self._queue_tab.set_status(msg)
        except Exception:
            pass

    def _pump_tab_log(self, msg: str):
        try:
            self._session_mgr.log(msg)
        except Exception:
            pass
        if self._pump_tab is not None:
            self._pump_tab.log(msg)

    def _on_session_state_change(self, *_):
        self._apply_session_gate()

    def _archive_session_after_queue(self, _summary):
        session_path = self._session_mgr.current_session_path
        if session_path is None or SESSION_ARCHIVE_DIR is None:
            return
        session_path = Path(session_path)
        destination = Path(SESSION_ARCHIVE_DIR)

        def worker():
            self._session_mgr.log(f"Archiving session after queue: {session_path.name}")
            try:
                uploaded = archive_session(session_path, destination)
                self._session_mgr.log(f"Session archive uploaded: {uploaded}")
            except Exception as exc:
                self._session_mgr.log(
                    f"Session archive upload failed; local data is unchanged: {exc}"
                )

        threading.Thread(target=worker, daemon=True).start()

    def _apply_session_gate(self):
        state = "normal" if self._session_mgr.has_session else "hidden"
        for tab in self._session_gated_tabs:
            self._nb.tab(tab, state=state)
    
    def _on_session_started(self):
        if self._pump_tab is None:
            return
        try:
            if self._pump_ctrl and self._pump_ctrl.connected:
                return
        except Exception:
            pass
        self._pump_tab.autoconnect()

    def _on_experiment_start_slack(self):
        if NGROK_AUTOSTART:
            self._start_ngrok_tunnel()
        if self._slack_bot is not None:
            self._slack_bot.start()

    def _on_experiment_end_slack(self):
        if self._slack_bot is not None:
            self._slack_bot.stop()
        if NGROK_AUTOSTART:
            self._stop_ngrok_tunnel()

    def _start_ngrok_tunnel(self):
        if self._ngrok_proc is not None:
            return
        if not NGROK_PATH:
            self._session_mgr.log("ngrok autostart skipped: EA_NGROK_PATH not set.")
            return
        args = [NGROK_PATH, "http"]
        if NGROK_DOMAIN:
            args.extend(["--domain", NGROK_DOMAIN])
        args.append(str(SLACK_PORT))
        try:
            self._ngrok_proc = threading.Thread  # keep type-checkers calm
            self._ngrok_proc = __import__("subprocess").Popen(
                args,
                stdout=__import__("subprocess").DEVNULL,
                stderr=__import__("subprocess").DEVNULL,
            )
            self._session_mgr.log("ngrok tunnel started.")
        except Exception as exc:
            self._ngrok_proc = None
            self._session_mgr.log(f"ngrok autostart failed: {exc}")

    def _stop_ngrok_tunnel(self):
        proc = self._ngrok_proc
        if proc is None:
            return
        try:
            proc.terminate()
        except Exception:
            pass
        self._ngrok_proc = None

    def _start_session_bar_resize(self, event):
        if self._session_bar_minimized:
            self._set_session_bar_minimized(False)
        self._resize_start_y = event.y_root
        self._resize_start_h = self._session_bar_frame.winfo_height()
        self.root.bind("<B1-Motion>", self._do_session_bar_resize)
        self.root.bind("<ButtonRelease-1>", self._stop_session_bar_resize)

    def _do_session_bar_resize(self, event):
        delta = self._resize_start_y - event.y_root
        new_h = self._resize_start_h + delta
        root_h = max(1, self.root.winfo_height())
        max_h = max(180, root_h - 24)
        new_h = max(self._session_bar_min_height, min(max_h, new_h))
        self._session_bar_frame.configure(height=new_h)

    def _stop_session_bar_resize(self, _event):
        self.root.unbind("<B1-Motion>")
        self.root.unbind("<ButtonRelease-1>")

    def _toggle_session_bar(self):
        self._set_session_bar_minimized(not self._session_bar_minimized)

    def _set_session_bar_minimized(self, minimized: bool):
        self._session_bar_minimized = bool(minimized)
        if self._session_bar_minimized:
            self._session_bar_body.pack_forget()
            self._session_bar_frame.configure(height=self._session_bar_collapsed_height)
            self._session_bar_toggle.configure(text="Show")
        else:
            self._session_bar_body.pack(side="top", fill="both", expand=True)
            self._session_bar_frame.configure(height=self._session_bar_min_height)
            self._session_bar_toggle.configure(text="Hide")

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

        if technique.endswith("_MUX_SEQ"):
            base_script = script_or_base
            if isinstance(extra, dict):
                channels = extra.get("channels", [])
                params = extra.get("params")
            else:
                channels = extra   # list[int]
                params = None
            tech        = technique[: -len("_MUX_SEQ")]
            self._run_mux_sequence(tech, base_script, channels, params=params)

        elif technique == "SWV_CYCLES":
            if isinstance(extra, dict):
                n_scans = extra.get("n_scans")
                delay = extra.get("delay")
                params = extra.get("params")
            else:
                n_scans, delay = extra
                params = None
            self._run_swv_cycles(script_or_base, n_scans, delay, params=params)

        elif technique == "SWV_MUX_CYCLES":
            if isinstance(extra, dict):
                channels = extra.get("channels", [])
                n_scans = extra.get("n_scans")
                delay = extra.get("delay")
                params = extra.get("params")
            else:
                channels, n_scans, delay = extra
                params = None
            self._run_mux_swv_cycles(script_or_base, channels, n_scans, delay, params=params)

        else:
            if isinstance(extra, dict):
                mux_channel = extra.get("mux_channel")
                params = extra.get("params")
            else:
                mux_channel = extra   # int or None
                params = None
            self._run_single(technique, script_or_base, mux_channel, params=params)

    def _require_run_data_folder(self):
        """Return the active experiment folder for immediate measurement runs."""
        if self._session_mgr is None:
            return None
        return self._session_mgr.require_experiment()

    # ── Single run ────────────────────────────────────────────────────────────

    def _run_single(self, technique: str, script: str, mux_channel=None, params=None):
        data_folder = self._require_run_data_folder()
        if data_folder is None:
            return
        try:
            fp, fn = self._session.registry.save_script(
                technique,
                script,
                params=params,
                mux_channel=mux_channel,
            )
        except Exception as exc:
            messagebox.showerror("File Error", f"Failed to save script: {exc}"); return

        self._queue_tab.clear_log()
        self._session.is_running = True
        self._queue_tab.set_status(f"Running: {technique} — {fn}")
        self._plotter_tab.start_live(f"{technique} (live)", label=technique)

        def worker():
            meas_tag = self._session.next_meas_tag_with_mux(mux_channel)
            self._session_mgr.log(f"[Tag] {meas_tag}")
            self.root.after(0, self._queue_tab.refresh_labels)
            runner = SerialMeasurementRunner(
                fp,
                log_callback  = self._session_mgr.log,
                data_callback = self._plotter_tab.push_live_point,
                data_folder = data_folder,
                save_raw_packets = self._session.save_raw_packets,
                simulate_measurements = self._session.simulate_measurements,
                invert_current = (technique == "SWV"),
                device_port = self._session.device_port,
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
                        allow_overlay=False,
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

    def _run_mux_sequence(self, technique: str, base_script: str, channels: list, params=None):
        data_folder = self._require_run_data_folder()
        if data_folder is None:
            return
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
                fp, fn = self._session.registry.save_script(
                    technique,
                    mux_script,
                    params=params,
                    mux_channel=ch,
                )
                color = self._session.next_plot_color()
                label = f"MUX ch {ch}"
                self.root.after(0, self._plotter_tab.start_live,
                                f"{technique} ch {ch} (live)", color, label)
                self.root.after(0, self._queue_tab.set_status,
                                f"Running: {technique} MUX ch {ch}")
                meas_tag = self._session.next_meas_tag_with_mux(ch)
                self._session_mgr.log(f"[Tag] {meas_tag}")
                self.root.after(0, self._queue_tab.refresh_labels)
                runner = SerialMeasurementRunner(
                    fp, log_callback=self._session_mgr.log,
                    data_callback=self._plotter_tab.push_live_point,
                    data_folder=data_folder,
                    save_raw_packets=self._session.save_raw_packets,
                    simulate_measurements=self._session.simulate_measurements,
                    invert_current=(technique == "SWV"),
                    device_port=self._session.device_port)
                self._session.current_runner = runner
                ok, csv_path = runner.execute(meas_tag=meas_tag)
                self._session.current_runner = None
                self.root.after(0, self._plotter_tab.stop_live)
                if csv_path:
                    last_csv = csv_path
                    self.root.after(0, self._plotter_tab.plot_data,
                                   csv_path, color, label, True, False)
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

    def _run_swv_cycles(self, base_script: str, n_scans: int, delay: float, params=None):
        data_folder = self._require_run_data_folder()
        if data_folder is None:
            return
        self._queue_tab.clear_log()
        self._session.is_running = True

        def worker():
            stopped = False; success = True; last_csv = None
            for scan in range(1, n_scans + 1):
                if not self._session.is_running:
                    stopped = True; success = False; break
                fp, fn = self._session.registry.save_script("SWV", base_script, params=params)
                color = self._session.next_plot_color()
                label = f"SWV scan {scan}"
                self.root.after(0, self._plotter_tab.start_live,
                                f"SWV (scan {scan}/{n_scans} live)", color, label)
                self.root.after(0, self._queue_tab.set_status,
                                f"Running: SWV scan {scan}/{n_scans}")
                meas_tag = self._session.next_meas_tag_with_mux(None)
                self._session_mgr.log(f"[Tag] {meas_tag}")
                self.root.after(0, self._queue_tab.refresh_labels)
                runner = SerialMeasurementRunner(
                    fp, log_callback=self._session_mgr.log,
                    data_callback=self._plotter_tab.push_live_point,
                    data_folder=data_folder,
                    save_raw_packets=self._session.save_raw_packets,
                    simulate_measurements=self._session.simulate_measurements,
                    invert_current=True,
                    device_port=self._session.device_port)
                self._session.current_runner = runner
                ok, csv_path = runner.execute(meas_tag=meas_tag)
                self._session.current_runner = None
                self.root.after(0, self._plotter_tab.stop_live)
                if csv_path:
                    last_csv = csv_path
                    self.root.after(0, self._plotter_tab.plot_data,
                                   csv_path, color, label, True, False)
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

    def _run_mux_swv_cycles(self, base_script, channels, n_scans, delay, params=None):
        data_folder = self._require_run_data_folder()
        if data_folder is None:
            return
        self._queue_tab.clear_log()
        self._session.is_running = True

        def worker():
            stopped = False; success = True; last_csv = None
            for scan in range(1, n_scans + 1):
                for ch in channels:
                    if not self._session.is_running:
                        stopped = True; success = False; break
                    mux_script = self._method_tab._wrap_mux(base_script, ch)
                    fp, fn = self._session.registry.save_script(
                        "SWV",
                        mux_script,
                        params=params,
                        mux_channel=ch,
                    )
                    color = self._session.next_plot_color()
                    label = f"MUX ch {ch} scan {scan}"
                    self.root.after(0, self._plotter_tab.start_live,
                                    f"SWV MUX ch {ch} ({scan}/{n_scans})", color, label)
                    self.root.after(0, self._queue_tab.set_status,
                                    f"Running: SWV MUX ch {ch} scan {scan}/{n_scans}")
                    meas_tag = self._session.next_meas_tag_with_mux(ch)
                    self._session_mgr.log(f"[Tag] {meas_tag}")
                    self.root.after(0, self._queue_tab.refresh_labels)
                    runner = SerialMeasurementRunner(
                        fp, log_callback=self._session_mgr.log,
                        data_callback=self._plotter_tab.push_live_point,
                        data_folder=data_folder,
                        save_raw_packets=self._session.save_raw_packets,
                        simulate_measurements=self._session.simulate_measurements,
                        invert_current=True,
                        device_port=self._session.device_port)
                    self._session.current_runner = runner
                    ok, csv_path = runner.execute(meas_tag=meas_tag)
                    self._session.current_runner = None
                    self.root.after(0, self._plotter_tab.stop_live)
                    if csv_path:
                        last_csv = csv_path
                        self.root.after(0, self._plotter_tab.plot_data,
                                       csv_path, color, label, True, False)
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
