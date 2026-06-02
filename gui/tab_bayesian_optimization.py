"""
gui/tab_bayesian_optimization.py - Optional SWV Bayesian optimization tab.

This is a UI shell around core.bo_session. It edits configuration, requests BO
suggestions, queues normal SWV methods, imports external analysis outputs, and
shows records. BO math lives in core.bo_session.
"""

from __future__ import annotations

import json
import time
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog
import tkinter as tk
from tkinter import ttk, scrolledtext

from config import (
    BO_ANALYSIS_APP_PATH,
    BO_ANALYSIS_FILE_GLOB,
    BO_ANALYSIS_OUTPUT_DIR,
    BO_ANALYSIS_POLL_SECONDS,
    BO_DEFAULT_CONFIG_PATH,
    BO_LOCAL_PATHS_CONFIG,
)
from core.bo_session import (
    BOIntegrationSession,
    PARAMETER_ORDER,
    active_parameters,
    build_swv_script,
    load_bo_config,
    normalize_bo_config,
    parse_channels,
    resolve_initial_parameters,
    save_bo_config,
    validate_bo_config,
)


class BayesianOptimizationTab:
    """Optional closed-loop Bayesian optimization UI."""

    ACCENT = "#155e63"
    ACCENT_DARK = "#0f3d44"
    ACCENT_LIGHT = "#dff7f5"

    def __init__(
        self,
        parent_frame,
        session,
        on_add_to_queue,
        on_refresh_queue,
        on_script_preview,
        on_run_queue=None,
    ):
        self._frame = parent_frame
        self._session = session
        self._add_to_queue = on_add_to_queue
        self._refresh_queue = on_refresh_queue
        self._script_preview = on_script_preview
        self._run_queue = on_run_queue

        self._config_path_var = tk.StringVar(value=str(BO_DEFAULT_CONFIG_PATH))
        self._analysis_dir_var = tk.StringVar(value=str(BO_ANALYSIS_OUTPUT_DIR))
        self._analysis_app_var = tk.StringVar(value=str(BO_ANALYSIS_APP_PATH or ""))
        self._analysis_glob_var = tk.StringVar(value=str(BO_ANALYSIS_FILE_GLOB))
        self._channels_var = tk.StringVar(value="")
        self._status_var = tk.StringVar(value="Load a BO config to begin.")
        self._record_dir_var = tk.StringVar(value="Record folder: (not started)")
        self._auto_target_var = tk.StringVar(value="5")
        self._auto_poll_var = tk.StringVar(value=str(BO_ANALYSIS_POLL_SECONDS))
        self._auto_status_var = tk.StringVar(value="Auto loop idle.")

        self._config = None
        self._bo_session = None
        self._suggestion = None
        self._auto_running = False
        self._auto_poll_after_id = None
        self._auto_analysis_cutoff = 0.0

        self._build()
        self._load_config(initial=True)

    def _build(self):
        root = ttk.Frame(self._frame)
        root.pack(fill="both", expand=True)

        banner = tk.Frame(root, bg=self.ACCENT_DARK, height=58)
        banner.pack(side="top", fill="x")
        banner.pack_propagate(False)
        tk.Label(
            banner,
            text="Bayesian Optimization",
            bg=self.ACCENT_DARK,
            fg="white",
            font=("Arial", 16, "bold"),
        ).pack(side="left", padx=(16, 10), pady=12)
        tk.Label(
            banner,
            text="Closed-loop SWV method search across the mux array",
            bg=self.ACCENT_DARK,
            fg=self.ACCENT_LIGHT,
            font=("Arial", 10),
        ).pack(side="left", padx=8, pady=15)

        self._tabs = ttk.Notebook(root)
        self._tabs.pack(fill="both", expand=True, padx=8, pady=8)

        setup = ttk.Frame(self._tabs)
        run = ttk.Frame(self._tabs)
        results = ttk.Frame(self._tabs)
        self._tabs.add(setup, text="Setup")
        self._tabs.add(run, text="Run")
        self._tabs.add(results, text="Results & Records")

        self._build_setup_tab(setup)
        self._build_run_tab(run)
        self._build_results_tab(results)

        status = ttk.Label(root, textvariable=self._status_var, relief="sunken")
        status.pack(side="bottom", fill="x", padx=8, pady=(0, 8))

    def _build_setup_tab(self, parent):
        pane = ttk.PanedWindow(parent, orient=tk.HORIZONTAL)
        pane.pack(fill="both", expand=True, padx=4, pady=4)
        left = ttk.Frame(pane)
        right = ttk.Frame(pane)
        pane.add(left, weight=2)
        pane.add(right, weight=3)

        cfg = ttk.LabelFrame(left, text="Configuration Files and Paths", padding=8)
        cfg.pack(fill="x", pady=(0, 8))
        cfg.columnconfigure(1, weight=1)
        ttk.Label(cfg, text="BO config:").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Entry(cfg, textvariable=self._config_path_var).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(cfg, text="Browse", command=self._browse_config).grid(row=0, column=2, padx=2)
        ttk.Button(cfg, text="Load", command=self._load_config).grid(row=0, column=3, padx=2)
        ttk.Button(cfg, text="Save", command=self._save_config).grid(row=0, column=4, padx=2)

        ttk.Label(cfg, text="Analysis output:").grid(row=1, column=0, sticky="w", pady=2)
        ttk.Entry(cfg, textvariable=self._analysis_dir_var).grid(row=1, column=1, sticky="ew", padx=4)
        ttk.Button(cfg, text="Browse", command=self._browse_analysis_dir).grid(row=1, column=2, padx=2)
        ttk.Button(cfg, text="Save Paths", command=self._save_local_paths).grid(row=1, column=3, columnspan=2, padx=2)

        ttk.Label(cfg, text="Analysis app:").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Entry(cfg, textvariable=self._analysis_app_var).grid(row=2, column=1, sticky="ew", padx=4)
        ttk.Button(cfg, text="Browse", command=self._browse_analysis_app).grid(row=2, column=2, padx=2)
        ttk.Label(cfg, text="Glob:").grid(row=2, column=3, sticky="e", padx=2)
        ttk.Entry(cfg, textvariable=self._analysis_glob_var, width=10).grid(row=2, column=4, sticky="w")

        ttk.Label(cfg, text="Mux channels:").grid(row=3, column=0, sticky="w", pady=2)
        channels = ttk.Entry(cfg, textvariable=self._channels_var)
        channels.grid(row=3, column=1, sticky="ew", padx=4)
        channels.bind("<FocusOut>", lambda _e: self._sync_channels_from_entry())
        channels.bind("<Return>", lambda _e: self._sync_channels_from_entry())
        ttk.Button(cfg, text="Validate", command=self._validate_config).grid(row=3, column=2, padx=2)
        ttk.Button(cfg, text="Start BO Session", command=self._start_bo_session).grid(row=3, column=3, columnspan=2, padx=2)

        clue = ttk.LabelFrame(left, text="Setup Cues", padding=8)
        clue.pack(fill="x", pady=(0, 8))
        ttk.Label(
            clue,
            text=(
                "Use the BO config for scientific choices: active parameters, hard constraints, "
                "initial design, channels, and scoring. Use Save Paths for machine-specific "
                "folders so you do not need to edit shell startup files."
            ),
            wraplength=460,
            justify="left",
        ).pack(fill="x")

        init_box = ttk.LabelFrame(left, text="Initial Parameters", padding=8)
        init_box.pack(fill="both", expand=True)
        init_toolbar = ttk.Frame(init_box)
        init_toolbar.pack(fill="x", pady=(0, 6))
        ttk.Button(init_toolbar, text="Edit Initial Parameters", command=self._edit_initial_parameters).pack(side="left", padx=2)
        ttk.Label(
            init_toolbar,
            text="BO runs this method first, then chooses additional early points automatically.",
            foreground=self.ACCENT,
        ).pack(side="left", padx=8)

        init_cols = ("Begin", "End", "Step", "Amp", "Freq", "Cond E", "Cond t")
        self._initial_tree = ttk.Treeview(init_box, columns=init_cols, show="tree headings", height=8)
        self._initial_tree.heading("#0", text="#")
        self._initial_tree.column("#0", width=38, anchor="center")
        for col in init_cols:
            self._initial_tree.heading(col, text=col)
            self._initial_tree.column(col, width=70, anchor="center")
        self._initial_tree.pack(fill="both", expand=True)
        self._initial_tree.bind("<Double-1>", lambda _e: self._edit_initial_parameters())

        params = ttk.LabelFrame(right, text="Parameter Space", padding=8)
        params.pack(fill="both", expand=True)
        toolbar = ttk.Frame(params)
        toolbar.pack(fill="x", pady=(0, 6))
        ttk.Button(toolbar, text="Edit Selected", command=self._edit_selected_parameter).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Active", command=lambda: self._set_selected_mode("active")).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Locked", command=lambda: self._set_selected_mode("locked")).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Tied", command=lambda: self._set_selected_mode("tied")).pack(side="left", padx=2)

        legend = ttk.Frame(params)
        legend.pack(fill="x", pady=(0, 6))
        tk.Label(legend, text="Active = optimized", bg="#dff7f5", fg="#0f3d44", padx=6).pack(side="left", padx=2)
        tk.Label(legend, text="Locked = fixed", bg="#eeeeee", fg="#333333", padx=6).pack(side="left", padx=2)
        tk.Label(legend, text="Tied = follows another", bg="#fff1d6", fg="#6b3b00", padx=6).pack(side="left", padx=2)

        cols = ("Mode", "Values", "Tie")
        self._param_tree = ttk.Treeview(params, columns=cols, show="tree headings", height=16)
        self._param_tree.heading("#0", text="Parameter")
        self._param_tree.heading("Mode", text="Mode")
        self._param_tree.heading("Values", text="Values / Lock")
        self._param_tree.heading("Tie", text="Tie")
        self._param_tree.column("#0", width=165)
        self._param_tree.column("Mode", width=75)
        self._param_tree.column("Values", width=330)
        self._param_tree.column("Tie", width=110)
        self._param_tree.tag_configure("active", background="#dff7f5", foreground="#0f3d44")
        self._param_tree.tag_configure("locked", background="#eeeeee", foreground="#333333")
        self._param_tree.tag_configure("tied", background="#fff1d6", foreground="#6b3b00")
        self._param_tree.pack(fill="both", expand=True)
        self._param_tree.bind("<Double-1>", lambda _e: self._edit_selected_parameter())

    def _build_run_tab(self, parent):
        controls = ttk.LabelFrame(parent, text="Manual Closed Loop", padding=8)
        controls.pack(fill="x", padx=4, pady=(4, 8))
        ttk.Button(controls, text="Suggest Next Method", command=self._suggest_next).pack(side="left", padx=3)
        ttk.Button(controls, text="Send Batch to Queue", command=self._send_to_queue).pack(side="left", padx=3)
        ttk.Button(controls, text="Preview Script", command=self._preview_suggestion).pack(side="left", padx=3)
        ttk.Separator(controls, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Button(controls, text="Import Analysis JSON", command=self._import_analysis_dialog).pack(side="left", padx=3)
        ttk.Button(controls, text="Use Latest Analysis", command=self._import_latest_analysis).pack(side="left", padx=3)

        auto = ttk.LabelFrame(parent, text="Auto Loop", padding=8)
        auto.pack(fill="x", padx=4, pady=(0, 8))
        ttk.Label(auto, text="Target completed iterations:").pack(side="left", padx=(0, 4))
        ttk.Entry(auto, textvariable=self._auto_target_var, width=6).pack(side="left", padx=(0, 10))
        ttk.Label(auto, text="Analysis poll (s):").pack(side="left", padx=(0, 4))
        ttk.Entry(auto, textvariable=self._auto_poll_var, width=6).pack(side="left", padx=(0, 10))
        ttk.Button(auto, text="Start Auto Loop", command=self._start_auto_loop).pack(side="left", padx=3)
        ttk.Button(auto, text="Stop Auto", command=self._stop_auto_loop).pack(side="left", padx=3)
        ttk.Label(auto, textvariable=self._auto_status_var, foreground=self.ACCENT).pack(side="left", padx=12)

        clue = ttk.LabelFrame(parent, text="Workflow Cues", padding=8)
        clue.pack(fill="x", padx=4, pady=(0, 8))
        ttk.Label(
            clue,
            text=(
                "Manual: suggest -> queue -> run queue -> import analysis. Auto: queue must be empty; "
                "the tab runs one BO batch at a time and imports the newest unused analysis JSON."
            ),
            wraplength=920,
            justify="left",
        ).pack(fill="x")

        ttk.Label(parent, textvariable=self._record_dir_var, foreground=self.ACCENT).pack(fill="x", padx=8, pady=(0, 6))

        suggest_box = ttk.LabelFrame(parent, text="Current Suggested Method", padding=6)
        suggest_box.pack(fill="both", expand=True, padx=4, pady=(0, 4))
        self._suggestion_text = scrolledtext.ScrolledText(suggest_box, height=16, wrap=tk.WORD)
        self._suggestion_text.pack(fill="both", expand=True)
        self._suggestion_text.config(state="disabled")

    def _build_results_tab(self, parent):
        pane = ttk.PanedWindow(parent, orient=tk.VERTICAL)
        pane.pack(fill="both", expand=True, padx=4, pady=4)
        top = ttk.PanedWindow(pane, orient=tk.HORIZONTAL)
        middle = ttk.PanedWindow(pane, orient=tk.HORIZONTAL)
        bottom = ttk.PanedWindow(pane, orient=tk.HORIZONTAL)
        pane.add(top, weight=1)
        pane.add(middle, weight=1)
        pane.add(bottom, weight=1)

        score_box = ttk.LabelFrame(top, text="Per-Channel Scores", padding=6)
        best_box = ttk.LabelFrame(top, text="Best Method So Far", padding=6)
        top.add(score_box, weight=1)
        top.add(best_box, weight=1)

        score_cols = ("Q", "SNR", "Shape", "Baseline", "Replicate", "Success")
        self._score_tree = ttk.Treeview(score_box, columns=score_cols, show="tree headings", height=10)
        self._score_tree.heading("#0", text="Ch")
        self._score_tree.column("#0", width=50, anchor="center")
        for col in score_cols:
            self._score_tree.heading(col, text=col)
            self._score_tree.column(col, width=82, anchor="center")
        self._score_tree.pack(fill="both", expand=True)

        self._best_text = scrolledtext.ScrolledText(best_box, height=10, wrap=tk.WORD)
        self._best_text.pack(fill="both", expand=True)
        self._best_text.config(state="disabled")

        hist_box = ttk.LabelFrame(middle, text="BO History", padding=6)
        model_box = ttk.LabelFrame(middle, text="Surrogate and Acquisition Artifacts", padding=6)
        middle.add(hist_box, weight=1)
        middle.add(model_box, weight=1)
        hist_cols = ("Q_run", "Mean", "Std", "Failed", "Low")
        self._history_tree = ttk.Treeview(hist_box, columns=hist_cols, show="tree headings", height=10)
        self._history_tree.heading("#0", text="Iter")
        self._history_tree.column("#0", width=55, anchor="center")
        for col in hist_cols:
            self._history_tree.heading(col, text=col)
            self._history_tree.column(col, width=90, anchor="center")
        self._history_tree.pack(fill="both", expand=True)

        cols = ("Type", "File")
        model_toolbar = ttk.Frame(model_box)
        model_toolbar.pack(fill="x", pady=(0, 4))
        ttk.Button(model_toolbar, text="Refresh", command=self._refresh_model_artifacts).pack(side="left", padx=2)
        self._model_tree = ttk.Treeview(model_box, columns=cols, show="tree headings", height=8)
        self._model_tree.heading("#0", text="#")
        self._model_tree.heading("Type", text="Type")
        self._model_tree.heading("File", text="File")
        self._model_tree.column("#0", width=45, anchor="center")
        self._model_tree.column("Type", width=115)
        self._model_tree.column("File", width=420)
        self._model_tree.pack(fill="both", expand=True)

        record_box = ttk.LabelFrame(bottom, text="BO Session Records", padding=6)
        bottom.add(record_box, weight=1)
        record_toolbar = ttk.Frame(record_box)
        record_toolbar.pack(fill="x", pady=(0, 4))
        ttk.Label(record_toolbar, textvariable=self._record_dir_var, foreground=self.ACCENT).pack(side="left", padx=2)
        ttk.Button(record_toolbar, text="Refresh", command=self._refresh_record_files).pack(side="right", padx=2)
        cols = ("Folder", "File")
        self._record_tree = ttk.Treeview(record_box, columns=cols, show="tree headings", height=8)
        self._record_tree.heading("#0", text="#")
        self._record_tree.heading("Folder", text="Folder")
        self._record_tree.heading("File", text="File")
        self._record_tree.column("#0", width=45, anchor="center")
        self._record_tree.column("Folder", width=150)
        self._record_tree.column("File", width=430)
        self._record_tree.pack(fill="both", expand=True)

    # Config and path actions
    def _browse_config(self):
        path = filedialog.askopenfilename(
            title="Choose BO config",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialfile=Path(self._config_path_var.get()).name,
        )
        if path:
            self._config_path_var.set(path)
            self._load_config()

    def _browse_analysis_dir(self):
        path = filedialog.askdirectory(title="Choose external analysis output folder")
        if path:
            self._analysis_dir_var.set(path)

    def _browse_analysis_app(self):
        path = filedialog.askopenfilename(title="Choose external analysis app")
        if path:
            self._analysis_app_var.set(path)

    def _save_local_paths(self):
        try:
            payload = {
                "analysis_output_dir": self._analysis_dir_var.get().strip(),
                "analysis_app_path": self._analysis_app_var.get().strip(),
                "analysis_file_glob": self._analysis_glob_var.get().strip() or "*.json",
                "analysis_poll_seconds": float(self._auto_poll_var.get() or BO_ANALYSIS_POLL_SECONDS),
            }
            BO_LOCAL_PATHS_CONFIG.parent.mkdir(parents=True, exist_ok=True)
            with open(BO_LOCAL_PATHS_CONFIG, "w", encoding="utf-8") as fh:
                json.dump(payload, fh, indent=2)
            self._status_var.set(f"Saved BO local paths: {BO_LOCAL_PATHS_CONFIG}")
        except Exception as exc:
            messagebox.showerror("Save BO Paths", str(exc))

    def _load_config(self, initial=False):
        try:
            self._config = load_bo_config(self._config_path_var.get())
            analysis_cfg = self._config.get("analysis", {})
            if analysis_cfg.get("file_glob"):
                self._analysis_glob_var.set(str(analysis_cfg.get("file_glob")))
            self._channels_var.set(", ".join(str(ch) for ch in self._config.get("channels", [])))
            self._refresh_parameter_table()
            self._refresh_initial_parameters_table()
            self._validate_config(show_dialog=False)
            if not initial:
                self._status_var.set(f"Loaded BO config: {self._config_path_var.get()}")
        except Exception as exc:
            self._config = None
            self._status_var.set(f"BO config load failed: {exc}")
            if not initial:
                messagebox.showerror("BO Config", str(exc))

    def _save_config(self):
        if self._config is None:
            return
        self._sync_channels_from_entry(show_error=False)
        self._config.setdefault("analysis", {})["file_glob"] = self._analysis_glob_var.get().strip() or "*.json"
        try:
            path = save_bo_config(self._config, self._config_path_var.get())
            self._status_var.set(f"Saved BO config: {path}")
        except Exception as exc:
            messagebox.showerror("Save BO Config", str(exc))

    def _sync_channels_from_entry(self, show_error=True):
        if self._config is None:
            return
        try:
            channels = parse_channels(self._channels_var.get())
            self._config["channels"] = channels
            self._channels_var.set(", ".join(str(ch) for ch in channels))
        except Exception as exc:
            if show_error:
                messagebox.showerror("Mux Channels", str(exc))

    def _validate_config(self, show_dialog=True):
        if self._config is None:
            return
        self._sync_channels_from_entry(show_error=False)
        errors = validate_bo_config(self._config)
        active = ", ".join(active_parameters(self._config)) or "(none)"
        if errors:
            msg = "Config invalid: " + "; ".join(errors)
            self._status_var.set(msg)
            if show_dialog:
                messagebox.showerror("BO Config", msg)
            return
        msg = f"Config valid. Active parameters: {active}"
        self._status_var.set(msg)
        if show_dialog:
            messagebox.showinfo("BO Config", msg)

    # Session and run actions
    def _start_bo_session(self):
        if self._config is None:
            messagebox.showwarning("BO Session", "Load a BO config first.")
            return
        session_mgr = getattr(self._session, "session_manager", None)
        exp_path = session_mgr.require_experiment() if session_mgr is not None else None
        if exp_path is None:
            return
        self._save_config()
        self._save_local_paths()
        try:
            self._bo_session = BOIntegrationSession.start(
                self._config_path_var.get(),
                exp_path,
                analysis_output_dir=self._analysis_dir_var.get(),
            )
            self._suggestion = None
            self._record_dir_var.set(f"Record folder: {self._bo_session.record_dir}")
            self._status_var.set(f"BO session started with {len(self._bo_session.candidates)} valid candidates.")
            self._refresh_history()
            self._render_best()
            self._refresh_model_artifacts()
            self._refresh_record_files()
            self._tabs.select(1)
        except Exception as exc:
            messagebox.showerror("Start BO Session", str(exc))

    def _suggest_next(self):
        if self._bo_session is None:
            self._start_bo_session()
            if self._bo_session is None:
                return
        try:
            self._suggestion = self._bo_session.ask_next()
            self._render_suggestion()
            self._status_var.set(
                f"Suggested BO iteration {self._suggestion.iteration}. Send it to queue when ready."
            )
        except Exception as exc:
            messagebox.showerror("BO Suggestion", str(exc))

    def _send_to_queue(self):
        if self._bo_session is None:
            messagebox.showwarning("BO Queue", "Start a BO session first.")
            return
        if self._suggestion is None:
            self._suggest_next()
            if self._suggestion is None:
                return
        try:
            items = self._bo_session.build_queue_items(self._session.registry, self._suggestion)
            for item in items:
                self._add_to_queue(item)
            self._bo_session.record_queued(self._suggestion, items)
            self._refresh_queue()
            self._refresh_record_files()
            self._status_var.set(
                f"Queued BO iteration {self._suggestion.iteration} for {len(items)} mux channel(s)."
            )
        except Exception as exc:
            messagebox.showerror("BO Queue", str(exc))

    def _preview_suggestion(self):
        if self._suggestion is None:
            self._suggest_next()
            if self._suggestion is None:
                return
        script = build_swv_script(self._suggestion.params, self._bo_session.config.get("method_options", {}))
        self._script_preview(script)
        self._status_var.set("Previewed the BO base SWV script in Script Preview.")

    def _import_analysis_dialog(self):
        if self._bo_session is None:
            messagebox.showwarning("BO Analysis", "Start a BO session first.")
            return
        path = filedialog.askopenfilename(
            title="Import external analysis JSON",
            initialdir=self._analysis_dir_var.get(),
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if path:
            self._import_analysis(path)

    def _import_latest_analysis(self):
        if self._bo_session is None:
            messagebox.showwarning("BO Analysis", "Start a BO session first.")
            return
        path = self._bo_session.latest_analysis_file()
        if path is None:
            messagebox.showwarning("BO Analysis", "No analysis JSON file found in the configured folder.")
            return
        self._import_analysis(path)

    def _import_analysis(self, path, notes=None, prompt=True):
        if prompt:
            notes = simpledialog.askstring("BO Analysis Notes", "Notes for this BO result:", parent=self._frame)
        try:
            obs = self._bo_session.import_analysis(path, notes=notes or "")
            self._suggestion = None
            self._render_scores(obs)
            self._refresh_history()
            self._render_best()
            self._refresh_model_artifacts()
            self._refresh_record_files()
            self._clear_text(self._suggestion_text)
            self._status_var.set(
                f"Imported analysis for iteration {obs['iteration']}. Q_run={obs['Q_run']:.3f}"
            )
            return obs
        except Exception as exc:
            if not prompt:
                self._auto_running = False
                self._auto_status_var.set(f"Auto loop stopped: analysis import failed ({exc})")
            messagebox.showerror("Import BO Analysis", str(exc))
            return None

    # Auto loop
    def _start_auto_loop(self):
        if self._run_queue is None:
            messagebox.showwarning("Auto Loop", "Queue runner is not wired for automation.")
            return
        if self._session.is_running:
            messagebox.showwarning("Auto Loop", "The measurement queue is already running.")
            return
        if self._session.measurement_queue:
            messagebox.showwarning(
                "Auto Loop",
                "Auto loop starts only from an empty queue so it cannot run unrelated items.",
            )
            return
        try:
            target = int(self._auto_target_var.get())
            poll = float(self._auto_poll_var.get())
        except ValueError:
            messagebox.showerror("Auto Loop", "Target iterations and poll interval must be numeric.")
            return
        if target < 1:
            messagebox.showerror("Auto Loop", "Target iterations must be at least 1.")
            return
        if poll < 1:
            messagebox.showerror("Auto Loop", "Analysis poll interval must be at least 1 second.")
            return
        if self._bo_session is None:
            self._start_bo_session()
            if self._bo_session is None:
                return
        self._auto_running = True
        self._auto_status_var.set(f"Auto loop running toward {target} completed iteration(s).")
        self._auto_submit_next()

    def _stop_auto_loop(self):
        self._auto_running = False
        if self._auto_poll_after_id is not None:
            try:
                self._frame.after_cancel(self._auto_poll_after_id)
            except Exception:
                pass
            self._auto_poll_after_id = None
        self._auto_status_var.set("Auto loop stopped.")

    def _auto_submit_next(self):
        if not self._auto_running:
            return
        target = int(self._auto_target_var.get())
        completed = len(self._bo_session.observations) if self._bo_session else 0
        if completed >= target:
            self._auto_running = False
            self._auto_status_var.set(f"Auto loop complete: {completed}/{target} iteration(s).")
            return
        if self._session.is_running:
            return
        if self._session.measurement_queue:
            if not self._clear_auto_queue_if_safe():
                self._auto_running = False
                self._auto_status_var.set("Auto loop stopped: queue contains non-BO items.")
                return
        try:
            self._suggestion = self._bo_session.ask_next()
            self._render_suggestion()
            items = self._bo_session.build_queue_items(self._session.registry, self._suggestion)
            for item in items:
                self._add_to_queue(item)
            self._bo_session.record_queued(self._suggestion, items)
            self._refresh_queue()
            self._refresh_record_files()
            self._auto_status_var.set(
                f"Queued BO iteration {self._suggestion.iteration}; starting queue."
            )
            self._auto_analysis_cutoff = time.time()
            self._run_queue()
        except Exception as exc:
            self._auto_running = False
            self._auto_status_var.set(f"Auto loop stopped: {exc}")
            messagebox.showerror("Auto Loop", str(exc))

    def on_queue_complete(self, summary):
        if self._bo_session is not None:
            self._bo_session.record_queue_completion(summary)
            self._refresh_record_files()
        if not self._auto_running:
            return
        if summary.get("failed", 0) or summary.get("stopped", 0):
            self._auto_running = False
            self._auto_status_var.set("Auto loop stopped: queue did not complete cleanly.")
            return
        self._auto_status_var.set("Queue complete. Waiting for external analysis JSON.")
        self._schedule_auto_analysis_poll()

    def _schedule_auto_analysis_poll(self):
        if not self._auto_running:
            return
        try:
            poll_ms = max(1000, int(float(self._auto_poll_var.get()) * 1000))
        except ValueError:
            poll_ms = 5000
        self._auto_poll_after_id = self._frame.after(poll_ms, self._poll_auto_analysis)

    def _poll_auto_analysis(self):
        self._auto_poll_after_id = None
        if not self._auto_running or self._bo_session is None:
            return
        path = self._latest_unused_analysis_file()
        if path is None:
            self._auto_status_var.set("Waiting for new analysis JSON...")
            self._schedule_auto_analysis_poll()
            return
        obs = self._import_analysis(path, notes="Imported by BO auto loop", prompt=False)
        if obs is None or not self._auto_running:
            return
        if self._clear_auto_queue_if_safe():
            self._auto_submit_next()

    def _latest_unused_analysis_file(self):
        path = self._bo_session.latest_analysis_file()
        if path is None:
            return None
        try:
            if path.stat().st_mtime < self._auto_analysis_cutoff:
                return None
        except OSError:
            return None
        used = set()
        for obs in self._bo_session.observations:
            used.add(str(obs.get("analysis_source")))
            used.add(str(obs.get("analysis_record")))
        if str(path) in used:
            return None
        return path

    def _clear_auto_queue_if_safe(self):
        queue = self._session.measurement_queue
        if not queue:
            return True
        session_id = self._bo_session.session_id if self._bo_session else None
        for item in queue:
            ref = item.get("bo_ref") or {}
            if ref.get("session_id") != session_id:
                return False
        queue.clear()
        self._refresh_queue()
        return True

    # Parameter and initial-design editing
    def _refresh_parameter_table(self):
        for row in self._param_tree.get_children():
            self._param_tree.delete(row)
        if self._config is None:
            return
        params = normalize_bo_config(self._config)["parameters"]
        for name in PARAMETER_ORDER:
            p = params[name]
            mode = str(p.get("mode", "locked"))
            value_text = self._values_text(p)
            tie = str(p.get("tie_to", "")) if mode == "tied" else ""
            label = p.get("label") or name
            self._param_tree.insert(
                "",
                "end",
                iid=name,
                text=label,
                values=(mode, value_text, tie),
                tags=(mode.lower(),),
            )

    @staticmethod
    def _values_text(param_cfg):
        mode = str(param_cfg.get("mode", "locked")).lower()
        if mode == "active":
            return ", ".join(str(v) for v in param_cfg.get("values", []))
        if mode == "tied":
            return ""
        return str(param_cfg.get("value", ""))

    def _edit_selected_parameter(self):
        if self._config is None:
            return
        selection = self._param_tree.selection()
        if not selection:
            messagebox.showwarning("Parameter Space", "Select a parameter first.")
            return
        name = selection[0]
        params = self._config.setdefault("parameters", {})
        current = dict(params.get(name) or {})

        win = tk.Toplevel(self._frame)
        win.title(f"Edit {name}")
        win.transient(self._frame)
        win.resizable(False, False)
        box = ttk.Frame(win, padding=12)
        box.pack(fill="both", expand=True)
        mode_var = tk.StringVar(value=str(current.get("mode", "locked")))
        values_var = tk.StringVar(value=", ".join(str(v) for v in current.get("values", [])))
        value_var = tk.StringVar(value=str(current.get("value", "")))
        tie_var = tk.StringVar(value=str(current.get("tie_to", "begin_potential")))

        ttk.Label(box, text="Mode:").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Combobox(box, textvariable=mode_var, values=("active", "locked", "tied"), state="readonly", width=16).grid(
            row=0, column=1, sticky="w", pady=4
        )
        ttk.Label(box, text="Active values:").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Entry(box, textvariable=values_var, width=48).grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Label(box, text="Locked value:").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(box, textvariable=value_var, width=18).grid(row=2, column=1, sticky="w", pady=4)
        ttk.Label(box, text="Tie to:").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Combobox(box, textvariable=tie_var, values=PARAMETER_ORDER, width=24).grid(row=3, column=1, sticky="w", pady=4)

        buttons = ttk.Frame(box)
        buttons.grid(row=4, column=0, columnspan=2, pady=(10, 0))

        def save():
            try:
                updated = dict(current)
                updated["mode"] = mode_var.get()
                updated["values"] = self._parse_float_list(values_var.get())
                if value_var.get().strip():
                    updated["value"] = float(value_var.get())
                updated["tie_to"] = tie_var.get()
                params[name] = updated
                self._config["parameters"] = params
                self._refresh_parameter_table()
                self._refresh_initial_parameters_table()
                self._validate_config(show_dialog=False)
                win.destroy()
            except Exception as exc:
                messagebox.showerror("Parameter Space", str(exc), parent=win)

        ttk.Button(buttons, text="Save", command=save).pack(side="left", padx=4)
        ttk.Button(buttons, text="Cancel", command=win.destroy).pack(side="left", padx=4)
        win.grab_set()
        win.focus_force()

    def _set_selected_mode(self, mode):
        if self._config is None:
            return
        selection = self._param_tree.selection()
        if not selection:
            return
        name = selection[0]
        self._config.setdefault("parameters", {}).setdefault(name, {})["mode"] = mode
        if mode == "tied" and name == "conditioning_potential":
            self._config["parameters"][name]["tie_to"] = "begin_potential"
        self._refresh_parameter_table()
        self._refresh_initial_parameters_table()
        self._validate_config(show_dialog=False)

    def _refresh_initial_parameters_table(self):
        for row in self._initial_tree.get_children():
            self._initial_tree.delete(row)
        if self._config is None:
            return
        try:
            point = resolve_initial_parameters(self._config)
        except Exception:
            return
        self._initial_tree.insert(
            "",
            "end",
            iid="0",
            text="1",
            values=(
                self._fmt_raw(point.get("begin_potential")),
                self._fmt_raw(point.get("end_potential")),
                self._fmt_raw(point.get("step_potential")),
                self._fmt_raw(point.get("amplitude")),
                self._fmt_raw(point.get("frequency")),
                self._fmt_raw(point.get("conditioning_potential")),
                self._fmt_raw(point.get("conditioning_time")),
            ),
        )

    def _edit_initial_parameters(self):
        if self._config is None:
            return
        current = resolve_initial_parameters(self._config)
        self._open_method_editor(
            "Edit Initial Parameters",
            current,
            lambda updated: self._save_initial_parameters(updated),
        )

    def _save_initial_parameters(self, updated):
        self._config["initial_parameters"] = updated
        self._config.pop("initial_method", None)
        self._config.pop("initial_design", None)
        self._refresh_initial_parameters_table()
        self._validate_config(show_dialog=False)

    def _open_method_editor(self, title, values, on_save):
        win = tk.Toplevel(self._frame)
        win.title(title)
        win.transient(self._frame)
        win.resizable(False, False)
        box = ttk.Frame(win, padding=12)
        box.pack(fill="both", expand=True)
        vars_by_name = {}
        labels = {
            "begin_potential": "Begin potential (V)",
            "end_potential": "End potential (V)",
            "step_potential": "Step potential (V)",
            "amplitude": "Amplitude (V)",
            "frequency": "Frequency (Hz)",
            "conditioning_potential": "Conditioning potential (V)",
            "conditioning_time": "Conditioning time (s)",
        }
        for row, name in enumerate(PARAMETER_ORDER):
            ttk.Label(box, text=labels.get(name, name)).grid(row=row, column=0, sticky="w", pady=3)
            var = tk.StringVar(value=str(values.get(name, "")))
            vars_by_name[name] = var
            ttk.Entry(box, textvariable=var, width=18).grid(row=row, column=1, sticky="w", pady=3)
        buttons = ttk.Frame(box)
        buttons.grid(row=len(PARAMETER_ORDER), column=0, columnspan=2, pady=(10, 0))

        def save():
            try:
                updated = {name: float(var.get()) for name, var in vars_by_name.items()}
                on_save(updated)
                win.destroy()
            except Exception as exc:
                messagebox.showerror(title, str(exc), parent=win)

        ttk.Button(buttons, text="Save", command=save).pack(side="left", padx=4)
        ttk.Button(buttons, text="Cancel", command=win.destroy).pack(side="left", padx=4)
        win.grab_set()
        win.focus_force()

    # Rendering and file lists
    def _render_suggestion(self):
        if self._suggestion is None:
            return
        lines = [
            f"Method ID: {self._suggestion.method_id}",
            f"Iteration: {self._suggestion.iteration}",
            f"Created: {self._suggestion.created_at}",
            "",
        ]
        for name in PARAMETER_ORDER:
            lines.append(f"{name}: {self._suggestion.params.get(name)}")
        self._write_text(self._suggestion_text, "\n".join(lines))

    def _render_scores(self, observation):
        for row in self._score_tree.get_children():
            self._score_tree.delete(row)
        components = observation["quality"].get("channel_components", {})
        for ch, data in sorted(components.items(), key=lambda item: int(item[0])):
            self._score_tree.insert(
                "",
                "end",
                text=str(ch),
                values=(
                    self._fmt(data.get("Q_channel")),
                    self._fmt(data.get("normalized_SNR")),
                    self._fmt(data.get("peak_shape_score")),
                    self._fmt(data.get("baseline_stability_score")),
                    self._fmt(data.get("replicate_consistency_score")),
                    self._fmt(data.get("success_score")),
                ),
            )

    def _refresh_history(self):
        for row in self._history_tree.get_children():
            self._history_tree.delete(row)
        if self._bo_session is None:
            return
        for obs in self._bo_session.observations:
            q = obs.get("quality", {})
            self._history_tree.insert(
                "",
                "end",
                text=str(obs.get("iteration")),
                values=(
                    self._fmt(obs.get("Q_run")),
                    self._fmt(q.get("mean_Q_channel")),
                    self._fmt(q.get("std_Q_channel")),
                    self._fmt(q.get("failed_channel_fraction")),
                    self._fmt(q.get("low_channel_fraction")),
                ),
            )

    def _render_best(self):
        best = self._bo_session.best_observation() if self._bo_session else None
        if not best:
            self._write_text(self._best_text, "No completed BO iterations yet.")
            return
        lines = [
            f"Iteration: {best['iteration']}",
            f"Q_run: {best['Q_run']:.4f}",
            f"Method ID: {best['method_id']}",
            "",
        ]
        for name in PARAMETER_ORDER:
            lines.append(f"{name}: {best['params'].get(name)}")
        self._write_text(self._best_text, "\n".join(lines))

    def _refresh_model_artifacts(self):
        if not hasattr(self, "_model_tree"):
            return
        for row in self._model_tree.get_children():
            self._model_tree.delete(row)
        if self._bo_session is None:
            return
        roots = [
            ("surrogate", self._bo_session.surrogate_dir),
            ("acquisition", self._bo_session.acquisition_dir),
            ("plot", self._bo_session.plots_dir),
        ]
        idx = 1
        for kind, folder in roots:
            for path in sorted(folder.glob("*")):
                if path.is_file():
                    self._model_tree.insert("", "end", text=str(idx), values=(kind, path.name))
                    idx += 1

    def _refresh_record_files(self):
        if not hasattr(self, "_record_tree"):
            return
        for row in self._record_tree.get_children():
            self._record_tree.delete(row)
        if self._bo_session is None:
            return
        idx = 1
        for path in sorted(self._bo_session.record_dir.rglob("*")):
            if path.is_file():
                rel = path.relative_to(self._bo_session.record_dir)
                folder = str(rel.parent) if str(rel.parent) != "." else "."
                self._record_tree.insert("", "end", text=str(idx), values=(folder, rel.name))
                idx += 1

    @staticmethod
    def _parse_float_list(text):
        values = []
        for token in text.replace(";", ",").split(","):
            token = token.strip()
            if token:
                values.append(float(token))
        if not values:
            raise ValueError("Enter at least one numeric value.")
        return values

    @staticmethod
    def _fmt(value):
        try:
            return f"{float(value):.3f}"
        except (TypeError, ValueError):
            return ""

    @staticmethod
    def _fmt_raw(value):
        try:
            return f"{float(value):.6g}"
        except (TypeError, ValueError):
            return ""

    @staticmethod
    def _write_text(widget, text):
        widget.config(state="normal")
        widget.delete("1.0", tk.END)
        widget.insert("1.0", text)
        widget.config(state="disabled")

    @staticmethod
    def _clear_text(widget):
        BayesianOptimizationTab._write_text(widget, "")
