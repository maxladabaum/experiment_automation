"""
gui/tab_bayesian_optimization.py - Optional SWV Bayesian optimization tab.

This is a UI shell around core.bo_session. It edits configuration, requests BO
suggestions, queues normal SWV methods, runs/imports analysis outputs, and
shows records. BO math lives in core.bo_session.
"""

from __future__ import annotations

import ast
import csv
import json
import re
import time
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog
import tkinter as tk
from tkinter import ttk, scrolledtext

from config import (
    BO_ANALYSIS_FILE_GLOB,
    BO_ANALYSIS_OUTPUT_DIR,
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
from core.bo_simulation import LANDSCAPE_TYPES, default_dimensions, run_optimizer_simulation
from core.analysis_io import load_swv_csv
from gui.widgets import FlowFrame, ScrollableFrame


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
        self._analysis_glob_var = tk.StringVar(value=str(BO_ANALYSIS_FILE_GLOB))
        self._analysis_crop_min_var = tk.StringVar(value="-0.6")
        self._analysis_crop_max_var = tk.StringVar(value="-0.1")
        self._analysis_smooth_window_var = tk.StringVar(value="15")
        self._analysis_smooth_polyorder_var = tk.StringVar(value="2")
        self._analysis_minima_window_var = tk.StringVar(value="0.30")
        self._analysis_min_peak_height_var = tk.StringVar(value="")
        self._analysis_min_start_voltage_var = tk.StringVar(value="-0.6")
        self._analysis_scan_windows_var = tk.StringVar(value="")
        self._analysis_use_prominent_var = tk.BooleanVar(value=False)
        self._analysis_double_correction_var = tk.BooleanVar(value=True)
        self._analysis_compute_skew_var = tk.BooleanVar(value=False)
        self._analysis_compute_wavelet_energy_var = tk.BooleanVar(value=False)
        self._analysis_wavelet_trace_var = tk.BooleanVar(value=False)
        self._analysis_wavelet_correction_var = tk.BooleanVar(value=False)
        self._channels_var = tk.StringVar(value="")
        self._exploration_var = tk.DoubleVar(value=0.35)
        self._exploration_text_var = tk.StringVar(value="0.35")
        self._candidate_pool_var = tk.StringVar(value="600")
        self._local_pool_var = tk.StringVar(value="120")
        self._initial_point_mode_var = tk.StringVar(value="specific")
        self._score_mode_var = tk.StringVar(value="classic_bounded")
        self._score_snr_weight_var = tk.StringVar(value="0.35")
        self._score_peak_height_weight_var = tk.StringVar(value="0.00")
        self._score_shape_weight_var = tk.StringVar(value="0.20")
        self._score_baseline_weight_var = tk.StringVar(value="0.20")
        self._score_replicate_weight_var = tk.StringVar(value="0.15")
        self._score_success_weight_var = tk.StringVar(value="0.10")
        self._score_snr_saturation_var = tk.StringVar(value="20.0")
        self._score_variability_penalty_var = tk.StringVar(value="0.20")
        self._score_failed_penalty_var = tk.StringVar(value="0.40")
        self._score_low_penalty_var = tk.StringVar(value="0.20")
        self._score_low_threshold_var = tk.StringVar(value="0.50")
        self._score_formula_var = tk.StringVar(value="")
        self._status_var = tk.StringVar(value="Load a BO config to begin.")
        self._record_dir_var = tk.StringVar(value="Record folder: (not started)")
        self._auto_target_var = tk.StringVar(value="5")
        self._analysis_trend_metric_var = tk.StringVar(value="Q_run")
        self._engine_iterations_var = tk.StringVar(value="20")
        self._engine_grid_var = tk.StringVar(value="25")
        self._engine_seed_var = tk.StringVar(value="42")
        self._engine_measurement_noise_var = tk.StringVar(value="0.03")
        self._engine_channel_noise_var = tk.StringVar(value="0.025")
        self._engine_peak_emphasis_var = tk.StringVar(value="0.70")
        self._engine_base_peak_var = tk.StringVar(value="0.45")
        self._engine_peak_gain_var = tk.StringVar(value="5.0")
        self._engine_base_noise_var = tk.StringVar(value="0.08")
        self._engine_noise_gain_var = tk.StringVar(value="0.45")
        self._engine_status_var = tk.StringVar(value="Simulation engine idle.")
        self._auto_status_var = tk.StringVar(value="Auto loop idle.")
        self._style = ttk.Style(self._frame)

        self._config = None
        self._bo_session = None
        self._suggestion = None
        self._auto_running = False
        self._simulation_result = None
        self._simulation_dims = []
        self._engine_plot_canvas = None
        self._engine_plot_figure = None
        self._engine_plot_ax = None
        self._engine_plot_colorbar = None
        self._engine_plot_signature = None
        self._engine_plot_overlay_artists = []
        self._engine_distribution_canvas = None
        self._engine_distribution_figure = None
        self._engine_distribution_ax = None
        self._engine_distribution_lines = {}
        self._engine_distribution_empty_text = None
        self._engine_selected_index = 0
        self._engine_page_index = 0
        self._engine_pages = []
        self._history_rows = {}
        self._selected_history_observation = None
        self._active_results_tree = None
        self._results_trace_panes_balanced = False

        self._build()
        self._load_config(initial=True)

    def _build(self):
        self._configure_styles()
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
        simulation = ttk.Frame(self._tabs)
        results = ttk.Frame(self._tabs)
        self._tabs.add(setup, text="Setup")
        self._tabs.add(run, text="Run")
        self._tabs.add(simulation, text="Simulation Engine")
        self._tabs.add(results, text="Results & Records")

        self._build_setup_tab(setup)
        self._build_run_tab(run)
        self._build_simulation_tab(simulation)
        self._build_results_tab(results)

        status = ttk.Label(root, textvariable=self._status_var, relief="sunken")
        status.pack(side="bottom", fill="x", padx=8, pady=(0, 8))

    def _configure_styles(self):
        self._style.configure(
            "BO.Treeview",
            background="white",
            fieldbackground="white",
        )
        self._style.map(
            "BO.Treeview",
            background=[("selected", self.ACCENT_DARK)],
            foreground=[("selected", "white")],
        )

    def _balance_results_trace_panes(self):
        if getattr(self, "_results_trace_panes_balanced", False):
            return
        panes = (
            getattr(self, "_results_top_pane", None),
            getattr(self, "_results_middle_pane", None),
        )
        if any(pane is None for pane in panes):
            return
        widths = []
        for pane in (
            getattr(self, "_results_top_pane", None),
            getattr(self, "_results_middle_pane", None),
        ):
            try:
                pane.update_idletasks()
                width = pane.winfo_width()
                if width <= 20:
                    return
                widths.append(width)
            except Exception:
                return
        for pane, width in zip(panes, widths):
            try:
                pane.sashpos(0, width // 2)
            except Exception:
                return
        self._results_trace_panes_balanced = True

    def _build_setup_tab(self, parent):
        scroller = ScrollableFrame(parent, min_width=1080)
        scroller.pack(fill="both", expand=True)
        parent = scroller.content
        pane = ttk.PanedWindow(parent, orient=tk.HORIZONTAL)
        pane.pack(fill="both", expand=True, padx=4, pady=4)
        left = ttk.Frame(pane)
        right = ttk.Frame(pane)
        pane.add(left, weight=2)
        pane.add(right, weight=3)

        right_pane = ttk.PanedWindow(right, orient=tk.VERTICAL)
        right_pane.pack(fill="both", expand=True)
        params_host = ttk.Frame(right_pane)
        controls_host = ttk.Frame(right_pane)
        right_pane.add(params_host, weight=3)
        right_pane.add(controls_host, weight=2)

        cfg = ttk.LabelFrame(left, text="Configuration and Analysis Output", padding=8)
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

        ttk.Label(cfg, text="Analysis glob:").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Entry(cfg, textvariable=self._analysis_glob_var, width=14).grid(row=2, column=1, sticky="w", padx=4)

        ttk.Label(cfg, text="Mux channels:").grid(row=3, column=0, sticky="w", pady=2)
        channels = ttk.Entry(cfg, textvariable=self._channels_var)
        channels.grid(row=3, column=1, sticky="ew", padx=4)
        channels.bind("<FocusOut>", lambda _e: self._sync_channels_from_entry())
        channels.bind("<Return>", lambda _e: self._sync_channels_from_entry())
        ttk.Button(cfg, text="Validate", command=self._validate_config).grid(row=3, column=2, padx=2)
        ttk.Button(cfg, text="Load BO Session", command=self._load_bo_session).grid(row=3, column=3, padx=2)
        ttk.Button(cfg, text="Start BO Session", command=self._start_bo_session).grid(row=3, column=4, padx=2)

        clue = ttk.LabelFrame(left, text="Setup Cues", padding=8)
        clue.pack(fill="x", pady=(0, 8))
        ttk.Label(
            clue,
            text=(
                "Use the BO config for scientific choices: active parameters, hard constraints, "
                "initial design, channels, and scoring. Use Save Paths for the machine-specific "
                "analysis output folder."
            ),
            wraplength=460,
            justify="left",
        ).pack(fill="x")

        algo_box = ttk.LabelFrame(left, text="Optimizer Behavior", padding=8)
        algo_box.pack(fill="x", pady=(0, 8))
        algo_box.columnconfigure(1, weight=1)
        ttk.Label(algo_box, text="Exploit <-> Explore").grid(row=0, column=0, sticky="w")
        ttk.Scale(
            algo_box,
            from_=0.0,
            to=1.0,
            orient=tk.HORIZONTAL,
            variable=self._exploration_var,
            command=lambda _v: self._sync_algorithm_config(show_error=False),
        ).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Label(algo_box, textvariable=self._exploration_text_var, foreground=self.ACCENT).grid(row=0, column=2, sticky="e")
        ttk.Label(algo_box, text="Global pool:").grid(row=1, column=0, sticky="w", pady=2)
        ttk.Entry(algo_box, textvariable=self._candidate_pool_var, width=8).grid(row=1, column=1, sticky="w", padx=6)
        ttk.Label(algo_box, text="Local pool:").grid(row=1, column=2, sticky="w", pady=2)
        ttk.Entry(algo_box, textvariable=self._local_pool_var, width=8).grid(row=1, column=2, sticky="e", padx=(6, 0))
        ttk.Label(algo_box, text="Start point:").grid(row=2, column=0, sticky="w", pady=2)
        start_mode = ttk.Combobox(
            algo_box,
            textvariable=self._initial_point_mode_var,
            values=("specific", "random"),
            state="readonly",
            width=12,
        )
        start_mode.grid(row=2, column=1, sticky="w", padx=6, pady=2)
        start_mode.bind("<<ComboboxSelected>>", lambda _e: self._sync_algorithm_config(show_error=False))
        ttk.Label(
            algo_box,
            text="`specific` uses Initial Parameters. `random` chooses one valid candidate as the first BO point.",
            foreground=self.ACCENT,
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=(2, 0))

        init_box = ttk.LabelFrame(left, text="Initial Parameters", padding=8)
        init_box.pack(fill="both", expand=True)
        init_toolbar = ttk.Frame(init_box)
        init_toolbar.pack(fill="x", pady=(0, 6))
        ttk.Button(init_toolbar, text="Edit Initial Parameters", command=self._edit_initial_parameters).pack(side="left", padx=2)
        ttk.Label(
            init_toolbar,
            text="Specific mode starts here. Random mode ignores this as the first BO point, but keeps it as the editable reference method.",
            foreground=self.ACCENT,
        ).pack(side="left", padx=8)

        init_cols = ("Begin", "End", "Step", "Amp", "Freq", "Cond E", "Cond t")
        self._initial_tree = ttk.Treeview(
            init_box, columns=init_cols, show="tree headings", height=8, style="BO.Treeview"
        )
        self._initial_tree.heading("#0", text="#")
        self._initial_tree.column("#0", width=38, anchor="center")
        for col in init_cols:
            self._initial_tree.heading(col, text=col)
            self._initial_tree.column(col, width=70, anchor="center")
        self._initial_tree.pack(fill="both", expand=True)
        self._initial_tree.bind("<Double-1>", lambda _e: self._edit_initial_parameters())

        params = ttk.LabelFrame(params_host, text="Parameter Space", padding=8)
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

        cols = ("Mode", "Space", "Values", "Tie")
        self._param_tree = ttk.Treeview(
            params, columns=cols, show="tree headings", height=16, style="BO.Treeview"
        )
        self._param_tree.heading("#0", text="Parameter")
        self._param_tree.heading("Mode", text="Mode")
        self._param_tree.heading("Space", text="Space")
        self._param_tree.heading("Values", text="Range / Values / Lock")
        self._param_tree.heading("Tie", text="Tie")
        self._param_tree.column("#0", width=165)
        self._param_tree.column("Mode", width=75)
        self._param_tree.column("Space", width=85)
        self._param_tree.column("Values", width=330)
        self._param_tree.column("Tie", width=110)
        self._param_tree.tag_configure("active", background="#dff7f5", foreground="#0f3d44")
        self._param_tree.tag_configure("locked", background="#eeeeee", foreground="#333333")
        self._param_tree.tag_configure("tied", background="#fff1d6", foreground="#6b3b00")
        self._param_tree.pack(fill="both", expand=True)
        self._param_tree.bind("<Double-1>", lambda _e: self._edit_selected_parameter())

        setup_tools = ttk.Notebook(controls_host)
        setup_tools.pack(fill="both", expand=True)
        analysis_tab = ttk.Frame(setup_tools)
        scoring_tab = ttk.Frame(setup_tools)
        setup_tools.add(analysis_tab, text="Analysis Settings")
        setup_tools.add(scoring_tab, text="Q Scoring")

        scoring_box = ttk.LabelFrame(scoring_tab, text="Q Score Decomposition", padding=8)
        scoring_box.pack(fill="both", expand=True, pady=(0, 8), padx=2)
        for idx in range(6):
            scoring_box.columnconfigure(idx, weight=1 if idx in (1, 3, 5) else 0)
        ttk.Label(scoring_box, text="Score mode:").grid(row=0, column=0, sticky="w", pady=2)
        mode_combo = ttk.Combobox(
            scoring_box,
            textvariable=self._score_mode_var,
            values=("classic_bounded", "signal_priority_unbounded"),
            state="readonly",
            width=26,
        )
        mode_combo.grid(row=0, column=1, columnspan=2, sticky="w", padx=(4, 10), pady=2)
        mode_combo.bind("<<ComboboxSelected>>", lambda _e: self._sync_scoring_config(show_error=False))
        ttk.Button(scoring_box, text="Signal Preset", command=self._apply_signal_priority_preset).grid(row=0, column=3, sticky="w", pady=2)
        entries = [
            ("SNR w:", self._score_snr_weight_var),
            ("Peak w:", self._score_peak_height_weight_var),
            ("Shape w:", self._score_shape_weight_var),
            ("Baseline w:", self._score_baseline_weight_var),
            ("Replicate w:", self._score_replicate_weight_var),
            ("Success w:", self._score_success_weight_var),
            ("SNR sat:", self._score_snr_saturation_var),
            ("Var penalty:", self._score_variability_penalty_var),
            ("Failed penalty:", self._score_failed_penalty_var),
            ("Low penalty:", self._score_low_penalty_var),
            ("Low threshold:", self._score_low_threshold_var),
        ]
        for idx, (label, var) in enumerate(entries):
            row = 1 + idx // 2
            base_col = (idx % 2) * 3
            ttk.Label(scoring_box, text=label).grid(row=row, column=base_col, sticky="w", pady=2)
            entry = ttk.Entry(scoring_box, textvariable=var, width=9)
            entry.grid(row=row, column=base_col + 1, sticky="w", padx=(4, 10), pady=2)
            entry.bind("<FocusOut>", lambda _e: self._sync_scoring_config(show_error=False))
            entry.bind("<Return>", lambda _e: self._sync_scoring_config(show_error=False))
        ttk.Label(scoring_box, textvariable=self._score_formula_var, foreground=self.ACCENT, wraplength=460, justify="left").grid(
            row=7,
            column=0,
            columnspan=6,
            sticky="w",
            pady=(4, 0),
        )
        ttk.Label(
            scoring_box,
            text=self._q_reference_text(),
            wraplength=460,
            justify="left",
        ).grid(
            row=8,
            column=0,
            columnspan=6,
            sticky="w",
            pady=(6, 0),
        )

        analysis_box = ttk.LabelFrame(analysis_tab, text="Headless Analysis Settings", padding=8)
        analysis_box.pack(fill="both", expand=True, pady=(0, 8), padx=2)
        for idx in range(4):
            analysis_box.columnconfigure(idx, weight=1 if idx in (1, 3) else 0)
        ttk.Label(analysis_box, text="Crop min/max (V):").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Entry(analysis_box, textvariable=self._analysis_crop_min_var, width=8).grid(row=0, column=1, sticky="w", padx=(4, 2))
        ttk.Entry(analysis_box, textvariable=self._analysis_crop_max_var, width=8).grid(row=0, column=1, sticky="e", padx=(2, 4))
        ttk.Label(analysis_box, text="Smooth win/poly:").grid(row=0, column=2, sticky="w", pady=2)
        ttk.Entry(analysis_box, textvariable=self._analysis_smooth_window_var, width=8).grid(row=0, column=3, sticky="w", padx=(4, 2))
        ttk.Entry(analysis_box, textvariable=self._analysis_smooth_polyorder_var, width=8).grid(row=0, column=3, sticky="e", padx=(2, 4))
        ttk.Label(analysis_box, text="Minima window (V):").grid(row=1, column=0, sticky="w", pady=2)
        ttk.Entry(analysis_box, textvariable=self._analysis_minima_window_var, width=10).grid(row=1, column=1, sticky="w", padx=4)
        ttk.Label(analysis_box, text="Min peak height (uA):").grid(row=1, column=2, sticky="w", pady=2)
        ttk.Entry(analysis_box, textvariable=self._analysis_min_peak_height_var, width=10).grid(row=1, column=3, sticky="w", padx=4)
        ttk.Label(analysis_box, text="Min start V:").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Entry(analysis_box, textvariable=self._analysis_min_start_voltage_var, width=10).grid(row=2, column=1, sticky="w", padx=4)
        ttk.Label(analysis_box, text="Scan windows:").grid(row=2, column=2, sticky="w", pady=2)
        ttk.Entry(analysis_box, textvariable=self._analysis_scan_windows_var).grid(row=2, column=3, sticky="ew", padx=4)
        ttk.Checkbutton(analysis_box, text="Prominent minima", variable=self._analysis_use_prominent_var).grid(row=3, column=0, sticky="w", pady=2)
        ttk.Checkbutton(analysis_box, text="Double correction", variable=self._analysis_double_correction_var).grid(row=3, column=1, sticky="w", pady=2)
        ttk.Checkbutton(analysis_box, text="Compute skew", variable=self._analysis_compute_skew_var).grid(row=3, column=2, sticky="w", pady=2)
        ttk.Checkbutton(analysis_box, text="Wavelet energy", variable=self._analysis_compute_wavelet_energy_var).grid(row=3, column=3, sticky="w", pady=2)
        ttk.Checkbutton(analysis_box, text="Wavelet trace", variable=self._analysis_wavelet_trace_var).grid(row=4, column=0, sticky="w", pady=2)
        ttk.Checkbutton(analysis_box, text="Wavelet correction", variable=self._analysis_wavelet_correction_var).grid(row=4, column=1, sticky="w", pady=2)
        ttk.Label(
            analysis_box,
            text="These settings are used by the in-repo BO analysis runner.",
            foreground=self.ACCENT,
        ).grid(row=5, column=0, columnspan=4, sticky="w", pady=(4, 0))

    def _build_run_tab(self, parent):
        scroller = ScrollableFrame(parent, min_width=1020)
        scroller.pack(fill="both", expand=True)
        parent = scroller.content
        controls = ttk.LabelFrame(parent, text="Manual Closed Loop", padding=8)
        controls.pack(fill="x", padx=4, pady=(4, 8))
        manual_bar = FlowFrame(controls)
        manual_bar.pack(fill="x")
        manual_bar.add(ttk.Button(manual_bar, text="Suggest Next Method", command=self._suggest_next))
        manual_bar.add(ttk.Button(manual_bar, text="Send Batch to Queue", command=self._send_to_queue))
        manual_bar.add(ttk.Button(manual_bar, text="Preview Script", command=self._preview_suggestion))
        manual_bar.separator()
        manual_bar.add(ttk.Button(manual_bar, text="Import Analysis JSON", command=self._import_analysis_dialog))
        manual_bar.add(ttk.Button(manual_bar, text="Use Latest Analysis", command=self._import_latest_analysis))
        manual_bar.add(ttk.Button(manual_bar, text="Run Analysis", command=self._run_analysis_for_pending))

        auto = ttk.LabelFrame(parent, text="Auto Loop", padding=8)
        auto.pack(fill="x", padx=4, pady=(0, 8))
        auto_bar = FlowFrame(auto)
        auto_bar.pack(fill="x")
        auto_bar.add(ttk.Label(auto_bar, text="Target completed iterations:"))
        auto_bar.add(ttk.Entry(auto_bar, textvariable=self._auto_target_var, width=6))
        auto_bar.add(ttk.Button(auto_bar, text="Start Auto Loop", command=self._start_auto_loop))
        auto_bar.add(ttk.Button(auto_bar, text="Stop Auto", command=self._stop_auto_loop))
        auto_bar.add(ttk.Label(auto_bar, textvariable=self._auto_status_var, foreground=self.ACCENT))

        clue = ttk.LabelFrame(parent, text="Workflow Cues", padding=8)
        clue.pack(fill="x", padx=4, pady=(0, 8))
        ttk.Label(
            clue,
            text=(
                "Manual: suggest -> queue -> run queue -> run/import analysis. Auto: queue must be empty; "
                "the tab runs one BO batch at a time and analyzes the active experiment folder."
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

    def _build_simulation_tab(self, parent):
        root = ttk.Frame(parent)
        root.pack(fill="both", expand=True, padx=4, pady=4)

        nav = ttk.Frame(root)
        nav.pack(fill="x", pady=(0, 6))
        self._engine_step_label = ttk.Label(nav, text="", font=("Arial", 11, "bold"))
        self._engine_step_label.pack(side="left", padx=(2, 10))
        self._engine_back_button = ttk.Button(nav, text="< Back", command=self._engine_prev_page)
        self._engine_back_button.pack(side="left", padx=2)
        self._engine_next_button = ttk.Button(nav, text="Next >", command=self._engine_next_page)
        self._engine_next_button.pack(side="right", padx=2)

        ttk.Label(root, textvariable=self._engine_status_var, foreground=self.ACCENT, wraplength=980, justify="left").pack(fill="x", pady=(0, 6))

        self._engine_page_container = ttk.Frame(root)
        self._engine_page_container.pack(fill="both", expand=True)
        landscape_page = ttk.Frame(self._engine_page_container)
        model_page = ttk.Frame(self._engine_page_container)
        results_page = ttk.Frame(self._engine_page_container)
        self._engine_pages = [
            ("1/3 Landscape", landscape_page),
            ("2/3 Signal Model", model_page),
            ("3/3 Results", results_page),
        ]

        dims_box = ttk.LabelFrame(landscape_page, text="Synthetic Parameter Landscape Map", padding=8)
        dims_box.pack(fill="both", expand=True)
        dims_split = ttk.PanedWindow(dims_box, orient=tk.HORIZONTAL)
        dims_split.pack(fill="both", expand=True)

        dims_left = ttk.Frame(dims_split)
        dims_right = ttk.Frame(dims_split)
        dims_split.add(dims_left, weight=2)
        dims_split.add(dims_right, weight=3)

        toolbar = ttk.Frame(dims_left)
        toolbar.pack(fill="x", pady=(0, 6))
        ttk.Button(toolbar, text="Load Active Parameters", command=self._engine_load_active_dimensions).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Edit Selected", command=self._engine_edit_dimension).pack(side="left", padx=2)
        ttk.Button(toolbar, text="Next: Signal Model", command=self._engine_next_page).pack(side="right", padx=2)

        dim_cols = ("Min", "Max", "Optimum", "Spread", "Shape", "Weight")
        self._engine_dim_tree = ttk.Treeview(dims_left, columns=dim_cols, show="tree headings", height=14, style="BO.Treeview")
        self._engine_dim_tree.heading("#0", text="Parameter")
        self._engine_dim_tree.column("#0", width=150)
        for col in dim_cols:
            self._engine_dim_tree.heading(col, text=col)
            self._engine_dim_tree.column(col, width=82, anchor="center")
        self._engine_dim_tree.pack(fill="both", expand=True)
        self._engine_dim_tree.bind("<Double-1>", lambda _e: self._engine_edit_dimension())
        self._engine_dim_tree.bind("<<TreeviewSelect>>", lambda _e: self._engine_refresh_landscape_inspector(refresh_cube=False))

        dist_box = ttk.LabelFrame(dims_right, text="Map Slice: Per-Dimension Success Distribution", padding=6)
        dist_box.pack(fill="both", expand=True)
        self._engine_distribution_frame = ttk.Frame(dist_box)
        self._engine_distribution_frame.pack(fill="both", expand=True)

        cube_box = ttk.LabelFrame(dims_right, text="Example Fake Data From Map Cells", padding=6)
        cube_box.pack(fill="both", expand=True, pady=(6, 0))
        cube_cols = ("True Q", "Success", "Peak", "Noise")
        self._engine_cube_tree = ttk.Treeview(cube_box, columns=cube_cols, show="tree headings", height=8, style="BO.Treeview")
        self._engine_cube_tree.heading("#0", text="Point")
        self._engine_cube_tree.column("#0", width=180)
        for col in cube_cols:
            self._engine_cube_tree.heading(col, text=col)
            self._engine_cube_tree.column(col, width=74, anchor="center")
        self._engine_cube_tree.pack(fill="both", expand=True)
        cube_x = ttk.Scrollbar(cube_box, orient=tk.HORIZONTAL, command=self._engine_cube_tree.xview)
        self._engine_cube_tree.configure(xscrollcommand=cube_x.set)
        cube_x.pack(fill="x")
        self._engine_cube_tree.bind("<<TreeviewSelect>>", lambda _e: self._engine_preview_selected_cube_point())

        model_box = ttk.LabelFrame(model_page, text="Synthetic SWV Model", padding=8)
        model_box.pack(fill="both", expand=True)
        model_grid = ttk.Frame(model_box)
        model_grid.pack(fill="x")
        model_entries = [
            ("Iterations", self._engine_iterations_var),
            ("Grid", self._engine_grid_var),
            ("Seed", self._engine_seed_var),
            ("Meas noise", self._engine_measurement_noise_var),
            ("Channel noise", self._engine_channel_noise_var),
            ("Peak emphasis", self._engine_peak_emphasis_var),
            ("Base peak uA", self._engine_base_peak_var),
            ("Peak gain uA", self._engine_peak_gain_var),
            ("Base noise uA", self._engine_base_noise_var),
            ("Noise gain uA", self._engine_noise_gain_var),
        ]
        for idx, (label, var) in enumerate(model_entries):
            row = idx // 2
            col = (idx % 2) * 2
            ttk.Label(model_grid, text=f"{label}:").grid(row=row, column=col, sticky="w", pady=2)
            ttk.Entry(model_grid, textvariable=var, width=10).grid(row=row, column=col + 1, sticky="w", padx=(4, 12), pady=2)
        run_bar = ttk.Frame(model_box)
        run_bar.pack(fill="x", pady=(8, 0))
        ttk.Button(run_bar, text="Draw Landscape", command=self._engine_draw_landscape).pack(side="left", padx=2)
        ttk.Button(run_bar, text="Run Optimizer Simulation", command=self._engine_run_optimizer).pack(side="left", padx=2)
        ttk.Button(run_bar, text="Next: Results", command=self._engine_next_page).pack(side="right", padx=2)
        ttk.Label(
            model_box,
            text=(
                "Peak emphasis controls how strongly the synthetic system rewards signal height relative "
                "to low noise and shape. Synthetic success is separate from Q and is driven mostly by peak "
                "height, then low noise. Think of the landscape as a topographic map over parameter space: "
                "high regions are good operating zones, and the optimizer is trying to climb into them."
            ),
            foreground=self.ACCENT,
            wraplength=760,
            justify="left",
        ).pack(fill="x", pady=(10, 0))

        results = ttk.PanedWindow(results_page, orient=tk.HORIZONTAL)
        results.pack(fill="both", expand=True)
        plot_box = ttk.LabelFrame(results, text="Optimizer Movement", padding=6)
        detail_box = ttk.LabelFrame(results, text="Iteration Window", padding=6)
        results.add(plot_box, weight=3)
        results.add(detail_box, weight=2)

        result_toolbar = ttk.Frame(plot_box)
        result_toolbar.pack(fill="x", pady=(0, 4))
        ttk.Button(result_toolbar, text="Run Optimizer Simulation", command=self._engine_run_optimizer).pack(side="left", padx=2)
        ttk.Button(result_toolbar, text="Apply Best To Setup", command=self._engine_apply_best_to_setup).pack(side="left", padx=2)
        ttk.Separator(result_toolbar, orient=tk.VERTICAL).pack(side="left", fill="y", padx=8)
        ttk.Button(result_toolbar, text="< Iter", command=lambda: self._engine_step_window(-1)).pack(side="left", padx=2)
        ttk.Button(result_toolbar, text="Iter >", command=lambda: self._engine_step_window(1)).pack(side="left", padx=2)
        ttk.Button(result_toolbar, text="Show All", command=self._engine_show_all).pack(side="left", padx=2)
        engine_output_tabs = ttk.Notebook(plot_box)
        engine_output_tabs.pack(fill="both", expand=True)
        movement_tab = ttk.Frame(engine_output_tabs)
        q_tab = ttk.Frame(engine_output_tabs)
        engine_output_tabs.add(movement_tab, text="Movement")
        engine_output_tabs.add(q_tab, text="Q Trend")
        self._engine_plot_frame = ttk.Frame(movement_tab)
        self._engine_plot_frame.pack(fill="both", expand=True)
        self._engine_q_plot_frame = ttk.Frame(q_tab)
        self._engine_q_plot_frame.pack(fill="both", expand=True)

        result_cols = ("Q_run", "True Q", "Distance", "Peak uA", "Raw SNR", "Begin", "End", "Step", "Amp", "Freq")
        self._engine_result_tree = ttk.Treeview(detail_box, columns=result_cols, show="tree headings", height=9, style="BO.Treeview")
        self._engine_result_tree.heading("#0", text="Iter")
        self._engine_result_tree.column("#0", width=46, anchor="center")
        for col in result_cols:
            self._engine_result_tree.heading(col, text=col)
            self._engine_result_tree.column(col, width=78, anchor="center")
        self._engine_result_tree.pack(fill="both", expand=True)
        self._engine_result_tree.bind("<<TreeviewSelect>>", lambda _e: self._engine_select_iteration_from_table())
        result_x = ttk.Scrollbar(detail_box, orient=tk.HORIZONTAL, command=self._engine_result_tree.xview)
        self._engine_result_tree.configure(xscrollcommand=result_x.set)
        result_x.pack(fill="x")

        trace_box = ttk.LabelFrame(detail_box, text="Synthetic SWV Trace Preview", padding=4)
        trace_box.pack(fill="both", expand=True, pady=(6, 0))
        self._engine_trace_plot_frame = ttk.Frame(trace_box)
        self._engine_trace_plot_frame.pack(fill="both", expand=True)
        self._engine_trace_text = scrolledtext.ScrolledText(trace_box, height=6, wrap=tk.WORD)
        self._engine_trace_text.pack(fill="both", expand=True, pady=(6, 0))
        self._engine_trace_text.config(state="disabled")
        self._engine_go_page(0)

    def _build_results_tab(self, parent):
        pane = ttk.PanedWindow(parent, orient=tk.VERTICAL)
        pane.pack(fill="both", expand=True, padx=4, pady=4)
        top = ttk.PanedWindow(pane, orient=tk.HORIZONTAL)
        middle = ttk.PanedWindow(pane, orient=tk.HORIZONTAL)
        self._results_top_pane = top
        self._results_middle_pane = middle
        top.bind("<Configure>", lambda _e: self._balance_results_trace_panes(), add="+")
        middle.bind("<Configure>", lambda _e: self._balance_results_trace_panes(), add="+")
        bottom = ttk.PanedWindow(pane, orient=tk.HORIZONTAL)
        pane.add(top, weight=1)
        pane.add(middle, weight=1)
        pane.add(bottom, weight=1)

        score_box = ttk.LabelFrame(top, text="Per-Channel Scores", padding=6)
        best_box = ttk.LabelFrame(top, text="Raw SWV Traces for Selected Iteration", padding=6)
        top.add(score_box, weight=1)
        top.add(best_box, weight=1)

        score_cols = ("Q", "Peak uA", "Raw SNR", "SNR Score", "Shape", "Baseline", "Replicate", "Success")
        self._score_tree = ttk.Treeview(score_box, columns=score_cols, show="tree headings", height=10, selectmode="extended")
        self._score_tree.heading("#0", text="Ch")
        self._score_tree.column("#0", width=50, anchor="center")
        for col in score_cols:
            self._score_tree.heading(col, text=col)
            self._score_tree.column(col, width=78, anchor="center")
        self._score_tree.pack(fill="both", expand=True)
        self._score_tree.configure(takefocus=True)
        self._score_tree.bind("<ButtonPress-1>", self._focus_tree_on_click, add="+")
        self._score_tree.bind("<Enter>", lambda _e: self._set_active_results_tree("score"), add="+")
        self._score_tree.bind("<ButtonPress-1>", lambda _e: self._set_active_results_tree("score"), add="+")
        self._score_tree.bind("<ButtonRelease-1>", lambda _e: self._restore_tree_focus(self._score_tree), add="+")
        self._score_tree.bind("<<TreeviewSelect>>", lambda _e: self._on_score_tree_select())
        self._score_tree.bind("<Up>", lambda event: self._move_score_selection(-1, event))
        self._score_tree.bind("<Down>", lambda event: self._move_score_selection(1, event))
        self._score_tree.bind("<Left>", lambda event: self._move_score_selection(-1, event))
        self._score_tree.bind("<Right>", lambda event: self._move_score_selection(1, event))
        score_x = ttk.Scrollbar(score_box, orient=tk.HORIZONTAL, command=self._score_tree.xview)
        self._score_tree.configure(xscrollcommand=score_x.set)
        score_x.pack(fill="x")
        self._q_equation_text = scrolledtext.ScrolledText(score_box, height=5, wrap=tk.WORD)
        self._q_equation_text.pack(fill="x", pady=(6, 0))
        self._q_equation_text.config(state="disabled")

        self._raw_trace_frame = ttk.Frame(best_box)
        self._raw_trace_frame.pack(fill="both", expand=True)

        history_host = ttk.Frame(middle)
        corrected_box = ttk.LabelFrame(middle, text="Smoothed Corrected Traces for Selected Iteration", padding=6)
        middle.add(history_host, weight=1)
        middle.add(corrected_box, weight=1)

        self._history_tabs = ttk.Notebook(history_host)
        self._history_tabs.pack(fill="both", expand=True)
        hist_box = ttk.Frame(self._history_tabs)
        q_plot_box = ttk.Frame(self._history_tabs)
        self._history_tabs.add(hist_box, text="History Table")
        self._history_tabs.add(q_plot_box, text="Trend Plot")
        trend_toolbar = ttk.Frame(q_plot_box)
        trend_toolbar.pack(fill="x", padx=6, pady=(6, 0))
        ttk.Label(trend_toolbar, text="Metric:").pack(side="left", padx=(0, 4))
        self._analysis_trend_combo = ttk.Combobox(
            trend_toolbar,
            textvariable=self._analysis_trend_metric_var,
            values=self._analysis_trend_metric_options(),
            state="readonly",
            width=24,
        )
        self._analysis_trend_combo.pack(side="left")
        self._analysis_trend_combo.bind("<<ComboboxSelected>>", lambda _e: self._refresh_analysis_q_trend())
        self._analysis_q_plot_frame = ttk.Frame(q_plot_box)
        self._analysis_q_plot_frame.pack(fill="both", expand=True, padx=6, pady=6)
        self._corrected_trace_frame = ttk.Frame(corrected_box)
        self._corrected_trace_frame.pack(fill="both", expand=True)
        model_box = ttk.LabelFrame(bottom, text="Surrogate and Acquisition Artifacts", padding=6)
        hist_cols = ("Q_run", "Mean", "Std", "Failed", "Low", "Peak uA", "RMS uA", "Begin", "End", "Step", "Amp", "Freq", "Cond E", "Cond t")
        self._history_tree = ttk.Treeview(hist_box, columns=hist_cols, show="tree headings", height=10)
        self._history_tree.heading("#0", text="Iter")
        self._history_tree.column("#0", width=55, anchor="center")
        for col in hist_cols:
            self._history_tree.heading(col, text=col)
            self._history_tree.column(col, width=76, anchor="center")
        self._history_tree.pack(fill="both", expand=True)
        self._history_tree.configure(takefocus=True)
        self._history_tree.bind("<ButtonPress-1>", self._focus_tree_on_click, add="+")
        self._history_tree.bind("<Enter>", lambda _e: self._set_active_results_tree("history"), add="+")
        self._history_tree.bind("<ButtonPress-1>", lambda _e: self._set_active_results_tree("history"), add="+")
        self._history_tree.bind("<ButtonRelease-1>", lambda _e: self._restore_tree_focus(self._history_tree), add="+")
        self._history_tree.bind("<<TreeviewSelect>>", lambda _e: self._select_history_iteration())
        self._history_tree.bind("<Double-1>", self._on_history_double_click)
        self._history_tree.bind("<Up>", lambda event: self._move_history_selection(-1, event))
        self._history_tree.bind("<Down>", lambda event: self._move_history_selection(1, event))
        self._history_tree.bind("<Left>", lambda event: self._move_history_selection(-1, event))
        self._history_tree.bind("<Right>", lambda event: self._move_history_selection(1, event))
        self._history_tree.bind_all("<Up>", lambda event: self._route_results_arrow(-1, event), add="+")
        self._history_tree.bind_all("<Down>", lambda event: self._route_results_arrow(1, event), add="+")
        self._history_tree.bind_all("<Left>", lambda event: self._route_results_arrow(-1, event), add="+")
        self._history_tree.bind_all("<Right>", lambda event: self._route_results_arrow(1, event), add="+")
        history_x = ttk.Scrollbar(hist_box, orient=tk.HORIZONTAL, command=self._history_tree.xview)
        self._history_tree.configure(xscrollcommand=history_x.set)
        history_x.pack(fill="x")

        cols = ("Type", "File")
        bottom.add(model_box, weight=1)
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
        parent.after_idle(self._balance_results_trace_panes)

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
        path = filedialog.askdirectory(title="Choose BO analysis output folder")
        if path:
            self._analysis_dir_var.set(path)

    def _save_local_paths(self):
        try:
            payload = {
                "analysis_output_dir": self._analysis_dir_var.get().strip(),
                "analysis_file_glob": self._analysis_glob_var.get().strip() or "*.json",
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
            self._set_analysis_vars_from_config(analysis_cfg)
            self._set_algorithm_vars_from_config(self._config)
            self._set_scoring_vars_from_config(self._config)
            self._engine_seed_var.set(str(self._config.get("random_seed", 42)))
            self._channels_var.set(", ".join(str(ch) for ch in self._config.get("channels", [])))
            self._refresh_parameter_table()
            self._refresh_initial_parameters_table()
            self._engine_load_active_dimensions()
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
        analysis_cfg = self._config.setdefault("analysis", {})
        analysis_cfg["file_glob"] = self._analysis_glob_var.get().strip() or "*.json"
        self._update_analysis_config_from_vars(analysis_cfg)
        self._sync_algorithm_config(show_error=False)
        self._sync_scoring_config(show_error=False)
        try:
            path = save_bo_config(self._config, self._config_path_var.get())
            self._status_var.set(f"Saved BO config: {path}")
        except Exception as exc:
            messagebox.showerror("Save BO Config", str(exc))

    def _set_analysis_vars_from_config(self, analysis_cfg: dict):
        self._analysis_crop_min_var.set(str(analysis_cfg.get("crop_min_v", -0.6)))
        self._analysis_crop_max_var.set(str(analysis_cfg.get("crop_max_v", -0.1)))
        self._analysis_smooth_window_var.set(str(analysis_cfg.get("smooth_window", 15)))
        self._analysis_smooth_polyorder_var.set(str(analysis_cfg.get("smooth_polyorder", 2)))
        self._analysis_minima_window_var.set(str(analysis_cfg.get("minima_search_window_v", 0.30)))
        self._analysis_min_peak_height_var.set("" if analysis_cfg.get("min_peak_height_ua") in (None, "") else str(analysis_cfg.get("min_peak_height_ua")))
        self._analysis_min_start_voltage_var.set(str(analysis_cfg.get("min_start_voltage_v", -0.6)))
        self._analysis_scan_windows_var.set(str(analysis_cfg.get("scan_windows", "")))
        self._analysis_use_prominent_var.set(bool(analysis_cfg.get("use_prominent_minima", False)))
        self._analysis_double_correction_var.set(bool(analysis_cfg.get("use_double_correction", True)))
        self._analysis_compute_skew_var.set(bool(analysis_cfg.get("compute_skew", False)))
        self._analysis_compute_wavelet_energy_var.set(bool(analysis_cfg.get("compute_wavelet_energy", False)))
        self._analysis_wavelet_trace_var.set(bool(analysis_cfg.get("compute_wavelet_denoised_trace", False)))
        self._analysis_wavelet_correction_var.set(bool(analysis_cfg.get("use_wavelet_for_correction", False)))

    def _update_analysis_config_from_vars(self, analysis_cfg: dict):
        analysis_cfg["crop_min_v"] = float(self._analysis_crop_min_var.get())
        analysis_cfg["crop_max_v"] = float(self._analysis_crop_max_var.get())
        analysis_cfg["smooth_window"] = int(self._analysis_smooth_window_var.get())
        analysis_cfg["smooth_polyorder"] = int(self._analysis_smooth_polyorder_var.get())
        analysis_cfg["minima_search_window_v"] = float(self._analysis_minima_window_var.get())
        peak_height_text = (self._analysis_min_peak_height_var.get() or "").strip()
        analysis_cfg["min_peak_height_ua"] = None if not peak_height_text else float(peak_height_text)
        analysis_cfg["min_start_voltage_v"] = float(self._analysis_min_start_voltage_var.get())
        analysis_cfg["scan_windows"] = (self._analysis_scan_windows_var.get() or "").strip()
        analysis_cfg["use_prominent_minima"] = bool(self._analysis_use_prominent_var.get())
        analysis_cfg["use_double_correction"] = bool(self._analysis_double_correction_var.get())
        analysis_cfg["compute_skew"] = bool(self._analysis_compute_skew_var.get())
        analysis_cfg["compute_wavelet_energy"] = bool(self._analysis_compute_wavelet_energy_var.get())
        analysis_cfg["compute_wavelet_denoised_trace"] = bool(self._analysis_wavelet_trace_var.get())
        analysis_cfg["use_wavelet_for_correction"] = bool(self._analysis_wavelet_correction_var.get())

    def _set_algorithm_vars_from_config(self, cfg: dict):
        acquisition = dict((cfg or {}).get("acquisition") or {})
        self._exploration_var.set(float(acquisition.get("exploration", 0.35)))
        self._exploration_text_var.set(f"{float(acquisition.get('exploration', 0.35)):.2f}")
        self._candidate_pool_var.set(str(acquisition.get("candidate_pool_size", 600)))
        self._local_pool_var.set(str(acquisition.get("local_candidate_pool_size", 120)))
        self._initial_point_mode_var.set(str(acquisition.get("initial_point_mode", "specific")))

    def _sync_algorithm_config(self, show_error=True):
        if self._config is None:
            return
        try:
            acquisition = self._config.setdefault("acquisition", {})
            acquisition["exploration"] = max(0.0, min(1.0, float(self._exploration_var.get())))
            self._exploration_text_var.set(f"{float(acquisition['exploration']):.2f}")
            acquisition["candidate_pool_size"] = max(50, int(self._candidate_pool_var.get() or 600))
            acquisition["local_candidate_pool_size"] = max(0, int(self._local_pool_var.get() or 120))
            mode = str(self._initial_point_mode_var.get() or "specific").strip().lower()
            acquisition["initial_point_mode"] = "random" if mode == "random" else "specific"
        except Exception as exc:
            if show_error:
                messagebox.showerror("Optimizer Behavior", str(exc))

    @staticmethod
    def _q_reference_text():
        return (
            "Q terms: Peak uA is the measured signal height. Raw SNR is peak height / background RMS. "
            "SNR Score is Raw SNR / SNR sat, clipped 0-1. Shape rewards a centered, stable peak. "
            "Baseline rewards low/stable background. Replicate rewards consistent peak heights across scans. "
            "Success is a separate outcome measure, not the same thing as Q; in simulation it is peak-first "
            "with noise as a secondary factor, while real analysis still uses OK scans / total scans. "
            "Classic mode bounds SNR and Q. Signal-priority mode removes the SNR ceiling and uses log1p(SNR) "
            "+ log1p(peak height) so stronger peaks keep separating instead of flattening."
        )

    def _set_scoring_vars_from_config(self, cfg: dict):
        scoring = dict((cfg or {}).get("scoring") or {})
        self._score_mode_var.set(str(scoring.get("mode", "classic_bounded")))
        channel = dict(scoring.get("channel_weights") or {})
        run = dict(scoring.get("run_weights") or {})
        self._score_snr_weight_var.set(str(channel.get("snr", 0.35)))
        self._score_peak_height_weight_var.set(str(channel.get("peak_height", 0.0)))
        self._score_shape_weight_var.set(str(channel.get("peak_shape", 0.20)))
        self._score_baseline_weight_var.set(str(channel.get("baseline", 0.20)))
        self._score_replicate_weight_var.set(str(channel.get("replicate_consistency", 0.15)))
        self._score_success_weight_var.set(str(channel.get("success", 0.10)))
        self._score_snr_saturation_var.set(str(channel.get("snr_saturation", 20.0)))
        self._score_variability_penalty_var.set(str(run.get("lambda_variability", 0.20)))
        self._score_failed_penalty_var.set(str(run.get("lambda_failed", 0.40)))
        self._score_low_penalty_var.set(str(run.get("lambda_low", 0.20)))
        self._score_low_threshold_var.set(str(run.get("low_channel_threshold", 0.50)))
        self._refresh_score_formula()

    def _sync_scoring_config(self, show_error=True):
        if self._config is None:
            return
        try:
            scoring = self._config.setdefault("scoring", {})
            mode = str(self._score_mode_var.get() or "classic_bounded").strip().lower()
            scoring["mode"] = "signal_priority_unbounded" if mode == "signal_priority_unbounded" else "classic_bounded"
            channel = scoring.setdefault("channel_weights", {})
            channel["snr"] = max(0.0, float(self._score_snr_weight_var.get() or 0.0))
            channel["peak_height"] = max(0.0, float(self._score_peak_height_weight_var.get() or 0.0))
            channel["peak_shape"] = max(0.0, float(self._score_shape_weight_var.get() or 0.0))
            channel["baseline"] = max(0.0, float(self._score_baseline_weight_var.get() or 0.0))
            channel["replicate_consistency"] = max(0.0, float(self._score_replicate_weight_var.get() or 0.0))
            channel["success"] = max(0.0, float(self._score_success_weight_var.get() or 0.0))
            channel["snr_saturation"] = max(1e-12, float(self._score_snr_saturation_var.get() or 20.0))
            run = scoring.setdefault("run_weights", {})
            run["lambda_variability"] = max(0.0, float(self._score_variability_penalty_var.get() or 0.0))
            run["lambda_failed"] = max(0.0, float(self._score_failed_penalty_var.get() or 0.0))
            run["lambda_low"] = max(0.0, float(self._score_low_penalty_var.get() or 0.0))
            run["low_channel_threshold"] = max(0.0, min(1.0, float(self._score_low_threshold_var.get() or 0.5)))
            self._refresh_score_formula()
        except Exception as exc:
            if show_error:
                messagebox.showerror("Q Score Decomposition", str(exc))

    def _refresh_score_formula(self):
        try:
            mode = str(self._score_mode_var.get() or "classic_bounded").strip().lower()
            total = (
                float(self._score_snr_weight_var.get() or 0.0)
                + float(self._score_peak_height_weight_var.get() or 0.0)
                + float(self._score_shape_weight_var.get() or 0.0)
                + float(self._score_baseline_weight_var.get() or 0.0)
                + float(self._score_replicate_weight_var.get() or 0.0)
                + float(self._score_success_weight_var.get() or 0.0)
            )
        except Exception:
            self._score_formula_var.set("Q_channel = weighted component score. Enter numeric weights.")
            return
        if mode == "signal_priority_unbounded":
            self._score_formula_var.set(
                "Q_channel = (w_snr*log1p(SNR) + w_peak*log1p(Peak uA) + w_base*baseline + "
                f"w_shape*shape + w_rep*replicate + w_success*success) / {total:.3g}; "
                "Q_run = mean channels - variability/failed/low penalties. No SNR cap, no Q clipping."
            )
        else:
            self._score_formula_var.set(
                "Q_channel = (SNR + shape + baseline + replicate + success weighted scores) / "
                f"{total:.3g}; Q_run = mean channels - variability/failed/low penalties."
            )

    def _apply_signal_priority_preset(self):
        self._score_mode_var.set("signal_priority_unbounded")
        self._score_snr_weight_var.set("0.45")
        self._score_peak_height_weight_var.set("0.35")
        self._score_shape_weight_var.set("0.05")
        self._score_baseline_weight_var.set("0.12")
        self._score_replicate_weight_var.set("0.03")
        self._score_success_weight_var.set("0.00")
        self._score_snr_saturation_var.set("20.0")
        self._score_variability_penalty_var.set("0.10")
        self._score_failed_penalty_var.set("0.10")
        self._score_low_penalty_var.set("0.05")
        self._score_low_threshold_var.set("1.50")
        self._sync_scoring_config(show_error=False)

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
        self._sync_scoring_config(show_error=False)
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
        if not messagebox.askyesno(
            "Start BO Session",
            "Start a new BO session for this experiment?\n\n"
            "This snapshots the current BO config and subsequent suggestions will be recorded in a new BO session folder.",
        ):
            return
        self._save_config()
        analysis_dir = Path(exp_path) / "bo_analysis"
        self._analysis_dir_var.set(str(analysis_dir))
        try:
            self._bo_session = BOIntegrationSession.start(
                self._config_path_var.get(),
                exp_path,
                analysis_output_dir=analysis_dir,
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

    def _load_bo_session(self):
        session_mgr = getattr(self._session, "session_manager", None)
        exp_path = session_mgr.require_experiment() if session_mgr is not None else None
        if exp_path is None:
            return
        base_dir = Path(exp_path) / "bo_sessions"
        path = filedialog.askdirectory(
            title="Choose saved BO session folder",
            initialdir=str(base_dir if base_dir.exists() else exp_path),
        )
        if not path:
            return
        try:
            path = self._resolve_bo_session_load_path(path)
            loaded = BOIntegrationSession.load(path)
            self._bo_session = loaded
            self._config = dict(loaded.config)
            if loaded.config_path is not None:
                self._config_path_var.set(str(loaded.config_path))
            self._analysis_dir_var.set(str(loaded.analysis_output_dir))
            self._set_analysis_vars_from_config(self._config.get("analysis", {}))
            self._set_algorithm_vars_from_config(self._config)
            self._set_scoring_vars_from_config(self._config)
            self._channels_var.set(", ".join(str(ch) for ch in self._config.get("channels", [])))
            self._refresh_parameter_table()
            self._refresh_initial_parameters_table()
            self._suggestion = None
            self._record_dir_var.set(f"Record folder: {loaded.record_dir}")
            self._refresh_history()
            self._render_best()
            self._refresh_model_artifacts()
            self._refresh_record_files()
            self._select_latest_history_iteration()
            self._tabs.select(3)
            self._status_var.set(
                f"Loaded BO session: {loaded.session_id} "
                f"({len(loaded.observations)} completed iterations)"
            )
        except Exception as exc:
            messagebox.showerror("Load BO Session", str(exc))

    def _resolve_bo_session_load_path(self, path):
        selected = Path(path)
        if (selected / BOIntegrationSession.STATE_FILE).exists():
            return selected

        candidates = []
        search_roots = []
        if selected.name == "bo_sessions":
            search_roots.append(selected)
        if (selected / "bo_sessions").is_dir():
            search_roots.append(selected / "bo_sessions")

        for root in search_roots:
            for child in root.iterdir():
                if child.is_dir() and (child / BOIntegrationSession.STATE_FILE).exists():
                    candidates.append(child)

        if len(candidates) == 1:
            return candidates[0]
        if len(candidates) > 1:
            latest = max(candidates, key=lambda p: (p / BOIntegrationSession.STATE_FILE).stat().st_mtime)
            if messagebox.askyesno(
                "Load BO Session",
                "The selected folder contains multiple BO sessions.\n\n"
                f"Load the most recently updated one?\n\n{latest}",
            ):
                return latest
            raise FileNotFoundError(
                "Choose a specific BO session folder inside bo_sessions "
                f"that contains {BOIntegrationSession.STATE_FILE}."
            )

        raise FileNotFoundError(
            "Selected folder is not a BO session. Choose the folder containing "
            f"{BOIntegrationSession.STATE_FILE}, usually:\n\n"
            "<experiment>/bo_sessions/<bo_session_folder>"
        )

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
            title="Import analysis JSON",
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

    def _run_analysis_for_pending(self, prompt=True):
        if self._bo_session is None:
            messagebox.showwarning("BO Analysis", "Start a BO session first.")
            return None
        if self._bo_session.pending is None:
            messagebox.showwarning("BO Analysis", "No pending BO suggestion is waiting for analysis.")
            return None
        try:
            path = self._run_in_repo_analysis()
        except Exception as exc:
            if prompt:
                messagebox.showerror("BO Analysis", str(exc))
            else:
                self._auto_running = False
                self._auto_status_var.set(f"Auto loop stopped: analysis failed ({exc})")
            return None
        return self._import_analysis(path, notes="Imported from in-repo BO analysis", prompt=prompt)

    # Simulation engine
    def _engine_go_page(self, index):
        if not self._engine_pages:
            return
        self._engine_page_index = max(0, min(len(self._engine_pages) - 1, int(index)))
        for _title, frame in self._engine_pages:
            frame.pack_forget()
        title, frame = self._engine_pages[self._engine_page_index]
        frame.pack(fill="both", expand=True)
        if hasattr(self, "_engine_step_label"):
            self._engine_step_label.config(text=title)
        if hasattr(self, "_engine_back_button"):
            self._engine_back_button.config(state="disabled" if self._engine_page_index == 0 else "normal")
        if hasattr(self, "_engine_next_button"):
            self._engine_next_button.config(
                text="Finish" if self._engine_page_index == len(self._engine_pages) - 1 else "Next >",
                state="disabled" if self._engine_page_index == len(self._engine_pages) - 1 else "normal",
            )

    def _engine_next_page(self):
        if self._engine_page_index == 0 and not self._simulation_dims:
            self._engine_load_active_dimensions()
        self._engine_go_page(self._engine_page_index + 1)

    def _engine_prev_page(self):
        self._engine_go_page(self._engine_page_index - 1)

    def _engine_load_active_dimensions(self):
        if self._config is None or not hasattr(self, "_engine_dim_tree"):
            return
        try:
            self._simulation_dims = default_dimensions(self._config, limit=3)
            self._engine_rebuild_landscape_cache(refresh_plot=False)
            self._engine_refresh_dimension_tree()
            if self._simulation_dims:
                self._engine_dim_tree.selection_set("0")
                self._engine_dim_tree.see("0")
            self._engine_refresh_landscape_inspector()
            self._engine_status_var.set(f"Loaded {len(self._simulation_dims)} active simulation dimension(s).")
        except Exception as exc:
            messagebox.showerror("Simulation Engine", str(exc))

    def _engine_refresh_dimension_tree(self):
        if not hasattr(self, "_engine_dim_tree"):
            return
        for row in self._engine_dim_tree.get_children():
            self._engine_dim_tree.delete(row)
        for idx, dim in enumerate(self._simulation_dims):
            self._engine_dim_tree.insert(
                "",
                "end",
                iid=str(idx),
                text=str(dim.get("name", "")),
                values=(
                    self._fmt_raw(dim.get("minimum")),
                    self._fmt_raw(dim.get("maximum")),
                    self._fmt_raw(dim.get("optimum")),
                    self._fmt_raw(dim.get("spread")),
                    dim.get("landscape", "gaussian"),
                    self._fmt_raw(dim.get("weight", 1.0)),
                ),
            )

    def _engine_edit_dimension(self):
        if not self._simulation_dims:
            self._engine_load_active_dimensions()
        selection = self._engine_dim_tree.selection() if hasattr(self, "_engine_dim_tree") else ()
        if not selection:
            messagebox.showwarning("Simulation Engine", "Select a simulation dimension first.")
            return
        idx = int(selection[0])
        dim = dict(self._simulation_dims[idx])
        win = tk.Toplevel(self._frame)
        win.title("Edit Simulation Dimension")
        win.transient(self._frame)
        win.resizable(False, False)
        box = ttk.Frame(win, padding=12)
        box.pack(fill="both", expand=True)
        vars_by_key = {
            "minimum": tk.StringVar(value=self._fmt_raw(dim.get("minimum"))),
            "maximum": tk.StringVar(value=self._fmt_raw(dim.get("maximum"))),
            "optimum": tk.StringVar(value=self._fmt_raw(dim.get("optimum"))),
            "spread": tk.StringVar(value=self._fmt_raw(dim.get("spread"))),
            "landscape": tk.StringVar(value=str(dim.get("landscape", "gaussian"))),
            "weight": tk.StringVar(value=self._fmt_raw(dim.get("weight", 1.0))),
        }
        ttk.Label(box, text=str(dim.get("name", "")), font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        labels = [
            ("minimum", "Minimum"),
            ("maximum", "Maximum"),
            ("optimum", "Optimum"),
            ("spread", "Spread"),
            ("weight", "Weight"),
        ]
        for row, (key, label) in enumerate(labels, start=1):
            ttk.Label(box, text=f"{label}:").grid(row=row, column=0, sticky="w", pady=3)
            ttk.Entry(box, textvariable=vars_by_key[key], width=16).grid(row=row, column=1, sticky="w", pady=3)
        ttk.Label(box, text="Shape:").grid(row=len(labels) + 1, column=0, sticky="w", pady=3)
        ttk.Combobox(
            box,
            textvariable=vars_by_key["landscape"],
            values=LANDSCAPE_TYPES,
            state="readonly",
            width=14,
        ).grid(row=len(labels) + 1, column=1, sticky="w", pady=3)
        buttons = ttk.Frame(box)
        buttons.grid(row=len(labels) + 2, column=0, columnspan=2, pady=(10, 0))

        def save():
            try:
                updated = dict(dim)
                updated["minimum"] = float(vars_by_key["minimum"].get())
                updated["maximum"] = float(vars_by_key["maximum"].get())
                updated["optimum"] = float(vars_by_key["optimum"].get())
                updated["spread"] = float(vars_by_key["spread"].get())
                updated["landscape"] = vars_by_key["landscape"].get()
                updated["weight"] = float(vars_by_key["weight"].get())
                if updated["maximum"] <= updated["minimum"]:
                    raise ValueError("Maximum must be greater than minimum.")
                if updated["spread"] <= 0:
                    raise ValueError("Spread must be positive.")
                updated["optimum"] = min(max(updated["optimum"], updated["minimum"]), updated["maximum"])
                self._simulation_dims[idx] = updated
                self._engine_rebuild_landscape_cache(refresh_plot=True)
                self._engine_refresh_dimension_tree()
                self._engine_dim_tree.selection_set(str(idx))
                self._engine_refresh_landscape_inspector()
                win.destroy()
            except Exception as exc:
                messagebox.showerror("Simulation Dimension", str(exc), parent=win)

        ttk.Button(buttons, text="Save", command=save).pack(side="left", padx=4)
        ttk.Button(buttons, text="Cancel", command=win.destroy).pack(side="left", padx=4)
        win.grab_set()
        win.focus_force()

    def _engine_sim_config(self):
        if not self._simulation_dims:
            self._engine_load_active_dimensions()
        return {
            "dimensions": [dict(dim) for dim in self._simulation_dims],
            "iterations": max(1, int(self._engine_iterations_var.get() or 1)),
            "grid_size": max(5, min(45, int(self._engine_grid_var.get() or 25))),
            "seed": int(self._engine_seed_var.get() or self._config.get("random_seed", 42)),
            "measurement_noise": max(0.0, float(self._engine_measurement_noise_var.get() or 0.03)),
            "channel_noise": max(0.0, float(self._engine_channel_noise_var.get() or 0.025)),
            "peak_emphasis": max(0.0, float(self._engine_peak_emphasis_var.get() or 0.70)),
            "base_peak_uA": max(0.0, float(self._engine_base_peak_var.get() or 0.45)),
            "peak_gain_uA": max(0.0, float(self._engine_peak_gain_var.get() or 5.0)),
            "base_noise_uA": max(1e-6, float(self._engine_base_noise_var.get() or 0.08)),
            "noise_gain_uA": max(0.0, float(self._engine_noise_gain_var.get() or 0.45)),
        }

    def _engine_draw_landscape(self):
        if self._config is None:
            messagebox.showwarning("Simulation Engine", "Load a BO config first.")
            return
        try:
            self._engine_rebuild_landscape_cache(refresh_plot=False)
            self._engine_selected_index = 0
            self._engine_refresh_results()
            self._engine_refresh_landscape_inspector()
            self._engine_render_plot(show_all=True)
            self._engine_go_page(2)
            self._engine_status_var.set("Drew synthetic landscape map. Run the optimizer to add a path.")
        except Exception as exc:
            messagebox.showerror("Simulation Engine", str(exc))

    def _engine_rebuild_landscape_cache(self, refresh_plot=False):
        if self._config is None:
            return
        sim_cfg = self._engine_sim_config()
        from core.bo_simulation import SyntheticSWVSimulationEngine

        engine = SyntheticSWVSimulationEngine(self._config, sim_cfg)
        rows = []
        session = None
        if isinstance(self._simulation_result, dict):
            rows = list(self._simulation_result.get("rows") or [])
            session = self._simulation_result.get("session")
        if rows:
            rows = []
            session = None
            self._engine_selected_index = 0
            self._engine_status_var.set("Simulation dimensions changed. Landscape cache rebuilt; rerun optimizer to regenerate the path.")
        self._simulation_result = {
            "session": session,
            "engine": engine,
            "rows": rows,
            "landscape": engine.sample_landscape(sim_cfg["grid_size"]),
            "distributions": engine.dimension_distributions(),
        }
        if refresh_plot and hasattr(self, "_engine_plot_frame"):
            self._engine_refresh_results()
            self._engine_render_plot(show_all=True)

    def _engine_run_optimizer(self):
        if self._config is None:
            messagebox.showwarning("Simulation Engine", "Load a BO config first.")
            return
        try:
            self._sync_channels_from_entry(show_error=False)
            self._sync_algorithm_config(show_error=False)
            self._sync_scoring_config(show_error=False)
            sim_cfg = self._engine_sim_config()
            output_root = Path("optimizer") / "bo_simulations"
            result = run_optimizer_simulation(
                self._config,
                sim_cfg,
                output_root=output_root,
                iterations=sim_cfg["iterations"],
            )
            self._simulation_result = result
            self._engine_selected_index = max(0, len(result.get("rows", [])) - 1)
            self._engine_refresh_results()
            self._engine_refresh_landscape_inspector()
            self._engine_render_plot(show_all=True)
            self._engine_update_trace_text()
            self._engine_go_page(2)
            best = min((row for row in result["rows"]), key=lambda r: r.get("distance", 1.0), default=None)
            if best:
                self._engine_status_var.set(
                    f"Completed {len(result['rows'])} simulated BO iteration(s). "
                    f"Closest distance={best['distance']:.3f}, computed Q={best['Q_run']:.3f}, true Q={best['true_Q']:.3f}."
                )
            else:
                self._engine_status_var.set("Simulation completed without optimizer rows.")
        except Exception as exc:
            messagebox.showerror("Simulation Engine", str(exc))

    def _engine_refresh_results(self):
        if not hasattr(self, "_engine_result_tree"):
            return
        for row in self._engine_result_tree.get_children():
            self._engine_result_tree.delete(row)
        rows = (self._simulation_result or {}).get("rows", [])
        session = (self._simulation_result or {}).get("session")
        observations = session.observations if session is not None else []
        for idx, row in enumerate(rows):
            obs = observations[idx] if idx < len(observations) else {}
            peak, snr = self._engine_peak_snr_for_obs(obs)
            self._engine_result_tree.insert(
                "",
                "end",
                iid=str(idx),
                text=str(row.get("iteration", idx + 1)),
                values=(
                    self._fmt(row.get("Q_run")),
                    self._fmt(row.get("true_Q")),
                    self._fmt(row.get("distance")),
                    self._fmt(peak),
                    self._fmt(snr),
                    self._fmt_raw(row.get("begin_potential")),
                    self._fmt_raw(row.get("end_potential")),
                    self._fmt_raw(row.get("step_potential")),
                    self._fmt_raw(row.get("amplitude")),
                    self._fmt_raw(row.get("frequency")),
                ),
            )
        if rows:
            self._engine_result_tree.selection_set(str(self._engine_selected_index))
            self._engine_result_tree.see(str(self._engine_selected_index))
        else:
            self._clear_text(self._engine_trace_text)
        self._refresh_engine_q_trend()

    def _engine_step_window(self, delta):
        rows = (self._simulation_result or {}).get("rows", [])
        if not rows:
            return
        self._engine_selected_index = max(0, min(len(rows) - 1, self._engine_selected_index + int(delta)))
        self._engine_result_tree.selection_set(str(self._engine_selected_index))
        self._engine_result_tree.see(str(self._engine_selected_index))
        self._engine_render_plot(show_all=False)
        self._refresh_engine_q_trend()
        self._engine_update_trace_text()

    def _engine_show_all(self):
        rows = (self._simulation_result or {}).get("rows", [])
        if rows:
            self._engine_selected_index = len(rows) - 1
            self._engine_result_tree.selection_set(str(self._engine_selected_index))
            self._engine_result_tree.see(str(self._engine_selected_index))
        self._engine_render_plot(show_all=True)
        self._refresh_engine_q_trend()
        self._engine_update_trace_text()

    def _engine_select_iteration_from_table(self):
        if not hasattr(self, "_engine_result_tree"):
            return
        selection = self._engine_result_tree.selection()
        if not selection:
            return
        try:
            self._engine_selected_index = int(selection[0])
            self._engine_render_plot(show_all=False)
            self._refresh_engine_q_trend()
            self._engine_update_trace_text()
        except Exception:
            pass

    def _engine_apply_best_to_setup(self):
        result = self._simulation_result or {}
        rows = result.get("rows", [])
        if self._config is None or not rows:
            messagebox.showwarning("Simulation Engine", "Run an optimizer simulation first.")
            return
        best = max(rows, key=lambda row: float(row.get("Q_run", 0.0)))
        if not messagebox.askyesno(
            "Apply Simulated Best",
            "Use the best simulated parameter set as the Setup initial parameters?",
        ):
            return
        updated = {name: best.get(name) for name in PARAMETER_ORDER if best.get(name) is not None}
        self._save_initial_parameters(updated)
        self._engine_status_var.set(f"Applied simulated iteration {best['iteration']} as Setup initial parameters.")

    def _engine_render_plot(self, show_all=False):
        if not hasattr(self, "_engine_plot_frame"):
            return
        for child in self._engine_plot_frame.winfo_children():
            child.destroy()
        result = self._simulation_result or {}
        landscape = result.get("landscape") or {}
        dims = landscape.get("dimensions") or []
        points = landscape.get("points") or []
        rows = result.get("rows") or []
        try:
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            from matplotlib.figure import Figure
        except Exception as exc:
            ttk.Label(self._engine_plot_frame, text=f"Matplotlib plot unavailable: {exc}").pack(fill="both", expand=True)
            return
        fig = Figure(figsize=(7.2, 4.5), dpi=100)
        path_rows = rows if show_all else rows[: self._engine_selected_index + 1]
        if not dims or not points:
            ax = fig.add_subplot(111)
            ax.text(0.5, 0.5, "No simulation landscape yet", ha="center", va="center")
            ax.set_axis_off()
        elif len(dims) == 1:
            name = dims[0]["name"]
            ax = fig.add_subplot(111)
            ordered = sorted(points, key=lambda p: p[name])
            ax.plot([p[name] for p in ordered], [p["true_Q"] for p in ordered], color=self.ACCENT_DARK)
            if path_rows:
                ax.scatter([r.get(name) for r in path_rows], [r.get("true_Q") for r in path_rows], color="#d67b32", s=20, zorder=3)
                selected = path_rows[-1]
                ax.scatter(
                    [selected.get(name)],
                    [selected.get("true_Q")],
                    color="#ffd166",
                    edgecolors="black",
                    linewidths=0.9,
                    s=95,
                    zorder=4,
                )
            ax.set_xlabel(name)
            ax.set_ylabel("True Q")
            ax.set_ylim(0.0, 1.02)
            ax.grid(alpha=0.25)
        elif len(dims) == 2:
            x_name, y_name = dims[0]["name"], dims[1]["name"]
            ax = fig.add_subplot(111)
            x_vals = [p[x_name] for p in points]
            y_vals = [p[y_name] for p in points]
            z_vals = [p["true_Q"] for p in points]
            contour = ax.tricontourf(x_vals, y_vals, z_vals, levels=12, cmap="viridis")
            ax.tricontour(x_vals, y_vals, z_vals, levels=12, colors="white", linewidths=0.45, alpha=0.55)
            if path_rows:
                path_x = [r.get(x_name) for r in path_rows]
                path_y = [r.get(y_name) for r in path_rows]
                ax.plot(path_x, path_y, color="#d67b32", linewidth=1.6, alpha=0.9, zorder=3)
                ax.scatter(path_x, path_y, color="#d67b32", s=14, zorder=4)
                ax.scatter(
                    [path_x[-1]],
                    [path_y[-1]],
                    color="#ffd166",
                    edgecolors="black",
                    linewidths=1.0,
                    s=85,
                    zorder=5,
                )
            fig.colorbar(contour, ax=ax, label="True Q")
            ax.set_xlabel(x_name)
            ax.set_ylabel(y_name)
            ax.set_title("2D landscape map")
            ax.grid(alpha=0.2)
        else:
            x_name, y_name, z_name = dims[0]["name"], dims[1]["name"], dims[2]["name"]
            ax = fig.add_subplot(111, projection="3d")
            scatter = ax.scatter(
                [p[x_name] for p in points],
                [p[y_name] for p in points],
                [p[z_name] for p in points],
                c=[p["true_Q"] for p in points],
                cmap="viridis",
                s=4,
                alpha=0.22,
            )
            if path_rows:
                path_x = [r.get(x_name) for r in path_rows]
                path_y = [r.get(y_name) for r in path_rows]
                path_z = [r.get(z_name) for r in path_rows]
                ax.plot(
                    path_x,
                    path_y,
                    path_z,
                    color="#d67b32",
                    linewidth=1.5,
                )
                ax.scatter(path_x, path_y, path_z, color="#d67b32", s=16, depthshade=False)
                ax.scatter(
                    [path_x[-1]],
                    [path_y[-1]],
                    [path_z[-1]],
                    color="#ffd166",
                    edgecolors="black",
                    linewidths=0.8,
                    s=85,
                    depthshade=False,
                )
            fig.colorbar(scatter, ax=ax, label="True Q", shrink=0.75)
            ax.set_xlabel(x_name)
            ax.set_ylabel(y_name)
            ax.set_zlabel(z_name)
            ax.set_title("3D landscape map")
        if path_rows:
            fig.suptitle(f"Optimizer path through iteration {path_rows[-1].get('iteration', len(path_rows))}")
        fig.tight_layout()
        canvas = FigureCanvasTkAgg(fig, master=self._engine_plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        self._engine_plot_canvas = canvas

    def _engine_refresh_landscape_inspector(self, refresh_cube=True):
        if not hasattr(self, "_engine_distribution_frame"):
            return
        if refresh_cube and hasattr(self, "_engine_cube_tree"):
            for row in self._engine_cube_tree.get_children():
                self._engine_cube_tree.delete(row)
        if self._config is None:
            return
        try:
            sim_cfg = self._engine_sim_config()
            engine = ((self._simulation_result or {}).get("engine")) or None
            if engine is None:
                from core.bo_simulation import SyntheticSWVSimulationEngine

                engine = SyntheticSWVSimulationEngine(self._config, sim_cfg)
            landscape = ((self._simulation_result or {}).get("landscape")) or engine.sample_landscape(sim_cfg["grid_size"])
            distributions = ((self._simulation_result or {}).get("distributions")) or engine.dimension_distributions()
        except Exception as exc:
            ttk.Label(self._engine_distribution_frame, text=f"Inspector unavailable: {exc}").pack(fill="both", expand=True)
            return

        selected_dim = None
        selection = self._engine_dim_tree.selection() if hasattr(self, "_engine_dim_tree") else ()
        if selection:
            try:
                selected_dim = self._simulation_dims[int(selection[0])]["name"]
            except Exception:
                selected_dim = None
        curves = {
            row.get("name"): row.get("curve", [])
            for row in (distributions.get("dimensions") or [])
            if isinstance(row, dict)
        }
        curve = curves.get(selected_dim) or next(iter(curves.values()), [])
        title = selected_dim or "No dimension selected"
        self._engine_render_distribution_plot(curve, title)
        if refresh_cube:
            self._engine_refresh_cube_matrix(landscape)

    def _engine_render_distribution_plot(self, curve, title):
        try:
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            from matplotlib.figure import Figure
        except Exception as exc:
            ttk.Label(self._engine_distribution_frame, text=f"Matplotlib plot unavailable: {exc}").pack(fill="both", expand=True)
            return
        if self._engine_distribution_canvas is None or self._engine_distribution_ax is None:
            for child in self._engine_distribution_frame.winfo_children():
                child.destroy()
            fig = Figure(figsize=(6.2, 3.1), dpi=100)
            ax = fig.add_subplot(111)
            success_line, = ax.plot([], [], color="#c96b1e", linewidth=2.2, label="Success")
            q_line, = ax.plot([], [], color=self.ACCENT_DARK, linewidth=2.0, label="True Q")
            peak_line, = ax.plot([], [], color="#2f7d32", linewidth=1.4, alpha=0.9, label="Peak")
            noise_line, = ax.plot([], [], color="#5a6b84", linewidth=1.2, alpha=0.85, label="Noise")
            self._engine_distribution_lines = {
                "success": success_line,
                "q": q_line,
                "peak": peak_line,
                "noise": noise_line,
            }
            self._engine_distribution_empty_text = ax.text(0.5, 0.5, "No distribution data yet", ha="center", va="center")
            self._engine_distribution_empty_text.set_visible(False)
            ax.set_ylabel("Score")
            ax.set_ylim(0.0, 1.02)
            ax.grid(alpha=0.2)
            ax.legend(loc="best", fontsize=8)
            canvas = FigureCanvasTkAgg(fig, master=self._engine_distribution_frame)
            canvas.get_tk_widget().pack(fill="both", expand=True)
            self._engine_distribution_figure = fig
            self._engine_distribution_ax = ax
            self._engine_distribution_canvas = canvas
        else:
            fig = self._engine_distribution_figure
            ax = self._engine_distribution_ax
            canvas = self._engine_distribution_canvas
        if not curve:
            for line in self._engine_distribution_lines.values():
                line.set_data([], [])
            if self._engine_distribution_empty_text is not None:
                self._engine_distribution_empty_text.set_visible(True)
            ax.set_xlabel(title)
            ax.set_xlim(0.0, 1.0)
        else:
            x = [row.get("value") for row in curve]
            q = [row.get("true_Q") for row in curve]
            success = [row.get("success_score") for row in curve]
            peak = [row.get("peak_score") for row in curve]
            noise = [row.get("noise_score") for row in curve]
            self._engine_distribution_lines["success"].set_data(x, success)
            self._engine_distribution_lines["q"].set_data(x, q)
            self._engine_distribution_lines["peak"].set_data(x, peak)
            self._engine_distribution_lines["noise"].set_data(x, noise)
            if self._engine_distribution_empty_text is not None:
                self._engine_distribution_empty_text.set_visible(False)
            ax.set_xlabel(title)
            x_min = min(x) if x else 0.0
            x_max = max(x) if x else 1.0
            if x_max <= x_min:
                x_max = x_min + 1.0
            ax.set_xlim(x_min, x_max)
        canvas.draw_idle()

    def _engine_refresh_cube_matrix(self, landscape):
        if not hasattr(self, "_engine_cube_tree"):
            return
        dims = landscape.get("dimensions") or []
        points = landscape.get("points") or []
        names = [str(dim.get("name", "")) for dim in dims[:3]]
        if names:
            self._engine_cube_tree.heading("#0", text=" / ".join(names))
        else:
            self._engine_cube_tree.heading("#0", text="Point")
        points = sorted(
            points,
            key=lambda row: (
                -float(row.get("success_score", 0.0)),
                -float(row.get("true_Q", 0.0)),
                float(row.get("distance", 1.0)),
            ),
        )[:12]
        for idx, point in enumerate(points, start=1):
            label = ", ".join(f"{name}={self._fmt_raw(point.get(name))}" for name in names if name)
            self._engine_cube_tree.insert(
                "",
                "end",
                iid=str(idx),
                text=label or f"Point {idx}",
                values=(
                    self._fmt(point.get("true_Q")),
                    self._fmt(point.get("success_score")),
                    self._fmt(point.get("peak_score")),
                    self._fmt(point.get("noise_score")),
                ),
            )

    def _engine_preview_selected_cube_point(self):
        if not hasattr(self, "_engine_cube_tree"):
            return
        selection = self._engine_cube_tree.selection()
        if not selection:
            return
        result = self._simulation_result or {}
        engine = result.get("engine")
        landscape = result.get("landscape") or {}
        if engine is None:
            return
        points = sorted(
            landscape.get("points") or [],
            key=lambda row: (
                -float(row.get("success_score", 0.0)),
                -float(row.get("true_Q", 0.0)),
                float(row.get("distance", 1.0)),
            ),
        )[:12]
        try:
            idx = max(0, int(selection[0]) - 1)
        except Exception:
            return
        if idx >= len(points):
            return
        point = points[idx]
        params = resolve_initial_parameters(self._config)
        for dim in landscape.get("dimensions") or []:
            name = str(dim.get("name", ""))
            if name and name in point:
                params[name] = point[name]
        payload = engine.analysis_payload(params, iteration=0)
        self._engine_write_fake_payload_preview(payload, title="Fake data preview from selected map cell")

    def _engine_write_fake_payload_preview(self, payload, title="Fake data preview"):
        truth = dict(payload.get("simulation_truth") or {})
        metrics = dict(payload.get("channel_metrics") or {})
        traces = dict(payload.get("swv_traces") or {})
        params = dict(((payload.get("simulation_engine") or {}).get("parameters")) or {})
        self._engine_render_trace_plot(traces, title="Fake SWV plot")
        lines = [title, ""]
        lines.append(f"True Q: {self._fmt(truth.get('true_Q'))}")
        lines.append(f"Success: {self._fmt(truth.get('success_score'))}")
        lines.append(f"Peak component: {self._fmt(truth.get('peak_score'))}")
        lines.append(f"Noise component: {self._fmt(truth.get('noise_score'))}")
        lines.append("")
        lines.append("Parameters:")
        for name in PARAMETER_ORDER:
            if name in params:
                lines.append(f"  {name}: {self._fmt_raw(params.get(name))}")
        if metrics:
            first_key = sorted(metrics.keys(), key=lambda value: int(value))[0]
            first = metrics.get(first_key) or {}
            lines.extend(
                [
                    "",
                    f"Example channel: {first_key}",
                    f"  Peak height: {self._fmt(first.get('mean_peak_current_uA'))} uA",
                    f"  Raw SNR: {self._fmt(first.get('snr'))}",
                    f"  Success score: {self._fmt(first.get('success_score'))}",
                ]
            )
        if traces:
            ch, trace = next(iter(traces.items()))
            volts = trace.get("voltage_v", [])[:8]
            currents = trace.get("current_uA", [])[:8]
            pairs = ", ".join(f"{v:g}V/{i:g}uA" for v, i in zip(volts, currents))
            lines.extend(["", f"Trace preview ch {ch}:", pairs])
        self._write_text(self._engine_trace_text, "\n".join(lines))

    def _engine_render_trace_plot(self, traces, title="Synthetic SWV trace"):
        if not hasattr(self, "_engine_trace_plot_frame"):
            return
        for child in self._engine_trace_plot_frame.winfo_children():
            child.destroy()
        if not isinstance(traces, dict) or not traces:
            ttk.Label(self._engine_trace_plot_frame, text="No synthetic SWV trace yet").pack(fill="both", expand=True)
            return
        try:
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            from matplotlib.figure import Figure
        except Exception as exc:
            ttk.Label(self._engine_trace_plot_frame, text=f"Matplotlib plot unavailable: {exc}").pack(fill="both", expand=True)
            return
        fig = Figure(figsize=(5.2, 2.6), dpi=100)
        ax = fig.add_subplot(111)
        palette = [self.ACCENT_DARK, "#d67b32", "#2f7d32"]
        for idx, (ch, trace) in enumerate(sorted(traces.items(), key=lambda item: int(item[0]))):
            volts = trace.get("voltage_v", [])
            currents = trace.get("current_uA", [])
            if volts and currents:
                ax.plot(volts, currents, color=palette[idx % len(palette)], linewidth=1.4, label=f"Ch {ch}")
        ax.set_title(title)
        ax.set_xlabel("Voltage (V)")
        ax.set_ylabel("Current (uA)")
        ax.grid(alpha=0.2)
        if len(traces) > 1:
            ax.legend(loc="best", fontsize=8)
        fig.tight_layout()
        canvas = FigureCanvasTkAgg(fig, master=self._engine_trace_plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    def _engine_update_trace_text(self):
        result = self._simulation_result or {}
        session = result.get("session")
        if session is None or not session.observations:
            self._engine_render_trace_plot({}, title="Synthetic SWV trace")
            self._clear_text(self._engine_trace_text)
            return
        idx = max(0, min(self._engine_selected_index, len(session.observations) - 1))
        obs = session.observations[idx]
        truth = obs.get("simulation_truth", {})
        traces = obs.get("swv_trace_preview", {})
        self._engine_render_trace_plot(traces, title=f"Iteration {obs.get('iteration')} synthetic SWV")
        lines = [
            f"Iteration {obs.get('iteration')}",
            f"Computed Q_run: {self._fmt(obs.get('Q_run'))}",
            f"True simulated Q: {self._fmt(truth.get('true_Q'))}",
            f"Simulated success: {self._fmt(truth.get('success_score'))}",
            f"Distance to true optimum: {self._fmt(truth.get('normalized_distance'))}",
            "",
            "Parameters:",
        ]
        for name in PARAMETER_ORDER:
            if name in obs.get("params", {}):
                lines.append(f"  {name}: {self._fmt_raw(obs['params'].get(name))}")
        peak, snr = self._engine_peak_snr_for_obs(obs)
        lines.extend(
            [
                "",
                f"Mean channel peak height: {self._fmt(peak)} uA",
                f"Mean raw SNR: {self._fmt(snr)}",
                f"Truth components: peak {self._fmt(truth.get('peak_score'))}, noise {self._fmt(truth.get('noise_score'))}, shape {self._fmt(truth.get('shape_score'))}",
            ]
        )
        lines.extend([""] + self._q_breakdown_lines(obs, config=session.config if session is not None else None))
        if traces:
            ch, trace = next(iter(traces.items()))
            volts = trace.get("voltage_v", [])[:8]
            currents = trace.get("current_uA", [])[:8]
            pairs = ", ".join(f"{v:g}V/{i:g}uA" for v, i in zip(volts, currents))
            lines.extend(["", f"Trace preview ch {ch}:", pairs])
        self._write_text(self._engine_trace_text, "\n".join(lines))

    def _refresh_engine_q_trend(self):
        if not hasattr(self, "_engine_q_plot_frame"):
            return
        rows = list((self._simulation_result or {}).get("rows") or [])
        self._render_q_trend_plot(
            self._engine_q_plot_frame,
            rows,
            empty_text="Run an optimizer simulation to see Q over iterations.",
            include_true_q=True,
            selected_index=self._engine_selected_index if rows else None,
        )

    def _refresh_analysis_q_trend(self):
        if not hasattr(self, "_analysis_q_plot_frame"):
            return
        rows = list(self._bo_session.observations) if self._bo_session is not None else []
        metric = self._analysis_trend_metric_var.get() or "Q_run"
        self._render_metric_trend_plot(
            self._analysis_q_plot_frame,
            rows,
            metric,
            empty_text="Import analysis results to see trends over iterations.",
        )

    def _render_metric_trend_plot(self, parent, rows, metric, empty_text):
        for child in parent.winfo_children():
            child.destroy()
        try:
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            from matplotlib.figure import Figure
        except Exception as exc:
            ttk.Label(parent, text=f"Matplotlib plot unavailable: {exc}").pack(fill="both", expand=True)
            return

        fig = Figure(figsize=(6.2, 2.8), dpi=100)
        ax = fig.add_subplot(111)
        values = []
        iterations = []
        for idx, row in enumerate(rows):
            value = self._analysis_trend_value(row, metric)
            if value is None:
                continue
            try:
                values.append(float(value))
                iterations.append(int(row.get("iteration", idx + 1)))
            except (TypeError, ValueError):
                continue

        if not values:
            ax.text(0.5, 0.5, empty_text, ha="center", va="center")
            ax.set_axis_off()
        else:
            ax.plot(iterations, values, marker="o", color=self.ACCENT_DARK, linewidth=1.8, label=metric)
            if metric == "Q_run":
                best_so_far = []
                running_best = 0.0
                for value in values:
                    running_best = max(running_best, value)
                    best_so_far.append(running_best)
                ax.plot(iterations, best_so_far, color="#d67b32", linewidth=1.6, label="Best so far")
                ax.set_ylim(bottom=0.0)
            ax.set_xlabel("BO iteration")
            ax.set_ylabel(metric)
            ax.set_title(f"{metric} over BO iterations")
            ax.grid(alpha=0.25)
            ax.legend(loc="best", fontsize=8)
        fig.tight_layout()
        canvas = FigureCanvasTkAgg(fig, master=parent)
        if values:
            canvas.mpl_connect(
                "button_press_event",
                lambda event, points=list(zip(iterations, values)): self._on_analysis_trend_click(event, points),
            )
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    def _on_analysis_trend_click(self, event, points):
        if event.inaxes is None or not points:
            return
        try:
            click_x, click_y = event.x, event.y
            transformed = event.inaxes.transData.transform(points)
            nearest_iteration = None
            nearest_distance = None
            for (iteration, _value), (px, py) in zip(points, transformed):
                distance = ((px - click_x) ** 2 + (py - click_y) ** 2) ** 0.5
                if nearest_distance is None or distance < nearest_distance:
                    nearest_distance = distance
                    nearest_iteration = iteration
            if nearest_iteration is None or nearest_distance is None or nearest_distance > 18:
                return
            item_id = str(int(nearest_iteration))
            if item_id not in self._history_tree.get_children():
                return
            self._history_tree.selection_set(item_id)
            self._history_tree.focus(item_id)
            self._history_tree.see(item_id)
            self._select_history_iteration(item_id)
        except Exception:
            return

    def _render_q_trend_plot(self, parent, rows, empty_text, include_true_q=False, selected_index=None):
        for child in parent.winfo_children():
            child.destroy()
        try:
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            from matplotlib.figure import Figure
        except Exception as exc:
            ttk.Label(parent, text=f"Matplotlib plot unavailable: {exc}").pack(fill="both", expand=True)
            return

        fig = Figure(figsize=(6.2, 2.8), dpi=100)
        ax = fig.add_subplot(111)
        if not rows:
            ax.text(0.5, 0.5, empty_text, ha="center", va="center")
            ax.set_axis_off()
        else:
            iterations = [int(row.get("iteration", idx + 1)) for idx, row in enumerate(rows)]
            q_values = [float(row.get("Q_run", 0.0) or 0.0) for row in rows]
            best_so_far = []
            running_best = 0.0
            for value in q_values:
                running_best = max(running_best, value)
                best_so_far.append(running_best)
            ax.plot(iterations, q_values, marker="o", color=self.ACCENT_DARK, linewidth=1.8, label="Q_run")
            ax.plot(iterations, best_so_far, color="#d67b32", linewidth=1.6, label="Best so far")
            if include_true_q:
                true_rows = [(iteration, row.get("true_Q")) for iteration, row in zip(iterations, rows) if row.get("true_Q") is not None]
                if true_rows:
                    ax.plot(
                        [iteration for iteration, _value in true_rows],
                        [float(value) for _iteration, value in true_rows],
                        color="#2f7d32",
                        linewidth=1.2,
                        linestyle="--",
                        label="True Q",
                    )
            if selected_index is not None and 0 <= int(selected_index) < len(rows):
                idx = int(selected_index)
                ax.scatter(
                    [iterations[idx]],
                    [q_values[idx]],
                    color="#ffd166",
                    edgecolors="black",
                    linewidths=0.8,
                    s=80,
                    zorder=4,
                    label="Selected",
                )
            ax.set_ylim(0.0, 1.02)
            ax.set_xlabel("BO iteration")
            ax.set_ylabel("Q")
            ax.set_title("Q improvement over iterations")
            ax.grid(alpha=0.25)
            ax.legend(loc="best", fontsize=8)
        fig.tight_layout()
        canvas = FigureCanvasTkAgg(fig, master=parent)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    @staticmethod
    def _engine_peak_snr_for_obs(obs):
        metrics = (obs or {}).get("channel_metrics", {})
        if not isinstance(metrics, dict) or not metrics:
            return None, None
        peaks = []
        snrs = []
        for data in metrics.values():
            if not isinstance(data, dict):
                continue
            if data.get("mean_peak_current_uA") is not None:
                peaks.append(float(data.get("mean_peak_current_uA")))
            elif data.get("median_peak_current_uA") is not None:
                peaks.append(float(data.get("median_peak_current_uA")))
            if data.get("snr") is not None:
                snrs.append(float(data.get("snr")))
        peak = sum(peaks) / len(peaks) if peaks else None
        snr = sum(snrs) / len(snrs) if snrs else None
        return peak, snr

    def _run_in_repo_analysis(self) -> Path:
        if self._bo_session is None or self._bo_session.pending is None:
            raise RuntimeError("No pending BO suggestion is waiting for analysis")
        session_mgr = getattr(self._session, "session_manager", None)
        exp_path = session_mgr.require_experiment() if session_mgr is not None else None
        if exp_path is None:
            raise RuntimeError("An active experiment folder is required for BO analysis")
        self._save_config()
        output_dir = Path(exp_path) / "bo_analysis"
        self._analysis_dir_var.set(str(output_dir))
        path = self._bo_session.run_pending_analysis(
            folders=[exp_path],
            output_dir=output_dir,
            analysis=dict(self._config.get("analysis") or {}),
        )
        self._status_var.set(f"Analysis completed: {path.name}")
        return path

    def _import_analysis(self, path, notes=None, prompt=True):
        if prompt:
            notes = simpledialog.askstring("BO Analysis Notes", "Notes for this BO result:", parent=self._frame)
        try:
            obs = self._bo_session.import_analysis(path, notes=notes or "")
            self._suggestion = None
            self._render_scores(obs)
            self._refresh_history()
            self._select_history_iteration(str(obs.get("iteration")))
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
        except ValueError:
            messagebox.showerror("Auto Loop", "Target iterations must be numeric.")
            return
        if target < 1:
            messagebox.showerror("Auto Loop", "Target iterations must be at least 1.")
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
        self._auto_status_var.set("Queue complete. Running BO analysis.")
        obs = self._run_analysis_for_pending(prompt=False)
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
                values=(mode, str(p.get("space", "discrete")), value_text, tie),
                tags=(mode.lower(),),
            )

    @staticmethod
    def _values_text(param_cfg):
        mode = str(param_cfg.get("mode", "locked")).lower()
        if mode == "active":
            if str(param_cfg.get("space", "discrete")).lower() == "continuous":
                step = param_cfg.get("step")
                step_text = "" if step in (None, "") else f", step {step}"
                return (
                    f"{param_cfg.get('min')}..{param_cfg.get('max')} "
                    f"{param_cfg.get('scale', 'linear')} sigma {param_cfg.get('proposal_sigma', '')}{step_text}"
                )
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
        space_var = tk.StringVar(value=str(current.get("space", "discrete")))
        values_var = tk.StringVar(value=", ".join(str(v) for v in current.get("values", [])))
        value_var = tk.StringVar(value=str(current.get("value", "")))
        tie_var = tk.StringVar(value=str(current.get("tie_to", "begin_potential")))
        min_var = tk.StringVar(value=str(current.get("min", "")))
        max_var = tk.StringVar(value=str(current.get("max", "")))
        step_var = tk.StringVar(value="" if current.get("step") in (None, "") else str(current.get("step")))
        scale_var = tk.StringVar(value=str(current.get("scale", current.get("encoding", "linear"))))
        sigma_var = tk.StringVar(value=str(current.get("proposal_sigma", "")))

        ttk.Label(box, text="Mode:").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Combobox(box, textvariable=mode_var, values=("active", "locked", "tied"), state="readonly", width=16).grid(
            row=0, column=1, sticky="w", pady=4
        )
        ttk.Label(box, text="Space:").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Combobox(box, textvariable=space_var, values=("discrete", "continuous"), state="readonly", width=16).grid(
            row=1, column=1, sticky="w", pady=4
        )
        ttk.Label(box, text="Active values:").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Entry(box, textvariable=values_var, width=48).grid(row=2, column=1, columnspan=3, sticky="ew", pady=4)
        ttk.Label(box, text="Continuous min/max:").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Entry(box, textvariable=min_var, width=12).grid(row=3, column=1, sticky="w", pady=4)
        ttk.Entry(box, textvariable=max_var, width=12).grid(row=3, column=1, padx=(96, 0), sticky="w", pady=4)
        ttk.Label(box, text="Scale / sigma:").grid(row=4, column=0, sticky="w", pady=4)
        ttk.Combobox(box, textvariable=scale_var, values=("linear", "log"), width=12).grid(row=4, column=1, sticky="w", pady=4)
        ttk.Entry(box, textvariable=sigma_var, width=12).grid(row=4, column=1, padx=(96, 0), sticky="w", pady=4)
        ttk.Label(box, text="Optional step:").grid(row=5, column=0, sticky="w", pady=4)
        ttk.Entry(box, textvariable=step_var, width=12).grid(row=5, column=1, sticky="w", pady=4)
        ttk.Label(box, text="Locked value:").grid(row=6, column=0, sticky="w", pady=4)
        ttk.Entry(box, textvariable=value_var, width=18).grid(row=6, column=1, sticky="w", pady=4)
        ttk.Label(box, text="Tie to:").grid(row=7, column=0, sticky="w", pady=4)
        ttk.Combobox(box, textvariable=tie_var, values=PARAMETER_ORDER, width=24).grid(row=7, column=1, sticky="w", pady=4)

        buttons = ttk.Frame(box)
        buttons.grid(row=8, column=0, columnspan=2, pady=(10, 0))

        def save():
            try:
                updated = dict(current)
                updated["mode"] = mode_var.get()
                updated["space"] = space_var.get()
                updated["values"] = self._parse_float_list(values_var.get())
                if min_var.get().strip():
                    updated["min"] = float(min_var.get())
                if max_var.get().strip():
                    updated["max"] = float(max_var.get())
                updated["scale"] = scale_var.get()
                updated["proposal_sigma"] = float(sigma_var.get() or 0.15)
                updated["step"] = None if not step_var.get().strip() else float(step_var.get())
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
        params = self._config.setdefault("parameters", {})
        resolved = {name: float(updated[name]) for name in PARAMETER_ORDER}
        for name in PARAMETER_ORDER:
            param_cfg = params.setdefault(name, {})
            mode = str(param_cfg.get("mode", "locked")).lower()
            if mode == "locked":
                param_cfg["value"] = resolved[name]
            elif mode == "tied":
                tie_to = param_cfg.get("tie_to") or "begin_potential"
                if tie_to in resolved:
                    resolved[name] = float(resolved[tie_to])
        self._config["initial_parameters"] = resolved
        self._config.pop("initial_method", None)
        self._config.pop("initial_design", None)
        self._refresh_parameter_table()
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
            entry_state = "normal"
            if title == "Edit Initial Parameters":
                param_cfg = (self._config or {}).get("parameters", {}).get(name, {})
                if str(param_cfg.get("mode", "")).lower() == "tied":
                    entry_state = "disabled"
            ttk.Entry(box, textvariable=var, width=18, state=entry_state).grid(row=row, column=1, sticky="w", pady=3)
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
            lines.append(f"{name}: {self._fmt_raw(self._suggestion.params.get(name))}")
        self._write_text(self._suggestion_text, "\n".join(lines))

    def _render_scores(self, observation):
        for row in self._score_tree.get_children():
            self._score_tree.delete(row)
        if hasattr(self, "_q_equation_text"):
            config = self._bo_session.config if self._bo_session else self._config
            self._write_text(self._q_equation_text, "\n".join(self._q_equation_lines(config)))
        components = observation["quality"].get("channel_components", {})
        channel_metrics = observation.get("channel_metrics", {})
        for ch, data in sorted(components.items(), key=lambda item: int(item[0])):
            metrics = channel_metrics.get(str(ch), {}) if isinstance(channel_metrics, dict) else {}
            self._score_tree.insert(
                "",
                "end",
                iid=str(ch),
                text=str(ch),
                values=(
                    self._fmt(data.get("Q_channel")),
                    self._fmt(self._channel_peak_height(metrics)),
                    self._fmt(data.get("snr_raw")),
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
        self._history_rows = {}
        if self._bo_session is None:
            self._selected_history_observation = None
            for row in self._score_tree.get_children():
                self._score_tree.delete(row)
            if hasattr(self, "_q_equation_text"):
                self._write_text(self._q_equation_text, "\n".join(self._q_equation_lines(self._config)))
            self._render_raw_traces(None)
            self._render_corrected_traces(None)
            self._refresh_analysis_q_trend()
            return
        for obs in self._bo_session.observations:
            q = obs.get("quality", {})
            params = obs.get("params", {})
            iteration = str(obs.get("iteration"))
            peak_uA, rms_uA = self._observation_peak_rms(obs)
            self._history_rows[iteration] = obs
            self._history_tree.insert(
                "",
                "end",
                iid=iteration,
                text=iteration,
                values=(
                    self._fmt(obs.get("Q_run")),
                    self._fmt(q.get("mean_Q_channel")),
                    self._fmt(q.get("std_Q_channel")),
                    self._fmt(q.get("failed_channel_fraction")),
                    self._fmt(q.get("low_channel_fraction")),
                    self._fmt(peak_uA),
                    self._fmt(rms_uA),
                    self._fmt_raw(params.get("begin_potential")),
                    self._fmt_raw(params.get("end_potential")),
                    self._fmt_raw(params.get("step_potential")),
                    self._fmt_raw(params.get("amplitude")),
                    self._fmt_raw(params.get("frequency")),
                    self._fmt_raw(params.get("conditioning_potential")),
                    self._fmt_raw(params.get("conditioning_time")),
                ),
            )
        if not self._history_rows:
            for row in self._score_tree.get_children():
                self._score_tree.delete(row)
            self._render_raw_traces(None)
            self._render_corrected_traces(None)
        self._refresh_analysis_q_trend()

    def _select_history_iteration(self, iteration=None):
        if self._bo_session is None:
            return
        target = iteration
        if target is None:
            selection = self._history_tree.selection()
            if not selection:
                return
            target = selection[0]
        obs = self._history_rows.get(str(target))
        if obs is None:
            return
        self._selected_history_observation = obs
        self._render_scores(obs)
        self._render_raw_traces(obs)
        self._render_corrected_traces(obs)
        self._status_var.set(
            f"Viewing BO iteration {obs.get('iteration')}: Q_run={float(obs.get('Q_run', 0.0)):.3f}"
        )
        self._restore_tree_focus(self._history_tree)

    def _on_score_tree_select(self):
        self._render_raw_traces(self._selected_history_observation)
        self._render_corrected_traces(self._selected_history_observation)
        self._restore_tree_focus(self._score_tree)

    def _select_latest_history_iteration(self):
        if not self._history_rows:
            return
        latest = max(self._history_rows, key=lambda value: int(value))
        self._history_tree.selection_set(latest)
        self._history_tree.focus(latest)
        self._history_tree.see(latest)
        self._select_history_iteration(latest)

    def _move_history_selection(self, direction, event=None):
        if self._bo_session is None:
            return "break"
        items = list(self._history_tree.get_children())
        if not items:
            return "break"
        selection = self._history_tree.selection()
        current = selection[0] if selection else self._history_tree.focus()
        try:
            idx = items.index(current)
        except ValueError:
            idx = 0 if direction > 0 else len(items) - 1
        idx = max(0, min(len(items) - 1, idx + int(direction)))
        target = items[idx]
        self._history_tree.selection_set(target)
        self._history_tree.focus(target)
        self._history_tree.see(target)
        self._select_history_iteration(target)
        return "break"

    @staticmethod
    def _focus_tree_on_click(event):
        try:
            event.widget.focus_set()
        except Exception:
            pass

    def _set_active_results_tree(self, name):
        self._active_results_tree = name

    def _restore_tree_focus(self, tree):
        try:
            selection = tree.selection()
            if selection:
                tree.focus(selection[-1])
            def restore():
                try:
                    tree.focus_set()
                    tree.focus_force()
                except Exception:
                    pass
            tree.after_idle(restore)
        except Exception:
            pass

    def _route_results_arrow(self, direction, event=None):
        widget_class = ""
        try:
            widget_class = str(event.widget.winfo_class())
        except Exception:
            pass
        if widget_class in ("Entry", "TEntry", "Text", "TCombobox", "Spinbox", "TSpinbox"):
            return None
        if self._active_results_tree == "history":
            return self._move_history_selection(direction, event)
        if self._active_results_tree == "score":
            return self._move_score_selection(direction, event)
        return None

    def _move_score_selection(self, direction, event=None):
        if self._selected_history_observation is None:
            return "break"
        items = list(self._score_tree.get_children())
        if not items:
            return "break"
        selection = list(self._score_tree.selection())
        current = selection[-1] if selection else self._score_tree.focus()
        try:
            idx = items.index(current)
        except ValueError:
            idx = 0 if direction > 0 else len(items) - 1
        idx = max(0, min(len(items) - 1, idx + int(direction)))
        target = items[idx]
        self._score_tree.selection_set(target)
        self._score_tree.focus(target)
        self._score_tree.see(target)
        self._render_raw_traces(self._selected_history_observation)
        self._render_corrected_traces(self._selected_history_observation)
        return "break"

    def _on_history_double_click(self, event):
        if self._history_tree.identify_region(event.x, event.y) != "heading":
            return None
        metric = self._history_metric_for_column(self._history_tree.identify_column(event.x))
        if metric is None:
            return "break"
        self._analysis_trend_metric_var.set(metric)
        self._refresh_analysis_q_trend()
        if hasattr(self, "_history_tabs"):
            self._history_tabs.select(1)
        self._status_var.set(f"Trend plot updated: {metric}")
        return "break"

    @staticmethod
    def _history_metric_for_column(column_id):
        try:
            index = int(str(column_id).lstrip("#")) - 1
        except ValueError:
            return None
        labels = (
            "Q_run",
            "Mean",
            "Std",
            "Failed",
            "Low",
            "Peak uA",
            "RMS uA",
            "Begin",
            "End",
            "Step",
            "Amp",
            "Freq",
            "Cond E",
            "Cond t",
        )
        metrics = {
            "Q_run": "Q_run",
            "Mean": "Mean channel Q",
            "Std": "Std channel Q",
            "Failed": "Failed fraction",
            "Low": "Low fraction",
            "Peak uA": "Mean peak uA",
            "RMS uA": "Mean RMS uA",
            "Begin": "Begin potential",
            "End": "End potential",
            "Step": "Step potential",
            "Amp": "Amplitude",
            "Freq": "Frequency",
            "Cond E": "Conditioning potential",
            "Cond t": "Conditioning time",
        }
        if index < 0 or index >= len(labels):
            return None
        return metrics.get(labels[index])

    def _render_best(self):
        if self._history_rows:
            return
        self._render_raw_traces(None)
        self._render_corrected_traces(None)

    def _render_raw_traces(self, observation):
        if not hasattr(self, "_raw_trace_frame"):
            return
        for child in self._raw_trace_frame.winfo_children():
            child.destroy()
        if not observation:
            ttk.Label(
                self._raw_trace_frame,
                text="Select a completed BO iteration in the history table to view raw SWV traces.",
                anchor="center",
                justify="center",
            ).pack(fill="both", expand=True)
            return

        rows = self._raw_trace_rows_for_observation(observation)
        selected_channels = self._selected_score_channels()
        if selected_channels:
            rows = [
                row for row in rows
                if str(row.get("channel") or "") in selected_channels
            ]
        if not rows:
            if selected_channels:
                message = (
                    "No raw SWV CSV paths were found for the selected channel(s) "
                    f"{', '.join(sorted(selected_channels, key=self._channel_sort_key))} in this iteration."
                )
            else:
                message = (
                    "No raw SWV CSV paths were found for this iteration.\n\n"
                    "Expected archived_measurements in iter_XXX_quality.json, or file_path rows "
                    "in the retained analysis results CSV."
                )
            ttk.Label(
                self._raw_trace_frame,
                text=message,
                anchor="center",
                justify="center",
            ).pack(fill="both", expand=True)
            return

        try:
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            from matplotlib.figure import Figure
        except Exception as exc:
            ttk.Label(self._raw_trace_frame, text=f"Matplotlib plot unavailable: {exc}").pack(fill="both", expand=True)
            return

        fig = Figure(figsize=(6.2, 3.2), dpi=100)
        ax = fig.add_subplot(111)
        plotted = 0
        errors = []
        palette = ["#155e63", "#d67b32", "#2f7d32", "#6d597a", "#5a6b84", "#b56576", "#457b9d", "#8a5a44"]
        for idx, row in enumerate(rows):
            path = Path(row["path"])
            try:
                volts, currents = load_swv_csv(str(path))
            except Exception as exc:
                errors.append(f"{path.name}: {exc}")
                continue
            if len(volts) == 0 or len(currents) == 0:
                continue
            channel = row.get("channel")
            scan = row.get("scan")
            label = f"Ch {channel}" if channel not in (None, "") else path.stem
            if scan not in (None, ""):
                label = f"{label} scan {scan}"
            ax.plot(volts, currents, linewidth=1.1, alpha=0.88, color=palette[idx % len(palette)], label=label)
            plotted += 1

        if plotted == 0:
            msg = "Raw SWV CSV paths were found, but none could be plotted."
            if errors:
                msg += "\n\n" + "\n".join(errors[:6])
            ttk.Label(self._raw_trace_frame, text=msg, anchor="center", justify="left").pack(fill="both", expand=True)
            return

        channel_suffix = ""
        if selected_channels:
            channel_suffix = f" | Ch {', '.join(sorted(selected_channels, key=self._channel_sort_key))}"
        ax.set_title(
            f"Iteration {observation.get('iteration')} raw SWV traces{channel_suffix} | Q_run={self._fmt(observation.get('Q_run'))}"
        )
        ax.set_xlabel("Voltage (V)")
        ax.set_ylabel("Current (uA)")
        ax.grid(alpha=0.25)
        if plotted <= 16:
            ax.legend(loc="best", fontsize=7)
        fig.tight_layout()
        canvas = FigureCanvasTkAgg(fig, master=self._raw_trace_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        if errors:
            ttk.Label(
                self._raw_trace_frame,
                text=f"{len(errors)} trace file(s) could not be loaded.",
                foreground="#8a5a44",
            ).pack(fill="x")

    def _render_corrected_traces(self, observation):
        if not hasattr(self, "_corrected_trace_frame"):
            return
        for child in self._corrected_trace_frame.winfo_children():
            child.destroy()
        if not observation:
            ttk.Label(
                self._corrected_trace_frame,
                text="Select a completed BO iteration to view smoothed corrected analysis traces.",
                anchor="center",
                justify="center",
            ).pack(fill="both", expand=True)
            return

        rows = self._corrected_trace_rows_for_observation(observation)
        selected_channels = self._selected_score_channels()
        if selected_channels:
            rows = [row for row in rows if str(row.get("channel") or "") in selected_channels]
        if not rows:
            message = (
                "No smoothed corrected traces were found for this iteration.\n\n"
                "Expected voltage and smoothed_corrected_current columns in "
                "bo_analysis/bo_iter_XXX_results.csv."
            )
            if selected_channels:
                message = (
                    "No smoothed corrected traces were found for the selected channel(s) "
                    f"{', '.join(sorted(selected_channels, key=self._channel_sort_key))} in this iteration."
                )
            ttk.Label(
                self._corrected_trace_frame,
                text=message,
                anchor="center",
                justify="center",
            ).pack(fill="both", expand=True)
            return

        try:
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            from matplotlib.figure import Figure
        except Exception as exc:
            ttk.Label(self._corrected_trace_frame, text=f"Matplotlib plot unavailable: {exc}").pack(fill="both", expand=True)
            return

        fig = Figure(figsize=(6.2, 3.2), dpi=100)
        ax = fig.add_subplot(111)
        palette = ["#155e63", "#d67b32", "#2f7d32", "#6d597a", "#5a6b84", "#b56576", "#457b9d", "#8a5a44"]
        for idx, row in enumerate(rows):
            channel = row.get("channel")
            scan = row.get("scan")
            label = f"Ch {channel}" if channel not in (None, "") else row.get("label", "trace")
            if scan not in (None, ""):
                label = f"{label} scan {scan}"
            ax.plot(
                row["voltage"],
                row["current"],
                linewidth=1.1,
                alpha=0.88,
                color=palette[idx % len(palette)],
                label=label,
            )

        channel_suffix = ""
        if selected_channels:
            channel_suffix = f" | Ch {', '.join(sorted(selected_channels, key=self._channel_sort_key))}"
        ax.set_title(f"Iteration {observation.get('iteration')} smoothed corrected traces{channel_suffix}")
        ax.set_xlabel("Voltage (V)")
        ax.set_ylabel("Corrected current (uA)")
        ax.grid(alpha=0.25)
        if len(rows) <= 16:
            ax.legend(loc="best", fontsize=7)
        fig.tight_layout()
        canvas = FigureCanvasTkAgg(fig, master=self._corrected_trace_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    def _corrected_trace_rows_for_observation(self, observation):
        rows = []
        seen = set()
        for results_path in self._analysis_results_paths_for_observation(observation):
            if not results_path.exists():
                continue
            try:
                with open(results_path, "r", encoding="utf-8-sig", newline="") as fh:
                    for record in csv.DictReader(fh):
                        voltage = self._parse_trace_array(record.get("voltage"))
                        current = self._parse_trace_array(record.get("smoothed_corrected_current"))
                        if not voltage or not current:
                            continue
                        n = min(len(voltage), len(current))
                        if n <= 0:
                            continue
                        file_path = self._resolve_observation_file_path(
                            record.get("file_path") or record.get("file_name"),
                            observation,
                        )
                        channel = record.get("channel") or self._infer_channel_from_path(file_path)
                        scan = record.get("scan_number")
                        key = (str(file_path), str(channel), str(scan or ""))
                        if key in seen:
                            continue
                        seen.add(key)
                        rows.append(
                            {
                                "voltage": voltage[:n],
                                "current": current[:n],
                                "channel": channel,
                                "scan": scan,
                                "label": record.get("file_name") or Path(file_path).stem,
                            }
                        )
            except Exception:
                continue
        return sorted(
            rows,
            key=lambda row: (self._channel_sort_key(row.get("channel")), str(row.get("scan") or ""), str(row.get("label") or "")),
        )

    @staticmethod
    def _parse_trace_array(value):
        if value in (None, ""):
            return []
        try:
            parsed = ast.literal_eval(str(value))
        except Exception:
            return []
        if not isinstance(parsed, (list, tuple)):
            return []
        out = []
        for item in parsed:
            try:
                out.append(float(item))
            except (TypeError, ValueError):
                continue
        return out

    def _analysis_results_paths_for_observation(self, observation):
        paths = []
        analysis_record_path = self._resolve_observation_file_path(observation.get("analysis_record"), observation)
        analysis_records = []
        if analysis_record_path.exists():
            analysis_records.append(analysis_record_path)
        fallback_analysis = self._fallback_bo_analysis_json(observation)
        if fallback_analysis is not None and fallback_analysis not in analysis_records:
            analysis_records.append(fallback_analysis)

        for analysis_record in analysis_records:
            try:
                with open(analysis_record, "r", encoding="utf-8") as fh:
                    payload = json.load(fh)
                results_csv = payload.get("results_csv")
                if results_csv:
                    results_path = self._resolve_observation_file_path(results_csv, observation)
                    if not results_path.exists() and not Path(str(results_csv)).is_absolute():
                        results_path = Path(analysis_record).parent / str(results_csv)
                    if results_path not in paths:
                        paths.append(results_path)
            except Exception:
                pass

        fallback_results = self._fallback_bo_analysis_results_csv(observation)
        if fallback_results is not None and fallback_results not in paths:
            paths.append(fallback_results)
        return paths

    def _selected_score_channels(self):
        if not hasattr(self, "_score_tree"):
            return set()
        selected = set()
        for item in self._score_tree.selection():
            channel = str(self._score_tree.item(item, "text") or item).strip()
            if channel:
                selected.add(channel)
        return selected

    def _raw_trace_rows_for_observation(self, observation):
        rows = []
        seen = set()

        def add(path, channel=None, scan=None):
            if not path:
                return
            p = self._resolve_observation_file_path(path, observation)
            key = str(p)
            if key in seen or not p.exists() or not p.is_file():
                return
            seen.add(key)
            rows.append(
                {
                    "path": str(p),
                    "channel": channel if channel not in (None, "") else self._infer_channel_from_path(p),
                    "scan": scan,
                }
            )

        for path in observation.get("archived_measurements") or []:
            add(path)

        analysis_record_path = self._resolve_observation_file_path(observation.get("analysis_record"), observation)
        analysis_records = []
        if analysis_record_path.exists():
            analysis_records.append(analysis_record_path)
        fallback_analysis = self._fallback_bo_analysis_json(observation)
        if fallback_analysis is not None and fallback_analysis not in analysis_records:
            analysis_records.append(fallback_analysis)

        for analysis_record in analysis_records:
            results_paths = []
            payload = {}
            try:
                with open(analysis_record, "r", encoding="utf-8") as fh:
                    payload = json.load(fh)
                results_csv = payload.get("results_csv")
                if results_csv:
                    results_path = self._resolve_observation_file_path(results_csv, observation)
                    if not results_path.exists() and not Path(str(results_csv)).is_absolute():
                        results_path = Path(analysis_record).parent / str(results_csv)
                    results_paths.append(results_path)
            except Exception:
                pass
            fallback_results = self._fallback_bo_analysis_results_csv(observation)
            if fallback_results is not None and fallback_results not in results_paths:
                results_paths.append(fallback_results)

            for results_path in results_paths:
                if not results_path.exists():
                    continue
                try:
                    with open(results_path, "r", encoding="utf-8-sig", newline="") as fh:
                        for row in csv.DictReader(fh):
                            add(row.get("file_path") or row.get("file_name"), channel=row.get("channel"), scan=row.get("scan_number"))
                except Exception:
                    pass

        legacy_dir = self._fallback_legacy_iteration_dir(observation)
        if legacy_dir is not None:
            for path in sorted(legacy_dir.glob("*.csv")):
                add(path)

        return sorted(rows, key=lambda row: (self._channel_sort_key(row.get("channel")), str(row.get("scan") or ""), Path(row["path"]).name))

    def _resolve_observation_file_path(self, raw_path, observation=None):
        if not raw_path:
            return Path("")
        text = str(raw_path).strip()
        direct = Path(text)
        if direct.exists():
            return direct

        normalized = text.replace("\\", "/")
        as_posix = Path(normalized)
        if as_posix.exists():
            return as_posix

        experiment_dir = getattr(self._bo_session, "experiment_dir", None) if self._bo_session is not None else None
        if experiment_dir is not None:
            experiment_dir = Path(experiment_dir)
            exp_name = experiment_dir.name
            marker = f"/{exp_name}/"
            if marker in normalized:
                suffix = normalized.split(marker, 1)[1]
                candidate = experiment_dir / suffix
                if candidate.exists():
                    return candidate
            if normalized.endswith(f"/{exp_name}"):
                return experiment_dir

            basename_candidate = experiment_dir / Path(normalized).name
            if basename_candidate.exists():
                return basename_candidate

            iteration = int((observation or {}).get("iteration", 0) or 0)
            if iteration:
                legacy_candidate = experiment_dir / "legacy" / f"iter_{iteration:03d}" / Path(normalized).name
                if legacy_candidate.exists():
                    return legacy_candidate

        return direct

    def _fallback_legacy_iteration_dir(self, observation):
        if self._bo_session is None or observation is None:
            return None
        iteration = int((observation or {}).get("iteration", 0) or 0)
        if not iteration:
            return None
        path = Path(self._bo_session.experiment_dir) / "legacy" / f"iter_{iteration:03d}"
        return path if path.is_dir() else None

    def _fallback_bo_analysis_json(self, observation):
        if self._bo_session is None or observation is None:
            return None
        iteration = int((observation or {}).get("iteration", 0) or 0)
        if not iteration:
            return None
        path = Path(self._bo_session.experiment_dir) / "bo_analysis" / f"bo_iter_{iteration:03d}.json"
        return path if path.is_file() else None

    def _fallback_bo_analysis_results_csv(self, observation):
        if self._bo_session is None or observation is None:
            return None
        iteration = int((observation or {}).get("iteration", 0) or 0)
        if not iteration:
            return None
        path = Path(self._bo_session.experiment_dir) / "bo_analysis" / f"bo_iter_{iteration:03d}_results.csv"
        return path if path.is_file() else None

    @staticmethod
    def _infer_channel_from_path(path):
        match = re.search(r"(?:^|[_\-\s])ch(?:annel)?\s*0*(\d+)(?:\D|$)", Path(path).stem, re.IGNORECASE)
        return match.group(1) if match else ""

    @staticmethod
    def _channel_sort_key(channel):
        try:
            return int(channel)
        except (TypeError, ValueError):
            return 9999

    def _q_breakdown_lines(self, observation, config=None):
        quality = dict((observation or {}).get("quality") or {})
        source_config = config if config is not None else (self._bo_session.config if self._bo_session else self._config or {})
        scoring = dict(source_config.get("scoring") or {})
        mode = str(scoring.get("mode", "classic_bounded") or "classic_bounded").strip().lower()
        channel_weights = dict(scoring.get("channel_weights") or {})
        run_weights = dict(scoring.get("run_weights") or {})
        q_run = float(observation.get("Q_run", quality.get("Q_run", 0.0)) or 0.0)
        mean_q = float(quality.get("mean_Q_channel", 0.0) or 0.0)
        std_q = float(quality.get("std_Q_channel", 0.0) or 0.0)
        failed = float(quality.get("failed_channel_fraction", 0.0) or 0.0)
        low = float(quality.get("low_channel_fraction", 0.0) or 0.0)
        lambda_var = float(run_weights.get("lambda_variability", 0.20))
        lambda_failed = float(run_weights.get("lambda_failed", 0.40))
        lambda_low = float(run_weights.get("lambda_low", 0.20))
        threshold = float(run_weights.get("low_channel_threshold", 0.50))
        total = sum(
            float(channel_weights.get(key, default))
            for key, default in (
                ("snr", 0.35),
                ("peak_height", 0.0),
                ("peak_shape", 0.20),
                ("baseline", 0.20),
                ("replicate_consistency", 0.15),
                ("success", 0.10),
            )
        )
        lines = [
            "Q_run breakdown:",
            f"  mean channel Q: {mean_q:.4f}",
            f"  variability penalty: {lambda_var:g} x std {std_q:.4f} = {lambda_var * std_q:.4f}",
            f"  failed-channel penalty: {lambda_failed:g} x fraction {failed:.4f} = {lambda_failed * failed:.4f}",
            f"  low-channel penalty: {lambda_low:g} x fraction {low:.4f} = {lambda_low * low:.4f} (low < {threshold:g})",
            f"  final Q_run: {q_run:.4f}",
            "",
        ]
        if mode == "signal_priority_unbounded":
            lines.extend(
                [
                    "Q_channel terms:",
                    (
                        "  "
                        f"log1p(SNR) {float(channel_weights.get('snr', 0.45)):g}, "
                        f"log1p(Peak) {float(channel_weights.get('peak_height', 0.35)):g}, "
                        f"Baseline {float(channel_weights.get('baseline', 0.12)):g}, "
                        f"Shape {float(channel_weights.get('peak_shape', 0.05)):g}, "
                        f"Replicate {float(channel_weights.get('replicate_consistency', 0.03)):g}, "
                        f"Success {float(channel_weights.get('success', 0.0)):g}; total {total:g}"
                    ),
                ]
            )
        else:
            lines.extend(
                [
                    "Q_channel weights:",
                    (
                        "  "
                        f"SNR {float(channel_weights.get('snr', 0.35)):g}, "
                        f"Peak {float(channel_weights.get('peak_height', 0.0)):g}, "
                        f"Shape {float(channel_weights.get('peak_shape', 0.20)):g}, "
                        f"Baseline {float(channel_weights.get('baseline', 0.20)):g}, "
                        f"Replicate {float(channel_weights.get('replicate_consistency', 0.15)):g}, "
                        f"Success {float(channel_weights.get('success', 0.10)):g}; total {total:g}"
                    ),
                ]
            )
        return lines

    def _q_equation_lines(self, config=None):
        source_config = config if config is not None else (self._bo_session.config if self._bo_session else self._config or {})
        scoring = dict((source_config or {}).get("scoring") or {})
        mode = str(scoring.get("mode", "classic_bounded") or "classic_bounded").strip().lower()
        channel_weights = dict(scoring.get("channel_weights") or {})
        run_weights = dict(scoring.get("run_weights") or {})

        if mode == "signal_priority_unbounded":
            terms = [
                ("log1p(SNR)", float(channel_weights.get("snr", 0.45))),
                ("log1p(Peak uA)", float(channel_weights.get("peak_height", 0.35))),
                ("Baseline", float(channel_weights.get("baseline", 0.12))),
                ("Shape", float(channel_weights.get("peak_shape", 0.05))),
                ("Replicate", float(channel_weights.get("replicate_consistency", 0.03))),
                ("Success", float(channel_weights.get("success", 0.0))),
            ]
        else:
            terms = [
                ("SNR score", float(channel_weights.get("snr", 0.35))),
                ("Peak", float(channel_weights.get("peak_height", 0.0))),
                ("Shape", float(channel_weights.get("peak_shape", 0.20))),
                ("Baseline", float(channel_weights.get("baseline", 0.20))),
                ("Replicate", float(channel_weights.get("replicate_consistency", 0.15))),
                ("Success", float(channel_weights.get("success", 0.10))),
            ]
        total = sum(weight for _label, weight in terms)
        numerator = " + ".join(f"{weight:g}*{label}" for label, weight in terms if weight)
        if not numerator:
            numerator = "0"
        lambda_var = float(run_weights.get("lambda_variability", 0.20))
        lambda_failed = float(run_weights.get("lambda_failed", 0.40))
        lambda_low = float(run_weights.get("lambda_low", 0.20))
        threshold = float(run_weights.get("low_channel_threshold", 0.50))
        lines = [
            f"Q_channel = ({numerator}) / {max(total, 1e-12):g}",
            (
                "Q_run = mean(Q_channel) "
                f"- {lambda_var:g}*std(Q_channel) "
                f"- {lambda_failed:g}*failed_fraction "
                f"- {lambda_low:g}*fraction(Q_channel < {threshold:g})"
            ),
        ]
        if mode != "signal_priority_unbounded":
            lines.append(f"SNR score = clip(raw SNR / {float(channel_weights.get('snr_saturation', 20.0)):g}, 0, 1); Q_channel and Q_run are clipped 0..1.")
        else:
            lines.append("Signal-priority mode uses log signal terms and does not clip Q_run.")
        return lines

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
    def _analysis_trend_metric_options():
        return (
            "Q_run",
            "Mean channel Q",
            "Std channel Q",
            "Failed fraction",
            "Low fraction",
            "Mean peak uA",
            "Mean RMS uA",
            "Mean raw SNR",
            "Step potential",
            "Amplitude",
            "Frequency",
            "Begin potential",
            "End potential",
            "Conditioning potential",
            "Conditioning time",
        )

    def _analysis_trend_value(self, observation, metric):
        quality = dict((observation or {}).get("quality") or {})
        params = dict((observation or {}).get("params") or {})
        if metric == "Q_run":
            return (observation or {}).get("Q_run")
        if metric == "Mean channel Q":
            return quality.get("mean_Q_channel")
        if metric == "Std channel Q":
            return quality.get("std_Q_channel")
        if metric == "Failed fraction":
            return quality.get("failed_channel_fraction")
        if metric == "Low fraction":
            return quality.get("low_channel_fraction")
        if metric == "Mean peak uA":
            peak, _rms = self._observation_peak_rms(observation)
            return peak
        if metric == "Mean RMS uA":
            _peak, rms = self._observation_peak_rms(observation)
            return rms
        if metric == "Mean raw SNR":
            metrics = (observation or {}).get("channel_metrics", {})
            if not isinstance(metrics, dict):
                return None
            values = []
            for data in metrics.values():
                if isinstance(data, dict) and data.get("snr") is not None:
                    values.append(float(data.get("snr")))
            return sum(values) / len(values) if values else None
        param_key_by_metric = {
            "Step potential": "step_potential",
            "Amplitude": "amplitude",
            "Frequency": "frequency",
            "Begin potential": "begin_potential",
            "End potential": "end_potential",
            "Conditioning potential": "conditioning_potential",
            "Conditioning time": "conditioning_time",
        }
        key = param_key_by_metric.get(metric)
        return params.get(key) if key else None

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
    def _channel_peak_height(metrics):
        if not isinstance(metrics, dict):
            return None
        for key in ("mean_peak_current_uA", "median_peak_current_uA", "peak_current"):
            value = metrics.get(key)
            if value is not None:
                return value
        return None

    @staticmethod
    def _channel_background_rms(metrics):
        if not isinstance(metrics, dict):
            return None
        for key in ("mean_background_rms_uA", "median_background_rms_uA", "background_current_rms", "baseline_noise"):
            value = metrics.get(key)
            if value is not None:
                return value
        return None

    @classmethod
    def _observation_peak_rms(cls, observation):
        metrics = (observation or {}).get("channel_metrics", {})
        if not isinstance(metrics, dict) or not metrics:
            return None, None
        peaks = []
        rms_values = []
        for data in metrics.values():
            peak = cls._channel_peak_height(data)
            rms = cls._channel_background_rms(data)
            if peak is not None:
                peaks.append(float(peak))
            if rms is not None:
                rms_values.append(float(rms))
        peak_avg = sum(peaks) / len(peaks) if peaks else None
        rms_avg = sum(rms_values) / len(rms_values) if rms_values else None
        return peak_avg, rms_avg

    @staticmethod
    def _write_text(widget, text):
        widget.config(state="normal")
        widget.delete("1.0", tk.END)
        widget.insert("1.0", text)
        widget.config(state="disabled")

    @staticmethod
    def _clear_text(widget):
        BayesianOptimizationTab._write_text(widget, "")
