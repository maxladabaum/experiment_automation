"""
gui/tab_bayesian_optimization.py - Optional SWV Bayesian optimization tab.

This is a UI shell around core.bo_session. It edits configuration, requests BO
suggestions, queues normal SWV methods, runs/imports analysis outputs, and
shows records. BO math lives in core.bo_session.
"""

from __future__ import annotations

import ast
import copy
import csv
import io
import json
import math
import pickle
import re
import threading
import time
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog
import tkinter as tk
from tkinter import ttk, scrolledtext

from config import (
    BO_ANALYSIS_FILE_GLOB,
    BO_ANALYSIS_OUTPUT_DIR,
    BO_DEFAULT_CONFIG_PATH,
    BO_EXTERNAL_ANALYSIS_PROJECT,
    BO_EXTERNAL_ANALYSIS_PYTHON,
    BO_LAST_SETUP_METADATA_PATH,
    BO_LOCAL_PATHS_CONFIG,
)
from core.bo_session import (
    BOIntegrationSession,
    DEFAULT_PARAMETER_RANGES,
    OPTIMIZER_ORDER,
    PARAMETER_ORDER,
    _acquisition_score,
    active_parameters,
    build_swv_script,
    compute_channel_quality,
    compute_paired_response_quality,
    compute_run_quality,
    encode_candidate,
    load_bo_config,
    load_bo_setup_metadata,
    normalize_bo_config,
    parse_channels,
    channel_groups,
    resolve_initial_parameters,
    save_bo_config,
    save_bo_setup_metadata,
    validate_bo_config,
)
from core.bo_analysis import _build_channel_metrics
from core.bo_simulation import LANDSCAPE_TYPES, default_dimensions, run_optimizer_simulation, run_paired_response_optimizer_simulation
from core.analysis import analyze_swv_arrays, partial_traces_for_failure_arrays
from core.analysis_io import load_swv_csv
from core.swv_method import EMSTAT_PICO_HIGH_SPEED_BA_RANGES, normalize_swv_ba_range_options, range_labels
from gui.widgets import FlowFrame, ScrollableFrame


class BayesianOptimizationTab:
    """Optional closed-loop Bayesian optimization UI."""

    ACCENT = "#155e63"
    ACCENT_DARK = "#0f3d44"
    ACCENT_LIGHT = "#dff7f5"

    LAST_SETUP_UI_VARS = {
        "config_path": "_config_path_var",
        "analysis_output_dir": "_analysis_dir_var",
        "analysis_project": "_analysis_project_var",
        "analysis_python": "_analysis_python_var",
        "analysis_file_glob": "_analysis_glob_var",
        "target_iterations": "_auto_target_var",
        "paired_target_exchange_block": "_paired_target_exchange_var",
        "paired_buffer_exchange_block": "_paired_buffer_exchange_var",
        "paired_target_equilibration_seconds": "_paired_target_equilibration_var",
        "paired_buffer_equilibration_seconds": "_paired_buffer_equilibration_var",
    }

    def __init__(
        self,
        parent_frame,
        session,
        on_add_to_queue,
        on_refresh_queue,
        on_script_preview,
        on_run_queue=None,
        on_run_queue_from_index=None,
        on_configure_auto_titration=None,
        is_auto_titration_locked=None,
        on_bo_finished=None,
    ):
        self._frame = parent_frame
        self._session = session
        self._add_to_queue = on_add_to_queue
        self._refresh_queue = on_refresh_queue
        self._script_preview = on_script_preview
        self._run_queue = on_run_queue
        self._run_queue_from_index = on_run_queue_from_index
        self._configure_auto_titration = on_configure_auto_titration
        self._is_auto_titration_locked = is_auto_titration_locked
        self._on_bo_finished = on_bo_finished

        self._config_path_var = tk.StringVar(value=str(BO_DEFAULT_CONFIG_PATH))
        self._analysis_dir_var = tk.StringVar(value=str(BO_ANALYSIS_OUTPUT_DIR))
        self._analysis_project_var = tk.StringVar(value=str(BO_EXTERNAL_ANALYSIS_PROJECT))
        self._analysis_python_var = tk.StringVar(value=str(BO_EXTERNAL_ANALYSIS_PYTHON))
        self._analysis_glob_var = tk.StringVar(value=str(BO_ANALYSIS_FILE_GLOB))
        self._analysis_crop_min_var = tk.StringVar(value="-0.61")
        self._analysis_crop_max_var = tk.StringVar(value="-0.30")
        self._analysis_smooth_window_var = tk.StringVar(value="15")
        self._analysis_smooth_polyorder_var = tk.StringVar(value="2")
        self._analysis_minima_window_var = tk.StringVar(value="0.30")
        self._analysis_min_peak_height_var = tk.StringVar(value="0.001")
        self._analysis_peak_voltage_min_var = tk.StringVar(value="")
        self._analysis_peak_voltage_max_var = tk.StringVar(value="")
        self._analysis_left_min_voltage_min_var = tk.StringVar(value="")
        self._analysis_left_min_voltage_max_var = tk.StringVar(value="")
        self._analysis_right_min_voltage_min_var = tk.StringVar(value="")
        self._analysis_right_min_voltage_max_var = tk.StringVar(value="")
        self._analysis_min_start_voltage_var = tk.StringVar(value="-0.70")
        self._analysis_scan_windows_var = tk.StringVar(value="")
        self._analysis_use_prominent_var = tk.BooleanVar(value=False)
        self._analysis_require_minima_var = tk.BooleanVar(value=False)
        self._analysis_double_correction_var = tk.BooleanVar(value=False)
        self._analysis_compute_skew_var = tk.BooleanVar(value=True)
        self._analysis_compute_wavelet_energy_var = tk.BooleanVar(value=True)
        self._analysis_wavelet_trace_var = tk.BooleanVar(value=False)
        self._analysis_wavelet_correction_var = tk.BooleanVar(value=False)
        self._channels_var = tk.StringVar(value="")
        self._channel_group_count_var = tk.StringVar(value="1")
        self._channel_group_vars = []
        self._channel_group_settings = []
        self._bo_bandwidth_var = tk.StringVar(value="4k")
        self._bo_ba_range_mode_var = tk.StringVar(value="fixed")
        self._bo_ba_fixed_range_var = tk.StringVar(value="100 nA")
        self._bo_ba_auto_min_var = tk.StringVar(value="100 nA")
        self._bo_ba_auto_max_var = tk.StringVar(value="100 nA")
        self._measurements_per_channel_var = tk.StringVar(value="1")
        self._exploration_var = tk.DoubleVar(value=0.35)
        self._exploration_text_var = tk.StringVar(value="0.35")
        self._gp_warmup_iterations_var = tk.StringVar(value="8")
        self._candidate_pool_var = tk.StringVar(value="600")
        self._local_pool_var = tk.StringVar(value="120")
        self._initial_point_mode_var = tk.StringVar(value="specific")
        self._optimization_direction_var = tk.StringVar(value="maximize")
        self._gp_length_scale_vars = {name: tk.StringVar(value="0.2") for name in PARAMETER_ORDER}
        self._gp_falloff_summary_var = tk.StringVar(value="GP falloff: fixed fractions of search range (0.2 each)")
        self._score_mode_var = tk.StringVar(value="classic")
        self._score_snr_weight_var = tk.StringVar(value="0.35")
        self._score_repeat_scan_snr_weight_var = tk.StringVar(value="0.00")
        self._score_peak_height_weight_var = tk.StringVar(value="0.00")
        self._score_shape_weight_var = tk.StringVar(value="0.20")
        self._score_baseline_weight_var = tk.StringVar(value="0.20")
        self._score_replicate_weight_var = tk.StringVar(value="0.15")
        self._score_success_weight_var = tk.StringVar(value="0.10")
        self._score_noise_penalty_var = tk.StringVar(value="0.00")
        self._score_snr_saturation_var = tk.StringVar(value="20.0")
        self._score_variability_penalty_var = tk.StringVar(value="0.20")
        self._score_repeat_std_penalty_var = tk.StringVar(value="0.00")
        self._score_failed_penalty_var = tk.StringVar(value="0.40")
        self._score_low_penalty_var = tk.StringVar(value="0.20")
        self._score_low_threshold_var = tk.StringVar(value="0.50")
        self._score_formula_var = tk.StringVar(value="")
        self._rescore_mode_var = tk.StringVar(value="classic")
        self._rescore_snr_weight_var = tk.StringVar(value="0.35")
        self._rescore_repeat_scan_snr_weight_var = tk.StringVar(value="0.00")
        self._rescore_peak_height_weight_var = tk.StringVar(value="0.00")
        self._rescore_shape_weight_var = tk.StringVar(value="0.20")
        self._rescore_baseline_weight_var = tk.StringVar(value="0.20")
        self._rescore_replicate_weight_var = tk.StringVar(value="0.15")
        self._rescore_success_weight_var = tk.StringVar(value="0.10")
        self._rescore_noise_penalty_var = tk.StringVar(value="0.00")
        self._rescore_snr_saturation_var = tk.StringVar(value="20.0")
        self._rescore_variability_penalty_var = tk.StringVar(value="0.20")
        self._rescore_repeat_std_penalty_var = tk.StringVar(value="0.00")
        self._rescore_failed_penalty_var = tk.StringVar(value="0.40")
        self._rescore_low_penalty_var = tk.StringVar(value="0.20")
        self._rescore_low_threshold_var = tk.StringVar(value="0.50")
        self._rescore_formula_var = tk.StringVar(value="")
        self._rescore_paired_score_vars = {
            key: tk.StringVar(value=value)
            for key, value in {
                "buffer_classic_Q": "0.25",
                "target_classic_Q": "0.25",
                "peak_prominence": "1.0",
                "repeat_scan_snr": "0.00",
                "lambda_repeat_std": "0.00",
            }.items()
        }
        self._rescore_paired_formula_var = tk.StringVar(value="")
        self._rescore_status_var = tk.StringVar(value="Load a BO session to rescore recorded data.")
        self._rescore_analysis_vars = {
            "crop_min_v": tk.StringVar(value="-0.61"),
            "crop_max_v": tk.StringVar(value="-0.30"),
            "smooth_window": tk.StringVar(value="15"),
            "smooth_polyorder": tk.StringVar(value="2"),
            "minima_search_window_v": tk.StringVar(value="0.30"),
            "min_peak_height_ua": tk.StringVar(value="0.001"),
            "peak_voltage_min_v": tk.StringVar(value=""),
            "peak_voltage_max_v": tk.StringVar(value=""),
            "left_min_voltage_min_v": tk.StringVar(value=""),
            "left_min_voltage_max_v": tk.StringVar(value=""),
            "right_min_voltage_min_v": tk.StringVar(value=""),
            "right_min_voltage_max_v": tk.StringVar(value=""),
            "min_start_voltage_v": tk.StringVar(value="-0.70"),
            "scan_windows": tk.StringVar(value=""),
            "use_prominent_minima": tk.BooleanVar(value=False),
            "require_local_minima_on_both_sides": tk.BooleanVar(value=False),
            "use_double_correction": tk.BooleanVar(value=False),
            "compute_skew": tk.BooleanVar(value=True),
            "compute_wavelet_energy": tk.BooleanVar(value=True),
            "compute_wavelet_denoised_trace": tk.BooleanVar(value=False),
            "use_wavelet_for_correction": tk.BooleanVar(value=False),
        }
        self._status_var = tk.StringVar(value="Load a BO config to begin.")
        self._record_dir_var = tk.StringVar(value="Record folder: (not started)")
        self._auto_target_var = tk.StringVar(value="5")
        self._run_auto_titration_var = tk.BooleanVar(value=False)
        self._post_bo_titration_started = False
        self._bo_objective_var = tk.StringVar(value="quality")
        self._paired_batch_size_var = tk.StringVar(value="4")
        self._paired_warmup_batch_size_var = tk.StringVar(value="4")
        self._paired_warmup_single_batch_var = tk.BooleanVar(value=False)
        self._paired_target_exchange_var = tk.StringVar(value="")
        self._paired_buffer_exchange_var = tk.StringVar(value="")
        self._paired_target_equilibration_var = tk.StringVar(value="0")
        self._paired_buffer_equilibration_var = tk.StringVar(value="0")
        self._paired_buffer_classic_q_weight_var = tk.StringVar(value="0.25")
        self._paired_target_classic_q_weight_var = tk.StringVar(value="0.25")
        self._paired_delta_peak_weight_var = tk.StringVar(value="1.0")
        self._paired_repeat_scan_snr_weight_var = tk.StringVar(value="0.00")
        self._paired_repeat_std_penalty_var = tk.StringVar(value="0.00")
        self._paired_formula_var = tk.StringVar(value="")
        self._analysis_trend_metric_var = tk.StringVar(value="Q_run")
        self._surrogate_iteration_var = tk.StringVar(value="")
        self._surrogate_value_var = tk.StringVar(value="predicted_mean_Q")
        self._surrogate_view_var = tk.StringVar(value="1D slice")
        self._surrogate_x_var = tk.StringVar(value="")
        self._surrogate_y_var = tk.StringVar(value="")
        self._surrogate_z_var = tk.StringVar(value="")
        self._surrogate_color_min_var = tk.StringVar(value="")
        self._surrogate_color_max_var = tk.StringVar(value="")
        self._engine_iterations_var = tk.StringVar(value="20")
        self._engine_channel_group_count_var = tk.StringVar(value="1")
        self._engine_channel_group_vars = []
        self._engine_channel_group_settings = []
        self._engine_paired_response_var = tk.BooleanVar(value=False)
        self._engine_paired_batch_size_var = tk.StringVar(value="4")
        self._engine_exploration_var = tk.DoubleVar(value=0.35)
        self._engine_exploration_text_var = tk.StringVar(value="0.35")
        self._engine_warmup_iterations_var = tk.StringVar(value="8")
        self._engine_candidate_pool_var = tk.StringVar(value="600")
        self._engine_local_pool_var = tk.StringVar(value="120")
        self._engine_initial_point_mode_var = tk.StringVar(value="specific")
        self._engine_optimization_direction_var = tk.StringVar(value="maximize")
        self._engine_gp_length_scale_vars = {
            name: tk.StringVar(value="0.2") for name in PARAMETER_ORDER
        }
        self._engine_score_vars = {
            key: tk.StringVar(value=value)
            for key, value in {
                "mode": "classic",
                "peak_prominence": "0.35",
                "repeat_scan_snr": "0.00",
                "peak_height": "0.00",
                "peak_shape": "0.20",
                "baseline": "0.20",
                "replicate_consistency": "0.15",
                "success": "0.10",
                "noise_penalty": "0.00",
                "peak_prominence_saturation": "20.0",
                "lambda_variability": "0.20",
                "lambda_repeat_std": "0.00",
                "lambda_failed": "0.40",
                "lambda_low": "0.20",
                "low_channel_threshold": "0.50",
            }.items()
        }
        self._engine_score_formula_var = tk.StringVar(value="")
        self._engine_paired_score_vars = {
            key: tk.StringVar(value=value)
            for key, value in {
                "buffer_classic_Q": "0.25",
                "target_classic_Q": "0.25",
                "peak_prominence": "1.0",
                "repeat_scan_snr": "0.00",
                "lambda_repeat_std": "0.00",
            }.items()
        }
        self._engine_paired_formula_var = tk.StringVar(value="")
        self._engine_analysis_vars = {
            "crop_min_v": tk.StringVar(value="-0.61"),
            "crop_max_v": tk.StringVar(value="-0.30"),
            "smooth_window": tk.StringVar(value="15"),
            "smooth_polyorder": tk.StringVar(value="2"),
            "minima_search_window_v": tk.StringVar(value="0.30"),
            "min_peak_height_ua": tk.StringVar(value="0.001"),
            "peak_voltage_min_v": tk.StringVar(value=""),
            "peak_voltage_max_v": tk.StringVar(value=""),
            "left_min_voltage_min_v": tk.StringVar(value=""),
            "left_min_voltage_max_v": tk.StringVar(value=""),
            "right_min_voltage_min_v": tk.StringVar(value=""),
            "right_min_voltage_max_v": tk.StringVar(value=""),
            "min_start_voltage_v": tk.StringVar(value="-0.70"),
            "scan_windows": tk.StringVar(value=""),
            "use_prominent_minima": tk.BooleanVar(value=False),
            "require_local_minima_on_both_sides": tk.BooleanVar(value=False),
            "use_double_correction": tk.BooleanVar(value=False),
            "compute_skew": tk.BooleanVar(value=True),
            "compute_wavelet_energy": tk.BooleanVar(value=True),
            "compute_wavelet_denoised_trace": tk.BooleanVar(value=False),
            "use_wavelet_for_correction": tk.BooleanVar(value=False),
        }
        self._engine_target_response_gain_var = tk.StringVar(value="2.0")
        self._engine_target_noise_multiplier_var = tk.StringVar(value="1.05")
        self._engine_delta_peak_floor_var = tk.StringVar(value="0.0")
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
        self._engine_progress_var = tk.DoubleVar(value=0.0)
        self._engine_progress_text_var = tk.StringVar(value="")
        self._auto_status_var = tk.StringVar(value="Auto loop idle.")
        self._style = ttk.Style(self._frame)

        self._config = None
        self._bo_session = None
        self._suggestion = None
        self._auto_running = False
        self._paired_queue_running = False
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
        self._loaded_original_config = None
        self._surrogate_plot_canvas = None
        self._selected_history_observation = None
        self._active_results_tree = None
        self._results_trace_panes_balanced = False
        self._results_render_deferred = False
        self._results_render_flush_job = None

        self._build()
        if not self._load_last_bo_setup():
            self._load_config(initial=True)
        setattr(self._session, "_bo_live_refresh_callback", self._on_live_paired_bo_update)

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

    def _visible_paned_window(self, parent, orient):
        return tk.PanedWindow(
            parent,
            orient=orient,
            sashwidth=8,
            sashpad=2,
            showhandle=True,
            handlesize=10,
            handlepad=8,
            opaqueresize=True,
            bd=0,
            relief=tk.FLAT,
            bg="#c4cbd3",
        )

    @staticmethod
    def _set_paned_sash_position(pane, index, position):
        try:
            if hasattr(pane, "sashpos"):
                pane.sashpos(index, position)
            elif str(pane.cget("orient")) == str(tk.HORIZONTAL):
                pane.sash_place(index, int(position), 0)
            else:
                pane.sash_place(index, 0, int(position))
        except Exception:
            pass

    @staticmethod
    def _fit_embedded_figure(fig, top=0.86, bottom=0.20, left=0.16, right=0.95):
        try:
            fig.set_layout_engine(None)
            fig.subplots_adjust(top=top, bottom=bottom, left=left, right=right)
        except Exception:
            pass

    def _balance_results_trace_panes(self):
        if getattr(self, "_results_trace_panes_balanced", False):
            return
        panes = (
            getattr(self, "_results_top_pane", None),
            getattr(self, "_results_middle_pane", None),
            getattr(self, "_results_bottom_pane", None),
        )
        if any(pane is None for pane in panes):
            return
        widths = []
        for pane in (
            getattr(self, "_results_top_pane", None),
            getattr(self, "_results_middle_pane", None),
            getattr(self, "_results_bottom_pane", None),
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
            self._set_paned_sash_position(pane, 0, width // 2)
        main_pane = getattr(self, "_results_main_pane", None)
        if main_pane is not None:
            try:
                main_pane.update_idletasks()
                height = main_pane.winfo_height()
                if height > 60:
                    self._set_paned_sash_position(main_pane, 0, height // 3)
                    self._set_paned_sash_position(main_pane, 1, 2 * height // 3)
            except Exception:
                pass
        self._results_trace_panes_balanced = True

    def _build_setup_tab(self, parent):
        scroller = ScrollableFrame(parent, min_width=1080)
        scroller.pack(fill="both", expand=True)
        parent = scroller.content
        pane = ttk.PanedWindow(parent, orient=tk.HORIZONTAL)
        pane.pack(fill="both", expand=True, padx=4, pady=4)
        left = ttk.Frame(pane)
        right = ttk.Frame(pane)
        pane.add(left, weight=1)
        pane.add(right, weight=1)
        split_initialized = {"done": False}
        def initialize_equal_split():
            try:
                pane.update_idletasks()
                width = pane.winfo_width()
                if width > 100:
                    self._set_paned_sash_position(pane, 0, width // 2)
                    split_initialized["done"] = True
            except Exception:
                pass
        def initialize_when_visible(_event=None):
            if not split_initialized["done"]:
                pane.after_idle(initialize_equal_split)
        pane.after_idle(initialize_equal_split)
        pane.after(150, initialize_equal_split)
        pane.bind("<Map>", initialize_when_visible, add="+")

        right_pane = ttk.PanedWindow(right, orient=tk.VERTICAL)
        right_pane.pack(fill="both", expand=True)
        params_host = ttk.Frame(right_pane)
        controls_host = ttk.Frame(right_pane)
        right_pane.add(params_host, weight=1)
        right_pane.add(controls_host, weight=5)
        right_pane.after_idle(lambda: self._set_paned_sash_position(right_pane, 0, 210))

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

        ttk.Label(cfg, text="64-bit analysis Python:").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Entry(cfg, textvariable=self._analysis_python_var).grid(row=2, column=1, sticky="ew", padx=4)
        ttk.Button(cfg, text="Browse", command=self._browse_analysis_python).grid(row=2, column=2, padx=2)

        ttk.Label(cfg, text="Application project:").grid(row=3, column=0, sticky="w", pady=2)
        ttk.Entry(cfg, textvariable=self._analysis_project_var).grid(row=3, column=1, sticky="ew", padx=4)
        ttk.Button(cfg, text="Browse", command=self._browse_analysis_project).grid(row=3, column=2, padx=2)

        ttk.Label(cfg, text="Analysis glob:").grid(row=4, column=0, sticky="w", pady=2)
        ttk.Entry(cfg, textvariable=self._analysis_glob_var, width=14).grid(row=4, column=1, sticky="w", padx=4)

        ttk.Label(cfg, text="Channel groups:").grid(row=5, column=0, sticky="nw", pady=2)
        group_controls = ttk.Frame(cfg)
        group_controls.grid(row=5, column=1, sticky="ew", padx=4)
        ttk.Label(group_controls, text="Number of groups").pack(side="left")
        group_count = ttk.Combobox(
            group_controls,
            textvariable=self._channel_group_count_var,
            values=[str(value) for value in range(1, 11)],
            state="readonly",
            width=4,
        )
        group_count.pack(side="left", padx=6)
        group_count.bind("<<ComboboxSelected>>", lambda _e: self._rebuild_channel_group_entries())
        self._channel_groups_frame = ttk.Frame(cfg)
        self._channel_groups_frame.grid(row=6, column=1, columnspan=4, sticky="ew", padx=4, pady=(0, 4))
        ttk.Button(cfg, text="Validate", command=self._validate_config).grid(row=5, column=2, padx=2)
        ttk.Button(cfg, text="Load BO Session", command=self._load_bo_session).grid(row=5, column=3, padx=2)
        ttk.Button(cfg, text="Start BO Session", command=self._start_bo_session).grid(row=5, column=4, padx=2)
        ttk.Checkbutton(
            cfg,
            text="Run autotitration when BO finishes",
            variable=self._run_auto_titration_var,
            command=self._toggle_auto_titration,
        ).grid(row=7, column=1, columnspan=4, sticky="w", padx=4, pady=(4, 0))
        self._rebuild_channel_group_entries()

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

        type_box = ttk.LabelFrame(left, text="BO Type", padding=8)
        type_box.pack(fill="x", pady=(0, 8))
        ttk.Radiobutton(
            type_box,
            text="Classic BO optimization",
            variable=self._bo_objective_var,
            value="quality",
            command=self._on_bo_type_changed,
        ).grid(row=0, column=0, sticky="w", pady=2)
        ttk.Radiobutton(
            type_box,
            text="Paired-response batched BO optimization",
            variable=self._bo_objective_var,
            value="paired_response",
            command=self._on_bo_type_changed,
        ).grid(row=1, column=0, sticky="w", pady=2)
        ttk.Label(
            type_box,
            text=(
                "Paired mode scores the target response as the target-minus-buffer peak change "
                "divided by the sum of target and buffer channel noise."
            ),
            foreground=self.ACCENT,
            wraplength=460,
            justify="left",
        ).grid(row=2, column=0, sticky="w", pady=(4, 0))

        method_box = ttk.LabelFrame(left, text="Method Settings", padding=8)
        method_box.pack(fill="x", pady=(0, 8))
        method_box.columnconfigure(1, weight=1)
        ttk.Label(method_box, text="Bandwidth:").grid(row=0, column=0, sticky="w", pady=2)
        bandwidth_combo = ttk.Combobox(
            method_box,
            textvariable=self._bo_bandwidth_var,
            values=("4k", "8k"),
            state="readonly",
            width=10,
        )
        bandwidth_combo.grid(row=0, column=1, sticky="w", padx=4, pady=2)
        bandwidth_combo.bind("<<ComboboxSelected>>", lambda _e: self._sync_method_options_config(show_error=False))

        ttk.Label(method_box, text="BA range:").grid(row=1, column=0, sticky="w", pady=(6, 2))
        range_mode_frame = ttk.Frame(method_box)
        range_mode_frame.grid(row=1, column=1, sticky="w", padx=4, pady=(6, 2))
        ttk.Radiobutton(
            range_mode_frame,
            text="Fixed",
            variable=self._bo_ba_range_mode_var,
            value="fixed",
            command=self._sync_bo_ba_range_controls,
        ).pack(side="left")
        ttk.Radiobutton(
            range_mode_frame,
            text="Autorange",
            variable=self._bo_ba_range_mode_var,
            value="auto",
            command=self._sync_bo_ba_range_controls,
        ).pack(side="left", padx=(8, 0))

        ba_labels = range_labels(EMSTAT_PICO_HIGH_SPEED_BA_RANGES)
        ttk.Label(method_box, text="Fixed range:").grid(row=2, column=0, sticky="w", pady=2)
        self._bo_ba_fixed_combo = ttk.Combobox(
            method_box,
            textvariable=self._bo_ba_fixed_range_var,
            values=ba_labels,
            state="readonly",
            width=14,
        )
        self._bo_ba_fixed_combo.grid(row=2, column=1, sticky="w", padx=4, pady=2)
        self._bo_ba_fixed_combo.bind("<<ComboboxSelected>>", lambda _e: self._sync_method_options_config(show_error=False))

        ttk.Label(method_box, text="Autorange min:").grid(row=3, column=0, sticky="w", pady=2)
        self._bo_ba_auto_min_combo = ttk.Combobox(
            method_box,
            textvariable=self._bo_ba_auto_min_var,
            values=ba_labels,
            state="readonly",
            width=14,
        )
        self._bo_ba_auto_min_combo.grid(row=3, column=1, sticky="w", padx=4, pady=2)
        self._bo_ba_auto_min_combo.bind("<<ComboboxSelected>>", lambda _e: self._sync_method_options_config(show_error=False))

        ttk.Label(method_box, text="Autorange max:").grid(row=4, column=0, sticky="w", pady=2)
        self._bo_ba_auto_max_combo = ttk.Combobox(
            method_box,
            textvariable=self._bo_ba_auto_max_var,
            values=ba_labels,
            state="readonly",
            width=14,
        )
        self._bo_ba_auto_max_combo.grid(row=4, column=1, sticky="w", padx=4, pady=2)
        self._bo_ba_auto_max_combo.bind("<<ComboboxSelected>>", lambda _e: self._sync_method_options_config(show_error=False))
        ttk.Label(method_box, text="Measurements per channel / point:").grid(row=5, column=0, sticky="w", pady=2)
        repeat_entry = ttk.Entry(
            method_box, textvariable=self._measurements_per_channel_var, width=8
        )
        repeat_entry.grid(row=5, column=1, sticky="w", padx=4, pady=2)
        repeat_entry.bind(
            "<FocusOut>", lambda _e: self._sync_method_options_config(show_error=False)
        )
        repeat_entry.bind(
            "<Return>", lambda _e: self._sync_method_options_config(show_error=False)
        )
        ttk.Label(
            method_box,
            text="These SWV settings are saved into BO config method_options and used for queued BO methods.",
            foreground=self.ACCENT,
            wraplength=460,
            justify="left",
        ).grid(row=6, column=0, columnspan=2, sticky="w", pady=(4, 0))
        self._sync_bo_ba_range_controls()

        algo_box = ttk.LabelFrame(left, text="Optimizer Behavior", padding=8)
        self._optimizer_behavior_frame = algo_box
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
        ttk.Label(algo_box, text="GP warmup iters:").grid(row=2, column=0, sticky="w", pady=2)
        warmup_entry = ttk.Entry(algo_box, textvariable=self._gp_warmup_iterations_var, width=8)
        warmup_entry.grid(row=2, column=1, sticky="w", padx=6, pady=2)
        warmup_entry.bind("<FocusOut>", lambda _e: self._sync_algorithm_config(show_error=False))
        warmup_entry.bind("<Return>", lambda _e: self._sync_algorithm_config(show_error=False))
        ttk.Label(algo_box, text="completed BO iterations before GP starts", foreground=self.ACCENT).grid(
            row=2, column=2, sticky="e", padx=(6, 0), pady=2
        )
        ttk.Label(algo_box, text="Start point:").grid(row=3, column=0, sticky="w", pady=2)
        start_mode = ttk.Combobox(
            algo_box,
            textvariable=self._initial_point_mode_var,
            values=("specific", "random"),
            state="readonly",
            width=12,
        )
        start_mode.grid(row=3, column=1, sticky="w", padx=6, pady=2)
        start_mode.bind("<<ComboboxSelected>>", lambda _e: self._sync_algorithm_config(show_error=False))
        ttk.Label(
            algo_box,
            text="`specific` uses Initial Parameters. `random` chooses one valid candidate as the first BO point. In paired mode, warmup cycles are consolidated into one buffer block and one target block.",
            foreground=self.ACCENT,
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(2, 0))
        for child in algo_box.winfo_children():
            child.destroy()
        algo_box.configure(text="Optimizer Behavior by Group")
        self._bo_group_optimizer_panels_frame = ttk.Frame(algo_box)
        self._bo_group_optimizer_panels_frame.pack(fill="x")
        self._rebuild_bo_group_optimizer_panels()

        paired_box = ttk.LabelFrame(left, text="Paired BO Fluid Exchange", padding=8)
        self._paired_behavior_frame = paired_box
        paired_box.columnconfigure(1, weight=1)
        ttk.Label(paired_box, text="Batch size:").grid(row=0, column=0, sticky="w", pady=2)
        paired_batch_entry = ttk.Entry(paired_box, textvariable=self._paired_batch_size_var, width=8)
        paired_batch_entry.grid(row=0, column=1, sticky="w", padx=4, pady=2)
        paired_batch_entry.bind("<FocusOut>", lambda _e: self._sync_algorithm_config(show_error=False))
        paired_batch_entry.bind("<Return>", lambda _e: self._sync_algorithm_config(show_error=False))
        ttk.Label(paired_box, text="Warmup batch size:").grid(row=1, column=0, sticky="w", pady=2)
        paired_warmup_batch_entry = ttk.Entry(
            paired_box, textvariable=self._paired_warmup_batch_size_var, width=8
        )
        paired_warmup_batch_entry.grid(row=1, column=1, sticky="w", padx=4, pady=2)
        paired_warmup_batch_entry.bind(
            "<FocusOut>", lambda _e: self._sync_algorithm_config(show_error=False)
        )
        paired_warmup_batch_entry.bind(
            "<Return>", lambda _e: self._sync_algorithm_config(show_error=False)
        )
        ttk.Checkbutton(
            paired_box,
            text="Use all warmup iterations as one batch",
            variable=self._paired_warmup_single_batch_var,
            command=lambda: self._sync_algorithm_config(show_error=False),
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=2)
        ttk.Label(paired_box, text="Buffer -> target block:").grid(row=3, column=0, sticky="w", pady=2)
        ttk.Entry(paired_box, textvariable=self._paired_target_exchange_var).grid(row=3, column=1, sticky="ew", padx=4, pady=2)
        ttk.Button(
            paired_box,
            text="Browse",
            command=lambda: self._browse_paired_block(self._paired_target_exchange_var, "Choose buffer-to-target exchange block"),
        ).grid(row=3, column=2, padx=2, pady=2)
        ttk.Label(paired_box, text="Target -> buffer block:").grid(row=4, column=0, sticky="w", pady=2)
        ttk.Entry(paired_box, textvariable=self._paired_buffer_exchange_var).grid(row=4, column=1, sticky="ew", padx=4, pady=2)
        ttk.Button(
            paired_box,
            text="Browse",
            command=lambda: self._browse_paired_block(self._paired_buffer_exchange_var, "Choose target-to-buffer exchange block"),
        ).grid(row=4, column=2, padx=2, pady=2)
        ttk.Label(paired_box, text="Target equilibration (s):").grid(row=5, column=0, sticky="w", pady=2)
        ttk.Entry(paired_box, textvariable=self._paired_target_equilibration_var, width=10).grid(row=5, column=1, sticky="w", padx=4, pady=2)
        ttk.Label(paired_box, text="Buffer equilibration (s):").grid(row=6, column=0, sticky="w", pady=2)
        ttk.Entry(paired_box, textvariable=self._paired_buffer_equilibration_var, width=10).grid(row=6, column=1, sticky="w", padx=4, pady=2)
        ttk.Label(
            paired_box,
            text=(
                "Warmup parameter sets count toward the Auto Loop target and run first in warmup-sized buffer/target batches. "
                "Remaining iterations then run in batches using the regular batch size."
            ),
            foreground=self.ACCENT,
            wraplength=460,
            justify="left",
        ).grid(row=7, column=0, columnspan=3, sticky="w", pady=(4, 0))
        ttk.Label(
            paired_box,
            text="Target equilibration runs after buffer-to-target exchange before target SWVs. Buffer equilibration runs after target-to-buffer exchange before the next cycle.",
            foreground=self.ACCENT,
            wraplength=460,
            justify="left",
        ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(4, 0))

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
            params, columns=cols, show="tree headings", height=6, style="BO.Treeview"
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

        scoring_scroll = ScrollableFrame(scoring_tab, min_width=820)
        scoring_scroll.pack(fill="both", expand=True)
        scoring_content = scoring_scroll.content
        scoring_scroll._bind_mousewheel(scoring_content)

        scoring_box = ttk.LabelFrame(scoring_content, text="Q Score Decomposition", padding=8)
        self._normal_scoring_frame = scoring_box
        scoring_scroll._bind_mousewheel(scoring_box)
        scoring_box.pack(fill="both", expand=True, pady=(0, 8), padx=2)
        self._build_q_scoring_controls(
            scoring_box,
            self._setup_scoring_vars(),
            self._score_formula_var,
            lambda: self._sync_scoring_config(show_error=False),
            preset_command=self._apply_signal_priority_preset,
        )
        paired_scoring_box = ttk.LabelFrame(scoring_content, text="Paired Q Scoring", padding=8)
        self._paired_scoring_frame = paired_scoring_box
        scoring_scroll._bind_mousewheel(paired_scoring_box)
        self._build_paired_q_scoring_controls(
            paired_scoring_box,
            vars_by_name=self._paired_scoring_vars(),
            formula_var=self._paired_formula_var,
            on_change=lambda: self._sync_scoring_config(show_error=False),
        )
        self._refresh_paired_score_formula()

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
        ttk.Label(analysis_box, text="Peak V min/max:").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Entry(analysis_box, textvariable=self._analysis_peak_voltage_min_var, width=8).grid(row=2, column=1, sticky="w", padx=(4, 2))
        ttk.Entry(analysis_box, textvariable=self._analysis_peak_voltage_max_var, width=8).grid(row=2, column=1, sticky="e", padx=(2, 4))
        ttk.Label(analysis_box, text="Min start V:").grid(row=2, column=2, sticky="w", pady=2)
        ttk.Entry(analysis_box, textvariable=self._analysis_min_start_voltage_var, width=10).grid(row=2, column=3, sticky="w", padx=4)
        ttk.Label(analysis_box, text="Left min V min/max:").grid(row=3, column=0, sticky="w", pady=2)
        ttk.Entry(analysis_box, textvariable=self._analysis_left_min_voltage_min_var, width=8).grid(row=3, column=1, sticky="w", padx=(4, 2))
        ttk.Entry(analysis_box, textvariable=self._analysis_left_min_voltage_max_var, width=8).grid(row=3, column=1, sticky="e", padx=(2, 4))
        ttk.Label(analysis_box, text="Right min V min/max:").grid(row=3, column=2, sticky="w", pady=2)
        ttk.Entry(analysis_box, textvariable=self._analysis_right_min_voltage_min_var, width=8).grid(row=3, column=3, sticky="w", padx=(4, 2))
        ttk.Entry(analysis_box, textvariable=self._analysis_right_min_voltage_max_var, width=8).grid(row=3, column=3, sticky="e", padx=(2, 4))
        ttk.Label(analysis_box, text="Scan windows:").grid(row=4, column=0, sticky="w", pady=2)
        ttk.Entry(analysis_box, textvariable=self._analysis_scan_windows_var).grid(row=4, column=1, columnspan=3, sticky="ew", padx=4)
        ttk.Checkbutton(analysis_box, text="Prominent minima", variable=self._analysis_use_prominent_var).grid(row=5, column=0, sticky="w", pady=2)
        ttk.Checkbutton(analysis_box, text="Require minima both sides", variable=self._analysis_require_minima_var).grid(row=5, column=1, sticky="w", pady=2)
        ttk.Checkbutton(analysis_box, text="Double correction", variable=self._analysis_double_correction_var).grid(row=5, column=2, sticky="w", pady=2)
        ttk.Checkbutton(analysis_box, text="Compute skew", variable=self._analysis_compute_skew_var).grid(row=5, column=3, sticky="w", pady=2)
        ttk.Checkbutton(analysis_box, text="Wavelet energy", variable=self._analysis_compute_wavelet_energy_var).grid(row=6, column=0, sticky="w", pady=2)
        ttk.Checkbutton(analysis_box, text="Wavelet trace", variable=self._analysis_wavelet_trace_var).grid(row=6, column=1, sticky="w", pady=2)
        ttk.Checkbutton(analysis_box, text="Wavelet correction", variable=self._analysis_wavelet_correction_var).grid(row=6, column=2, sticky="w", pady=2)
        ttk.Label(
            analysis_box,
            text="These settings are sent to the external 64-bit BO analysis worker.",
            foreground=self.ACCENT,
        ).grid(row=7, column=0, columnspan=4, sticky="w", pady=(4, 0))
        self._on_bo_type_changed(sync=False)

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
        auto_bar.add(ttk.Label(auto_bar, text="Total target iterations:"))
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
        setup_page = ttk.Frame(self._engine_page_container)
        landscape_page = ttk.Frame(self._engine_page_container)
        model_page = ttk.Frame(self._engine_page_container)
        results_page = ttk.Frame(self._engine_page_container)
        self._engine_pages = [
            ("0/3 Setup", setup_page),
            ("1/3 Landscape", landscape_page),
            ("2/3 Signal Model", model_page),
            ("3/3 Results", results_page),
        ]

        setup_scroll = ScrollableFrame(setup_page, min_width=900)
        setup_scroll.pack(fill="both", expand=True)
        setup_box = ttk.LabelFrame(setup_scroll.content, text="Simulation Setup", padding=10)
        setup_box.pack(fill="both", expand=True)
        mode_box = ttk.LabelFrame(setup_box, text="Simulation Type", padding=8)
        mode_box.pack(fill="x", pady=(0, 8))
        ttk.Radiobutton(
            mode_box,
            text="Classic BO simulation",
            variable=self._engine_paired_response_var,
            value=False,
            command=self._on_engine_bo_type_changed,
        ).grid(row=0, column=0, sticky="w", padx=4, pady=3)
        ttk.Radiobutton(
            mode_box,
            text="Paired-response batched BO simulation",
            variable=self._engine_paired_response_var,
            value=True,
            command=self._on_engine_bo_type_changed,
        ).grid(row=1, column=0, sticky="w", padx=4, pady=3)

        schedule_box = ttk.LabelFrame(setup_box, text="Simulation Schedule", padding=8)
        schedule_box.pack(fill="x", pady=(0, 8))
        ttk.Label(schedule_box, text="Iterations / cycles:").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Entry(schedule_box, textvariable=self._engine_iterations_var, width=10).grid(row=0, column=1, sticky="w", padx=(6, 18), pady=3)
        ttk.Label(schedule_box, text="Batch size:").grid(row=0, column=2, sticky="w", pady=3)
        ttk.Entry(schedule_box, textvariable=self._engine_paired_batch_size_var, width=10).grid(row=0, column=3, sticky="w", padx=(6, 18), pady=3)

        groups_box = ttk.LabelFrame(setup_box, text="Simulation Channel Groups", padding=8)
        groups_box.pack(fill="x", pady=(0, 8))
        group_header = ttk.Frame(groups_box)
        group_header.pack(fill="x")
        ttk.Label(group_header, text="Number of groups:").pack(side="left")
        engine_group_count = ttk.Combobox(
            group_header,
            textvariable=self._engine_channel_group_count_var,
            values=[str(value) for value in range(1, 11)],
            state="readonly",
            width=4,
        )
        engine_group_count.pack(side="left", padx=6)
        engine_group_count.bind(
            "<<ComboboxSelected>>",
            lambda _e: self._rebuild_engine_channel_group_entries(),
        )
        self._engine_channel_groups_frame = ttk.Frame(groups_box)
        self._engine_channel_groups_frame.pack(fill="x", pady=(4, 0))
        self._rebuild_engine_channel_group_entries()
        ttk.Label(
            groups_box,
            text="Each group runs an independent optimizer using only its simulated channel results.",
            foreground=self.ACCENT,
        ).pack(fill="x", pady=(4, 0))
        ttk.Label(
            setup_box,
            text=(
                "Classic mode treats Iterations as BO observations and ignores Batch size. "
                "Paired-response mode treats Iterations as buffer-to-target cycles and Batch size as hyperparameter sets per cycle."
            ),
            foreground=self.ACCENT,
            wraplength=760,
            justify="left",
        ).pack(fill="x", pady=(0, 8))

        optimizer_box = ttk.LabelFrame(setup_box, text="Optimizer Behavior", padding=8)
        optimizer_box.pack(fill="x", pady=(0, 8))
        optimizer_box.columnconfigure(1, weight=1)
        ttk.Label(optimizer_box, text="Exploit <-> Explore:").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Scale(
            optimizer_box,
            from_=0.0,
            to=1.0,
            orient=tk.HORIZONTAL,
            variable=self._engine_exploration_var,
            command=lambda _v: self._refresh_engine_scoring_formulas(),
        ).grid(row=0, column=1, sticky="ew", padx=6, pady=2)
        ttk.Label(
            optimizer_box,
            textvariable=self._engine_exploration_text_var,
            foreground=self.ACCENT,
            width=5,
        ).grid(row=0, column=2, sticky="e")
        optimizer_entries = [
            ("Global pool:", self._engine_candidate_pool_var),
            ("Local pool:", self._engine_local_pool_var),
            ("GP warmup iters/cycles:", self._engine_warmup_iterations_var),
        ]
        for idx, (label, var) in enumerate(optimizer_entries, start=1):
            ttk.Label(optimizer_box, text=label).grid(row=idx, column=0, sticky="w", pady=2)
            ttk.Entry(optimizer_box, textvariable=var, width=12).grid(
                row=idx, column=1, sticky="w", padx=6, pady=2
            )
        ttk.Label(optimizer_box, text="Start point:").grid(row=4, column=0, sticky="w", pady=2)
        ttk.Combobox(
            optimizer_box,
            textvariable=self._engine_initial_point_mode_var,
            values=("specific", "random"),
            state="readonly",
            width=12,
        ).grid(row=4, column=1, sticky="w", padx=6, pady=2)
        ttk.Label(
            optimizer_box,
            text=(
                "In paired mode, warmup is measured in cycles and total warmup points equal "
                "warmup cycles × batch size."
            ),
            foreground=self.ACCENT,
            justify="left",
        ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(4, 0))
        for child in optimizer_box.winfo_children():
            child.destroy()
        optimizer_box.configure(text="Optimizer Behavior by Group")
        self._engine_group_optimizer_panels_frame = ttk.Frame(optimizer_box)
        self._engine_group_optimizer_panels_frame.pack(fill="x")
        self._rebuild_engine_group_optimizer_panels()

        analysis_box = ttk.LabelFrame(setup_box, text="External Analysis Settings", padding=8)
        analysis_box.pack(fill="x", pady=(0, 8))
        for idx in range(4):
            analysis_box.columnconfigure(idx, weight=1 if idx in (1, 3) else 0)
        analysis_vars = self._engine_analysis_vars
        ttk.Label(analysis_box, text="Crop min/max (V):").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Entry(analysis_box, textvariable=analysis_vars["crop_min_v"], width=8).grid(row=0, column=1, sticky="w", padx=(4, 2))
        ttk.Entry(analysis_box, textvariable=analysis_vars["crop_max_v"], width=8).grid(row=0, column=1, sticky="e", padx=(2, 4))
        ttk.Label(analysis_box, text="Smooth win/poly:").grid(row=0, column=2, sticky="w", pady=2)
        ttk.Entry(analysis_box, textvariable=analysis_vars["smooth_window"], width=8).grid(row=0, column=3, sticky="w", padx=(4, 2))
        ttk.Entry(analysis_box, textvariable=analysis_vars["smooth_polyorder"], width=8).grid(row=0, column=3, sticky="e", padx=(2, 4))
        ttk.Label(analysis_box, text="Minima window (V):").grid(row=1, column=0, sticky="w", pady=2)
        ttk.Entry(analysis_box, textvariable=analysis_vars["minima_search_window_v"], width=10).grid(row=1, column=1, sticky="w", padx=4)
        ttk.Label(analysis_box, text="Min peak height (uA):").grid(row=1, column=2, sticky="w", pady=2)
        ttk.Entry(analysis_box, textvariable=analysis_vars["min_peak_height_ua"], width=10).grid(row=1, column=3, sticky="w", padx=4)
        ttk.Label(analysis_box, text="Peak V min/max:").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Entry(analysis_box, textvariable=analysis_vars["peak_voltage_min_v"], width=8).grid(row=2, column=1, sticky="w", padx=(4, 2))
        ttk.Entry(analysis_box, textvariable=analysis_vars["peak_voltage_max_v"], width=8).grid(row=2, column=1, sticky="e", padx=(2, 4))
        ttk.Label(analysis_box, text="Min start V:").grid(row=2, column=2, sticky="w", pady=2)
        ttk.Entry(analysis_box, textvariable=analysis_vars["min_start_voltage_v"], width=10).grid(row=2, column=3, sticky="w", padx=4)
        ttk.Label(analysis_box, text="Left min V min/max:").grid(row=3, column=0, sticky="w", pady=2)
        ttk.Entry(analysis_box, textvariable=analysis_vars["left_min_voltage_min_v"], width=8).grid(row=3, column=1, sticky="w", padx=(4, 2))
        ttk.Entry(analysis_box, textvariable=analysis_vars["left_min_voltage_max_v"], width=8).grid(row=3, column=1, sticky="e", padx=(2, 4))
        ttk.Label(analysis_box, text="Right min V min/max:").grid(row=3, column=2, sticky="w", pady=2)
        ttk.Entry(analysis_box, textvariable=analysis_vars["right_min_voltage_min_v"], width=8).grid(row=3, column=3, sticky="w", padx=(4, 2))
        ttk.Entry(analysis_box, textvariable=analysis_vars["right_min_voltage_max_v"], width=8).grid(row=3, column=3, sticky="e", padx=(2, 4))
        ttk.Label(analysis_box, text="Scan windows:").grid(row=4, column=0, sticky="w", pady=2)
        ttk.Entry(analysis_box, textvariable=analysis_vars["scan_windows"]).grid(row=4, column=1, columnspan=3, sticky="ew", padx=4)
        ttk.Checkbutton(analysis_box, text="Prominent minima", variable=analysis_vars["use_prominent_minima"]).grid(row=5, column=0, sticky="w", pady=2)
        ttk.Checkbutton(analysis_box, text="Require minima both sides", variable=analysis_vars["require_local_minima_on_both_sides"]).grid(row=5, column=1, sticky="w", pady=2)
        ttk.Checkbutton(analysis_box, text="Double correction", variable=analysis_vars["use_double_correction"]).grid(row=5, column=2, sticky="w", pady=2)
        ttk.Checkbutton(analysis_box, text="Compute skew", variable=analysis_vars["compute_skew"]).grid(row=5, column=3, sticky="w", pady=2)
        ttk.Checkbutton(analysis_box, text="Wavelet energy", variable=analysis_vars["compute_wavelet_energy"]).grid(row=6, column=0, sticky="w", pady=2)
        ttk.Checkbutton(analysis_box, text="Wavelet trace", variable=analysis_vars["compute_wavelet_denoised_trace"]).grid(row=6, column=1, sticky="w", pady=2)
        ttk.Checkbutton(analysis_box, text="Wavelet correction", variable=analysis_vars["use_wavelet_for_correction"]).grid(row=6, column=2, sticky="w", pady=2)
        ttk.Label(
            analysis_box,
            text="These values are sent with every simulated raw trace to the same 64-bit analysis worker used by real BO.",
            foreground=self.ACCENT,
            wraplength=760,
            justify="left",
        ).grid(row=7, column=0, columnspan=4, sticky="w", pady=(4, 0))

        scoring_tabs = ttk.Notebook(setup_box)
        scoring_tabs.pack(fill="both", expand=True, pady=(0, 8))
        classic_scoring_tab = ttk.Frame(scoring_tabs)
        paired_scoring_tab = ttk.Frame(scoring_tabs)
        self._engine_scoring_tabs = scoring_tabs
        self._engine_classic_scoring_tab = classic_scoring_tab
        self._engine_paired_scoring_tab = paired_scoring_tab
        scoring_tabs.add(classic_scoring_tab, text="Classic Q Scoring")
        scoring_tabs.add(paired_scoring_tab, text="Paired Q Scoring")
        self._build_q_scoring_controls(
            classic_scoring_tab,
            self._engine_score_vars,
            self._engine_score_formula_var,
            self._refresh_engine_scoring_formulas,
        )
        self._build_paired_q_scoring_controls(
            paired_scoring_tab,
            vars_by_name=self._engine_paired_score_vars,
            formula_var=self._engine_paired_formula_var,
            on_change=self._refresh_engine_scoring_formulas,
        )
        self._refresh_engine_scoring_formulas()
        self._on_engine_bo_type_changed()
        ttk.Button(setup_box, text="Next: Landscape", command=self._engine_next_page).pack(anchor="e", pady=(8, 0))

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

        dim_cols = ("Min", "Max", "Q Opt", "Q Spread", "Q Shape", "Q Wt", "Delta", "D Opt", "D Spread", "D Shape", "D Wt")
        self._engine_dim_tree = ttk.Treeview(dims_left, columns=dim_cols, show="tree headings", height=14, style="BO.Treeview")
        self._engine_dim_tree.heading("#0", text="Parameter")
        self._engine_dim_tree.column("#0", width=150)
        for col in dim_cols:
            self._engine_dim_tree.heading(col, text=col)
            self._engine_dim_tree.column(col, width=76, anchor="center")
        self._engine_dim_tree.pack(fill="both", expand=True)
        self._engine_dim_tree.bind("<Double-1>", lambda _e: self._engine_edit_dimension())
        self._engine_dim_tree.bind("<<TreeviewSelect>>", lambda _e: self._engine_refresh_landscape_inspector(refresh_cube=False))

        dist_box = ttk.LabelFrame(dims_right, text="Map Slice: Per-Dimension Success Distribution", padding=6)
        dist_box.pack(fill="both", expand=True)
        self._engine_distribution_frame = ttk.Frame(dist_box)
        self._engine_distribution_frame.pack(fill="both", expand=True)

        cube_box = ttk.LabelFrame(dims_right, text="Example Fake Data From Map Cells", padding=6)
        cube_box.pack(fill="both", expand=True, pady=(6, 0))
        cube_cols = ("True Q", "Success", "Peak", "Noise", "Delta Peak")
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
            ("Grid", self._engine_grid_var),
            ("Seed", self._engine_seed_var),
            ("Meas noise", self._engine_measurement_noise_var),
            ("Channel noise", self._engine_channel_noise_var),
            ("Target response gain", self._engine_target_response_gain_var),
            ("Target noise multiplier", self._engine_target_noise_multiplier_var),
            ("Delta peak floor", self._engine_delta_peak_floor_var),
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
        gp_box = ttk.LabelFrame(model_box, text="GP Falloff Fractions", padding=6)
        gp_box.pack(fill="x", pady=(8, 0))
        for idx, name in enumerate(PARAMETER_ORDER):
            row = idx // 2
            col = (idx % 2) * 2
            ttk.Label(gp_box, text=f"{name.replace('_', ' ').title()}:").grid(row=row, column=col, sticky="w", pady=2)
            entry = ttk.Entry(gp_box, textvariable=self._engine_gp_length_scale_vars[name], width=10)
            entry.grid(row=row, column=col + 1, sticky="w", padx=(4, 12), pady=2)
        run_bar = ttk.Frame(model_box)
        run_bar.pack(fill="x", pady=(8, 0))
        ttk.Button(run_bar, text="Draw Landscape", command=self._engine_draw_landscape).pack(side="left", padx=2)
        ttk.Button(run_bar, text="Run Optimizer Simulation", command=self._engine_run_optimizer).pack(side="left", padx=2)
        ttk.Button(run_bar, text="Next: Results", command=self._engine_next_page).pack(side="right", padx=2)
        progress_row = ttk.Frame(model_box)
        progress_row.pack(fill="x", pady=(8, 0))
        self._engine_progress_bar = ttk.Progressbar(
            progress_row,
            orient=tk.HORIZONTAL,
            mode="determinate",
            variable=self._engine_progress_var,
            maximum=100.0,
        )
        self._engine_progress_bar.pack(side="left", fill="x", expand=True)
        ttk.Label(progress_row, textvariable=self._engine_progress_text_var, width=28, anchor="e").pack(side="left", padx=(8, 0))
        ttk.Label(
            model_box,
            textvariable=self._gp_falloff_summary_var,
            foreground=self.ACCENT,
            wraplength=760,
            justify="left",
        ).pack(fill="x", pady=(8, 0))
        ttk.Label(
            model_box,
            text=(
                "Grid is the number of sampled points per simulated dimension used when drawing the landscape; it is clamped to 5-45, so 25 means 25 points in 1D, 25x25 in 2D, or 25x25x25 in 3D.\n"
                "Seed is any integer used to make the random landscape and noise reproducible; using the same seed repeats the same synthetic system.\n"
                "Global pool is the minimum number of broad BO candidates sampled from the full search space; it is clamped to at least 50 and larger values give the surrogate more candidate points to rank.\n"
                "Local pool is the number of extra BO candidates sampled near promising regions; 0 disables local candidates and larger values bias the candidate set toward local refinement.\n"
                "Meas noise is a nonnegative current-noise scale applied to simulated peak/background measurements; 0 means no measurement jitter, about 0.03 is mild, about 0.1 is noticeable, and there is no hard maximum although values near 1 uA can dominate the default signal model.\n"
                "Channel noise is a unitless Gaussian variation added to each channel's underlying quality before clipping to 0-1; 0 means identical channels, about 0.025 is mild, about 0.2 is large, and 1 is effectively the maximum useful range because channel quality is clipped.\n"
                "Target response gain is the maximum target-induced peak-height increase in uA. Delta peak optimum/spread/shape define an independent response-vs-frequency curve, separate from the normal Q landscape. Target noise multiplier scales background RMS in the target phase.\n"
                "Peak emphasis is usually 0-1 and controls how much Q rewards signal height relative to noise and shape; higher values make the optimizer chase taller peaks more aggressively.\n"
                "Base peak uA is the minimum simulated peak current.\n"
                "Peak gain uA is the extra peak current available near good parameter regions.\n"
                "Base noise uA is the minimum background RMS noise.\n"
                "Noise gain uA is the extra background noise added when parameters are far from the synthetic optimum.\n"
                "GP Falloff is a positive unitless fraction of each parameter's search range; 0.2 means points about 20% of that range apart are still meaningfully correlated, larger values make smoother GP predictions, and blank values let the GP learn the falloff."
            ),
            foreground=self.ACCENT,
            wraplength=760,
            justify="left",
        ).pack(fill="x", pady=(10, 0))

        results = ttk.PanedWindow(results_page, orient=tk.HORIZONTAL)
        results.pack(fill="both", expand=True)
        plot_box = ttk.LabelFrame(results, text="Optimizer Movement", padding=6)
        detail_box = ttk.LabelFrame(results, text="Simulation Window", padding=6)
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
        ttk.Progressbar(
            result_toolbar,
            orient=tk.HORIZONTAL,
            mode="determinate",
            variable=self._engine_progress_var,
            maximum=100.0,
            length=160,
        ).pack(side="left", padx=(10, 4))
        ttk.Label(result_toolbar, textvariable=self._engine_progress_text_var, width=14).pack(side="left")
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

        result_cols = ("Group", "Set", "BO Iter", "Buffer Trace", "Target Trace", "Q_run", "True Q", "Paired Q", "Delta Peak", "Distance", "Peak uA", "Peak Prominence", "Begin", "End", "Step", "Amp", "Freq")
        self._engine_result_tree = ttk.Treeview(detail_box, columns=result_cols, show="tree headings", height=9, style="BO.Treeview")
        self._engine_result_tree.heading("#0", text="Iter")
        self._engine_result_tree.column("#0", width=62, anchor="center")
        for col in result_cols:
            self._engine_result_tree.heading(col, text=col)
            self._engine_result_tree.column(col, width=82, anchor="center")
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
        pane = self._visible_paned_window(parent, orient=tk.VERTICAL)
        pane.pack(fill="both", expand=True, padx=4, pady=4)
        top = self._visible_paned_window(pane, orient=tk.HORIZONTAL)
        middle = self._visible_paned_window(pane, orient=tk.HORIZONTAL)
        self._results_main_pane = pane
        self._results_top_pane = top
        self._results_middle_pane = middle
        bottom = self._visible_paned_window(pane, orient=tk.HORIZONTAL)
        self._results_bottom_pane = bottom
        top.bind("<Configure>", lambda _e: self._balance_results_trace_panes(), add="+")
        middle.bind("<Configure>", lambda _e: self._balance_results_trace_panes(), add="+")
        bottom.bind("<Configure>", lambda _e: self._balance_results_trace_panes(), add="+")
        pane.add(top, minsize=160, stretch="always")
        pane.add(middle, minsize=180, stretch="always")
        pane.add(bottom, minsize=120, stretch="always")

        score_box = ttk.LabelFrame(top, text="Per-Channel Scores", padding=6)
        best_box = ttk.LabelFrame(top, text="Raw SWV Traces for Selected Iteration", padding=6)
        top.add(score_box, minsize=260, stretch="always")
        top.add(best_box, minsize=260, stretch="always")

        score_tree_frame = ttk.Frame(score_box)
        score_tree_frame.pack(fill="both", expand=True)
        score_cols = (
            "Classic Q", "Prom. Term", "Repeat-SNR Term", "Peak Term",
            "Shape Term", "Baseline Term", "Replicate Term", "Success Term",
            "Noise Adj", "Clip Adj", "Peak uA", "Peak Prominence",
            "Repeat-scan SNR", "Shape", "Baseline", "Replicate", "Success",
        )
        self._score_tree = ttk.Treeview(score_tree_frame, columns=score_cols, show="tree headings", height=10, selectmode="extended")
        self._score_tree.heading("#0", text="Ch")
        self._score_tree.column("#0", width=50, anchor="center", stretch=False)
        for col in score_cols:
            self._score_tree.heading(col, text=col)
            self._score_tree.column(col, width=78, anchor="center", stretch=False)
        self._score_tree.grid(row=0, column=0, sticky="nsew")
        score_tree_frame.columnconfigure(0, weight=1)
        score_tree_frame.rowconfigure(0, weight=1)
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
        score_y = ttk.Scrollbar(score_tree_frame, orient=tk.VERTICAL, command=self._score_tree.yview)
        score_y.grid(row=0, column=1, sticky="ns")
        self._score_tree.configure(yscrollcommand=score_y.set)
        score_x = ttk.Scrollbar(score_tree_frame, orient=tk.HORIZONTAL, command=self._score_tree.xview)
        score_x.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        self._score_tree.configure(xscrollcommand=score_x.set)
        score_scroll_controls = ttk.Frame(score_box)
        score_scroll_controls.pack(fill="x", pady=(4, 0))
        ttk.Label(score_scroll_controls, text="Score scroll:").pack(side="left", padx=(0, 6))
        ttk.Button(
            score_scroll_controls,
            text="<",
            width=3,
            command=lambda: self._score_tree.xview_scroll(-3, "units"),
        ).pack(side="left")
        ttk.Button(
            score_scroll_controls,
            text=">",
            width=3,
            command=lambda: self._score_tree.xview_scroll(3, "units"),
        ).pack(side="left", padx=(4, 8))
        ttk.Label(
            score_scroll_controls,
            text="Use the bar above or Shift+mouse wheel to view hidden columns.",
        ).pack(side="left")
        self._q_equation_text = scrolledtext.ScrolledText(score_box, height=5, wrap=tk.WORD)
        self._q_equation_text.pack(fill="x", pady=(6, 0))
        self._q_equation_text.config(state="disabled")

        self._raw_trace_frame = ttk.Frame(best_box)
        self._raw_trace_frame.pack(fill="both", expand=True)

        history_host = ttk.LabelFrame(middle, text="History and Trend", padding=6)
        corrected_box = ttk.LabelFrame(middle, text="Smoothed Corrected Traces for Selected Iteration", padding=6)
        middle.add(history_host, minsize=260, stretch="always")
        middle.add(corrected_box, minsize=260, stretch="always")

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
        hist_cols = (
            "Group", "Q_run", "Mean", "Std", "Failed", "Poor",
            "Peak uA", "Noise uA", "Peak Prominence", "Repeat-scan SNR", "Prominence Score", "Shape", "Baseline", "Replicate", "Success",
            "Begin", "End", "Step", "Amp", "Freq", "Cond E", "Cond t",
        )
        history_tree_frame = ttk.Frame(hist_box)
        history_tree_frame.pack(fill="both", expand=True)
        self._history_tree = ttk.Treeview(history_tree_frame, columns=hist_cols, show="tree headings", height=10)
        self._history_tree.heading("#0", text="Iter")
        self._history_tree.column("#0", width=55, anchor="center", stretch=False)
        for col in hist_cols:
            self._history_tree.heading(col, text=col)
            self._history_tree.column(col, width=76, anchor="center", stretch=False)
        self._history_tree.grid(row=0, column=0, sticky="nsew")
        history_tree_frame.columnconfigure(0, weight=1)
        history_tree_frame.rowconfigure(0, weight=1)
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
        self._history_tree.bind("<Shift-MouseWheel>", self._scroll_history_horizontally, add="+")
        self._history_tree.bind_all("<Up>", lambda event: self._route_results_arrow(-1, event), add="+")
        self._history_tree.bind_all("<Down>", lambda event: self._route_results_arrow(1, event), add="+")
        self._history_tree.bind_all("<Left>", lambda event: self._route_results_arrow(-1, event), add="+")
        self._history_tree.bind_all("<Right>", lambda event: self._route_results_arrow(1, event), add="+")
        history_y = ttk.Scrollbar(history_tree_frame, orient=tk.VERTICAL, command=self._history_tree.yview)
        history_y.grid(row=0, column=1, sticky="ns")
        self._history_tree.configure(yscrollcommand=history_y.set)
        history_x = ttk.Scrollbar(history_tree_frame, orient=tk.HORIZONTAL, command=self._history_tree.xview)
        history_x.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        self._history_tree.configure(xscrollcommand=history_x.set)
        history_tree_frame.columnconfigure(0, weight=1)
        history_scroll_controls = ttk.Frame(hist_box)
        history_scroll_controls.pack(fill="x", pady=(4, 0))
        ttk.Label(history_scroll_controls, text="History scroll:").pack(side="left", padx=(0, 6))
        ttk.Button(
            history_scroll_controls,
            text="<",
            width=3,
            command=lambda: self._history_tree.xview_scroll(-3, "units"),
        ).pack(side="left")
        ttk.Button(
            history_scroll_controls,
            text=">",
            width=3,
            command=lambda: self._history_tree.xview_scroll(3, "units"),
        ).pack(side="left", padx=(4, 8))
        ttk.Label(
            history_scroll_controls,
            text="Use the bar above or Shift+mouse wheel to view hidden columns.",
        ).pack(side="left")

        q_rescore_left = ttk.Frame(bottom)
        bottom.add(q_rescore_left, minsize=420, stretch="always")

        current_equation_box = ttk.LabelFrame(q_rescore_left, text="Current Q Equation", padding=6)
        current_equation_box.pack(fill="both", expand=True, pady=(0, 6))
        self._current_q_equation_text = scrolledtext.ScrolledText(current_equation_box, height=8, wrap=tk.WORD)
        self._current_q_equation_text.pack(fill="both", expand=True)
        self._current_q_equation_text.config(state="disabled")

        rescore_box = ttk.LabelFrame(q_rescore_left, text="Rescore Recorded Data", padding=8)
        rescore_box.pack(fill="both", expand=True)
        rescore_canvas = tk.Canvas(rescore_box, highlightthickness=0, borderwidth=0)
        rescore_scrollbar = ttk.Scrollbar(rescore_box, orient="vertical", command=rescore_canvas.yview)
        rescore_content = ttk.Frame(rescore_canvas)
        rescore_window = rescore_canvas.create_window((0, 0), window=rescore_content, anchor="nw")
        rescore_canvas.configure(yscrollcommand=rescore_scrollbar.set)
        rescore_content.bind(
            "<Configure>",
            lambda _event: rescore_canvas.configure(scrollregion=rescore_canvas.bbox("all")),
        )
        rescore_canvas.bind(
            "<Configure>",
            lambda event: rescore_canvas.itemconfigure(rescore_window, width=event.width),
        )
        rescore_canvas.pack(side="left", fill="both", expand=True)
        rescore_scrollbar.pack(side="right", fill="y")
        rescore_canvas.bind(
            "<MouseWheel>",
            lambda event: rescore_canvas.yview_scroll(-1 if event.delta > 0 else 1, "units"),
        )

        button_bar = ttk.Frame(rescore_content)
        button_bar.pack(fill="x", pady=(0, 4))
        ttk.Button(button_bar, text="Apply Rescore", command=self._apply_rescore_to_loaded_session).pack(side="left", padx=2)
        self._reanalyze_rescore_button = ttk.Button(
            button_bar,
            text="Reanalyze & Rescore",
            command=self._reanalyze_and_rescore_loaded_session,
        )
        self._reanalyze_rescore_button.pack(side="left", padx=2)
        ttk.Button(button_bar, text="Reset Original", command=self._reset_rescore_to_original).pack(side="left", padx=2)
        ttk.Button(button_bar, text="Save Rescored Session", command=self._save_rescored_session).pack(side="left", padx=2)
        ttk.Label(button_bar, textvariable=self._rescore_status_var, foreground=self.ACCENT, wraplength=430, justify="left").pack(side="left", padx=(8, 2))
        analysis_controls = ttk.LabelFrame(
            rescore_content,
            text="Analysis Parameters Used by Reanalyze & Rescore",
            padding=6,
        )
        analysis_controls.pack(fill="x", pady=(2, 6))
        self._build_reanalysis_controls(analysis_controls)
        controls = ttk.Frame(rescore_content)
        controls.pack(fill="both", expand=True)
        self._build_q_scoring_controls(
            controls,
            self._rescore_scoring_vars(),
            self._rescore_formula_var,
            self._preview_rescore_equation,
            preset_command=self._apply_rescore_signal_priority_preset,
        )
        paired_rescore_controls = ttk.LabelFrame(rescore_content, text="Paired Q Scoring", padding=6)
        paired_rescore_controls.pack(fill="x", pady=(6, 0))
        self._build_paired_q_scoring_controls(
            paired_rescore_controls,
            vars_by_name=self._rescore_paired_scoring_vars(),
            formula_var=self._rescore_paired_formula_var,
            on_change=self._preview_rescore_equation,
        )
        def bind_rescore_wheel(widget):
            widget.bind(
                "<MouseWheel>",
                lambda event: rescore_canvas.yview_scroll(-1 if event.delta > 0 else 1, "units"),
                add="+",
            )
            widget.bind("<Button-4>", lambda _event: rescore_canvas.yview_scroll(-1, "units"), add="+")
            widget.bind("<Button-5>", lambda _event: rescore_canvas.yview_scroll(1, "units"), add="+")
            for child in widget.winfo_children():
                bind_rescore_wheel(child)

        bind_rescore_wheel(rescore_content)

        surrogate_box = ttk.LabelFrame(bottom, text="Surrogate View", padding=6)
        bottom.add(surrogate_box, minsize=260, stretch="always")
        self._build_surrogate_view(surrogate_box)
        parent.after_idle(self._balance_results_trace_panes)

    def _build_surrogate_view(self, parent):
        surrogate_toolbar = FlowFrame(parent)
        surrogate_toolbar.pack(fill="x", pady=(0, 4))
        surrogate_toolbar.add(ttk.Label(surrogate_toolbar, text="Artifact iter:"))
        self._surrogate_iteration_combo = ttk.Combobox(
            surrogate_toolbar,
            textvariable=self._surrogate_iteration_var,
            state="readonly",
            width=8,
        )
        surrogate_toolbar.add(self._surrogate_iteration_combo)
        self._surrogate_iteration_combo.bind("<<ComboboxSelected>>", lambda _e: self._refresh_surrogate_view())
        surrogate_toolbar.add(ttk.Label(surrogate_toolbar, text="Value:"))
        self._surrogate_value_combo = ttk.Combobox(
            surrogate_toolbar,
            textvariable=self._surrogate_value_var,
            values=("predicted_mean_Q", "predicted_std_Q", "acquisition_value"),
            state="readonly",
            width=18,
        )
        surrogate_toolbar.add(self._surrogate_value_combo)
        self._surrogate_value_combo.bind("<<ComboboxSelected>>", lambda _e: self._refresh_surrogate_view())
        surrogate_toolbar.add(ttk.Label(surrogate_toolbar, text="View:"))
        self._surrogate_view_combo = ttk.Combobox(
            surrogate_toolbar,
            textvariable=self._surrogate_view_var,
            values=("1D slice", "2D map", "3D tensor", "Correlation falloff"),
            state="readonly",
            width=18,
        )
        surrogate_toolbar.add(self._surrogate_view_combo)
        self._surrogate_view_combo.bind("<<ComboboxSelected>>", lambda _e: self._refresh_surrogate_view())
        surrogate_toolbar.add(ttk.Label(surrogate_toolbar, text="X:"))
        self._surrogate_x_combo = ttk.Combobox(surrogate_toolbar, textvariable=self._surrogate_x_var, state="readonly", width=18)
        surrogate_toolbar.add(self._surrogate_x_combo)
        self._surrogate_x_combo.bind("<<ComboboxSelected>>", lambda _e: self._refresh_surrogate_view())
        surrogate_toolbar.add(ttk.Label(surrogate_toolbar, text="Y:"))
        self._surrogate_y_combo = ttk.Combobox(surrogate_toolbar, textvariable=self._surrogate_y_var, state="readonly", width=18)
        surrogate_toolbar.add(self._surrogate_y_combo)
        self._surrogate_y_combo.bind("<<ComboboxSelected>>", lambda _e: self._refresh_surrogate_view())
        surrogate_toolbar.add(ttk.Label(surrogate_toolbar, text="Z:"))
        self._surrogate_z_combo = ttk.Combobox(surrogate_toolbar, textvariable=self._surrogate_z_var, state="readonly", width=18)
        surrogate_toolbar.add(self._surrogate_z_combo)
        self._surrogate_z_combo.bind("<<ComboboxSelected>>", lambda _e: self._refresh_surrogate_view())
        surrogate_toolbar.add(ttk.Label(surrogate_toolbar, text="Color min:"))
        color_min_entry = ttk.Entry(surrogate_toolbar, textvariable=self._surrogate_color_min_var, width=8)
        surrogate_toolbar.add(color_min_entry)
        color_min_entry.bind("<Return>", lambda _e: self._refresh_surrogate_view())
        surrogate_toolbar.add(ttk.Label(surrogate_toolbar, text="max:"))
        color_max_entry = ttk.Entry(surrogate_toolbar, textvariable=self._surrogate_color_max_var, width=8)
        surrogate_toolbar.add(color_max_entry)
        color_max_entry.bind("<Return>", lambda _e: self._refresh_surrogate_view())
        surrogate_toolbar.add(ttk.Button(surrogate_toolbar, text="Auto Color", command=self._clear_surrogate_color_range))
        surrogate_toolbar.add(ttk.Button(surrogate_toolbar, text="Refresh", command=self._refresh_surrogate_view))
        self._surrogate_plot_frame = ttk.Frame(parent)
        self._surrogate_plot_frame.pack(fill="both", expand=True)

    def _clear_surrogate_color_range(self):
        self._surrogate_color_min_var.set("")
        self._surrogate_color_max_var.set("")
        self._refresh_surrogate_view()

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

    def _browse_analysis_python(self):
        path = filedialog.askopenfilename(title="Choose 64-bit Python executable")
        if path:
            self._analysis_python_var.set(path)

    def _browse_analysis_project(self):
        path = filedialog.askdirectory(title="Choose experiment automation project")
        if path:
            self._analysis_project_var.set(path)

    def _browse_paired_block(self, variable, title):
        path = filedialog.askopenfilename(
            title=title,
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if path:
            variable.set(path)

    def _save_local_paths(self):
        try:
            payload = {
                "analysis_output_dir": self._analysis_dir_var.get().strip(),
                "analysis_file_glob": self._analysis_glob_var.get().strip() or "*.json",
                "analysis_project": self._analysis_project_var.get().strip(),
                "analysis_script": str(
                    Path(self._analysis_project_var.get().strip())
                    / "analysis_worker"
                    / "bo_headless.py"
                ),
                "analysis_python": self._analysis_python_var.get().strip(),
                "analysis_mode": "external",
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
            self._apply_config_to_setup_vars()
            if not initial:
                self._status_var.set(f"Loaded BO config: {self._config_path_var.get()}")
        except Exception as exc:
            self._config = None
            self._status_var.set(f"BO config load failed: {exc}")
            if not initial:
                messagebox.showerror("BO Config", str(exc))

    def _apply_config_to_setup_vars(self):
        analysis_cfg = self._config.get("analysis", {})
        if analysis_cfg.get("file_glob"):
            self._analysis_glob_var.set(str(analysis_cfg.get("file_glob")))
        self._set_analysis_vars_from_config(analysis_cfg)
        self._set_method_option_vars_from_config(self._config)
        self._set_algorithm_vars_from_config(self._config)
        self._set_scoring_vars_from_config(self._config)
        self._set_engine_tuning_vars_from_config(self._config)
        if self._bo_session is None:
            self._loaded_original_config = None
            self._set_rescore_vars_from_config(self._config)
        self._engine_seed_var.set(str(self._config.get("random_seed", 42)))
        self._channels_var.set(", ".join(str(ch) for ch in self._config.get("channels", [])))
        self._set_channel_group_vars(self._config)
        self._refresh_parameter_table()
        self._refresh_initial_parameters_table()
        self._engine_load_active_dimensions()
        self._validate_config(show_dialog=False)

    def _last_bo_setup_ui_settings(self) -> dict:
        return {
            key: getattr(self, variable_name).get()
            for key, variable_name in self.LAST_SETUP_UI_VARS.items()
        }

    def _load_last_bo_setup(self) -> bool:
        metadata = load_bo_setup_metadata(BO_LAST_SETUP_METADATA_PATH)
        if metadata is None:
            return False
        try:
            ui_settings = dict(metadata.get("ui_settings") or {})
            config_path = str(ui_settings.get("config_path") or "").strip()
            if config_path:
                self._config_path_var.set(config_path)
            self._config = normalize_bo_config(metadata["bo_config"])
            self._apply_config_to_setup_vars()
            for key, variable_name in self.LAST_SETUP_UI_VARS.items():
                if key == "config_path" or key not in ui_settings:
                    continue
                getattr(self, variable_name).set(ui_settings[key])
            self._status_var.set(
                f"Loaded last BO setup defaults: {BO_LAST_SETUP_METADATA_PATH}"
            )
            return True
        except Exception as exc:
            self._status_var.set(f"Last BO setup could not be loaded: {exc}")
            return False

    def _save_config(self):
        if self._config is None:
            return
        self._sync_channel_groups(show_error=False)
        self._config["objective"] = self._bo_objective_var.get()
        analysis_cfg = self._config.setdefault("analysis", {})
        analysis_cfg["file_glob"] = self._analysis_glob_var.get().strip() or "*.json"
        self._update_analysis_config_from_vars(analysis_cfg)
        self._sync_method_options_config(show_error=False)
        self._sync_algorithm_config(show_error=False)
        self._sync_scoring_config(show_error=False)
        try:
            path = save_bo_config(self._config, self._config_path_var.get())
            metadata_path = save_bo_setup_metadata(
                self._config,
                self._last_bo_setup_ui_settings(),
                BO_LAST_SETUP_METADATA_PATH,
            )
            self._status_var.set(
                f"Saved BO config: {path} | next-run defaults: {metadata_path}"
            )
        except Exception as exc:
            messagebox.showerror("Save BO Config", str(exc))

    def _set_analysis_vars_from_config(self, analysis_cfg: dict):
        self._analysis_crop_min_var.set(str(analysis_cfg.get("crop_min_v", -0.6)))
        self._analysis_crop_max_var.set(str(analysis_cfg.get("crop_max_v", -0.1)))
        self._analysis_smooth_window_var.set(str(analysis_cfg.get("smooth_window", 15)))
        self._analysis_smooth_polyorder_var.set(str(analysis_cfg.get("smooth_polyorder", 2)))
        self._analysis_minima_window_var.set(str(analysis_cfg.get("minima_search_window_v", 0.30)))
        self._analysis_min_peak_height_var.set("" if analysis_cfg.get("min_peak_height_ua") in (None, "") else str(analysis_cfg.get("min_peak_height_ua")))
        self._analysis_peak_voltage_min_var.set("" if analysis_cfg.get("peak_voltage_min_v") in (None, "") else str(analysis_cfg.get("peak_voltage_min_v")))
        self._analysis_peak_voltage_max_var.set("" if analysis_cfg.get("peak_voltage_max_v") in (None, "") else str(analysis_cfg.get("peak_voltage_max_v")))
        self._analysis_left_min_voltage_min_var.set("" if analysis_cfg.get("left_min_voltage_min_v") in (None, "") else str(analysis_cfg.get("left_min_voltage_min_v")))
        self._analysis_left_min_voltage_max_var.set("" if analysis_cfg.get("left_min_voltage_max_v") in (None, "") else str(analysis_cfg.get("left_min_voltage_max_v")))
        self._analysis_right_min_voltage_min_var.set("" if analysis_cfg.get("right_min_voltage_min_v") in (None, "") else str(analysis_cfg.get("right_min_voltage_min_v")))
        self._analysis_right_min_voltage_max_var.set("" if analysis_cfg.get("right_min_voltage_max_v") in (None, "") else str(analysis_cfg.get("right_min_voltage_max_v")))
        self._analysis_min_start_voltage_var.set(str(analysis_cfg.get("min_start_voltage_v", -0.6)))
        self._analysis_scan_windows_var.set(str(analysis_cfg.get("scan_windows", "")))
        self._analysis_use_prominent_var.set(bool(analysis_cfg.get("use_prominent_minima", False)))
        self._analysis_require_minima_var.set(bool(analysis_cfg.get("require_local_minima_on_both_sides", False)))
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
        peak_voltage_min_text = (self._analysis_peak_voltage_min_var.get() or "").strip()
        peak_voltage_max_text = (self._analysis_peak_voltage_max_var.get() or "").strip()
        analysis_cfg["peak_voltage_min_v"] = None if not peak_voltage_min_text else float(peak_voltage_min_text)
        analysis_cfg["peak_voltage_max_v"] = None if not peak_voltage_max_text else float(peak_voltage_max_text)
        left_min_text = (self._analysis_left_min_voltage_min_var.get() or "").strip()
        left_max_text = (self._analysis_left_min_voltage_max_var.get() or "").strip()
        right_min_text = (self._analysis_right_min_voltage_min_var.get() or "").strip()
        right_max_text = (self._analysis_right_min_voltage_max_var.get() or "").strip()
        analysis_cfg["left_min_voltage_min_v"] = None if not left_min_text else float(left_min_text)
        analysis_cfg["left_min_voltage_max_v"] = None if not left_max_text else float(left_max_text)
        analysis_cfg["right_min_voltage_min_v"] = None if not right_min_text else float(right_min_text)
        analysis_cfg["right_min_voltage_max_v"] = None if not right_max_text else float(right_max_text)
        self._validate_voltage_window(analysis_cfg, "peak_voltage_min_v", "peak_voltage_max_v", "Peak V")
        self._validate_voltage_window(analysis_cfg, "left_min_voltage_min_v", "left_min_voltage_max_v", "Left min V")
        self._validate_voltage_window(analysis_cfg, "right_min_voltage_min_v", "right_min_voltage_max_v", "Right min V")
        analysis_cfg["min_start_voltage_v"] = float(self._analysis_min_start_voltage_var.get())
        analysis_cfg["scan_windows"] = (self._analysis_scan_windows_var.get() or "").strip()
        analysis_cfg["use_prominent_minima"] = bool(self._analysis_use_prominent_var.get())
        analysis_cfg["require_local_minima_on_both_sides"] = bool(self._analysis_require_minima_var.get())
        analysis_cfg["use_double_correction"] = bool(self._analysis_double_correction_var.get())
        analysis_cfg["compute_skew"] = bool(self._analysis_compute_skew_var.get())
        analysis_cfg["compute_wavelet_energy"] = bool(self._analysis_compute_wavelet_energy_var.get())
        analysis_cfg["compute_wavelet_denoised_trace"] = bool(self._analysis_wavelet_trace_var.get())
        analysis_cfg["use_wavelet_for_correction"] = bool(self._analysis_wavelet_correction_var.get())

    @staticmethod
    def _validate_voltage_window(config: dict, min_key: str, max_key: str, label: str):
        if (
            config[min_key] is not None
            and config[max_key] is not None
            and config[min_key] > config[max_key]
        ):
            raise ValueError(f"{label} min must be less than or equal to {label} max.")

    def _set_method_option_vars_from_config(self, cfg: dict):
        method_options = dict((cfg or {}).get("method_options") or {})
        normalized = normalize_swv_ba_range_options(method_options)
        bandwidth = str(method_options.get("bandwidth", "4k")).strip().lower() or "4k"
        self._bo_bandwidth_var.set(bandwidth if bandwidth in ("4k", "8k") else "4k")
        self._bo_ba_range_mode_var.set(normalized["mode"])
        self._bo_ba_fixed_range_var.set(normalized["fixed_label"])
        self._bo_ba_auto_min_var.set(normalized["auto_min_label"])
        self._bo_ba_auto_max_var.set(normalized["auto_max_label"])
        self._measurements_per_channel_var.set(
            str(max(1, int((cfg or {}).get("measurements_per_channel", 1) or 1)))
        )
        self._sync_bo_ba_range_controls(save=False)

    def _sync_bo_ba_range_controls(self, save=True):
        mode = (self._bo_ba_range_mode_var.get() or "fixed").strip().lower()
        fixed_state = "readonly" if mode == "fixed" else "disabled"
        auto_state = "readonly" if mode == "auto" else "disabled"
        self._bo_ba_fixed_combo.configure(state=fixed_state)
        self._bo_ba_auto_min_combo.configure(state=auto_state)
        self._bo_ba_auto_max_combo.configure(state=auto_state)
        if save:
            self._sync_method_options_config(show_error=False)

    def _sync_method_options_config(self, show_error=True):
        if self._config is None:
            return
        try:
            method_options = self._config.setdefault("method_options", {})
            bandwidth = (self._bo_bandwidth_var.get() or "4k").strip().lower()
            if bandwidth not in ("4k", "8k"):
                raise ValueError(f"Unsupported SWV bandwidth: {bandwidth}")
            normalized = normalize_swv_ba_range_options(
                {
                    "ba_range": {
                        "mode": self._bo_ba_range_mode_var.get(),
                        "fixed": self._bo_ba_fixed_range_var.get(),
                        "auto_min": self._bo_ba_auto_min_var.get(),
                        "auto_max": self._bo_ba_auto_max_var.get(),
                    }
                }
            )
            method_options["bandwidth"] = bandwidth
            method_options["ba_range"] = {
                "mode": normalized["mode"],
                "fixed": normalized["fixed_label"],
                "auto_min": normalized["auto_min_label"],
                "auto_max": normalized["auto_max_label"],
            }
            measurement_count = int(self._measurements_per_channel_var.get() or 1)
            if measurement_count < 1:
                raise ValueError("Measurements per channel / point must be at least 1.")
            self._config["measurements_per_channel"] = measurement_count
            self._measurements_per_channel_var.set(str(measurement_count))
            self._bo_ba_range_mode_var.set(normalized["mode"])
            self._bo_ba_fixed_range_var.set(normalized["fixed_label"])
            self._bo_ba_auto_min_var.set(normalized["auto_min_label"])
            self._bo_ba_auto_max_var.set(normalized["auto_max_label"])
        except Exception as exc:
            if show_error:
                messagebox.showerror("Method Settings", str(exc))

    def _set_algorithm_vars_from_config(self, cfg: dict):
        acquisition = dict((cfg or {}).get("acquisition") or {})
        self._exploration_var.set(float(acquisition.get("exploration", 0.35)))
        self._exploration_text_var.set(f"{float(acquisition.get('exploration', 0.35)):.2f}")
        if str((cfg or {}).get("objective") or "").lower() == "paired_response":
            self._gp_warmup_iterations_var.set(str(int((cfg or {}).get("paired_warmup_cycles", (cfg or {}).get("n_initial_points", 8)))))
            if (cfg or {}).get("paired_batch_size") is not None:
                self._paired_batch_size_var.set(str(max(1, int((cfg or {}).get("paired_batch_size") or 1))))
            self._paired_warmup_batch_size_var.set(
                str(
                    max(
                        1,
                        int(
                            (cfg or {}).get(
                                "paired_warmup_batch_size",
                                (cfg or {}).get("paired_batch_size", 1),
                            )
                            or 1
                        ),
                    )
                )
            )
            self._paired_warmup_single_batch_var.set(
                bool((cfg or {}).get("paired_warmup_single_batch", False))
            )
        else:
            self._gp_warmup_iterations_var.set(str(int((cfg or {}).get("n_initial_points", 8))))
        self._candidate_pool_var.set(str(acquisition.get("candidate_pool_size", 600)))
        self._local_pool_var.set(str(acquisition.get("local_candidate_pool_size", 120)))
        self._initial_point_mode_var.set(str(acquisition.get("initial_point_mode", "specific")))
        self._optimization_direction_var.set(self._display_optimization_direction(acquisition.get("optimization_direction", "maximize")))
        length_scales = dict(acquisition.get("gp_falloff_fractions") or acquisition.get("gp_length_scales") or {})
        if not length_scales:
            length_scales = {name: 0.2 for name in PARAMETER_ORDER}
        for name, var in self._gp_length_scale_vars.items():
            var.set(str(length_scales.get(name, 0.2)))
        self._refresh_gp_falloff_summary()

    def _set_engine_tuning_vars_from_config(self, cfg: dict):
        acquisition = dict((cfg or {}).get("acquisition") or {})
        exploration = float(acquisition.get("exploration", 0.35))
        self._engine_exploration_var.set(exploration)
        self._engine_exploration_text_var.set(f"{exploration:.2f}")
        self._engine_candidate_pool_var.set(str(acquisition.get("candidate_pool_size", 600)))
        self._engine_local_pool_var.set(str(acquisition.get("local_candidate_pool_size", 120)))
        self._engine_initial_point_mode_var.set(str(acquisition.get("initial_point_mode", "specific")))
        self._engine_optimization_direction_var.set(
            self._display_optimization_direction(acquisition.get("optimization_direction", "maximize"))
        )
        paired = str((cfg or {}).get("objective") or "").lower() == "paired_response"
        if paired:
            warmup = int((cfg or {}).get("paired_warmup_cycles", 0) or 0)
            batch_size = max(1, int((cfg or {}).get("paired_batch_size", 1) or 1))
            self._engine_paired_batch_size_var.set(str(batch_size))
        else:
            warmup = int((cfg or {}).get("n_initial_points", 8) or 0)
        self._engine_warmup_iterations_var.set(str(warmup))
        if hasattr(self, "_engine_channel_group_count_var"):
            groups = channel_groups(cfg)
            self._engine_channel_group_count_var.set(str(len(groups)))
            self._rebuild_engine_channel_group_entries(
                [", ".join(str(ch) for ch in group["channels"]) for group in groups],
                group_configs=groups,
            )
        length_scales = dict(
            acquisition.get("gp_falloff_fractions")
            or acquisition.get("gp_length_scales")
            or {}
        )
        for name, var in self._engine_gp_length_scale_vars.items():
            var.set("" if not length_scales else str(length_scales.get(name, 0.2)))

        scoring = dict((cfg or {}).get("scoring") or {})
        self._set_scoring_vars(cfg, self._engine_score_vars, self._engine_score_formula_var)
        self._set_paired_scoring_vars(scoring, self._engine_paired_score_vars)
        if hasattr(self, "_engine_analysis_vars"):
            self._set_engine_analysis_vars(dict((cfg or {}).get("analysis") or {}))
        self._refresh_engine_scoring_formulas()

    def _set_engine_analysis_vars(self, analysis: dict):
        values = {
            "crop_min_v": analysis.get("crop_min_v", -0.61),
            "crop_max_v": analysis.get("crop_max_v", -0.30),
            "smooth_window": analysis.get("smooth_window", 15),
            "smooth_polyorder": analysis.get("smooth_polyorder", 2),
            "minima_search_window_v": analysis.get("minima_search_window_v", 0.30),
            "min_peak_height_ua": analysis.get("min_peak_height_ua", 0.001),
            "peak_voltage_min_v": analysis.get("peak_voltage_min_v"),
            "peak_voltage_max_v": analysis.get("peak_voltage_max_v"),
            "left_min_voltage_min_v": analysis.get("left_min_voltage_min_v"),
            "left_min_voltage_max_v": analysis.get("left_min_voltage_max_v"),
            "right_min_voltage_min_v": analysis.get("right_min_voltage_min_v"),
            "right_min_voltage_max_v": analysis.get("right_min_voltage_max_v"),
            "min_start_voltage_v": analysis.get("min_start_voltage_v", -0.70),
            "scan_windows": analysis.get("scan_windows", ""),
        }
        for key, value in values.items():
            self._engine_analysis_vars[key].set("" if value is None else str(value))
        for key in (
            "use_prominent_minima",
            "require_local_minima_on_both_sides",
            "use_double_correction",
            "compute_skew",
            "compute_wavelet_energy",
            "compute_wavelet_denoised_trace",
            "use_wavelet_for_correction",
        ):
            self._engine_analysis_vars[key].set(bool(analysis.get(key, False)))

    def _engine_analysis_config(self):
        variables = self._engine_analysis_vars

        def optional_float(key):
            variable = variables.get(key)
            if variable is None:
                return None
            text = (variable.get() or "").strip()
            return None if not text else float(text)

        analysis = {
            "crop_min_v": float(variables["crop_min_v"].get()),
            "crop_max_v": float(variables["crop_max_v"].get()),
            "smooth_window": int(variables["smooth_window"].get()),
            "smooth_polyorder": int(variables["smooth_polyorder"].get()),
            "minima_search_window_v": float(variables["minima_search_window_v"].get()),
            "min_peak_height_ua": optional_float("min_peak_height_ua"),
            "peak_voltage_min_v": optional_float("peak_voltage_min_v"),
            "peak_voltage_max_v": optional_float("peak_voltage_max_v"),
            "left_min_voltage_min_v": optional_float("left_min_voltage_min_v"),
            "left_min_voltage_max_v": optional_float("left_min_voltage_max_v"),
            "right_min_voltage_min_v": optional_float("right_min_voltage_min_v"),
            "right_min_voltage_max_v": optional_float("right_min_voltage_max_v"),
            "min_start_voltage_v": float(variables["min_start_voltage_v"].get()),
            "scan_windows": (variables["scan_windows"].get() or "").strip(),
            "use_prominent_minima": bool(variables["use_prominent_minima"].get()),
            "require_local_minima_on_both_sides": bool(variables["require_local_minima_on_both_sides"].get()),
            "use_double_correction": bool(variables["use_double_correction"].get()),
            "compute_skew": bool(variables["compute_skew"].get()),
            "compute_wavelet_energy": bool(variables["compute_wavelet_energy"].get()),
            "compute_wavelet_denoised_trace": bool(variables["compute_wavelet_denoised_trace"].get()),
            "use_wavelet_for_correction": bool(variables["use_wavelet_for_correction"].get()),
        }
        if analysis["crop_min_v"] >= analysis["crop_max_v"]:
            raise ValueError("Simulation analysis crop minimum must be below crop maximum.")
        if analysis["smooth_window"] < 0 or analysis["smooth_polyorder"] < 0:
            raise ValueError("Simulation smoothing window and polynomial order must be nonnegative.")
        if analysis["minima_search_window_v"] <= 0:
            raise ValueError("Simulation minima window must be positive.")
        self._validate_voltage_window(analysis, "peak_voltage_min_v", "peak_voltage_max_v", "Simulation peak V")
        self._validate_voltage_window(analysis, "left_min_voltage_min_v", "left_min_voltage_max_v", "Simulation left min V")
        self._validate_voltage_window(analysis, "right_min_voltage_min_v", "right_min_voltage_max_v", "Simulation right min V")
        return analysis

    def _refresh_engine_scoring_formulas(self):
        self._engine_exploration_text_var.set(f"{float(self._engine_exploration_var.get()):.2f}")
        self._refresh_formula_from_vars(self._engine_score_vars, self._engine_score_formula_var)
        self._engine_paired_formula_var.set(
            self._paired_formula_text(
                self._paired_scoring_from_var_map(self._engine_paired_score_vars)
            )
        )

    def _on_engine_bo_type_changed(self):
        tabs = getattr(self, "_engine_scoring_tabs", None)
        paired_tab = getattr(self, "_engine_paired_scoring_tab", None)
        classic_tab = getattr(self, "_engine_classic_scoring_tab", None)
        if tabs is None or paired_tab is None:
            return
        if bool(self._engine_paired_response_var.get()):
            tabs.add(paired_tab, text="Paired Q Scoring")
        else:
            try:
                if tabs.select() == str(paired_tab) and classic_tab is not None:
                    tabs.select(classic_tab)
            except Exception:
                pass
            tabs.hide(paired_tab)

    def _engine_bo_config(self, sim_cfg=None):
        if self._config is None:
            raise ValueError("Load a BO config first.")
        sim_cfg = sim_cfg or self._engine_sim_config()
        cfg = json.loads(json.dumps(self._config))
        if hasattr(self, "_engine_channel_group_vars"):
            groups = self._engine_channel_groups_from_vars()
            cfg["channel_groups"] = groups
            cfg["channels"] = [
                channel for group in groups for channel in group["channels"]
            ]
        acquisition = cfg.setdefault("acquisition", {})
        acquisition["exploration"] = max(0.0, min(1.0, float(self._engine_exploration_var.get())))
        acquisition["candidate_pool_size"] = max(50, int(self._engine_candidate_pool_var.get() or 600))
        acquisition["local_candidate_pool_size"] = max(0, int(self._engine_local_pool_var.get() or 120))
        mode = str(self._engine_initial_point_mode_var.get() or "specific").strip().lower()
        acquisition["initial_point_mode"] = "random" if mode == "random" else "specific"
        engine_direction_var = getattr(self, "_engine_optimization_direction_var", None)
        acquisition["optimization_direction"] = self._display_optimization_direction(
            engine_direction_var.get() if engine_direction_var is not None else "maximize"
        )
        falloffs = {
            name: (var.get() or "").strip()
            for name, var in self._engine_gp_length_scale_vars.items()
        }
        populated = {name: float(value) for name, value in falloffs.items() if value}
        if populated and len(populated) != len(PARAMETER_ORDER):
            missing = [name for name in PARAMETER_ORDER if not falloffs[name]]
            raise ValueError(
                "Fill every simulation GP falloff fraction, or clear every field. "
                f"Missing: {', '.join(missing)}"
            )
        if populated and any(value <= 0 for value in populated.values()):
            raise ValueError("Simulation GP falloff fractions must be positive.")
        acquisition["gp_falloff_fractions"] = populated
        acquisition["gp_length_scales"] = populated

        scoring = self._scoring_from_vars(self._engine_score_vars)
        scoring["paired_response_weights"] = {
            key: max(0.0, float(var.get() or 0.0))
            for key, var in self._engine_paired_score_vars.items()
        }
        cfg["scoring"] = scoring
        if hasattr(self, "_engine_analysis_vars"):
            cfg["analysis"] = self._engine_analysis_config()
        warmup = max(0, int(self._engine_warmup_iterations_var.get() or 0))
        if bool(sim_cfg.get("paired_response")):
            batch_size = max(1, int(sim_cfg.get("paired_batch_size", 1)))
            cfg["objective"] = "paired_response"
            cfg["paired_warmup_cycles"] = warmup
            cfg["paired_batch_size"] = batch_size
            cfg["n_initial_points"] = warmup * batch_size
        else:
            cfg["objective"] = "quality"
            cfg["n_initial_points"] = warmup
            cfg.pop("paired_warmup_cycles", None)
            cfg.pop("paired_batch_size", None)
        return normalize_bo_config(cfg)

    def _sync_algorithm_config(self, show_error=True):
        if self._config is None:
            return
        try:
            acquisition = self._config.setdefault("acquisition", {})
            acquisition["exploration"] = max(0.0, min(1.0, float(self._exploration_var.get())))
            self._exploration_text_var.set(f"{float(acquisition['exploration']):.2f}")
            warmup_value = max(0, int(self._gp_warmup_iterations_var.get() or 8))
            if self._bo_objective_var.get() == "paired_response":
                batch_size = max(1, int(self._paired_batch_size_var.get() or 1))
                warmup_batch_size = max(
                    1, int(self._paired_warmup_batch_size_var.get() or batch_size)
                )
                self._config["paired_warmup_cycles"] = warmup_value
                self._config["paired_batch_size"] = batch_size
                self._config["paired_warmup_batch_size"] = warmup_batch_size
                self._config["paired_warmup_single_batch"] = bool(
                    self._paired_warmup_single_batch_var.get()
                )
                self._config["n_initial_points"] = warmup_value * batch_size
            else:
                self._config.pop("paired_warmup_cycles", None)
                self._config.pop("paired_batch_size", None)
                self._config.pop("paired_warmup_batch_size", None)
                self._config.pop("paired_warmup_single_batch", None)
                self._config["n_initial_points"] = warmup_value
            self._gp_warmup_iterations_var.set(str(warmup_value))
            acquisition["candidate_pool_size"] = max(50, int(self._candidate_pool_var.get() or 600))
            acquisition["local_candidate_pool_size"] = max(0, int(self._local_pool_var.get() or 120))
            mode = str(self._initial_point_mode_var.get() or "specific").strip().lower()
            acquisition["initial_point_mode"] = "random" if mode == "random" else "specific"
            direction_var = getattr(self, "_optimization_direction_var", None)
            acquisition["optimization_direction"] = self._display_optimization_direction(
                direction_var.get() if direction_var is not None else "maximize"
            )
            falloff_fractions = self._gp_length_scales_from_vars()
            acquisition["gp_falloff_fractions"] = falloff_fractions
            acquisition["gp_length_scales"] = falloff_fractions
            self._refresh_gp_falloff_summary()
        except Exception as exc:
            if show_error:
                messagebox.showerror("Optimizer Behavior", str(exc))

    def _gp_length_scales_from_vars(self):
        raw = {name: (var.get() or "").strip() for name, var in self._gp_length_scale_vars.items()}
        filled = {name: text for name, text in raw.items() if text}
        if not filled:
            return {}
        missing = [name for name in PARAMETER_ORDER if not raw.get(name)]
        if missing:
            raise ValueError(
                "Fill every GP falloff fraction, or clear every field to let the GP learn them. "
                f"Missing: {', '.join(missing)}"
            )
        parsed = {}
        for name, text in raw.items():
            value = float(text)
            if value <= 0:
                raise ValueError(f"{name} GP falloff fraction must be > 0")
            parsed[name] = value
        return parsed

    def _edit_gp_length_scales(self):
        win = tk.Toplevel(self._frame)
        win.title("GP Correlation Falloff")
        win.transient(self._frame.winfo_toplevel())
        box = ttk.Frame(win, padding=12)
        box.pack(fill="both", expand=True)
        ttk.Label(
            box,
            text=(
                "Set fixed GP correlation falloff as a fraction of each parameter's search range. "
                "Example: 0.2 means about 20% of the search range, so a 500 Hz frequency range gives "
                "roughly a 100 Hz falloff. Larger values make the GP smoother; smaller values make "
                "correlation fall off faster. Clear all fields to let the GP learn them."
            ),
            foreground=self.ACCENT,
            wraplength=560,
            justify="left",
        ).grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        local_vars = {name: tk.StringVar(value=self._gp_length_scale_vars[name].get()) for name in PARAMETER_ORDER}
        for row, name in enumerate(PARAMETER_ORDER, start=1):
            ttk.Label(box, text=name.replace("_", " ").title() + ":").grid(row=row, column=0, sticky="w", pady=3)
            ttk.Entry(box, textvariable=local_vars[name], width=12).grid(row=row, column=1, sticky="w", padx=(8, 4), pady=3)
            ttk.Label(box, text="fraction of range; blank = learned", foreground="#666666").grid(row=row, column=2, sticky="w", pady=3)
        buttons = ttk.Frame(box)
        buttons.grid(row=len(PARAMETER_ORDER) + 1, column=0, columnspan=3, pady=(10, 0))

        def clear_all():
            for var in local_vars.values():
                var.set("")

        def save():
            old_values = {name: var.get() for name, var in self._gp_length_scale_vars.items()}
            try:
                for name, var in local_vars.items():
                    self._gp_length_scale_vars[name].set(var.get().strip())
                if self._config is not None:
                    self._sync_algorithm_config(show_error=False)
                self._refresh_gp_falloff_summary()
                win.destroy()
            except Exception as exc:
                for name, value in old_values.items():
                    self._gp_length_scale_vars[name].set(value)
                messagebox.showerror("GP Correlation Falloff", str(exc), parent=win)

        ttk.Button(buttons, text="Clear All", command=clear_all).pack(side="left", padx=4)
        ttk.Button(buttons, text="Save", command=save).pack(side="left", padx=4)
        ttk.Button(buttons, text="Cancel", command=win.destroy).pack(side="left", padx=4)
        win.grab_set()
        win.focus_force()

    def _refresh_gp_falloff_summary(self):
        if not hasattr(self, "_gp_falloff_summary_var"):
            return
        values = {name: (var.get() or "").strip() for name, var in self._gp_length_scale_vars.items()}
        filled = {name: value for name, value in values.items() if value}
        if not filled:
            self._gp_falloff_summary_var.set("GP falloff: learned by GP")
            return
        missing = [name for name in PARAMETER_ORDER if not values.get(name)]
        if missing:
            self._gp_falloff_summary_var.set(f"GP falloff: incomplete fixed values; missing {', '.join(missing)}")
            return
        summary = ", ".join(f"{name}={value}" for name, value in filled.items())
        self._gp_falloff_summary_var.set(f"GP falloff: fixed fractions of search range ({summary})")

    @staticmethod
    def _q_reference_text():
        return (
            "Metric definitions:\n"
            "  Peak uA = average measured signal height.\n"
            "  Peak prominence = average peak height / average RMS trace noise.\n"
            "  Repeat-scan SNR = average peak height / peak-height STD across repeat scans (0 for one scan).\n"
            "  Noise uA = RMS noise estimated from neighboring-point current differences / sqrt(2).\n\n"
            "Bounded components:\n"
            "  Shape = centered, stable peak quality.\n"
            "  Baseline = low and stable background quality.\n"
            "  Replicate = consistency of peak heights across repeat scans.\n"
            "  Success = fraction of measurements analyzed successfully; it is separate from Q.\n\n"
            "Run-level terms:\n"
            "  Run STD = variation in Q_channel across channels.\n"
            "  Repeat relative STD = repeat measurement variability.\n"
            "  Failed fraction = fraction of channels without a successful result.\n"
            "  Poor-channel fraction = fraction outside the direction-specific Q threshold.\n\n"
            "Modes:\n"
            "  Classic uses a direct weighted sum and does not normalize by summed weights.\n"
            "  Signal-priority uses log1p(signal metrics) and log1p(peak height), with weight normalization."
        )

    def _setup_scoring_vars(self):
        return {
            "mode": self._score_mode_var,
            "peak_prominence": self._score_snr_weight_var,
            "repeat_scan_snr": self._score_repeat_scan_snr_weight_var,
            "peak_height": self._score_peak_height_weight_var,
            "peak_shape": self._score_shape_weight_var,
            "baseline": self._score_baseline_weight_var,
            "replicate_consistency": self._score_replicate_weight_var,
            "success": self._score_success_weight_var,
            "noise_penalty": self._score_noise_penalty_var,
            "peak_prominence_saturation": self._score_snr_saturation_var,
            "lambda_variability": self._score_variability_penalty_var,
            "lambda_repeat_std": self._score_repeat_std_penalty_var,
            "lambda_failed": self._score_failed_penalty_var,
            "lambda_low": self._score_low_penalty_var,
            "low_channel_threshold": self._score_low_threshold_var,
        }

    def _paired_scoring_vars(self):
        return {
            "buffer_classic_Q": self._paired_buffer_classic_q_weight_var,
            "target_classic_Q": self._paired_target_classic_q_weight_var,
            "peak_prominence": self._paired_delta_peak_weight_var,
            "repeat_scan_snr": self._paired_repeat_scan_snr_weight_var,
            "lambda_repeat_std": self._paired_repeat_std_penalty_var,
        }

    @staticmethod
    def _default_paired_response_weights() -> dict:
        return {
            "buffer_classic_Q": 0.25,
            "target_classic_Q": 0.25,
            "peak_prominence": 1.0,
            "repeat_scan_snr": 0.0,
            "lambda_repeat_std": 0.0,
        }

    @staticmethod
    def _display_score_mode(mode) -> str:
        mode_text = str(mode or "classic").strip().lower()
        return "signal_priority_unbounded" if mode_text == "signal_priority_unbounded" else "classic"

    @staticmethod
    def _display_optimization_direction(direction) -> str:
        text = str(direction or "maximize").strip().lower()
        if text in {"minimize", "min", "more_negative", "negative"}:
            return "minimize"
        if text in {"survey", "either", "absolute", "magnitude"}:
            return "survey"
        return "maximize"

    def _optimization_objective_value(self, q_run, config=None) -> float:
        direction = self._display_optimization_direction(
            dict((config or self._config or {}).get("acquisition") or {}).get("optimization_direction", "maximize")
        )
        value = float(q_run or 0.0)
        if direction == "minimize":
            return -value
        if direction == "survey":
            return abs(value)
        return value

    def _rescore_scoring_vars(self):
        return {
            "mode": self._rescore_mode_var,
            "peak_prominence": self._rescore_snr_weight_var,
            "repeat_scan_snr": self._rescore_repeat_scan_snr_weight_var,
            "peak_height": self._rescore_peak_height_weight_var,
            "peak_shape": self._rescore_shape_weight_var,
            "baseline": self._rescore_baseline_weight_var,
            "replicate_consistency": self._rescore_replicate_weight_var,
            "success": self._rescore_success_weight_var,
            "noise_penalty": self._rescore_noise_penalty_var,
            "peak_prominence_saturation": self._rescore_snr_saturation_var,
            "lambda_variability": self._rescore_variability_penalty_var,
            "lambda_repeat_std": self._rescore_repeat_std_penalty_var,
            "lambda_failed": self._rescore_failed_penalty_var,
            "lambda_low": self._rescore_low_penalty_var,
            "low_channel_threshold": self._rescore_low_threshold_var,
        }

    def _rescore_paired_scoring_vars(self):
        return self._rescore_paired_score_vars

    def _build_reanalysis_controls(self, parent):
        variables = self._rescore_analysis_vars
        entries = (
            ("Crop min V", "crop_min_v", "Crop max V", "crop_max_v"),
            ("Smooth window", "smooth_window", "Polynomial order", "smooth_polyorder"),
            ("Minima window V", "minima_search_window_v", "Min peak height uA", "min_peak_height_ua"),
            ("Peak V min", "peak_voltage_min_v", "Peak V max", "peak_voltage_max_v"),
            ("Left min V min", "left_min_voltage_min_v", "Left min V max", "left_min_voltage_max_v"),
            ("Right min V min", "right_min_voltage_min_v", "Right min V max", "right_min_voltage_max_v"),
            ("Min start V", "min_start_voltage_v", "Scan windows", "scan_windows"),
        )
        for row, (left_label, left_key, right_label, right_key) in enumerate(entries):
            ttk.Label(parent, text=f"{left_label}:").grid(row=row, column=0, sticky="w", pady=2)
            ttk.Entry(parent, textvariable=variables[left_key], width=13).grid(
                row=row, column=1, sticky="ew", padx=(4, 12), pady=2
            )
            ttk.Label(parent, text=f"{right_label}:").grid(row=row, column=2, sticky="w", pady=2)
            ttk.Entry(parent, textvariable=variables[right_key], width=16).grid(
                row=row, column=3, sticky="ew", padx=4, pady=2
            )
        checks = (
            ("Prominent minima", "use_prominent_minima"),
            ("Require minima both sides", "require_local_minima_on_both_sides"),
            ("Double correction", "use_double_correction"),
            ("Compute skew", "compute_skew"),
            ("Wavelet energy", "compute_wavelet_energy"),
            ("Wavelet trace", "compute_wavelet_denoised_trace"),
            ("Wavelet correction", "use_wavelet_for_correction"),
        )
        for index, (label, key) in enumerate(checks):
            row = 7 + index // 2
            column = (index % 2) * 2
            ttk.Checkbutton(parent, text=label, variable=variables[key]).grid(
                row=row, column=column, columnspan=2, sticky="w", pady=2
            )
        parent.columnconfigure(1, weight=1)
        parent.columnconfigure(3, weight=1)

    def _set_reanalysis_vars(self, analysis):
        values = {
            "crop_min_v": analysis.get("crop_min_v", -0.61),
            "crop_max_v": analysis.get("crop_max_v", -0.30),
            "smooth_window": analysis.get("smooth_window", 15),
            "smooth_polyorder": analysis.get("smooth_polyorder", 2),
            "minima_search_window_v": analysis.get("minima_search_window_v", 0.30),
            "min_peak_height_ua": analysis.get("min_peak_height_ua", 0.001),
            "peak_voltage_min_v": analysis.get("peak_voltage_min_v"),
            "peak_voltage_max_v": analysis.get("peak_voltage_max_v"),
            "left_min_voltage_min_v": analysis.get("left_min_voltage_min_v"),
            "left_min_voltage_max_v": analysis.get("left_min_voltage_max_v"),
            "right_min_voltage_min_v": analysis.get("right_min_voltage_min_v"),
            "right_min_voltage_max_v": analysis.get("right_min_voltage_max_v"),
            "min_start_voltage_v": analysis.get("min_start_voltage_v", -0.70),
            "scan_windows": analysis.get("scan_windows", ""),
        }
        for key, value in values.items():
            self._rescore_analysis_vars[key].set("" if value is None else str(value))
        for key in (
            "use_prominent_minima",
            "require_local_minima_on_both_sides",
            "use_double_correction",
            "compute_skew",
            "compute_wavelet_energy",
            "compute_wavelet_denoised_trace",
            "use_wavelet_for_correction",
        ):
            self._rescore_analysis_vars[key].set(bool(analysis.get(key, False)))

    def _reanalysis_config(self):
        variables = self._rescore_analysis_vars

        def optional_float(key):
            variable = variables.get(key)
            if variable is None:
                return None
            text = (variable.get() or "").strip()
            return None if not text else float(text)

        analysis = {
            "crop_min_v": float(variables["crop_min_v"].get()),
            "crop_max_v": float(variables["crop_max_v"].get()),
            "smooth_window": int(variables["smooth_window"].get()),
            "smooth_polyorder": int(variables["smooth_polyorder"].get()),
            "minima_search_window_v": float(variables["minima_search_window_v"].get()),
            "min_peak_height_ua": optional_float("min_peak_height_ua"),
            "peak_voltage_min_v": optional_float("peak_voltage_min_v"),
            "peak_voltage_max_v": optional_float("peak_voltage_max_v"),
            "left_min_voltage_min_v": optional_float("left_min_voltage_min_v"),
            "left_min_voltage_max_v": optional_float("left_min_voltage_max_v"),
            "right_min_voltage_min_v": optional_float("right_min_voltage_min_v"),
            "right_min_voltage_max_v": optional_float("right_min_voltage_max_v"),
            "min_start_voltage_v": float(variables["min_start_voltage_v"].get()),
            "scan_windows": (variables["scan_windows"].get() or "").strip(),
            "use_prominent_minima": bool(variables["use_prominent_minima"].get()),
            "require_local_minima_on_both_sides": bool(variables["require_local_minima_on_both_sides"].get()),
            "use_double_correction": bool(variables["use_double_correction"].get()),
            "compute_skew": bool(variables["compute_skew"].get()),
            "compute_wavelet_energy": bool(variables["compute_wavelet_energy"].get()),
            "compute_wavelet_denoised_trace": bool(variables["compute_wavelet_denoised_trace"].get()),
            "use_wavelet_for_correction": bool(variables["use_wavelet_for_correction"].get()),
        }
        if analysis["crop_min_v"] >= analysis["crop_max_v"]:
            raise ValueError("Reanalysis crop minimum must be below crop maximum.")
        if analysis["smooth_window"] < 0 or analysis["smooth_polyorder"] < 0:
            raise ValueError("Reanalysis smoothing window and polynomial order must be nonnegative.")
        if analysis["minima_search_window_v"] <= 0:
            raise ValueError("Reanalysis minima window must be positive.")
        self._validate_voltage_window(analysis, "peak_voltage_min_v", "peak_voltage_max_v", "Reanalysis peak V")
        self._validate_voltage_window(analysis, "left_min_voltage_min_v", "left_min_voltage_max_v", "Reanalysis left min V")
        self._validate_voltage_window(analysis, "right_min_voltage_min_v", "right_min_voltage_max_v", "Reanalysis right min V")
        return analysis

    def _build_q_scoring_controls(self, scoring_box, vars_by_name, formula_var, on_change, preset_command=None):
        def refresh_explanation():
            self._refresh_formula_from_vars(vars_by_name, formula_var)
            on_change()

        for idx in range(6):
            scoring_box.columnconfigure(idx, weight=1 if idx in (1, 3, 5) else 0)
        ttk.Label(scoring_box, text="Score mode:").grid(row=0, column=0, sticky="w", pady=2)
        mode_combo = ttk.Combobox(
            scoring_box,
            textvariable=vars_by_name["mode"],
            values=("classic", "signal_priority_unbounded"),
            state="readonly",
            width=26,
        )
        mode_combo.grid(row=0, column=1, columnspan=2, sticky="w", padx=(4, 10), pady=2)
        mode_combo.bind("<<ComboboxSelected>>", lambda _e: on_change())
        if preset_command is not None:
            ttk.Button(scoring_box, text="Signal Preset", command=preset_command).grid(row=0, column=3, sticky="w", pady=2)
        ttk.Button(
            scoring_box,
            text="Refresh Scoring Explanation",
            command=refresh_explanation,
        ).grid(row=0, column=4, columnspan=2, sticky="e", pady=2)
        entries = [
            ("Channel peak prominence weight:", vars_by_name["peak_prominence"]),
            ("Channel repeat-scan SNR weight:", vars_by_name["repeat_scan_snr"]),
            ("Channel peak weight:", vars_by_name["peak_height"]),
            ("Channel shape weight:", vars_by_name["peak_shape"]),
            ("Channel baseline weight:", vars_by_name["baseline"]),
            ("Channel replicate weight:", vars_by_name["replicate_consistency"]),
            ("Channel success weight:", vars_by_name["success"]),
            ("Channel noise penalty:", vars_by_name["noise_penalty"]),
            ("Peak prominence saturation:", vars_by_name["peak_prominence_saturation"]),
            ("Run std penalty:", vars_by_name["lambda_variability"]),
            ("Classic run repeat relative-std penalty:", vars_by_name["lambda_repeat_std"]),
            ("Run failed penalty:", vars_by_name["lambda_failed"]),
            ("Run poor-channel penalty:", vars_by_name["lambda_low"]),
            ("Poor-channel threshold:", vars_by_name["low_channel_threshold"]),
        ]
        for idx, (label, var) in enumerate(entries):
            row = 1 + idx // 2
            base_col = (idx % 2) * 3
            ttk.Label(scoring_box, text=label).grid(row=row, column=base_col, sticky="w", pady=2)
            entry = ttk.Entry(scoring_box, textvariable=var, width=9)
            entry.grid(row=row, column=base_col + 1, sticky="w", padx=(4, 10), pady=2)
            entry.bind("<FocusOut>", lambda _e: on_change())
            entry.bind("<Return>", lambda _e: on_change())
        formula_row = 2 + (len(entries) - 1) // 2
        ttk.Label(scoring_box, textvariable=formula_var, foreground=self.ACCENT, wraplength=460, justify="left").grid(
            row=formula_row,
            column=0,
            columnspan=6,
            sticky="w",
            pady=(4, 0),
        )

    def _build_paired_q_scoring_controls(
        self,
        scoring_box,
        vars_by_name=None,
        formula_var=None,
        on_change=None,
    ):
        vars_by_name = vars_by_name or self._paired_scoring_vars()
        formula_var = formula_var or self._paired_formula_var
        on_change = on_change or (lambda: None)

        def refresh_explanation():
            try:
                formula_var.set(
                    self._paired_formula_text(
                        self._paired_scoring_from_var_map(vars_by_name)
                    )
                )
            except Exception:
                formula_var.set(self._paired_formula_fallback_text())
            on_change()

        for idx in range(4):
            scoring_box.columnconfigure(idx, weight=1 if idx in (1, 3) else 0)
        entries = (
            ("Buffer classic Q weight:", "buffer_classic_Q"),
            ("Target classic Q weight:", "target_classic_Q"),
            ("Repeat-scan SNR weight:", "repeat_scan_snr"),
            ("Repeat-scan peak prominence weight:", "peak_prominence"),
            ("Paired run repeat relative-std penalty:", "lambda_repeat_std"),
        )
        for idx, (label, key) in enumerate(entries):
            row = idx // 2
            base_col = (idx % 2) * 2
            ttk.Label(scoring_box, text=label).grid(row=row, column=base_col, sticky="w", pady=2)
            entry = ttk.Entry(scoring_box, textvariable=vars_by_name[key], width=9)
            entry.grid(row=row, column=base_col + 1, sticky="w", padx=(4, 12), pady=2)
            entry.bind("<FocusOut>", lambda _e: on_change())
            entry.bind("<Return>", lambda _e: on_change())
        formula_row = 1 + (len(entries) - 1) // 2
        ttk.Button(
            scoring_box,
            text="Refresh Scoring Explanation",
            command=refresh_explanation,
        ).grid(row=formula_row, column=0, columnspan=4, sticky="w", pady=(6, 2))
        ttk.Label(
            scoring_box,
            textvariable=formula_var,
            foreground=self.ACCENT,
            wraplength=760,
            justify="left",
        ).grid(row=formula_row + 1, column=0, columnspan=4, sticky="w", pady=(6, 8))

    def _set_scoring_vars_from_config(self, cfg: dict):
        scoring = dict((cfg or {}).get("scoring") or {})
        self._bo_objective_var.set("paired_response" if str((cfg or {}).get("objective") or "").lower() == "paired_response" else "quality")
        self._score_mode_var.set(self._display_score_mode(scoring.get("mode", "classic")))
        channel = dict(scoring.get("channel_weights") or {})
        run = dict(scoring.get("run_weights") or {})
        self._score_snr_weight_var.set(str(channel.get("peak_prominence", channel.get("snr", 0.35))))
        self._score_repeat_scan_snr_weight_var.set(str(channel.get("repeat_scan_snr", 0.0)))
        self._score_peak_height_weight_var.set(str(channel.get("peak_height", 0.0)))
        self._score_shape_weight_var.set(str(channel.get("peak_shape", 0.20)))
        self._score_baseline_weight_var.set(str(channel.get("baseline", 0.20)))
        self._score_replicate_weight_var.set(str(channel.get("replicate_consistency", 0.15)))
        self._score_success_weight_var.set(str(channel.get("success", 0.10)))
        self._score_noise_penalty_var.set(str(channel.get("noise_penalty", 0.0)))
        self._score_snr_saturation_var.set(str(channel.get("peak_prominence_saturation", channel.get("snr_saturation", 20.0))))
        self._score_variability_penalty_var.set(str(run.get("lambda_variability", 0.20)))
        self._score_repeat_std_penalty_var.set(str(run.get("lambda_repeat_std", 0.0)))
        self._score_failed_penalty_var.set(str(run.get("lambda_failed", 0.40)))
        self._score_low_penalty_var.set(str(run.get("lambda_low", 0.20)))
        self._score_low_threshold_var.set(str(run.get("low_channel_threshold", 0.50)))
        self._set_paired_scoring_vars(scoring, self._paired_scoring_vars())
        self._refresh_score_formula()
        self._refresh_paired_score_formula()
        self._on_bo_type_changed(sync=False)

    def _set_paired_scoring_vars(self, scoring: dict, vars_by_name: dict):
        paired = self._default_paired_response_weights()
        saved_paired = dict((scoring or {}).get("paired_response_weights") or {})
        if "standard_quality" in saved_paired and "buffer_classic_Q" not in saved_paired and "target_classic_Q" not in saved_paired:
            legacy_quality_weight = max(0.0, float(saved_paired.get("standard_quality", 0.0) or 0.0))
            saved_paired["buffer_classic_Q"] = legacy_quality_weight / 2.0
            saved_paired["target_classic_Q"] = legacy_quality_weight / 2.0
        if "peak_prominence" not in saved_paired and "delta_peak" in saved_paired:
            saved_paired["peak_prominence"] = saved_paired["delta_peak"]
        paired.update(saved_paired)
        for key, var in vars_by_name.items():
            var.set(str(paired.get(key, self._default_paired_response_weights().get(key, 0.0))))

    def _sync_scoring_config(self, show_error=True):
        if self._config is None:
            return
        try:
            scoring = self._scoring_from_vars(self._setup_scoring_vars())
            scoring["paired_response_weights"] = self._paired_scoring_from_vars()
            self._config["scoring"] = scoring
            self._config["objective"] = self._bo_objective_var.get()
            self._refresh_score_formula()
            self._refresh_paired_score_formula()
            self._refresh_current_q_equation(self._config)
        except Exception as exc:
            if show_error:
                messagebox.showerror("Q Score Decomposition", str(exc))

    def _refresh_score_formula(self):
        self._refresh_formula_from_vars(self._setup_scoring_vars(), self._score_formula_var)
        if self._bo_objective_var.get() == "paired_response":
            self._normal_scoring_frame.configure(text="Classic Trace Q Decomposition")
        elif hasattr(self, "_normal_scoring_frame"):
            self._normal_scoring_frame.configure(text="Q Score Decomposition")

    def _paired_scoring_from_vars(self):
        return self._paired_scoring_from_var_map(self._paired_scoring_vars())

    def _paired_scoring_from_var_map(self, vars_by_name):
        return {
            key: max(0.0, float(var.get() or 0.0))
            for key, var in vars_by_name.items()
        }

    def _refresh_paired_score_formula(self):
        try:
            self._paired_formula_var.set(
                self._paired_formula_text(self._paired_scoring_from_vars())
            )
        except Exception:
            self._paired_formula_var.set(self._paired_formula_fallback_text())

    @staticmethod
    def _paired_formula_text(weights=None):
        weights = dict(weights or {})
        terms = []
        definitions = [
            "  delta_peak = average target peak - average buffer peak."
        ]
        if float(weights.get("repeat_scan_snr", 0.0) or 0.0) != 0.0:
            weight = float(weights["repeat_scan_snr"])
            terms.append(f"  + {weight:g}*repeat_scan_SNR")
            definitions.append(
                "  repeat_scan_SNR = delta_peak / (buffer peak STD + target peak STD)."
            )
        if float(weights.get("peak_prominence", 0.0) or 0.0) != 0.0:
            weight = float(weights["peak_prominence"])
            terms.append(f"  + {weight:g}*peak_prominence")
            definitions.append(
                "  peak_prominence = delta_peak / (average buffer RMS + average target RMS)."
            )
        if float(weights.get("buffer_classic_Q", 0.0) or 0.0) != 0.0:
            weight = float(weights["buffer_classic_Q"])
            terms.append(f"  + sign(delta_peak)*{weight:g}*buffer_classic_Q")
            definitions.append(
                "  buffer_classic_Q = classic trace Q for the averaged buffer measurements."
            )
        if float(weights.get("target_classic_Q", 0.0) or 0.0) != 0.0:
            weight = float(weights["target_classic_Q"])
            terms.append(f"  + sign(delta_peak)*{weight:g}*target_classic_Q")
            definitions.append(
                "  target_classic_Q = classic trace Q for the averaged target measurements."
            )

        if terms:
            terms[0] = terms[0].replace("  + ", "  ", 1)
            channel_text = "paired_Q_channel =\n" + "\n".join(terms)
        else:
            channel_text = "paired_Q_channel = 0 (no paired weights are active)"

        sections = [channel_text]
        if terms:
            sections.append("Active paired Q terms:\n" + "\n".join(definitions))
        repeat_penalty = float(weights.get("lambda_repeat_std", 0.0) or 0.0)
        if repeat_penalty != 0.0:
            sections.append(
                "Active paired Q_run penalty:\n"
                f"  {repeat_penalty:g}*mean repeat relative STD"
            )
        sections.append(
            "Maximize/minimize clips the undesired Q_run sign to 0.\n"
            "Survey preserves both signs and optimizes |Q_run|."
        )
        return "\n\n".join(sections)

    @staticmethod
    def _paired_formula_fallback_text():
        return "Enter numeric paired weights, then refresh the scoring explanation."

    def _on_bo_type_changed(self, sync=True):
        paired = self._bo_objective_var.get() == "paired_response"
        if hasattr(self, "_paired_behavior_frame"):
            if paired:
                self._paired_behavior_frame.pack(fill="x", pady=(0, 8))
            else:
                self._paired_behavior_frame.pack_forget()
        if hasattr(self, "_optimizer_behavior_frame"):
            self._optimizer_behavior_frame.configure(
                text="Optimizer Behavior by Group (paired cycles)"
                if paired else "Optimizer Behavior by Group"
            )
        if hasattr(self, "_paired_scoring_frame") and hasattr(self, "_normal_scoring_frame"):
            if paired:
                self._paired_scoring_frame.pack(fill="both", expand=True, pady=(0, 8), padx=2)
                self._normal_scoring_frame.configure(text="Classic Trace Q Decomposition")
            else:
                self._paired_scoring_frame.pack_forget()
                self._normal_scoring_frame.configure(text="Q Score Decomposition")
        if self._config is not None:
            self._config["objective"] = "paired_response" if paired else "quality"
        self._refresh_score_formula()
        self._refresh_paired_score_formula()
        if sync:
            self._sync_scoring_config(show_error=False)

    def _scoring_from_vars(self, vars_by_name):
        mode = str(vars_by_name["mode"].get() or "classic").strip().lower()
        prominence_var = vars_by_name.get("peak_prominence") or vars_by_name.get("snr")
        repeat_scan_snr_var = vars_by_name.get("repeat_scan_snr")
        prominence_saturation_var = (
            vars_by_name.get("peak_prominence_saturation")
            or vars_by_name.get("snr_saturation")
        )
        repeat_std_var = vars_by_name.get("lambda_repeat_std")
        repeat_std_weight = (
            max(0.0, float(repeat_std_var.get() or 0.0))
            if repeat_std_var is not None
            else 0.0
        )
        return {
            "mode": "signal_priority_unbounded" if mode == "signal_priority_unbounded" else "classic",
            "channel_weights": {
                "peak_prominence": max(0.0, float(prominence_var.get() or 0.0)),
                "repeat_scan_snr": max(0.0, float(repeat_scan_snr_var.get() or 0.0)) if repeat_scan_snr_var is not None else 0.0,
                "peak_height": max(0.0, float(vars_by_name["peak_height"].get() or 0.0)),
                "peak_shape": max(0.0, float(vars_by_name["peak_shape"].get() or 0.0)),
                "baseline": max(0.0, float(vars_by_name["baseline"].get() or 0.0)),
                "replicate_consistency": max(0.0, float(vars_by_name["replicate_consistency"].get() or 0.0)),
                "success": max(0.0, float(vars_by_name["success"].get() or 0.0)),
                "noise_penalty": max(0.0, float(vars_by_name["noise_penalty"].get() or 0.0)),
                "peak_prominence_saturation": max(1e-12, float(prominence_saturation_var.get() or 20.0)),
            },
            "run_weights": {
                "lambda_variability": max(0.0, float(vars_by_name["lambda_variability"].get() or 0.0)),
                "lambda_repeat_std": repeat_std_weight,
                "lambda_failed": max(0.0, float(vars_by_name["lambda_failed"].get() or 0.0)),
                "lambda_low": max(0.0, float(vars_by_name["lambda_low"].get() or 0.0)),
                "low_channel_threshold": max(0.0, min(1.0, float(vars_by_name["low_channel_threshold"].get() or 0.5))),
            },
        }

    def _refresh_formula_from_vars(self, vars_by_name, formula_var):
        try:
            mode = str(vars_by_name["mode"].get() or "classic").strip().lower()
            weights = {
                key: float(vars_by_name[key].get() or 0.0)
                for key in (
                    "peak_prominence",
                    "repeat_scan_snr",
                    "peak_height",
                    "peak_shape",
                    "baseline",
                    "replicate_consistency",
                    "success",
                )
            }
            noise_penalty = float(vars_by_name["noise_penalty"].get() or 0.0)
            run_weights = {
                key: float(vars_by_name[key].get() or 0.0)
                for key in (
                    "lambda_variability",
                    "lambda_repeat_std",
                    "lambda_failed",
                    "lambda_low",
                )
                if key in vars_by_name
            }
            poor_threshold = float(
                vars_by_name["low_channel_threshold"].get() or 0.5
            )
        except Exception:
            formula_var.set("Q_channel = weighted component score. Enter numeric weights.")
            return

        metric_info = {
            "peak_prominence": ("Peak prominence", "average peak height / average RMS trace noise"),
            "repeat_scan_snr": ("Repeat-scan SNR", "average peak height / repeat peak-height STD"),
            "peak_height": ("Peak uA", "average measured peak height"),
            "peak_shape": ("Shape", "centered, stable peak quality"),
            "baseline": ("Baseline", "low and stable background quality"),
            "replicate_consistency": ("Replicate", "repeat peak-height consistency"),
            "success": ("Success", "fraction of measurements analyzed successfully"),
        }
        channel_terms = []
        definition_lines = []
        for key, weight in weights.items():
            if weight == 0.0:
                continue
            label, definition = metric_info[key]
            expression = label
            if mode == "signal_priority_unbounded" and key in {
                "peak_prominence", "repeat_scan_snr", "peak_height"
            }:
                expression = f"log1p({label})"
            channel_terms.append(f"  + {weight:g}*{expression}")
            definition_lines.append(f"  {label} = {definition}.")

        if channel_terms:
            channel_terms[0] = channel_terms[0].replace("  + ", "  ", 1)
            if mode == "signal_priority_unbounded":
                total = sum(weight for weight in weights.values() if weight != 0.0)
                channel_text = "Q_channel = (\n" + "\n".join(channel_terms) + f"\n) / {total:g}"
            else:
                channel_text = "Q_channel =\n" + "\n".join(channel_terms)
        else:
            channel_text = "Q_channel = 0 (no channel weights are active)"

        if mode != "signal_priority_unbounded" and noise_penalty != 0.0:
            channel_text += f"\n  -/+ {noise_penalty:g}*Noise uA"
            definition_lines.append(
                "  Noise uA = RMS noise estimated from neighboring-point current differences / sqrt(2)."
            )

        penalty_info = (
            ("lambda_variability", "std(Q_channel)"),
            ("lambda_repeat_std", "mean repeat relative STD"),
            ("lambda_failed", "failed-channel fraction"),
            (
                "lambda_low",
                f"poor-channel fraction (threshold {poor_threshold:g})",
            ),
        )
        penalty_lines = [
            f"  {run_weights[key]:g}*{label}"
            for key, label in penalty_info
            if run_weights.get(key, 0.0) != 0.0
        ]

        sections = [channel_text]
        if definition_lines:
            sections.append("Active Q terms:\n" + "\n".join(definition_lines))
        if penalty_lines:
            sections.append("Active Q_run penalties:\n" + "\n".join(penalty_lines))
        sections.append(
            "Maximize/minimize clips the undesired Q_run sign to 0.\n"
            "Survey retains both signs and optimizes |Q_run|."
        )
        if mode != "signal_priority_unbounded" and noise_penalty != 0.0:
            sections.append(
                "Noise penalty is subtracted for maximize/survey and added for minimize."
            )
        formula_var.set("\n\n".join(sections))

    def _apply_signal_priority_preset(self):
        self._score_mode_var.set("signal_priority_unbounded")
        self._score_snr_weight_var.set("0.45")
        self._score_repeat_scan_snr_weight_var.set("0.00")
        self._score_peak_height_weight_var.set("0.35")
        self._score_shape_weight_var.set("0.05")
        self._score_baseline_weight_var.set("0.12")
        self._score_replicate_weight_var.set("0.03")
        self._score_success_weight_var.set("0.00")
        self._score_noise_penalty_var.set("0.00")
        self._score_snr_saturation_var.set("20.0")
        self._score_variability_penalty_var.set("0.10")
        self._score_failed_penalty_var.set("0.10")
        self._score_low_penalty_var.set("0.05")
        self._score_low_threshold_var.set("1.50")
        self._sync_scoring_config(show_error=False)

    def _set_scoring_vars(self, cfg, vars_by_name, formula_var):
        scoring = dict((cfg or {}).get("scoring") or {})
        channel = dict(scoring.get("channel_weights") or {})
        run = dict(scoring.get("run_weights") or {})
        vars_by_name["mode"].set(self._display_score_mode(scoring.get("mode", "classic")))
        vars_by_name["peak_prominence"].set(str(channel.get("peak_prominence", channel.get("snr", 0.35))))
        vars_by_name["repeat_scan_snr"].set(str(channel.get("repeat_scan_snr", 0.0)))
        vars_by_name["peak_height"].set(str(channel.get("peak_height", 0.0)))
        vars_by_name["peak_shape"].set(str(channel.get("peak_shape", 0.20)))
        vars_by_name["baseline"].set(str(channel.get("baseline", 0.20)))
        vars_by_name["replicate_consistency"].set(str(channel.get("replicate_consistency", 0.15)))
        vars_by_name["success"].set(str(channel.get("success", 0.10)))
        vars_by_name["noise_penalty"].set(str(channel.get("noise_penalty", 0.0)))
        vars_by_name["peak_prominence_saturation"].set(str(channel.get("peak_prominence_saturation", channel.get("snr_saturation", 20.0))))
        vars_by_name["lambda_variability"].set(str(run.get("lambda_variability", 0.20)))
        if "lambda_repeat_std" in vars_by_name:
            vars_by_name["lambda_repeat_std"].set(str(run.get("lambda_repeat_std", 0.0)))
        vars_by_name["lambda_failed"].set(str(run.get("lambda_failed", 0.40)))
        vars_by_name["lambda_low"].set(str(run.get("lambda_low", 0.20)))
        vars_by_name["low_channel_threshold"].set(str(run.get("low_channel_threshold", 0.50)))
        self._refresh_formula_from_vars(vars_by_name, formula_var)

    def _set_rescore_vars_from_config(self, cfg):
        scoring = dict((cfg or {}).get("scoring") or {})
        self._set_scoring_vars(cfg, self._rescore_scoring_vars(), self._rescore_formula_var)
        self._set_paired_scoring_vars(scoring, self._rescore_paired_scoring_vars())
        self._rescore_paired_formula_var.set(
            self._paired_formula_text(
                self._paired_scoring_from_var_map(
                    self._rescore_paired_scoring_vars()
                )
            )
        )
        self._set_reanalysis_vars(dict((cfg or {}).get("analysis") or {}))
        self._refresh_current_q_equation(cfg)

    def _preview_rescore_equation(self):
        try:
            scoring = self._scoring_from_vars(self._rescore_scoring_vars())
            scoring["paired_response_weights"] = self._paired_scoring_from_var_map(
                self._rescore_paired_scoring_vars()
            )
            self._refresh_formula_from_vars(self._rescore_scoring_vars(), self._rescore_formula_var)
            self._rescore_paired_formula_var.set(
                self._paired_formula_text(
                    self._paired_scoring_from_var_map(
                        self._rescore_paired_scoring_vars()
                    )
                )
            )
            self._refresh_current_q_equation(self._config_with_scoring(scoring))
            if self._bo_session is not None:
                self._rescore_status_var.set("Edited scoring values. Click Apply Rescore to update recorded data.")
        except Exception as exc:
            self._rescore_status_var.set(f"Q equation preview failed: {exc}")

    def _apply_rescore_signal_priority_preset(self):
        self._rescore_mode_var.set("signal_priority_unbounded")
        self._rescore_snr_weight_var.set("0.45")
        self._rescore_repeat_scan_snr_weight_var.set("0.00")
        self._rescore_peak_height_weight_var.set("0.35")
        self._rescore_shape_weight_var.set("0.05")
        self._rescore_baseline_weight_var.set("0.12")
        self._rescore_replicate_weight_var.set("0.03")
        self._rescore_success_weight_var.set("0.00")
        self._rescore_noise_penalty_var.set("0.00")
        self._rescore_snr_saturation_var.set("20.0")
        self._rescore_variability_penalty_var.set("0.10")
        self._rescore_failed_penalty_var.set("0.10")
        self._rescore_low_penalty_var.set("0.05")
        self._rescore_low_threshold_var.set("1.50")
        self._preview_rescore_equation()

    def _apply_rescore_to_loaded_session(self, show_error=True):
        try:
            scoring = self._scoring_from_vars(self._rescore_scoring_vars())
            scoring["paired_response_weights"] = self._paired_scoring_from_var_map(
                self._rescore_paired_scoring_vars()
            )
            self._refresh_formula_from_vars(self._rescore_scoring_vars(), self._rescore_formula_var)
            self._rescore_paired_formula_var.set(
                self._paired_formula_text(scoring["paired_response_weights"])
            )
            self._refresh_current_q_equation(self._config_with_scoring(scoring))
            if self._bo_session is None:
                self._rescore_status_var.set("Load a BO session to rescore recorded data.")
                return
            self._bo_session.config["scoring"] = scoring
            if self._config is not None:
                self._config["scoring"] = dict(scoring)
            rescored = 0
            rebuilt_metrics = 0
            for obs in self._bo_session.observations:
                optimization_direction = self._bo_session._group_optimization_direction(
                    int(obs.get("group_id", 1) or 1)
                )
                if self._is_paired_observation(obs):
                    buffer_metrics = obs.get("buffer_channel_metrics")
                    target_metrics = obs.get("target_channel_metrics")
                    if not isinstance(buffer_metrics, dict) or not isinstance(target_metrics, dict):
                        continue
                    quality = compute_paired_response_quality(
                        buffer_metrics,
                        target_metrics,
                        scoring,
                        optimization_direction,
                    )
                else:
                    channel_metrics = self._rebuilt_channel_metrics_for_observation(obs)
                    if isinstance(channel_metrics, dict):
                        obs["channel_metrics"] = channel_metrics
                        rebuilt_metrics += 1
                    else:
                        channel_metrics = obs.get("channel_metrics")
                    if not isinstance(channel_metrics, dict):
                        continue
                    quality = compute_run_quality(
                        channel_metrics,
                        scoring,
                        optimization_direction,
                    )
                obs["quality"] = quality
                obs["Q_run"] = quality["Q_run"]
                for record in self._bo_session.suggestions:
                    if record.get("method_id") == obs.get("method_id"):
                        record["Q_run"] = obs["Q_run"]
                rescored += 1
            selected = None
            if hasattr(self, "_history_tree"):
                selection = self._history_tree.selection()
                selected = selection[0] if selection else None
            self._refresh_history()
            if selected in self._history_rows:
                self._history_tree.selection_set(selected)
                self._history_tree.focus(selected)
                self._select_history_iteration(selected)
            else:
                self._select_latest_history_iteration()
            self._render_best()
            self._rescore_status_var.set(
                f"Preview rescored {rescored} completed iteration(s); rebuilt metrics for {rebuilt_metrics}. "
                "Use Save Rescored Session to persist."
            )
        except Exception as exc:
            self._rescore_status_var.set(f"Rescore failed: {exc}")
            if show_error:
                messagebox.showerror("Rescore Q Scores", str(exc))

    def _reanalyze_and_rescore_loaded_session(self):
        if self._bo_session is None:
            self._rescore_status_var.set("Load a BO session before reanalyzing.")
            return
        if getattr(self, "_reanalyze_rescore_running", False):
            return
        try:
            scoring = self._scoring_from_vars(self._rescore_scoring_vars())
            scoring["paired_response_weights"] = self._paired_scoring_from_var_map(
                self._rescore_paired_scoring_vars()
            )
            analysis = self._reanalysis_config()
        except Exception as exc:
            messagebox.showerror("Reanalyze & Rescore", str(exc))
            return
        observations = list(self._bo_session.observations)
        if not observations:
            self._rescore_status_var.set("The loaded BO session has no observations.")
            return
        self._reanalyze_rescore_running = True
        self._reanalyze_rescore_button.configure(state="disabled")
        output_dir = (
            Path(self._bo_session.record_dir)
            / "reanalysis"
            / datetime.now().strftime("%Y%m%d_%H%M%S")
        )
        self._rescore_status_var.set(
            f"Reanalyzing 0/{len(observations)} iteration(s) with the external worker..."
        )
        selection = self._history_tree.selection() if hasattr(self, "_history_tree") else ()
        selected_iteration = selection[0] if selection else None

        def update_progress(done, iteration):
            self._frame.after(
                0,
                lambda: self._rescore_status_var.set(
                    f"Reanalyzing {done}/{len(observations)} iteration(s); completed iteration {iteration}."
                ),
            )

        def failed_reanalysis_update(observation, exc):
            optimization_direction = self._bo_session._group_optimization_direction(
                int(observation.get("group_id", 1) or 1)
            )
            channels = observation.get("channels") or []
            if not channels:
                metrics = observation.get("channel_metrics")
                if isinstance(metrics, dict):
                    channels = list(metrics.keys())
            failed_metrics = {
                str(channel): {
                    "snr": 0.0,
                    "peak_shape_score": 0.0,
                    "baseline_stability_score": 0.0,
                    "replicate_consistency_score": 0.0,
                    "success_score": 0.0,
                    "ok_scan_count": 0,
                    "total_scan_count": 1,
                }
                for channel in channels
            }
            if self._is_paired_observation(observation):
                quality = compute_paired_response_quality(
                    failed_metrics,
                    failed_metrics,
                    scoring,
                    optimization_direction,
                )
                return {
                    "buffer_channel_metrics": failed_metrics,
                    "target_channel_metrics": failed_metrics,
                    "channel_metrics": failed_metrics,
                    "quality": quality,
                    "Q_run": 0.0,
                    "reanalysis_error": str(exc),
                }
            quality = compute_run_quality(
                failed_metrics,
                scoring,
                optimization_direction,
            )
            return {
                "channel_metrics": failed_metrics,
                "quality": quality,
                "Q_run": 0.0,
                "reanalysis_error": str(exc),
            }

        def work():
            updates = []
            failures = []
            for index, observation in enumerate(observations, start=1):
                try:
                    rebuilt = self._bo_session.reanalyze_observation(
                        observation,
                        analysis=analysis,
                        scoring=scoring,
                        output_dir=output_dir,
                    )
                except Exception as exc:
                    rebuilt = failed_reanalysis_update(observation, exc)
                    failures.append((observation.get("iteration"), str(exc)))
                updates.append((observation, rebuilt))
                update_progress(index, observation.get("iteration"))
            self._frame.after(0, lambda: finish_success(updates, failures))

        def finish_error(exc):
            self._reanalyze_rescore_running = False
            self._reanalyze_rescore_button.configure(state="normal")
            self._rescore_status_var.set(f"Reanalysis stopped without applying changes: {exc}")
            messagebox.showerror("Reanalyze & Rescore", str(exc))

        def finish_success(updates, failures):
            for observation, rebuilt in updates:
                observation.update(rebuilt)
                for record in self._bo_session.suggestions:
                    if record.get("method_id") == observation.get("method_id"):
                        record["Q_run"] = observation["Q_run"]
            self._bo_session.config["scoring"] = dict(scoring)
            self._bo_session.config["analysis"] = dict(analysis)
            if self._config is not None:
                self._config["scoring"] = dict(scoring)
                self._config["analysis"] = dict(analysis)
            self._reanalyze_rescore_running = False
            self._reanalyze_rescore_button.configure(state="normal")
            self._refresh_history()
            self._refresh_record_files()
            if selected_iteration and selected_iteration in self._history_rows:
                self._history_tree.selection_set(selected_iteration)
                self._history_tree.focus(selected_iteration)
                self._select_history_iteration(selected_iteration)
            else:
                self._select_latest_history_iteration()
            failure_text = f" {len(failures)} failed iteration(s) were assigned Q=0." if failures else ""
            self._rescore_status_var.set(
                f"Reanalyzed and rescored {len(updates)} iteration(s).{failure_text} "
                "Use Save Rescored Session to persist the updated session."
            )

        threading.Thread(target=work, daemon=True).start()

    def _reset_rescore_to_original(self):
        source = self._loaded_original_config or self._bo_session.config if self._bo_session else self._config
        self._set_rescore_vars_from_config(source or {})
        self._preview_rescore_equation()

    def _save_rescored_session(self):
        if self._bo_session is None:
            messagebox.showwarning("Save Rescored Session", "Load a BO session first.")
            return
        try:
            for obs in self._bo_session.observations:
                iteration = int(obs.get("iteration", 0) or 0)
                if iteration > 0:
                    group_id = int(obs.get("group_id", 1) or 1)
                    objective = str(obs.get("objective") or "").lower()
                    suffix = "_paired_quality.json" if objective == "paired_response" else "_quality.json"
                    stem = self._bo_session._group_iteration_stem(iteration, group_id=group_id)
                    self._bo_session._write_json(self._bo_session.analysis_dir / f"{stem}{suffix}", obs)
            self._bo_session._write_json(self._bo_session.record_dir / "bo_config_snapshot.json", self._bo_session.config)
            self._bo_session._write_history_csv()
            self._bo_session.save_state()
            self._rescore_status_var.set(f"Saved rescored Q values to {self._bo_session.STATE_FILE}.")
            self._status_var.set("Saved rescored BO session state.")
        except Exception as exc:
            messagebox.showerror("Save Rescored Session", str(exc))

    def _config_with_scoring(self, scoring):
        source = self._bo_session.config if self._bo_session else self._config
        cfg = dict(source or {})
        cfg["scoring"] = scoring
        return cfg

    def _refresh_current_q_equation(self, config=None):
        if not hasattr(self, "_current_q_equation_text"):
            return
        source = config if config is not None else (self._bo_session.config if self._bo_session else self._config)
        self._write_text(self._current_q_equation_text, "\n".join(self._q_equation_lines(source)))

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

    def _set_channel_group_vars(self, config):
        groups = channel_groups(config)
        self._channel_group_count_var.set(str(len(groups)))
        self._rebuild_channel_group_entries(
            [", ".join(str(ch) for ch in group["channels"]) for group in groups],
            group_configs=groups,
        )

    def _rebuild_channel_group_entries(self, values=None, group_configs=None):
        old_values = [var.get() for var in self._channel_group_vars]
        old_settings = list(self._channel_group_settings)
        requested = max(1, min(10, int(self._channel_group_count_var.get() or 1)))
        if values is None:
            channels = []
            for value in old_values:
                try:
                    channels.extend(parse_channels(value))
                except Exception:
                    pass
            if not channels:
                channels = list(range(1, 11))
            sizes = [len(channels) // requested + (1 if i < len(channels) % requested else 0) for i in range(requested)]
            values = []
            offset = 0
            for size in sizes:
                values.append(", ".join(str(ch) for ch in channels[offset:offset + size]))
                offset += size
        for child in self._channel_groups_frame.winfo_children():
            child.destroy()
        self._channel_group_vars = []
        self._channel_group_settings = []
        for index in range(requested):
            var = tk.StringVar(value=values[index] if index < len(values) else "")
            self._channel_group_vars.append(var)
            source = (
                group_configs[index] if group_configs and index < len(group_configs)
                else old_settings[index] if index < len(old_settings)
                else {}
            )
            def source_value(key, variable_key, default):
                if source.get(key) is not None:
                    return source[key]
                variable = source.get(variable_key)
                return variable.get() if variable is not None else default
            exploration_var = tk.StringVar(
                value=str(source_value("exploration", "exploration_var", self._exploration_var.get()))
            )
            warmup_var = tk.StringVar(
                value=str(source_value("n_initial_points", "warmup_var", self._gp_warmup_iterations_var.get()))
            )
            candidate_pool_var = tk.StringVar(
                value=str(source_value("candidate_pool_size", "candidate_pool_var", self._candidate_pool_var.get()))
            )
            local_pool_var = tk.StringVar(
                value=str(source_value("local_candidate_pool_size", "local_pool_var", self._local_pool_var.get()))
            )
            start_mode_var = tk.StringVar(
                value=str(source_value("initial_point_mode", "start_mode_var", self._initial_point_mode_var.get()))
            )
            optimization_direction_var = tk.StringVar(
                value=self._display_optimization_direction(
                    source_value("optimization_direction", "optimization_direction_var", self._optimization_direction_var.get())
                )
            )
            initial_parameters = dict(
                source.get("initial_parameters")
                or resolve_initial_parameters(self._config or {})
            )
            gp_falloffs = dict(
                source.get("gp_falloff_fractions")
                or source.get("gp_length_scales")
                or {
                    name: float(variable.get() or 0.2)
                    for name, variable in self._gp_length_scale_vars.items()
                }
            )
            settings = {
                "exploration_var": exploration_var,
                "warmup_var": warmup_var,
                "candidate_pool_var": candidate_pool_var,
                "local_pool_var": local_pool_var,
                "start_mode_var": start_mode_var,
                "optimization_direction_var": optimization_direction_var,
                "initial_parameters": initial_parameters,
                "gp_falloff_fractions": gp_falloffs,
            }
            self._channel_group_settings.append(settings)
            ttk.Label(self._channel_groups_frame, text=f"Group {index + 1}").grid(
                row=index, column=0, sticky="w", pady=2
            )
            entry = ttk.Entry(self._channel_groups_frame, textvariable=var)
            entry.grid(row=index, column=1, sticky="ew", padx=(6, 0), pady=2)
            entry.bind("<FocusOut>", lambda _e: self._sync_channel_groups(show_error=False))
        self._channel_groups_frame.columnconfigure(1, weight=1)
        self._rebuild_bo_group_optimizer_panels()
        self._sync_channel_groups(show_error=False)

    def _sync_channel_groups(self, show_error=True):
        if self._config is None:
            return
        try:
            groups = []
            all_channels = []
            for index, var in enumerate(self._channel_group_vars, 1):
                channels = parse_channels(var.get())
                if not channels:
                    raise ValueError(f"Group {index} must contain at least one channel")
                settings = self._channel_group_settings[index - 1]
                exploration = float(settings["exploration_var"].get())
                warmup = int(settings["warmup_var"].get())
                if not 0.0 <= exploration <= 1.0:
                    raise ValueError(f"Group {index} exploration must be between 0 and 1")
                if warmup < 0:
                    raise ValueError(f"Group {index} warmup must be zero or greater")
                start_mode = "random" if settings["start_mode_var"].get() == "random" else "specific"
                direction_var = settings.get("optimization_direction_var")
                optimization_direction = self._display_optimization_direction(
                    direction_var.get() if direction_var is not None else "maximize"
                )
                group_payload = {
                    "id": index,
                    "name": f"Group {index}",
                    "channels": channels,
                    "exploration": exploration,
                    "n_initial_points": warmup,
                    "candidate_pool_size": max(50, int(settings["candidate_pool_var"].get())),
                    "local_candidate_pool_size": max(0, int(settings["local_pool_var"].get())),
                    "initial_point_mode": start_mode,
                    "optimization_direction": optimization_direction,
                    "gp_falloff_fractions": dict(settings.get("gp_falloff_fractions") or {}),
                }
                if start_mode != "random":
                    group_payload["initial_parameters"] = dict(settings["initial_parameters"])
                groups.append(group_payload)
                all_channels.extend(channels)
            duplicates = sorted({ch for ch in all_channels if all_channels.count(ch) > 1})
            if duplicates:
                raise ValueError(f"Channels may only appear in one group: {duplicates}")
            self._config["channel_groups"] = groups
            self._config["channels"] = all_channels
            self._channels_var.set(", ".join(str(ch) for ch in all_channels))
        except Exception as exc:
            if show_error:
                messagebox.showerror("Channel Groups", str(exc))

    def _edit_bo_group_initial_parameters(self, group_index):
        settings = self._channel_group_settings[group_index]

        def save(updated):
            temp_config = json.loads(json.dumps(self._config or {}))
            temp_config["channel_groups"] = [{
                "name": f"Group {group_index + 1}",
                "channels": [1],
                "initial_parameters": updated,
            }]
            temp_config["channels"] = [1]
            errors = validate_bo_config(temp_config)
            if errors:
                raise ValueError("; ".join(errors))
            settings["initial_parameters"] = {
                name: float(updated[name]) for name in PARAMETER_ORDER
            }
            self._sync_channel_groups(show_error=False)
            self._status_var.set(
                f"Updated starting parameters for Group {group_index + 1}."
            )

        self._open_method_editor(
            f"Group {group_index + 1} Starting Parameters",
            settings["initial_parameters"],
            save,
            start_mode_var=settings["start_mode_var"],
            on_mode_change=lambda _mode: self._sync_channel_groups(show_error=False),
        )

    def _rebuild_bo_group_optimizer_panels(self):
        parent = getattr(self, "_bo_group_optimizer_panels_frame", None)
        if parent is None:
            return
        for child in parent.winfo_children():
            child.destroy()
        for index, settings in enumerate(self._channel_group_settings):
            panel = ttk.LabelFrame(parent, text=f"Group {index + 1}", padding=6)
            panel.pack(fill="x", pady=(0, 6))
            panel.columnconfigure(1, weight=1)
            ttk.Label(panel, text="Exploit ↔ Explore:").grid(row=0, column=0, sticky="w")

            def update_exploration(value, variable=settings["exploration_var"]):
                try:
                    variable.set(f"{float(value):.3f}")
                    self._sync_channel_groups(show_error=False)
                except Exception:
                    pass

            ttk.Scale(
                panel,
                from_=0.0,
                to=1.0,
                orient=tk.HORIZONTAL,
                variable=settings["exploration_var"],
                command=update_exploration,
            ).grid(row=0, column=1, sticky="ew", padx=6)
            exploration_entry = ttk.Entry(panel, textvariable=settings["exploration_var"], width=6)
            exploration_entry.grid(
                row=0, column=2, sticky="e"
            )
            exploration_entry.bind("<FocusOut>", lambda _e: self._sync_channel_groups(show_error=False))
            exploration_entry.bind("<Return>", lambda _e: self._sync_channel_groups(show_error=False))
            ttk.Label(panel, text="Global pool:").grid(row=1, column=0, sticky="w", pady=2)
            ttk.Entry(panel, textvariable=settings["candidate_pool_var"], width=8).grid(
                row=1, column=1, sticky="w", padx=6
            )
            ttk.Label(panel, text="Local pool:").grid(row=1, column=2, sticky="e", pady=2)
            ttk.Entry(panel, textvariable=settings["local_pool_var"], width=8).grid(
                row=1, column=3, sticky="w", padx=6
            )
            ttk.Label(panel, text="GP warmup iterations:").grid(row=2, column=0, sticky="w", pady=2)
            ttk.Entry(panel, textvariable=settings["warmup_var"], width=8).grid(
                row=2, column=1, sticky="w", padx=6
            )
            ttk.Label(panel, text="Start point:").grid(row=2, column=2, sticky="e", pady=2)
            start_mode = ttk.Combobox(
                panel,
                textvariable=settings["start_mode_var"],
                values=("specific", "random"),
                state="readonly",
                width=10,
            )
            start_mode.grid(row=2, column=3, sticky="w", padx=6)
            start_mode.bind("<<ComboboxSelected>>", lambda _e: self._sync_channel_groups(show_error=False))
            ttk.Label(panel, text="Optimize:").grid(row=3, column=0, sticky="w", pady=2)
            direction = ttk.Combobox(
                panel,
                textvariable=settings["optimization_direction_var"],
                values=("maximize", "minimize", "survey"),
                state="readonly",
                width=10,
            )
            direction.grid(row=3, column=1, sticky="w", padx=6)
            direction.bind("<<ComboboxSelected>>", lambda _e: self._sync_channel_groups(show_error=False))
            ttk.Button(
                panel,
                text="Edit Starting Parameters…",
                command=lambda group_index=index: self._edit_bo_group_initial_parameters(group_index),
            ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(4, 0))
            ttk.Button(
                panel,
                text="Edit GP Falloff…",
                command=lambda group_index=index: self._edit_group_gp_falloff(group_index, simulator=False),
            ).grid(row=4, column=2, columnspan=2, sticky="w", pady=(4, 0))

    def _rebuild_engine_channel_group_entries(self, values=None, group_configs=None):
        old_values = [var.get() for var in self._engine_channel_group_vars]
        old_settings = list(self._engine_channel_group_settings)
        requested = max(1, min(10, int(self._engine_channel_group_count_var.get() or 1)))
        if values is None:
            channels = []
            for value in old_values:
                try:
                    channels.extend(parse_channels(value))
                except Exception:
                    pass
            if not channels:
                source = self._config or {}
                channels = parse_channels(source.get("channels", list(range(1, 11))))
            sizes = [
                len(channels) // requested + (1 if index < len(channels) % requested else 0)
                for index in range(requested)
            ]
            values = []
            offset = 0
            for size in sizes:
                values.append(", ".join(str(ch) for ch in channels[offset:offset + size]))
                offset += size
        for child in self._engine_channel_groups_frame.winfo_children():
            child.destroy()
        self._engine_channel_group_vars = []
        self._engine_channel_group_settings = []
        for index in range(requested):
            var = tk.StringVar(value=values[index] if index < len(values) else "")
            self._engine_channel_group_vars.append(var)
            source = (
                group_configs[index] if group_configs and index < len(group_configs)
                else old_settings[index] if index < len(old_settings)
                else {}
            )
            def source_value(key, variable_key, default):
                if source.get(key) is not None:
                    return source[key]
                variable = source.get(variable_key)
                return variable.get() if variable is not None else default
            exploration_var = tk.StringVar(
                value=str(source_value("exploration", "exploration_var", self._engine_exploration_var.get()))
            )
            warmup_var = tk.StringVar(
                value=str(source_value("n_initial_points", "warmup_var", self._engine_warmup_iterations_var.get()))
            )
            candidate_pool_var = tk.StringVar(
                value=str(source_value("candidate_pool_size", "candidate_pool_var", self._engine_candidate_pool_var.get()))
            )
            local_pool_var = tk.StringVar(
                value=str(source_value("local_candidate_pool_size", "local_pool_var", self._engine_local_pool_var.get()))
            )
            start_mode_var = tk.StringVar(
                value=str(source_value("initial_point_mode", "start_mode_var", self._engine_initial_point_mode_var.get()))
            )
            optimization_direction_var = tk.StringVar(
                value=self._display_optimization_direction(
                    source_value(
                        "optimization_direction",
                        "optimization_direction_var",
                        self._engine_optimization_direction_var.get(),
                    )
                )
            )
            initial_parameters = dict(
                source.get("initial_parameters")
                or resolve_initial_parameters(self._config or {})
            )
            gp_falloffs = dict(
                source.get("gp_falloff_fractions")
                or source.get("gp_length_scales")
                or {
                    name: float(variable.get() or 0.2)
                    for name, variable in self._engine_gp_length_scale_vars.items()
                }
            )
            settings = {
                "exploration_var": exploration_var,
                "warmup_var": warmup_var,
                "candidate_pool_var": candidate_pool_var,
                "local_pool_var": local_pool_var,
                "start_mode_var": start_mode_var,
                "optimization_direction_var": optimization_direction_var,
                "initial_parameters": initial_parameters,
                "gp_falloff_fractions": gp_falloffs,
            }
            self._engine_channel_group_settings.append(settings)
            ttk.Label(
                self._engine_channel_groups_frame,
                text=f"Group {index + 1}",
            ).grid(row=index, column=0, sticky="w", pady=2)
            ttk.Entry(
                self._engine_channel_groups_frame,
                textvariable=var,
            ).grid(row=index, column=1, sticky="ew", padx=(6, 0), pady=2)
        self._engine_channel_groups_frame.columnconfigure(1, weight=1)
        self._rebuild_engine_group_optimizer_panels()

    def _engine_channel_groups_from_vars(self):
        groups = []
        assigned = []
        for index, var in enumerate(self._engine_channel_group_vars, 1):
            channels = parse_channels(var.get())
            if not channels:
                raise ValueError(f"Simulation group {index} must contain at least one channel.")
            settings = self._engine_channel_group_settings[index - 1]
            exploration = float(settings["exploration_var"].get())
            warmup = int(settings["warmup_var"].get())
            if not 0.0 <= exploration <= 1.0:
                raise ValueError(f"Simulation group {index} exploration must be between 0 and 1.")
            if warmup < 0:
                raise ValueError(f"Simulation group {index} warmup must be zero or greater.")
            start_mode = "random" if settings["start_mode_var"].get() == "random" else "specific"
            direction_var = settings.get("optimization_direction_var")
            optimization_direction = self._display_optimization_direction(
                direction_var.get() if direction_var is not None else "maximize"
            )
            group_payload = {
                "id": index,
                "name": f"Group {index}",
                "channels": channels,
                "exploration": exploration,
                "n_initial_points": warmup,
                "candidate_pool_size": max(50, int(settings["candidate_pool_var"].get())),
                "local_candidate_pool_size": max(0, int(settings["local_pool_var"].get())),
                "initial_point_mode": start_mode,
                "optimization_direction": optimization_direction,
                "gp_falloff_fractions": dict(settings.get("gp_falloff_fractions") or {}),
            }
            if start_mode != "random":
                group_payload["initial_parameters"] = dict(settings["initial_parameters"])
            groups.append(group_payload)
            assigned.extend(channels)
        duplicates = sorted({channel for channel in assigned if assigned.count(channel) > 1})
        if duplicates:
            raise ValueError(
                "Simulation channels may only belong to one group: "
                + ", ".join(str(channel) for channel in duplicates)
            )
        return groups

    def _edit_engine_group_initial_parameters(self, group_index):
        settings = self._engine_channel_group_settings[group_index]

        def save(updated):
            config = self._engine_bo_config(self._engine_sim_config())
            resolved = resolve_initial_parameters({
                **config,
                "initial_parameters": updated,
            })
            errors = validate_bo_config({
                **config,
                "channel_groups": [{
                    "name": "Preview",
                    "channels": [1],
                    "initial_parameters": resolved,
                }],
            })
            if errors:
                raise ValueError("; ".join(errors))
            settings["initial_parameters"] = resolved
            self._engine_status_var.set(
                f"Updated starting parameters for Group {group_index + 1}."
            )

        self._open_method_editor(
            f"Group {group_index + 1} Starting Parameters",
            settings["initial_parameters"],
            save,
            start_mode_var=settings["start_mode_var"],
        )

    def _rebuild_engine_group_optimizer_panels(self):
        parent = getattr(self, "_engine_group_optimizer_panels_frame", None)
        if parent is None:
            return
        for child in parent.winfo_children():
            child.destroy()
        for index, settings in enumerate(self._engine_channel_group_settings):
            panel = ttk.LabelFrame(parent, text=f"Group {index + 1}", padding=6)
            panel.pack(fill="x", pady=(0, 6))
            panel.columnconfigure(1, weight=1)
            ttk.Label(panel, text="Exploit ↔ Explore:").grid(row=0, column=0, sticky="w")

            def update_exploration(value, variable=settings["exploration_var"]):
                try:
                    variable.set(f"{float(value):.3f}")
                except Exception:
                    pass

            ttk.Scale(
                panel,
                from_=0.0,
                to=1.0,
                orient=tk.HORIZONTAL,
                variable=settings["exploration_var"],
                command=update_exploration,
            ).grid(row=0, column=1, sticky="ew", padx=6)
            ttk.Entry(panel, textvariable=settings["exploration_var"], width=6).grid(
                row=0, column=2, sticky="e"
            )
            ttk.Label(panel, text="Global pool:").grid(row=1, column=0, sticky="w", pady=2)
            ttk.Entry(panel, textvariable=settings["candidate_pool_var"], width=8).grid(
                row=1, column=1, sticky="w", padx=6
            )
            ttk.Label(panel, text="Local pool:").grid(row=1, column=2, sticky="e", pady=2)
            ttk.Entry(panel, textvariable=settings["local_pool_var"], width=8).grid(
                row=1, column=3, sticky="w", padx=6
            )
            ttk.Label(panel, text="GP warmup iterations:").grid(row=2, column=0, sticky="w", pady=2)
            ttk.Entry(panel, textvariable=settings["warmup_var"], width=8).grid(
                row=2, column=1, sticky="w", padx=6
            )
            ttk.Label(panel, text="Start point:").grid(row=2, column=2, sticky="e", pady=2)
            start_mode = ttk.Combobox(
                panel,
                textvariable=settings["start_mode_var"],
                values=("specific", "random"),
                state="readonly",
                width=10,
            )
            start_mode.grid(row=2, column=3, sticky="w", padx=6)
            ttk.Label(panel, text="Optimize:").grid(row=3, column=0, sticky="w", pady=2)
            ttk.Combobox(
                panel,
                textvariable=settings["optimization_direction_var"],
                values=("maximize", "minimize", "survey"),
                state="readonly",
                width=10,
            ).grid(row=3, column=1, sticky="w", padx=6)
            ttk.Button(
                panel,
                text="Edit Starting Parameters…",
                command=lambda group_index=index: self._edit_engine_group_initial_parameters(group_index),
            ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(4, 0))
            ttk.Button(
                panel,
                text="Edit GP Falloff…",
                command=lambda group_index=index: self._edit_group_gp_falloff(group_index, simulator=True),
            ).grid(row=4, column=2, columnspan=2, sticky="w", pady=(4, 0))

    def _edit_group_gp_falloff(self, group_index, simulator=False):
        settings_list = (
            self._engine_channel_group_settings
            if simulator else self._channel_group_settings
        )
        settings = settings_list[group_index]
        current = dict(settings.get("gp_falloff_fractions") or {})
        win = tk.Toplevel(self._frame)
        win.title(f"Group {group_index + 1} GP Falloff")
        win.transient(self._frame)
        box = ttk.Frame(win, padding=12)
        box.pack(fill="both", expand=True)
        variables = {}
        for row, name in enumerate(PARAMETER_ORDER):
            ttk.Label(box, text=name.replace("_", " ").title()).grid(
                row=row, column=0, sticky="w", pady=3
            )
            variable = tk.StringVar(value=str(current.get(name, 0.2)))
            variables[name] = variable
            ttk.Entry(box, textvariable=variable, width=12).grid(
                row=row, column=1, sticky="w", padx=(8, 0), pady=3
            )

        def save():
            try:
                values = {name: float(variable.get()) for name, variable in variables.items()}
                if any(value <= 0 for value in values.values()):
                    raise ValueError("Every GP falloff fraction must be greater than zero.")
                settings["gp_falloff_fractions"] = values
                if not simulator:
                    self._sync_channel_groups(show_error=False)
                status = self._engine_status_var if simulator else self._status_var
                status.set(f"Updated GP falloff settings for Group {group_index + 1}.")
                win.destroy()
            except Exception as exc:
                messagebox.showerror("GP Falloff", str(exc), parent=win)

        buttons = ttk.Frame(box)
        buttons.grid(row=len(PARAMETER_ORDER), column=0, columnspan=2, pady=(10, 0))
        ttk.Button(buttons, text="Save", command=save).pack(side="left", padx=4)
        ttk.Button(buttons, text="Cancel", command=win.destroy).pack(side="left", padx=4)
        win.grab_set()
        win.focus_force()

    def _toggle_auto_titration(self):
        enabled = bool(self._run_auto_titration_var.get())
        if not callable(self._configure_auto_titration):
            if enabled:
                messagebox.showwarning(
                    "BO Autotitration", "Automated titration is not available."
                )
                self._run_auto_titration_var.set(False)
            return
        if not enabled:
            self._configure_auto_titration(False, [])
            return
        if self._config is None:
            messagebox.showwarning("BO Autotitration", "Load a BO config first.")
            self._run_auto_titration_var.set(False)
            return
        self._sync_channel_groups(show_error=False)
        groups = []
        for group in channel_groups(self._config):
            effective = copy.deepcopy(self._config)
            effective["initial_parameters"] = copy.deepcopy(
                group.get("initial_parameters")
                or resolve_initial_parameters(self._config)
            )
            groups.append(
                {
                    "id": int(group["id"]),
                    "name": str(group.get("name") or f"Group {group['id']}"),
                    "channels": list(group.get("channels") or []),
                    "params": resolve_initial_parameters(effective),
                    "method_options": copy.deepcopy(
                        self._config.get("method_options") or {}
                    ),
                }
            )
        self._configure_auto_titration(True, groups)

    def select_setup_tab(self):
        self._tabs.select(0)

    def _start_post_bo_titration(self):
        if (
            not self._run_auto_titration_var.get()
            or self._post_bo_titration_started
            or not callable(self._on_bo_finished)
        ):
            return
        self._post_bo_titration_started = True
        try:
            self._on_bo_finished()
            self._auto_status_var.set(
                "BO complete; starting the locked automated titration."
            )
        except Exception as exc:
            self._post_bo_titration_started = False
            self._auto_status_var.set(
                f"BO complete, but autotitration could not start: {exc}"
            )
            messagebox.showerror("BO Autotitration", str(exc))

    def _validate_config(self, show_dialog=True):
        if self._config is None:
            return
        self._sync_channel_groups(show_error=False)
        self._config["objective"] = self._bo_objective_var.get()
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
            self._loaded_original_config = json.loads(json.dumps(self._bo_session.config))
            self._set_rescore_vars_from_config(self._bo_session.config)
            self._suggestion = None
            self._record_dir_var.set(f"Record folder: {self._bo_session.record_dir}")
            self._status_var.set(f"BO session started with {len(self._bo_session.candidates)} valid candidates.")
            self._refresh_history()
            self._render_best()
            self._refresh_model_artifacts()
            self._refresh_record_files()
            self._refresh_surrogate_view()
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
            self._loaded_original_config = json.loads(json.dumps(loaded.config))
            self._config = dict(loaded.config)
            if loaded.config_path is not None:
                self._config_path_var.set(str(loaded.config_path))
            self._analysis_dir_var.set(str(loaded.analysis_output_dir))
            self._set_analysis_vars_from_config(self._config.get("analysis", {}))
            self._set_method_option_vars_from_config(self._config)
            self._set_algorithm_vars_from_config(self._config)
            self._set_scoring_vars_from_config(self._config)
            self._set_engine_tuning_vars_from_config(self._config)
            self._set_rescore_vars_from_config(self._config)
            self._channels_var.set(", ".join(str(ch) for ch in self._config.get("channels", [])))
            self._set_channel_group_vars(self._config)
            self._refresh_parameter_table()
            self._refresh_initial_parameters_table()
            self._suggestion = None
            self._record_dir_var.set(f"Record folder: {loaded.record_dir}")
            self._refresh_history()
            self._render_best()
            self._refresh_model_artifacts()
            self._refresh_record_files()
            self._select_latest_history_iteration()
            self._refresh_surrogate_view()
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
            groups = channel_groups(self._bo_session.config)
            group = min(
                groups,
                key=lambda candidate: sum(
                    1 for obs in self._bo_session.observations
                    if int(obs.get("group_id", 1)) == int(candidate["id"])
                ),
            )
            self._suggestion = self._bo_session.ask_next_for_group(group["id"])
            self._render_suggestion()
            self._status_var.set(
                f"Suggested BO iteration {self._suggestion.iteration}. Send it to queue when ready."
            )
        except Exception as exc:
            messagebox.showerror("BO Suggestion", str(exc))

    def get_best_parameter_groups(self):
        """Return the best completed observation for every active BO group."""
        if self._bo_session is None:
            raise RuntimeError("Load or run a Bayesian Optimization session first.")

        results = []
        missing = []
        for group in channel_groups(self._bo_session.config):
            observations = [
                observation
                for observation in self._bo_session.observations
                if int(observation.get("group_id", 1)) == int(group["id"])
                and isinstance(observation.get("params"), dict)
                and observation.get("Q_run") is not None
            ]
            if not observations:
                missing.append(str(group.get("name") or f"Group {group['id']}"))
                continue

            effective_config = self._bo_session._config_for_group(group["id"])
            direction = self._display_optimization_direction(
                effective_config.get("acquisition", {}).get(
                    "optimization_direction", "maximize"
                )
            )
            if direction == "minimize":
                best = min(
                    observations,
                    key=lambda observation: float(observation["Q_run"]),
                )
            elif direction == "survey":
                best = max(
                    observations,
                    key=lambda observation: abs(float(observation["Q_run"])),
                )
            else:
                best = max(
                    observations,
                    key=lambda observation: float(observation["Q_run"]),
                )
            results.append(
                {
                    "id": int(group["id"]),
                    "name": str(group.get("name") or f"Group {group['id']}"),
                    "channels": list(group.get("channels") or []),
                    "score": float(best["Q_run"]),
                    "params": copy.deepcopy(best["params"]),
                    "method_options": copy.deepcopy(
                        effective_config.get("method_options") or {}
                    ),
                    "session_id": self._bo_session.session_id,
                    "iteration": int(best.get("iteration", 0) or 0),
                    "optimization_direction": direction,
                }
            )

        if missing:
            raise RuntimeError(
                "No completed BO observation is available for: "
                + ", ".join(missing)
                + "."
            )
        return results

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
        if self._engine_page_index == 0:
            if not self._simulation_dims:
                self._engine_load_active_dimensions()
            elif bool(self._engine_paired_response_var.get()):
                self._ensure_frequency_simulation_dimension()
                self._engine_rebuild_landscape_cache(refresh_plot=False)
                self._engine_refresh_dimension_tree()
                self._engine_refresh_landscape_inspector()
        self._engine_go_page(self._engine_page_index + 1)

    def _engine_prev_page(self):
        self._engine_go_page(self._engine_page_index - 1)

    def _engine_load_active_dimensions(self):
        if self._config is None or not hasattr(self, "_engine_dim_tree"):
            return
        try:
            self._simulation_dims = default_dimensions(self._config, limit=3)
            if bool(self._engine_paired_response_var.get()):
                self._ensure_frequency_simulation_dimension()
            self._engine_rebuild_landscape_cache(refresh_plot=False)
            self._engine_refresh_dimension_tree()
            if self._simulation_dims:
                self._engine_dim_tree.selection_set("0")
                self._engine_dim_tree.see("0")
            self._engine_refresh_landscape_inspector()
            self._engine_status_var.set(f"Loaded {len(self._simulation_dims)} active simulation dimension(s).")
        except Exception as exc:
            messagebox.showerror("Simulation Engine", str(exc))

    def _ensure_frequency_simulation_dimension(self):
        if self._config is None:
            return
        if any(str(dim.get("name")) == "frequency" for dim in self._simulation_dims):
            return
        cfg = self._config
        initial = resolve_initial_parameters(cfg)
        p_cfg = dict((cfg.get("parameters") or {}).get("frequency") or {})
        values = [float(v) for v in p_cfg.get("values", []) if v not in (None, "")]
        minimum = float(p_cfg.get("min", min(values) if values else 1.0))
        maximum = float(p_cfg.get("max", max(values) if values else 1000.0))
        if maximum <= minimum:
            minimum, maximum = min(minimum, maximum), max(minimum, maximum) + 1.0
        span = max(maximum - minimum, 1e-12)
        optimum = min(max(float(initial.get("frequency", (minimum + maximum) / 2.0)), minimum), maximum)
        frequency_dim = {
            "name": "frequency",
            "minimum": minimum,
            "maximum": maximum,
            "optimum": optimum,
            "spread": span * 0.22,
            "landscape": "gaussian",
            "weight": 1.0,
        }
        if len(self._simulation_dims) >= 3:
            self._simulation_dims[-1] = frequency_dim
        else:
            self._simulation_dims.append(frequency_dim)

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
                    "Follow Q" if bool(dim.get("delta_follow_q", True)) else "Separate",
                    self._fmt_raw(dim.get("delta_optimum", dim.get("optimum"))),
                    self._fmt_raw(dim.get("delta_spread", dim.get("spread"))),
                    dim.get("delta_landscape", dim.get("landscape", "gaussian")),
                    self._fmt_raw(dim.get("delta_weight", 1.0)),
                ),
            )

    def _engine_edit_dimension(self):
        if not self._simulation_dims:
            self._engine_load_active_dimensions()
        selection = self._engine_dim_tree.selection() if hasattr(self, "_engine_dim_tree") else ()
        if not selection:
            messagebox.showwarning("Simulation Engine", "Select a simulation dimension first.")
            return
        try:
            idx = int(selection[0])
        except Exception:
            messagebox.showwarning("Simulation Engine", "Select a simulation dimension first.")
            return
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
            "delta_optimum": tk.StringVar(value=self._fmt_raw(dim.get("delta_optimum", dim.get("optimum")))),
            "delta_spread": tk.StringVar(value=self._fmt_raw(dim.get("delta_spread", dim.get("spread")))),
            "delta_landscape": tk.StringVar(value=str(dim.get("delta_landscape", dim.get("landscape", "gaussian")))),
            "delta_weight": tk.StringVar(value=self._fmt_raw(dim.get("delta_weight", 1.0))),
        }
        delta_follow_var = tk.BooleanVar(value=bool(dim.get("delta_follow_q", True)))
        ttk.Label(box, text=str(dim.get("name", "")), font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 8))
        q_box = ttk.LabelFrame(box, text="Traditional Q distribution", padding=8)
        q_box.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=(0, 8))
        delta_box = ttk.LabelFrame(box, text="Delta peak distribution", padding=8)
        delta_box.grid(row=1, column=2, columnspan=2, sticky="nsew")
        labels = [
            ("minimum", "Minimum"),
            ("maximum", "Maximum"),
            ("optimum", "Optimum"),
            ("spread", "Spread"),
            ("weight", "Weight"),
        ]
        for row, (key, label) in enumerate(labels):
            ttk.Label(q_box, text=f"{label}:").grid(row=row, column=0, sticky="w", pady=3)
            ttk.Entry(q_box, textvariable=vars_by_key[key], width=16).grid(row=row, column=1, sticky="w", pady=3)
        ttk.Label(q_box, text="Shape:").grid(row=len(labels), column=0, sticky="w", pady=3)
        ttk.Combobox(
            q_box,
            textvariable=vars_by_key["landscape"],
            values=LANDSCAPE_TYPES,
            state="readonly",
            width=14,
        ).grid(row=len(labels), column=1, sticky="w", pady=3)
        ttk.Checkbutton(delta_box, text="Follow traditional Q distribution", variable=delta_follow_var).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))
        delta_labels = [
            ("delta_optimum", "Delta optimum"),
            ("delta_spread", "Delta spread"),
            ("delta_weight", "Delta weight"),
        ]
        for row, (key, label) in enumerate(delta_labels, start=1):
            ttk.Label(delta_box, text=f"{label}:").grid(row=row, column=0, sticky="w", pady=3)
            ttk.Entry(delta_box, textvariable=vars_by_key[key], width=16).grid(row=row, column=1, sticky="w", pady=3)
        ttk.Label(delta_box, text="Delta shape:").grid(row=len(delta_labels) + 1, column=0, sticky="w", pady=3)
        ttk.Combobox(
            delta_box,
            textvariable=vars_by_key["delta_landscape"],
            values=LANDSCAPE_TYPES,
            state="readonly",
            width=14,
        ).grid(row=len(delta_labels) + 1, column=1, sticky="w", pady=3)
        buttons = ttk.Frame(box)
        buttons.grid(row=2, column=0, columnspan=4, pady=(10, 0))

        def save():
            try:
                updated = dict(dim)
                updated["minimum"] = float(vars_by_key["minimum"].get())
                updated["maximum"] = float(vars_by_key["maximum"].get())
                updated["optimum"] = float(vars_by_key["optimum"].get())
                updated["spread"] = float(vars_by_key["spread"].get())
                updated["landscape"] = vars_by_key["landscape"].get()
                updated["weight"] = float(vars_by_key["weight"].get())
                updated["delta_follow_q"] = bool(delta_follow_var.get())
                updated["delta_optimum"] = float(vars_by_key["delta_optimum"].get())
                updated["delta_spread"] = float(vars_by_key["delta_spread"].get())
                updated["delta_landscape"] = vars_by_key["delta_landscape"].get()
                updated["delta_weight"] = float(vars_by_key["delta_weight"].get())
                if updated["maximum"] <= updated["minimum"]:
                    raise ValueError("Maximum must be greater than minimum.")
                if updated["spread"] <= 0:
                    raise ValueError("Spread must be positive.")
                if updated["delta_spread"] <= 0:
                    raise ValueError("Delta spread must be positive.")
                if updated["landscape"] not in LANDSCAPE_TYPES:
                    updated["landscape"] = "gaussian"
                if updated["delta_landscape"] not in LANDSCAPE_TYPES:
                    updated["delta_landscape"] = "gaussian"
                updated["optimum"] = min(max(updated["optimum"], updated["minimum"]), updated["maximum"])
                updated["delta_optimum"] = min(max(updated["delta_optimum"], updated["minimum"]), updated["maximum"])
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
        iterations = max(1, int(self._engine_iterations_var.get() or 1))
        paired_batch_size = max(1, int(self._engine_paired_batch_size_var.get() or 1))
        return {
            "dimensions": [dict(dim) for dim in self._simulation_dims],
            "iterations": iterations,
            "paired_response": bool(self._engine_paired_response_var.get()),
            "paired_batch_size": paired_batch_size,
            "grid_size": max(5, min(45, int(self._engine_grid_var.get() or 25))),
            "seed": int(self._engine_seed_var.get() or self._config.get("random_seed", 42)),
            "measurement_noise": max(0.0, float(self._engine_measurement_noise_var.get() or 0.03)),
            "channel_noise": max(0.0, float(self._engine_channel_noise_var.get() or 0.025)),
            "target_response_gain_uA": max(0.0, float(self._engine_target_response_gain_var.get() or 2.0)),
            "target_noise_multiplier": max(0.05, float(self._engine_target_noise_multiplier_var.get() or 1.05)),
            "target_response_floor": max(0.0, min(1.0, float(self._engine_delta_peak_floor_var.get() or 0.0))),
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
            self._engine_go_page(3)
            self._engine_status_var.set("Drew synthetic landscape map. Run the optimizer to add a path.")
        except Exception as exc:
            messagebox.showerror("Simulation Engine", str(exc))

    def _engine_rebuild_landscape_cache(self, refresh_plot=False):
        if self._config is None:
            return
        sim_cfg = self._engine_sim_config()
        sim_bo_config = self._engine_bo_config(sim_cfg)
        from core.bo_simulation import SyntheticSWVSimulationEngine

        engine = SyntheticSWVSimulationEngine(sim_bo_config, sim_cfg)
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
            self._sync_channel_groups(show_error=False)
            sim_cfg = self._engine_sim_config()
            sim_bo_config = self._engine_bo_config(sim_cfg)
            session_mgr = getattr(self._session, "session_manager", None)
            exp_path = session_mgr.require_experiment() if session_mgr is not None else None
            if exp_path is None:
                return
            output_root = Path(exp_path)
            analysis_dir = output_root / "bo_analysis"
            self._analysis_dir_var.set(str(analysis_dir))
            self._engine_progress_var.set(0.0)
            self._engine_progress_text_var.set("Starting...")
            self._engine_status_var.set("Running optimizer simulation...")
            self._frame.update_idletasks()

            def progress(done, total, message):
                total = max(1, int(total))
                self._engine_progress_var.set(100.0 * float(done) / float(total))
                self._engine_progress_text_var.set(f"{int(done)}/{int(total)}")
                self._engine_status_var.set(message)
                self._frame.update_idletasks()

            if bool(sim_cfg.get("paired_response")):
                result = run_paired_response_optimizer_simulation(
                    sim_bo_config,
                    sim_cfg,
                    output_root=output_root,
                    cycles=sim_cfg["iterations"],
                    batch_size=sim_cfg["paired_batch_size"],
                    analysis_output_dir=analysis_dir,
                    progress_callback=progress,
                )
            else:
                result = run_optimizer_simulation(
                    sim_bo_config,
                    sim_cfg,
                    output_root=output_root,
                    iterations=sim_cfg["iterations"],
                    analysis_output_dir=analysis_dir,
                    progress_callback=progress,
                )
            self._simulation_result = result
            self._bo_session = result.get("session")
            if self._bo_session is not None:
                self._loaded_original_config = json.loads(json.dumps(self._bo_session.config))
                self._set_rescore_vars_from_config(self._bo_session.config)
                self._record_dir_var.set(f"Record folder: {self._bo_session.record_dir}")
                self._refresh_history()
                self._render_best()
                self._refresh_model_artifacts()
                self._refresh_record_files()
                self._select_latest_history_iteration()
                self._refresh_surrogate_view()
            self._engine_selected_index = max(0, len(result.get("rows", [])) - 1)
            self._engine_refresh_results()
            self._engine_refresh_landscape_inspector()
            self._engine_render_plot(show_all=True)
            self._engine_update_trace_text()
            self._engine_go_page(3)
            best = min((row for row in result["rows"]), key=lambda r: r.get("distance", 1.0), default=None)
            if best:
                self._engine_progress_var.set(100.0)
                if bool(sim_cfg.get("paired_response")):
                    cycles = int(result.get("cycles", sim_cfg["iterations"]) or sim_cfg["iterations"])
                    batch_size = int(result.get("batch_size", sim_cfg["paired_batch_size"]) or sim_cfg["paired_batch_size"])
                    total_traces = int(result.get("total_swv_traces", cycles * batch_size * 2) or 0)
                    self._engine_progress_text_var.set(f"{cycles}/{cycles} cycles")
                    self._engine_status_var.set(
                        f"Completed {cycles} paired cycle(s): {batch_size} parameter set(s)/cycle, "
                        f"{len(result['rows'])} paired Q comparison(s), {total_traces} simulated SWV trace(s). "
                        f"Best computed Q={best['Q_run']:.3f}, true Q={best['true_Q']:.3f}."
                    )
                else:
                    group_count = len(channel_groups(sim_bo_config))
                    expected_rows = sim_cfg["iterations"] * group_count
                    self._engine_progress_text_var.set(f"{len(result['rows'])}/{expected_rows}")
                    self._engine_status_var.set(
                        f"Completed {sim_cfg['iterations']} iteration(s) for {group_count} group(s), "
                        f"{len(result['rows'])} group observation(s). "
                        f"Closest distance={best['distance']:.3f}, computed Q={best['Q_run']:.3f}, true Q={best['true_Q']:.3f}."
                    )
            else:
                self._engine_progress_text_var.set("")
                self._engine_status_var.set("Simulation completed without optimizer rows.")
        except Exception as exc:
            self._engine_progress_text_var.set("Error")
            messagebox.showerror("Simulation Engine", str(exc))

    def _engine_refresh_results(self):
        if not hasattr(self, "_engine_result_tree"):
            return
        for row in self._engine_result_tree.get_children():
            self._engine_result_tree.delete(row)
        rows = (self._simulation_result or {}).get("rows", [])
        paired_result = bool((self._simulation_result or {}).get("paired_response"))
        self._engine_result_tree.heading("#0", text="Cycle" if paired_result else "Iter")
        session = (self._simulation_result or {}).get("session")
        observations = session.observations if session is not None else []
        for idx, row in enumerate(rows):
            obs = observations[idx] if idx < len(observations) else {}
            peak, snr = self._engine_peak_snr_for_obs(obs)
            cycle = row.get("paired_cycle") or obs.get("paired_cycle") or ""
            parameter_set = row.get("paired_batch_index") or obs.get("paired_batch_index") or ""
            buffer_trace = row.get("buffer_trace_number") or obs.get("buffer_trace_number") or ""
            target_trace = row.get("target_trace_number") or obs.get("target_trace_number") or ""
            self._engine_result_tree.insert(
                "",
                "end",
                iid=str(idx),
                text=str(cycle if paired_result else row.get("iteration", idx + 1)),
                values=(
                    str(row.get("group_name") or obs.get("group_name") or "Group 1"),
                    str(parameter_set) if paired_result else "",
                    str(row.get("iteration", idx + 1)) if paired_result else "",
                    str(buffer_trace) if paired_result else "",
                    str(target_trace) if paired_result else "",
                    self._fmt(row.get("Q_run")),
                    self._fmt(row.get("true_Q")),
                    self._fmt(row.get("paired_Q_score") if paired_result else None),
                    self._fmt(row.get("delta_peak")),
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
        best = max(rows, key=lambda row: self._optimization_objective_value(row.get("Q_run", 0.0), self._config))
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
        paired_landscape = bool(landscape.get("paired_response"))
        value_key = "paired_Q_score" if paired_landscape else "true_Q"
        value_label = "Paired Q" if paired_landscape else "True Q"
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
            ax.plot([p[name] for p in ordered], [p[value_key] for p in ordered], color=self.ACCENT_DARK)
            if path_rows:
                ax.scatter([r.get(name) for r in path_rows], [r.get(value_key) for r in path_rows], color="#d67b32", s=20, zorder=3)
                selected = path_rows[-1]
                ax.scatter(
                    [selected.get(name)],
                    [selected.get(value_key)],
                    color="#ffd166",
                    edgecolors="black",
                    linewidths=0.9,
                    s=95,
                    zorder=4,
                )
            ax.set_xlabel(name)
            ax.set_ylabel(value_label)
            if not paired_landscape:
                ax.set_ylim(0.0, 1.02)
            ax.grid(alpha=0.25)
        elif len(dims) == 2:
            x_name, y_name = dims[0]["name"], dims[1]["name"]
            ax = fig.add_subplot(111)
            x_vals = [p[x_name] for p in points]
            y_vals = [p[y_name] for p in points]
            z_vals = [p[value_key] for p in points]
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
            fig.colorbar(contour, ax=ax, label=value_label)
            ax.set_xlabel(x_name)
            ax.set_ylabel(y_name)
            ax.set_title("2D response landscape map" if paired_landscape else "2D landscape map")
            ax.grid(alpha=0.2)
        else:
            x_name, y_name, z_name = dims[0]["name"], dims[1]["name"], dims[2]["name"]
            ax = fig.add_subplot(111, projection="3d")
            scatter = ax.scatter(
                [p[x_name] for p in points],
                [p[y_name] for p in points],
                [p[z_name] for p in points],
                c=[p[value_key] for p in points],
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
            fig.colorbar(scatter, ax=ax, label=value_label, shrink=0.75)
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
            delta_line, = ax.plot([], [], color="#7b3f98", linewidth=1.8, alpha=0.9, label="Delta peak")
            self._engine_distribution_lines = {
                "success": success_line,
                "q": q_line,
                "peak": peak_line,
                "noise": noise_line,
                "delta": delta_line,
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
            paired_curve = any(row.get("paired_Q_score") is not None for row in curve)
            success = [row.get("paired_Q_score", row.get("success_score")) if paired_curve else row.get("success_score") for row in curve]
            peak = [row.get("peak_score") for row in curve]
            noise = [row.get("noise_score") for row in curve]
            delta = [row.get("delta_peak_score") for row in curve]
            self._engine_distribution_lines["success"].set_label("Paired Q" if paired_curve else "Success")
            self._engine_distribution_lines["success"].set_data(x, success)
            self._engine_distribution_lines["q"].set_data(x, q)
            self._engine_distribution_lines["peak"].set_data(x, peak)
            self._engine_distribution_lines["noise"].set_data(x, noise)
            self._engine_distribution_lines["delta"].set_data(x, delta)
            if self._engine_distribution_empty_text is not None:
                self._engine_distribution_empty_text.set_visible(False)
            ax.set_xlabel(title)
            ax.legend(loc="best", fontsize=8)
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
        paired_landscape = bool(landscape.get("paired_response"))
        names = [str(dim.get("name", "")) for dim in dims[:3]]
        if names:
            self._engine_cube_tree.heading("#0", text=" / ".join(names))
        else:
            self._engine_cube_tree.heading("#0", text="Point")
        self._engine_cube_tree.heading("Success", text="Paired Q" if paired_landscape else "Success")
        points = sorted(
            points,
            key=lambda row: (
                -float((row.get("paired_Q_score") if paired_landscape and row.get("paired_Q_score") is not None else row.get("success_score", 0.0)) or 0.0),
                -float(row.get("true_Q", 0.0) or 0.0),
                -float(row.get("delta_peak_score", 0.0) or 0.0) if paired_landscape else 0.0,
                float(row.get("distance", 1.0) or 1.0),
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
                    self._fmt(point.get("paired_Q_score") if paired_landscape and point.get("paired_Q_score") is not None else point.get("success_score")),
                    self._fmt(point.get("peak_score")),
                    self._fmt(point.get("noise_score")),
                    self._fmt(point.get("delta_peak")),
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
                -float((row.get("paired_Q_score") if bool(landscape.get("paired_response")) and row.get("paired_Q_score") is not None else row.get("success_score", 0.0)) or 0.0),
                -float(row.get("true_Q", 0.0) or 0.0),
                -float(row.get("delta_peak_score", 0.0) or 0.0) if bool(landscape.get("paired_response")) else 0.0,
                float(row.get("distance", 1.0) or 1.0),
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
        if bool(landscape.get("paired_response")):
            payload = engine.paired_analysis_payload(params, iteration=0, phase="target")
        else:
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
        if truth.get("paired_Q_score") is not None:
            lines.append(f"Paired Q: {self._fmt(truth.get('paired_Q_score'))}")
        lines.append(f"Peak component: {self._fmt(truth.get('peak_score'))}")
        lines.append(f"Noise component: {self._fmt(truth.get('noise_score'))}")
        if truth.get("expected_delta_peak_uA") is not None:
            lines.append(f"Expected delta peak: {self._fmt(truth.get('expected_delta_peak_uA'))} uA")
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
                    f"  Peak prominence: {self._fmt(first.get('snr'))}",
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
        def trace_sort_key(item):
            label = str(item[0])
            digits = "".join(ch for ch in label if ch.isdigit())
            return (int(digits) if digits else 10**9, label)

        for idx, (ch, trace) in enumerate(sorted(traces.items(), key=trace_sort_key)):
            volts = trace.get("voltage_v", [])
            currents = trace.get("current_uA", [])
            if volts and currents:
                label = f"Ch {ch}" if str(ch).isdigit() else str(ch)
                ax.plot(volts, currents, color=palette[idx % len(palette)], linewidth=1.4, label=label)
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
        quality = dict(obs.get("quality") or {})
        paired_obs = str(obs.get("objective") or quality.get("objective") or "").lower() == "paired_response"
        if paired_obs:
            buffer_traces = obs.get("buffer_swv_trace_preview", {}) or {}
            target_traces = obs.get("target_swv_trace_preview", {}) or {}
            traces = {}
            for ch, trace in buffer_traces.items():
                traces[f"buffer ch {ch}"] = trace
            for ch, trace in target_traces.items():
                traces[f"target ch {ch}"] = trace
            title = f"Cycle {obs.get('paired_cycle')} set {obs.get('paired_batch_index')} buffer/target SWVs"
        else:
            traces = obs.get("swv_trace_preview", {})
            title = f"Iteration {obs.get('iteration')} synthetic SWV"
        self._engine_render_trace_plot(traces, title=title)
        lines = [
            f"Cycle {obs.get('paired_cycle')}, parameter set {obs.get('paired_batch_index')}" if paired_obs else f"Iteration {obs.get('iteration')}",
            f"BO suggestion iteration: {obs.get('iteration')}" if paired_obs else "",
            f"Buffer trace: {obs.get('buffer_trace_number')}; target trace: {obs.get('target_trace_number')}" if paired_obs else "",
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
                f"Mean peak prominence: {self._fmt(snr)}",
                f"Truth components: peak {self._fmt(truth.get('peak_score'))}, noise {self._fmt(truth.get('noise_score'))}, shape {self._fmt(truth.get('shape_score'))}",
            ]
        )
        lines = [line for line in lines if line != ""]
        if paired_obs:
            lines.extend(
                [
                    "",
                    "Paired comparison:",
                    f"  Cycle {obs.get('paired_cycle')} set {obs.get('paired_batch_index')}: buffer trace {obs.get('buffer_trace_number')} compared with target trace {obs.get('target_trace_number')}",
                    f"  Paired Q score: {self._fmt(truth.get('paired_Q_score'))}",
                    f"Expected target response: {self._fmt(truth.get('expected_delta_peak_uA'))} uA",
                    f"Mean signed delta peak: {self._fmt(quality.get('mean_delta_peak_height_uA'))} uA",
                    f"Mean absolute delta peak: {self._fmt(quality.get('mean_abs_delta_peak_height_uA'))} uA",
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
        if self._defer_results_render(
            self._analysis_q_plot_frame,
            "Trend plot rendering is paused while a measurement is collecting data.\nAcquisition has priority.",
        ):
            return
        rows = list(self._bo_session.observations) if self._bo_session is not None else []
        metric = self._analysis_trend_metric_var.get() or "Q_run"
        self._render_metric_trend_plot(
            self._analysis_q_plot_frame,
            rows,
            metric,
            empty_text="Import analysis results to see trends over iterations.",
        )

    def get_slack_q_trend_image(self):
        """Return a PNG Q-score trend while a BO auto loop is active."""
        if not (self._auto_running or self._paired_queue_running):
            return None
        rows = list(self._bo_session.observations) if self._bo_session is not None else []
        series = self._grouped_trend_series(rows, "Q_run")
        if not series:
            return None

        from matplotlib.figure import Figure

        fig = Figure(figsize=(7.2, 4.0), dpi=120)
        ax = fig.add_subplot(111)
        for group_name, points in series:
            ax.plot(
                [point[1] for point in points],
                [point[2] for point in points],
                marker="o",
                linewidth=2,
                label=group_name,
            )
        ax.set_ylim(bottom=0.0)
        ax.set_xlabel("BO iteration")
        ax.set_ylabel("Q_run")
        ax.set_title("BO Q-score trend")
        ax.grid(alpha=0.25)
        ax.legend(loc="best", fontsize=8)
        fig.tight_layout()

        output = io.BytesIO()
        fig.savefig(output, format="png")
        return output.getvalue(), "bo_q_score_trend.png", "BO Q-score trend"

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
        series = self._grouped_trend_series(rows, metric)
        all_points = []

        if not series:
            ax.text(0.5, 0.5, empty_text, ha="center", va="center")
            ax.set_axis_off()
        else:
            for group_name, points in series:
                x_values = [point[1] for point in points]
                values = [point[2] for point in points]
                all_points.extend(points)
                ax.plot(
                    x_values,
                    values,
                    marker="o",
                    linewidth=1.8,
                    label=group_name,
                )
            if metric == "Q_run":
                ax.set_ylim(bottom=0.0)
            ax.set_xlabel("BO iteration")
            ax.set_ylabel(metric)
            ax.set_title(f"{metric} over BO iterations")
            ax.grid(alpha=0.25)
            ax.legend(loc="best", fontsize=8)
        fig.tight_layout()
        canvas = FigureCanvasTkAgg(fig, master=parent)
        if all_points:
            canvas.mpl_connect(
                "button_press_event",
                lambda event, points=all_points: self._on_analysis_trend_click(event, points),
            )
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    def _grouped_trend_series(self, rows, metric):
        grouped = {}
        for index, row in enumerate(rows):
            value = self._analysis_trend_value(row, metric)
            if value is None:
                continue
            try:
                group_id = int(row.get("group_id", 1))
                group_name = str(row.get("group_name") or f"Group {group_id}")
                iteration = int(row.get("iteration", index + 1))
                history_key = f"g{group_id}:i{iteration}"
                grouped.setdefault((group_id, group_name), []).append(
                    (history_key, iteration, float(value))
                )
            except (TypeError, ValueError):
                continue
        return [
            (group_name, sorted(points, key=lambda point: point[1]))
            for (_group_id, group_name), points in sorted(grouped.items())
        ]

    def _on_analysis_trend_click(self, event, points):
        if event.inaxes is None or not points:
            return
        try:
            click_x, click_y = event.x, event.y
            transformed = event.inaxes.transData.transform([(x_value, value) for _item_id, x_value, value in points])
            nearest_iteration = None
            nearest_distance = None
            for (item_id, _x_value, _value), (px, py) in zip(points, transformed):
                distance = ((px - click_x) ** 2 + (py - click_y) ** 2) ** 0.5
                if nearest_distance is None or distance < nearest_distance:
                    nearest_distance = distance
                    nearest_iteration = item_id
            if nearest_iteration is None or nearest_distance is None or nearest_distance > 18:
                return
            item_id = str(nearest_iteration)
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
            grouped = {}
            for row in rows:
                grouped.setdefault(
                    (int(row.get("group_id", 1)), str(row.get("group_name", "Group 1"))),
                    [],
                ).append(row)
            for (_group_id, group_name), group_rows in sorted(grouped.items()):
                group_rows.sort(key=lambda row: int(row.get("iteration", 0)))
                iterations = [int(row.get("iteration", idx + 1)) for idx, row in enumerate(group_rows)]
                q_values = [float(row.get("Q_run", 0.0) or 0.0) for row in group_rows]
                line, = ax.plot(iterations, q_values, marker="o", linewidth=1.8, label=group_name)
                true_rows = [
                    (int(row.get("iteration", index + 1)), row.get("true_Q"))
                    for index, row in enumerate(group_rows)
                    if include_true_q and row.get("true_Q") is not None
                ]
                if true_rows:
                    ax.plot(
                        [iteration for iteration, _value in true_rows],
                        [float(value) for _iteration, value in true_rows],
                        color=line.get_color(),
                        linewidth=1.2,
                        linestyle="--",
                        alpha=0.65,
                        label=f"{group_name} true Q",
                    )
            if selected_index is not None and 0 <= int(selected_index) < len(rows):
                idx = int(selected_index)
                selected_row = rows[idx]
                ax.scatter(
                    [int(selected_row.get("iteration", idx + 1))],
                    [float(selected_row.get("Q_run", 0.0) or 0.0)],
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
            refresh_errors = []
            for label, action in (
                ("score table", lambda: self._render_scores(obs)),
                ("history refresh", self._refresh_history),
                ("history selection", lambda: self._select_history_iteration(str(obs.get("iteration")))),
                ("best summary", self._render_best),
                ("model artifacts", self._refresh_model_artifacts),
                ("record files", self._refresh_record_files),
            ):
                try:
                    action()
                except Exception as exc:
                    refresh_errors.append(f"{label}: {exc}")
            self._clear_text(self._suggestion_text)
            status = f"Imported analysis for iteration {obs['iteration']}. Q_run={obs['Q_run']:.3f}"
            if refresh_errors:
                status += f" | UI refresh warnings: {len(refresh_errors)}"
                self._status_var.set(status)
                message = "Analysis imported, but some BO plots/tables could not refresh:\n\n" + "\n".join(refresh_errors[:8])
                if prompt:
                    messagebox.showwarning("Import BO Analysis", message)
            else:
                self._status_var.set(status)
            if self._results_render_deferred and not self._measurement_priority_active():
                self._flush_deferred_results_render(preferred_iteration=obs.get("iteration"))
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
        if (
            self._run_auto_titration_var.get()
            and (
                not callable(self._is_auto_titration_locked)
                or not self._is_auto_titration_locked()
            )
        ):
            messagebox.showwarning(
                "BO Autotitration",
                "Configure and lock the automatic titration settings before starting BO.",
            )
            return
        self._post_bo_titration_started = False
        if self._bo_objective_var.get() == "paired_response":
            self._start_paired_auto_loop()
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
        self._paired_queue_running = False
        self._auto_status_var.set("Auto loop stopped.")

    def _paired_bo_block_from_setup(self, target_iterations: int) -> dict:
        self._save_config()
        channels = self._channels_var.get().strip()
        batch_size = max(1, int(self._paired_batch_size_var.get() or 1))
        warmup_batch_size = max(
            1, int(self._paired_warmup_batch_size_var.get() or batch_size)
        )
        warmup_single_batch = bool(self._paired_warmup_single_batch_var.get())
        warmup_cycles = max(0, int(self._gp_warmup_iterations_var.get() or 0))
        cfg = json.loads(json.dumps(self._config or {}))
        cfg["objective"] = "paired_response"
        cfg["paired_warmup_cycles"] = warmup_cycles
        cfg["paired_batch_size"] = batch_size
        cfg["paired_warmup_batch_size"] = warmup_batch_size
        cfg["paired_warmup_single_batch"] = warmup_single_batch
        cfg["n_initial_points"] = warmup_cycles * batch_size
        self._config = cfg
        groups = channel_groups(cfg)
        warmup_iterations = min(
            (
                max(
                    0,
                    int(
                        group.get("n_initial_points", cfg.get("n_initial_points", 0))
                        or 0
                    ),
                )
                for group in groups
            ),
            default=max(0, int(cfg.get("n_initial_points", 0) or 0)),
        )
        if warmup_single_batch and warmup_iterations > 0:
            warmup_batch_size = warmup_iterations
            cfg["paired_warmup_batch_size"] = warmup_batch_size
        return {
            "bo_config_path": self._config_path_var.get().strip(),
            "analysis_output_dir": self._analysis_dir_var.get().strip(),
            "analysis_file_glob": self._analysis_glob_var.get().strip() or "*.json",
            "target_iterations": int(target_iterations),
            "objective": "paired_response",
            "batch_size": batch_size,
            "warmup_batch_size": warmup_batch_size,
            "warmup_iterations": warmup_iterations,
            "warmup_single_batch": warmup_single_batch,
            "target_exchange_block_path": self._paired_target_exchange_var.get().strip(),
            "buffer_exchange_block_path": self._paired_buffer_exchange_var.get().strip(),
            "target_equilibration_seconds": max(0.0, float(self._paired_target_equilibration_var.get() or 0.0)),
            "buffer_equilibration_seconds": max(0.0, float(self._paired_buffer_equilibration_var.get() or 0.0)),
            "channels_override": channels,
            "paired_warmup_cycles": warmup_cycles,
            "scoring": dict(cfg.get("scoring") or {}),
            "analysis": dict(cfg.get("analysis") or {}),
            "config_overrides": {
                "n_initial_points": warmup_cycles * batch_size,
                "paired_warmup_cycles": warmup_cycles,
                "paired_batch_size": batch_size,
                "paired_warmup_batch_size": warmup_batch_size,
                "paired_warmup_single_batch": warmup_single_batch,
            },
        }

    def _format_paired_bo_block_details(self, block: dict) -> str:
        target = int(block.get("target_iterations", 1) or 1)
        batch = max(1, int(block.get("batch_size", 1) or 1))
        warmup_batch = max(1, int(block.get("warmup_batch_size", batch) or batch))
        warmup_text = (
            "all warmups in one batch"
            if bool(block.get("warmup_single_batch", False))
            else f"warmup batches x {warmup_batch}"
        )
        channels = (block.get("channels_override") or "").strip() or "config channels"
        config_name = Path(str(block.get("bo_config_path") or "BO config")).name
        target_eq = max(0.0, float(block.get("target_equilibration_seconds", 0.0) or 0.0))
        buffer_eq = max(0.0, float(block.get("buffer_equilibration_seconds", 0.0) or 0.0))
        eq_text = f" | eq target {target_eq:g}s, buffer {buffer_eq:g}s" if (target_eq or buffer_eq) else ""
        return (
            f"{config_name} | paired {target} total iter, batches x {batch} methods "
            f"({warmup_text}){eq_text} | {channels}"
        )

    def _start_paired_auto_loop(self):
        if self._session.is_running:
            messagebox.showwarning("Auto Loop", "The measurement queue is already running.")
            return
        if self._session.measurement_queue:
            messagebox.showwarning(
                "Auto Loop",
                "Paired BO starts only from an empty queue so it cannot run unrelated items.",
            )
            return
        try:
            target_iterations = int(self._auto_target_var.get())
            if target_iterations < 1:
                raise ValueError("Total target iterations must be at least 1.")
            block = self._paired_bo_block_from_setup(target_iterations)
        except Exception as exc:
            messagebox.showerror("Paired BO Auto Loop", str(exc))
            return
        item = {
            "type": "BO_AUTO_LOOP",
            "status": "pending",
            "details": self._format_paired_bo_block_details(block),
            "bo_block": block,
        }
        self._paired_queue_running = True
        self._auto_running = False
        self._add_to_queue(item)
        self._refresh_queue()
        self._auto_status_var.set(
            f"Queued paired BO: {target_iterations} total iteration(s), including warmups; "
            f"regular batch size {block['batch_size']}. Starting queue."
        )
        self._run_queue()

    def _auto_submit_next(self):
        if not self._auto_running:
            return
        target = int(self._auto_target_var.get())
        groups = channel_groups(self._bo_session.config) if self._bo_session else []
        completed = len(self._bo_session.observations) if self._bo_session else 0
        expected = target * max(1, len(groups))
        if completed >= expected:
            self._auto_running = False
            self._auto_status_var.set(f"Auto loop complete: {completed}/{expected} group iteration(s).")
            session_mgr = getattr(self._session, "session_manager", None)
            if session_mgr is not None:
                best = self._bo_session.best_observation()
                best_text = ""
                if best is not None:
                    best_text = (
                        f" Best Q_run={float(best.get('Q_run', 0.0)):.3f} "
                        f"at iter {int(best.get('iteration', 0) or 0)}."
                    )
                experiment_name = (
                    session_mgr.current_experiment_path.name
                    if session_mgr.current_experiment_path is not None
                    else "(none)"
                )
                session_mgr.notify_slack(
                    f"BO auto loop completed: {completed}/{expected} group iterations. "
                    f"Session={self._bo_session.session_id}; "
                    f"Experiment={experiment_name}.{best_text}"
                )
            self._start_post_bo_titration()
            return
        if self._session.is_running:
            return
        queue_start_index = len(self._session.measurement_queue)
        if self._session.measurement_queue:
            if not self._auto_queue_is_safe():
                self._auto_running = False
                self._auto_status_var.set("Auto loop stopped: queue contains non-BO items.")
                return
        try:
            group = min(
                groups,
                key=lambda candidate: sum(
                    1 for obs in self._bo_session.observations
                    if int(obs.get("group_id", 1)) == int(candidate["id"])
                ),
            )
            self._suggestion = self._bo_session.ask_next_for_group(group["id"])
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
            if queue_start_index > 0 and callable(self._run_queue_from_index):
                self._run_queue_from_index(queue_start_index)
            else:
                self._run_queue()
        except Exception as exc:
            self._auto_running = False
            self._auto_status_var.set(f"Auto loop stopped: {exc}")
            messagebox.showerror("Auto Loop", str(exc))

    def on_queue_complete(self, summary):
        if self._paired_queue_running:
            self._paired_queue_running = False
            if summary.get("failed", 0) or summary.get("stopped", 0):
                self._auto_status_var.set("Paired BO stopped: queue did not complete cleanly.")
                return
            loaded = self._load_latest_paired_queue_session()
            if loaded is not None:
                self._bo_session = loaded
                self._loaded_original_config = json.loads(json.dumps(self._bo_session.config))
                self._config = dict(self._bo_session.config)
                self._sync_suggestion_from_session()
                self._set_rescore_vars_from_config(self._config)
                self._record_dir_var.set(f"Record folder: {self._bo_session.record_dir}")
                self._flush_deferred_results_render()
                self._auto_status_var.set(
                    f"Paired BO complete: {len(self._bo_session.observations)} paired comparison(s)."
                )
                self._tabs.select(3)
                self._start_post_bo_titration()
            else:
                self._auto_status_var.set("Paired BO complete, but session folder could not be loaded.")
            return
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
        self._auto_submit_next()

    def _on_live_paired_bo_update(self, payload):
        if not isinstance(payload, dict):
            return
        record_dir = str(payload.get("record_dir") or "").strip()
        if not record_dir:
            return
        selected_iteration = None
        if self._selected_history_observation is not None:
            selected_iteration = self._selected_history_observation.get("iteration")
        if selected_iteration is None:
            selection = self._history_tree.selection()
            if selection:
                selected_iteration = selection[0]
        try:
            loaded = BOIntegrationSession.load(record_dir)
        except Exception:
            return
        self._bo_session = loaded
        self._loaded_original_config = json.loads(json.dumps(self._bo_session.config))
        self._config = dict(self._bo_session.config)
        self._sync_suggestion_from_session()
        self._set_rescore_vars_from_config(self._config)
        self._record_dir_var.set(f"Record folder: {self._bo_session.record_dir}")
        if self._measurement_priority_active():
            self._results_render_deferred = True
            self._refresh_history()
            self._refresh_model_artifacts()
            self._refresh_record_files()
            self._schedule_deferred_results_render()
        else:
            self._flush_deferred_results_render(preferred_iteration=selected_iteration)
        iteration = payload.get("iteration")
        completed = len(self._bo_session.observations)
        if iteration is not None:
            self._auto_status_var.set(
                f"Paired BO progress: imported iteration {iteration}. {completed} paired comparison(s) recorded."
            )

    def _measurement_priority_active(self) -> bool:
        if getattr(self._session, "current_runner", None) is not None:
            return True
        try:
            status = self._session.get_queue_status()
        except Exception:
            status = {}
        active_type = str(status.get("active_step_type") or "").strip().upper()
        measurement_types = {"CV", "SWV", "DPV", "LSV", "EIS", "CUSTOM", "CUSTOM_MUX"}
        return active_type in measurement_types

    def _defer_results_render(self, frame, message: str) -> bool:
        if not self._measurement_priority_active():
            return False
        self._results_render_deferred = True
        if frame is not None:
            for child in frame.winfo_children():
                child.destroy()
            ttk.Label(
                frame,
                text=message,
                anchor="center",
                justify="center",
            ).pack(fill="both", expand=True)
        self._status_var.set("Measurement is collecting data; heavy BO result plots are deferred until acquisition finishes.")
        return True

    def _schedule_deferred_results_render(self, delay_ms: int = 250) -> None:
        if self._results_render_flush_job is not None:
            return

        def retry():
            self._results_render_flush_job = None
            self._flush_deferred_results_render()

        try:
            self._results_render_flush_job = self._frame.after(int(delay_ms), retry)
        except Exception:
            self._results_render_flush_job = None

    def _flush_deferred_results_render(self, preferred_iteration=None) -> bool:
        if self._measurement_priority_active():
            if self._results_render_deferred:
                self._schedule_deferred_results_render()
            return False
        if self._results_render_flush_job is not None:
            try:
                self._frame.after_cancel(self._results_render_flush_job)
            except Exception:
                pass
            self._results_render_flush_job = None
        if self._bo_session is None:
            self._results_render_deferred = False
            return False

        selected_iteration = preferred_iteration
        if selected_iteration is None and self._selected_history_observation is not None:
            selected_iteration = self._selected_history_observation.get("iteration")
        if selected_iteration is None:
            try:
                selection = self._history_tree.selection()
            except Exception:
                selection = ()
            if selection:
                selected_iteration = selection[0]

        self._results_render_deferred = False
        self._refresh_history()
        self._refresh_model_artifacts()
        self._refresh_record_files()
        history_key = self._resolve_history_key(selected_iteration)
        if history_key is not None:
            self._history_tree.selection_set(history_key)
            self._history_tree.focus(history_key)
            self._history_tree.see(history_key)
            self._select_history_iteration(history_key)
        else:
            self._select_latest_history_iteration()
        self._refresh_surrogate_view()
        return True

    def _load_latest_paired_queue_session(self):
        for item in self._session.measurement_queue:
            if str(item.get("type") or "").upper() == "BO_AUTO_LOOP" and item.get("bo_record_dir"):
                try:
                    return BOIntegrationSession.load(item.get("bo_record_dir"))
                except Exception:
                    pass
        session_mgr = getattr(self._session, "session_manager", None)
        exp_path = session_mgr.require_experiment() if session_mgr is not None else None
        if exp_path is None:
            return None
        root = Path(exp_path) / "bo_sessions"
        if not root.exists():
            return None
        candidates = [path for path in root.iterdir() if path.is_dir() and (path / BOIntegrationSession.STATE_FILE).exists()]
        if not candidates:
            return None
        latest = max(candidates, key=lambda path: path.stat().st_mtime)
        try:
            return BOIntegrationSession.load(latest)
        except Exception:
            return None

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

    def _auto_queue_is_safe(self):
        queue = self._session.measurement_queue
        if not queue:
            return True
        session_id = self._bo_session.session_id if self._bo_session else None
        for item in queue:
            ref = item.get("bo_ref") or {}
            if ref.get("session_id") != session_id:
                return False
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
        gp_falloff_var = tk.StringVar(value=self._gp_length_scale_vars[name].get())

        ttk.Label(box, text="Mode:").grid(row=0, column=0, sticky="w", pady=4)
        mode_combo = ttk.Combobox(
            box, textvariable=mode_var, values=("active", "locked", "tied"), state="readonly", width=16
        )
        mode_combo.grid(row=0, column=1, sticky="w", pady=4)
        space_label = ttk.Label(box, text="Space:")
        space_label.grid(row=1, column=0, sticky="w", pady=4)
        space_combo = ttk.Combobox(
            box,
            textvariable=space_var,
            values=("discrete", "continuous"),
            state="readonly",
            width=16,
        )
        space_combo.grid(row=1, column=1, sticky="w", pady=4)
        active_values_label = ttk.Label(box, text="Discrete values:")
        active_values_label.grid(row=2, column=0, sticky="w", pady=4)
        active_values_entry = ttk.Entry(box, textvariable=values_var, width=48)
        active_values_entry.grid(row=2, column=1, columnspan=3, sticky="ew", pady=4)

        continuous_range_label = ttk.Label(box, text="Continuous min/max:")
        continuous_range_label.grid(row=3, column=0, sticky="w", pady=4)
        continuous_min_entry = ttk.Entry(box, textvariable=min_var, width=12)
        continuous_min_entry.grid(row=3, column=1, sticky="w", pady=4)
        continuous_max_entry = ttk.Entry(box, textvariable=max_var, width=12)
        continuous_max_entry.grid(row=3, column=1, padx=(96, 0), sticky="w", pady=4)
        system_range = DEFAULT_PARAMETER_RANGES[name]
        unit = str(current.get("unit") or "").strip()
        unit_text = f" {unit}" if unit else ""
        system_bounds_label = ttk.Label(
            box,
            text=(
                f"System bounds: {system_range['min']:g} to "
                f"{system_range['max']:g}{unit_text}"
            ),
            foreground="#666666",
        )
        system_bounds_label.grid(row=3, column=2, sticky="w", padx=(8, 0), pady=4)
        quantization_label = ttk.Label(box, text="Continuous quantization:")
        quantization_label.grid(row=4, column=0, sticky="w", pady=4)
        quantization_entry = ttk.Entry(box, textvariable=step_var, width=12)
        quantization_entry.grid(row=4, column=1, sticky="w", pady=4)
        potential_parameters = {
            "begin_potential",
            "end_potential",
            "step_potential",
            "amplitude",
            "conditioning_potential",
        }
        hardware_quantization_min = 0.000932 if name in potential_parameters else None
        if hardware_quantization_min is not None:
            quantization_help = (
                f"Instrument minimum (PGStat mode 3): "
                f"{hardware_quantization_min:g} V; blank = unquantized"
            )
        else:
            quantization_help = (
                "No verified instrument minimum; must be >0 and no larger "
                "than the selected range; blank = unquantized"
            )
        quantization_bounds_label = ttk.Label(
            box,
            text=quantization_help,
            foreground="#666666",
        )
        quantization_bounds_label.grid(row=4, column=2, sticky="w", padx=(8, 0), pady=4)
        scale_label = ttk.Label(box, text="Scale:")
        scale_label.grid(row=5, column=0, sticky="w", pady=4)
        scale_combo = ttk.Combobox(box, textvariable=scale_var, values=("linear", "log"), width=12)
        scale_combo.grid(row=5, column=1, sticky="w", pady=4)
        sigma_label = ttk.Label(box, text="Proposal sigma:")
        sigma_label.grid(row=6, column=0, sticky="w", pady=4)
        sigma_entry = ttk.Entry(box, textvariable=sigma_var, width=12)
        sigma_entry.grid(row=6, column=1, sticky="w", pady=4)
        gp_falloff_label = ttk.Label(box, text="GP falloff:")
        gp_falloff_label.grid(row=7, column=0, sticky="w", pady=4)
        gp_falloff_entry = ttk.Entry(box, textvariable=gp_falloff_var, width=12)
        gp_falloff_entry.grid(row=7, column=1, sticky="w", pady=4)
        gp_falloff_help = ttk.Label(
            box,
            text="fraction of range; blank = learn all GP falloffs",
            foreground="#666666",
        )
        gp_falloff_help.grid(row=7, column=2, sticky="w", padx=(8, 0), pady=4)
        locked_value_label = ttk.Label(box, text="Locked value:")
        locked_value_label.grid(row=8, column=0, sticky="w", pady=4)
        locked_value_entry = ttk.Entry(box, textvariable=value_var, width=18)
        locked_value_entry.grid(row=8, column=1, sticky="w", pady=4)
        tie_to_label = ttk.Label(box, text="Tie to:")
        tie_to_label.grid(row=9, column=0, sticky="w", pady=4)
        tie_to_combo = ttk.Combobox(box, textvariable=tie_var, values=PARAMETER_ORDER, width=24)
        tie_to_combo.grid(row=9, column=1, sticky="w", pady=4)

        def set_visible(widgets, visible):
            for widget in widgets:
                widget.grid() if visible else widget.grid_remove()

        def refresh_relevant_fields(*_args):
            active = mode_var.get() == "active"
            continuous = active and space_var.get() == "continuous"
            discrete = active and space_var.get() == "discrete"
            set_visible((space_label, space_combo), active)
            set_visible((active_values_label, active_values_entry), discrete)
            set_visible(
                (
                    continuous_range_label,
                    continuous_min_entry,
                    continuous_max_entry,
                    system_bounds_label,
                ),
                continuous,
            )
            set_visible(
                (quantization_label, quantization_entry, quantization_bounds_label),
                continuous,
            )
            set_visible((scale_label, scale_combo), active)
            set_visible((sigma_label, sigma_entry), continuous)
            set_visible((gp_falloff_label, gp_falloff_entry, gp_falloff_help), active)
            set_visible((locked_value_label, locked_value_entry), mode_var.get() == "locked")
            set_visible((tie_to_label, tie_to_combo), mode_var.get() == "tied")

        mode_combo.bind("<<ComboboxSelected>>", refresh_relevant_fields)
        space_combo.bind("<<ComboboxSelected>>", refresh_relevant_fields)
        refresh_relevant_fields()

        buttons = ttk.Frame(box)
        buttons.grid(row=10, column=0, columnspan=2, pady=(10, 0))

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
                if space_var.get() == "continuous":
                    system_min = float(system_range["min"])
                    system_max = float(system_range["max"])
                    continuous_min = float(updated["min"])
                    continuous_max = float(updated["max"])
                    if continuous_min < system_min or continuous_max > system_max:
                        raise ValueError(
                            f"{name} continuous range must stay within the system bounds "
                            f"{system_min:g} to {system_max:g}{unit_text}."
                        )
                    if continuous_min > continuous_max:
                        raise ValueError("Continuous minimum cannot exceed continuous maximum.")
                updated["scale"] = scale_var.get()
                updated["proposal_sigma"] = float(sigma_var.get() or 0.15)
                updated["step"] = (
                    None
                    if space_var.get() != "continuous" or not step_var.get().strip()
                    else float(step_var.get())
                )
                if updated["step"] is not None:
                    continuous_span = float(updated["max"]) - float(updated["min"])
                    if updated["step"] <= 0:
                        raise ValueError("Continuous quantization must be greater than zero.")
                    if (
                        hardware_quantization_min is not None
                        and updated["step"] < hardware_quantization_min
                    ):
                        raise ValueError(
                            f"{name} continuous quantization cannot be smaller than the "
                            f"PGStat mode 3 applied-potential resolution "
                            f"({hardware_quantization_min:g} V)."
                        )
                    if updated["step"] > continuous_span:
                        raise ValueError(
                            "Continuous quantization cannot be larger "
                            f"than the selected range span ({continuous_span:g}{unit_text})."
                        )
                if value_var.get().strip():
                    updated["value"] = float(value_var.get())
                updated["tie_to"] = tie_var.get()
                falloff_text = gp_falloff_var.get().strip()
                if falloff_text:
                    falloff = float(falloff_text)
                    if falloff <= 0:
                        raise ValueError("GP falloff must be greater than zero.")
                    self._gp_length_scale_vars[name].set(str(falloff))
                    for other_name, falloff_setting in self._gp_length_scale_vars.items():
                        if not falloff_setting.get().strip():
                            falloff_setting.set("0.2")
                else:
                    for falloff_var in self._gp_length_scale_vars.values():
                        falloff_var.set("")
                falloffs = self._gp_length_scales_from_vars()
                acquisition = self._config.setdefault("acquisition", {})
                acquisition["gp_falloff_fractions"] = falloffs
                acquisition["gp_length_scales"] = falloffs
                params[name] = updated
                self._config["parameters"] = params
                self._refresh_gp_falloff_summary()
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
        if not hasattr(self, "_initial_tree"):
            return
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
            start_mode_var=self._initial_point_mode_var,
            on_mode_change=lambda _mode: self._sync_algorithm_config(show_error=False),
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

    def _open_method_editor(self, title, values, on_save, start_mode_var=None, on_mode_change=None):
        win = tk.Toplevel(self._frame)
        win.title(title)
        win.transient(self._frame)
        win.resizable(False, False)
        box = ttk.Frame(win, padding=12)
        box.pack(fill="both", expand=True)
        vars_by_name = {}
        entries_by_name = {}
        labels = {
            "begin_potential": "Begin potential (V)",
            "end_potential": "End potential (V)",
            "step_potential": "Step potential (V)",
            "amplitude": "Amplitude (V)",
            "frequency": "Frequency (Hz)",
            "conditioning_potential": "Conditioning potential (V)",
            "conditioning_time": "Conditioning time (s)",
        }
        row_offset = 0
        dialog_start_mode_var = None
        if start_mode_var is not None:
            row_offset = 1
            ttk.Label(box, text="Start point").grid(row=0, column=0, sticky="w", pady=3)
            dialog_start_mode_var = tk.StringVar(
                value="random" if str(start_mode_var.get()).strip().lower() == "random" else "specific"
            )
            start_mode_combo = ttk.Combobox(
                box,
                textvariable=dialog_start_mode_var,
                values=("specific", "random"),
                state="readonly",
                width=16,
            )
            start_mode_combo.grid(row=0, column=1, sticky="w", pady=3)

        def entry_state_for(name):
            if dialog_start_mode_var is not None and dialog_start_mode_var.get() == "random":
                return "disabled"
            if title == "Edit Initial Parameters" or "Starting Parameters" in title:
                param_cfg = (self._config or {}).get("parameters", {}).get(name, {})
                if str(param_cfg.get("mode", "")).lower() == "tied":
                    return "disabled"
            return "normal"

        def refresh_entry_states(*_args):
            for param_name, entry in entries_by_name.items():
                entry.configure(state=entry_state_for(param_name))

        for row, name in enumerate(PARAMETER_ORDER, start=row_offset):
            ttk.Label(box, text=labels.get(name, name)).grid(row=row, column=0, sticky="w", pady=3)
            var = tk.StringVar(value=str(values.get(name, "")))
            vars_by_name[name] = var
            entry = ttk.Entry(box, textvariable=var, width=18, state=entry_state_for(name))
            entries_by_name[name] = entry
            entry.grid(row=row, column=1, sticky="w", pady=3)
        if dialog_start_mode_var is not None:
            dialog_start_mode_var.trace_add("write", refresh_entry_states)
        buttons = ttk.Frame(box)
        buttons.grid(row=len(PARAMETER_ORDER) + row_offset, column=0, columnspan=2, pady=(10, 0))

        def save():
            try:
                if dialog_start_mode_var is not None:
                    mode = "random" if dialog_start_mode_var.get() == "random" else "specific"
                    start_mode_var.set(mode)
                    if on_mode_change is not None:
                        on_mode_change(mode)
                    if mode == "random":
                        win.destroy()
                        return
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
            self._write_text(
                self._suggestion_text,
                "No active BO suggestion.\n\nStart a BO run to populate the current suggested method.",
            )
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

    def _sync_suggestion_from_session(self):
        if self._bo_session is None:
            self._suggestion = None
            self._render_suggestion()
            return
        pending_batch = list(getattr(self._bo_session, "pending_batch", []) or [])
        if pending_batch:
            try:
                self._suggestion = self._bo_session.ask_batch(len(pending_batch))[0]
            except Exception:
                record = pending_batch[0]
                self._suggestion = type("SuggestionView", (), record)()
            self._render_suggestion()
            return
        pending = getattr(self._bo_session, "pending", None)
        if isinstance(pending, dict):
            self._suggestion = type("SuggestionView", (), dict(pending))()
        else:
            self._suggestion = None
        self._render_suggestion()

    def _render_scores(self, observation):
        for row in self._score_tree.get_children():
            self._score_tree.delete(row)
        if hasattr(self, "_q_equation_text"):
            config = self._bo_session.config if self._bo_session else self._config
            self._write_text(self._q_equation_text, "\n".join(self._q_equation_lines(config)))
        paired = self._is_paired_observation(observation)
        classic_cols = (
            "Classic Q", "Prom. Term", "Repeat-SNR Term", "Peak Term",
            "Shape Term", "Baseline Term", "Replicate Term", "Success Term",
            "Noise Adj", "Clip Adj",
        )
        if paired:
            score_cols = (
                "Phase", *classic_cols, "Classic Pair Q", "Buffer Term", "Target Term",
                "Paired Prom. Term", "Paired Repeat Term", "Peak uA", "Trace Prominence", "Prominence Score", "Shape", "Success",
                "Paired Q", "Delta Peak", "Paired Prominence", "Paired Repeat SNR",
            )
        else:
            score_cols = (
                *classic_cols, "Peak uA", "Peak Prominence", "Repeat-scan SNR",
                "Prominence Score", "Shape", "Baseline", "Replicate", "Success",
            )
        self._score_tree.configure(columns=score_cols)
        for col in score_cols:
            self._score_tree.heading(col, text=col)
            self._score_tree.column(col, width=86, anchor="center", stretch=False)
        components = observation["quality"].get("channel_components", {})
        channel_metrics = observation.get("channel_metrics", {})
        buffer_channel_metrics = observation.get("buffer_channel_metrics", {})
        target_channel_metrics = observation.get("target_channel_metrics", channel_metrics)
        source_config = self._bo_session.config if self._bo_session else self._config or {}
        scoring = dict(source_config.get("scoring") or {})
        direction = str(
            observation.get("optimization_direction")
            or dict(source_config.get("acquisition") or {}).get("optimization_direction")
            or "maximize"
        )

        def classic_components(data, metrics, key=None):
            components = data.get(key) if key else data
            if not isinstance(components, dict) or "peak_prominence_contribution" not in components:
                # Older session records predate stored term contributions. Rebuild
                # them from the retained per-channel metrics for display.
                components = compute_channel_quality(
                    metrics,
                    scoring,
                    "maximize" if paired else direction,
                )
            return components

        def classic_values(components):
            return (
                self._fmt(components.get("Q_channel")),
                self._fmt(components.get("peak_prominence_contribution")),
                self._fmt(components.get("repeat_scan_snr_contribution")),
                self._fmt(components.get("peak_height_contribution")),
                self._fmt(components.get("peak_shape_contribution")),
                self._fmt(components.get("baseline_contribution")),
                self._fmt(components.get("replicate_consistency_contribution")),
                self._fmt(components.get("success_contribution")),
                self._fmt(components.get("noise_penalty_adjustment")),
                self._fmt(components.get("clip_adjustment")),
            )

        for ch, data in sorted(components.items(), key=lambda item: int(item[0])):
            metrics = channel_metrics.get(str(ch), {}) if isinstance(channel_metrics, dict) else {}
            if paired:
                buffer_metrics = buffer_channel_metrics.get(str(ch), {}) if isinstance(buffer_channel_metrics, dict) else {}
                target_metrics = target_channel_metrics.get(str(ch), {}) if isinstance(target_channel_metrics, dict) else metrics
                paired_q = self._fmt(data.get("paired_Q_channel", data.get("Q_channel")))
                delta_peak = self._fmt(data.get("delta_peak_height_uA"))
                paired_prominence = self._fmt(data.get("peak_prominence"))
                paired_repeat_snr = self._fmt(data.get("repeat_scan_snr"))
                buffer_classic = classic_components(
                    data, buffer_metrics, "buffer_classic_components"
                )
                target_classic = classic_components(
                    data, target_metrics, "target_classic_components"
                )
                row_specs = (
                    (
                        f"{ch}_buffer",
                        (
                            "Buffer",
                            *classic_values(buffer_classic),
                            self._fmt(data.get("classic_pair_Q")),
                            self._fmt(data.get("buffer_classic_Q_contribution")),
                            self._fmt(data.get("target_classic_Q_contribution")),
                            self._fmt(data.get("peak_prominence_contribution")),
                            self._fmt(data.get("repeat_scan_snr_contribution")),
                            self._fmt(self._channel_peak_height(buffer_metrics)),
                            self._fmt(data.get("buffer_peak_prominence_raw", data.get("buffer_snr_raw"))),
                            self._fmt(data.get("buffer_peak_prominence_score", data.get("buffer_snr_score"))),
                            "",
                            self._fmt(data.get("success_score")),
                            paired_q,
                            delta_peak,
                            paired_prominence,
                            paired_repeat_snr,
                        ),
                    ),
                    (
                        f"{ch}_target",
                        (
                            "Target",
                            *classic_values(target_classic),
                            self._fmt(data.get("classic_pair_Q")),
                            self._fmt(data.get("buffer_classic_Q_contribution")),
                            self._fmt(data.get("target_classic_Q_contribution")),
                            self._fmt(data.get("peak_prominence_contribution")),
                            self._fmt(data.get("repeat_scan_snr_contribution")),
                            self._fmt(self._channel_peak_height(target_metrics)),
                            self._fmt(data.get("target_peak_prominence_raw", data.get("target_snr_raw"))),
                            self._fmt(data.get("target_peak_prominence_score", data.get("target_snr_score"))),
                            self._fmt(data.get("target_shape_score")),
                            self._fmt(data.get("success_score")),
                            paired_q,
                            delta_peak,
                            paired_prominence,
                            paired_repeat_snr,
                        ),
                    ),
                )
                for iid, values in row_specs:
                    self._score_tree.insert(
                        "",
                        "end",
                        iid=iid,
                        text=str(ch),
                        values=values,
                    )
            else:
                classic = classic_components(data, metrics)
                values = (
                    *classic_values(classic),
                    self._fmt(self._channel_peak_height(metrics)),
                    self._fmt(data.get("peak_prominence_raw", data.get("snr_raw"))),
                    self._fmt(data.get("repeat_scan_snr_raw")),
                    self._fmt(data.get("normalized_peak_prominence", data.get("normalized_SNR"))),
                    self._fmt(data.get("peak_shape_score")),
                    self._fmt(data.get("baseline_stability_score")),
                    self._fmt(data.get("replicate_consistency_score")),
                    self._fmt(data.get("success_score")),
                )
                self._score_tree.insert(
                    "",
                    "end",
                    iid=str(ch),
                    text=str(ch),
                    values=values,
                )

    @staticmethod
    def _is_paired_observation(observation):
        quality = dict((observation or {}).get("quality") or {})
        return str((observation or {}).get("objective") or quality.get("objective") or "").lower() == "paired_response"

    def _configure_history_table(self, paired=False):
        if not hasattr(self, "_history_tree"):
            return
        if paired:
            columns = (
                "Group", "Set", "BO Iter", "Buffer Trace", "Target Trace",
                "Q_run", "Paired Q", "Buffer Q", "Target Q", "Classic Pair Q",
                "Buffer Term", "Target Term", "Prominence Term", "Repeat-SNR Term",
                "Delta Peak", "Paired Prominence", "Paired Repeat SNR", "Frac Delta", "Distance",
                "Buffer Prominence", "Target Prominence", "Buffer Noise", "Target Noise", "Combined Noise",
                "Target Shape", "Success",
                "Begin", "End", "Step", "Amp", "Freq", "Cond E", "Cond t",
            )
            self._history_tree.configure(columns=columns)
            self._history_tree.heading("#0", text="Cycle")
            self._history_tree.column("#0", width=62, anchor="center", stretch=False)
            widths = {
                "Group": 90,
                "Buffer Trace": 96,
                "Target Trace": 96,
                "Paired Q": 82,
                "Buffer Term": 92,
                "Target Term": 92,
                "Prominence Term": 110,
                "Repeat-SNR Term": 108,
                "Delta Peak": 88,
                "Classic Pair Q": 104,
                "Buffer Prominence": 112,
                "Target Prominence": 112,
                "Buffer Noise": 94,
                "Target Noise": 94,
                "Combined Noise": 106,
                "BO Iter": 78,
            }
        else:
            columns = (
                "Group", "Q_run", "Mean", "Std", "Failed", "Poor",
                "Peak uA", "Noise uA", "Peak Prominence", "Repeat-scan SNR", "Prominence Score", "Shape", "Baseline", "Replicate", "Success",
                "Begin", "End", "Step", "Amp", "Freq", "Cond E", "Cond t",
            )
            self._history_tree.configure(columns=columns)
            self._history_tree.heading("#0", text="Iter")
            self._history_tree.column("#0", width=55, anchor="center", stretch=False)
            widths = {"Group": 90}
        for col in columns:
            self._history_tree.heading(col, text=col)
            self._history_tree.column(col, width=widths.get(col, 76), anchor="center", stretch=False)

    def _paired_history_values(self, obs):
        params = obs.get("params", {})
        quality = dict(obs.get("quality") or {})
        truth = dict(obs.get("simulation_truth") or {})
        paired_q = truth.get("paired_Q_score")
        delta_peak = quality.get("mean_abs_delta_peak_height_uA", truth.get("expected_delta_peak_uA"))
        batch_size = self._paired_batch_size_for_observation(obs)
        paired_batch_index = self._paired_batch_index_for_observation(obs, batch_size=batch_size)
        return (
            str(obs.get("group_name") or f"Group {int(obs.get('group_id', 1))}"),
            str(paired_batch_index) if paired_batch_index is not None else "",
            str(obs.get("iteration") or ""),
            self._string_or_empty(obs.get("buffer_trace_number")),
            self._string_or_empty(obs.get("target_trace_number")),
            self._fmt(obs.get("Q_run")),
            self._fmt(paired_q if paired_q is not None else quality.get("mean_paired_Q_channel", quality.get("mean_Q_channel"))),
            self._fmt(quality.get("mean_buffer_classic_Q")),
            self._fmt(quality.get("mean_target_classic_Q")),
            self._fmt(quality.get("mean_classic_pair_Q")),
            self._fmt(quality.get("mean_buffer_classic_Q_contribution")),
            self._fmt(quality.get("mean_target_classic_Q_contribution")),
            self._fmt(quality.get("mean_peak_prominence_contribution", quality.get("mean_delta_peak_contribution"))),
            self._fmt(quality.get("mean_repeat_scan_snr_contribution")),
            self._fmt(delta_peak),
            self._fmt(quality.get("mean_peak_prominence", quality.get("mean_delta_peak_score"))),
            self._fmt(quality.get("mean_repeat_scan_snr")),
            self._fmt(quality.get("mean_fractional_delta_peak")),
            self._fmt(truth.get("normalized_distance")),
            self._fmt(quality.get("mean_buffer_peak_prominence", quality.get("mean_buffer_snr_raw"))),
            self._fmt(quality.get("mean_target_peak_prominence", quality.get("mean_target_snr_raw"))),
            self._fmt(quality.get("mean_buffer_channel_noise")),
            self._fmt(quality.get("mean_target_channel_noise")),
            self._fmt(quality.get("mean_combined_channel_noise")),
            self._fmt(quality.get("mean_target_shape_score")),
            self._fmt(quality.get("mean_success_score")),
            self._fmt_raw(params.get("begin_potential")),
            self._fmt_raw(params.get("end_potential")),
            self._fmt_raw(params.get("step_potential")),
            self._fmt_raw(params.get("amplitude")),
            self._fmt_raw(params.get("frequency")),
            self._fmt_raw(params.get("conditioning_potential")),
            self._fmt_raw(params.get("conditioning_time")),
        )

    @staticmethod
    def _string_or_empty(value):
        return "" if value is None else str(value)

    def _paired_batch_size_for_observation(self, obs) -> int:
        for source in (
            dict((obs or {}).get("quality") or {}),
            dict(getattr(self._bo_session, "config", {}) or {}) if self._bo_session is not None else {},
            dict(self._config or {}),
        ):
            try:
                value = source.get("paired_batch_size")
                if value is not None and value != "":
                    return max(1, int(value))
            except Exception:
                pass
        try:
            return max(1, int(self._paired_batch_size_var.get() or 1))
        except Exception:
            return 1

    def _paired_cycle_for_observation(self, obs, batch_size: int | None = None):
        value = (obs or {}).get("paired_cycle")
        if value is not None and value != "":
            return value
        try:
            iteration = int((obs or {}).get("iteration") or 0)
        except Exception:
            return None
        if iteration < 1:
            return None
        batch = max(1, int(batch_size or self._paired_batch_size_for_observation(obs)))
        return ((iteration - 1) // batch) + 1

    def _paired_batch_index_for_observation(self, obs, batch_size: int | None = None):
        value = (obs or {}).get("paired_batch_index")
        if value is not None and value != "":
            return value
        try:
            iteration = int((obs or {}).get("iteration") or 0)
        except Exception:
            return None
        if iteration < 1:
            return None
        batch = max(1, int(batch_size or self._paired_batch_size_for_observation(obs)))
        return ((iteration - 1) % batch) + 1

    def _refresh_history(self):
        self._refresh_current_q_equation()
        for row in self._history_tree.get_children():
            self._history_tree.delete(row)
        self._history_rows = {}
        self._selected_history_observation = None
        for row in self._score_tree.get_children():
            self._score_tree.delete(row)
        if self._bo_session is None:
            if hasattr(self, "_q_equation_text"):
                self._write_text(self._q_equation_text, "\n".join(self._q_equation_lines(self._config)))
            self._render_raw_traces(None)
            self._render_corrected_traces(None)
            self._refresh_analysis_q_trend()
            self._refresh_surrogate_view()
            return
        paired_history = any(self._is_paired_observation(obs) for obs in self._bo_session.observations)
        self._configure_history_table(paired=paired_history)
        for obs in self._bo_session.observations:
            q = obs.get("quality", {})
            params = obs.get("params", {})
            iteration = str(obs.get("iteration"))
            history_key = f"g{int(obs.get('group_id', 1))}:i{iteration}"
            peak_uA, rms_uA = self._observation_peak_rms(obs)
            prominence_raw = self._observation_component_mean(obs, "peak_prominence_raw")
            repeat_scan_snr = self._observation_component_mean(obs, "repeat_scan_snr_raw")
            prominence_score = self._observation_component_mean(obs, "normalized_peak_prominence")
            shape_score = self._observation_component_mean(obs, "peak_shape_score")
            baseline_score = self._observation_component_mean(obs, "baseline_stability_score")
            replicate_score = self._observation_component_mean(obs, "replicate_consistency_score")
            success_score = self._observation_component_mean(obs, "success_score")
            self._history_rows[history_key] = obs
            if paired_history:
                values = self._paired_history_values(obs)
                cycle = self._paired_cycle_for_observation(obs)
                text = str(cycle) if cycle is not None else ""
            else:
                values = (
                    str(obs.get("group_name") or f"Group {int(obs.get('group_id', 1))}"),
                    self._fmt(obs.get("Q_run")),
                    self._fmt(q.get("mean_Q_channel")),
                    self._fmt(q.get("std_Q_channel")),
                    self._fmt(q.get("failed_channel_fraction")),
                    self._fmt(q.get("low_channel_fraction")),
                    self._fmt(peak_uA),
                    self._fmt(rms_uA),
                    self._fmt(prominence_raw),
                    self._fmt(repeat_scan_snr),
                    self._fmt(prominence_score),
                    self._fmt(shape_score),
                    self._fmt(baseline_score),
                    self._fmt(replicate_score),
                    self._fmt(success_score),
                    self._fmt_raw(params.get("begin_potential")),
                    self._fmt_raw(params.get("end_potential")),
                    self._fmt_raw(params.get("step_potential")),
                    self._fmt_raw(params.get("amplitude")),
                    self._fmt_raw(params.get("frequency")),
                    self._fmt_raw(params.get("conditioning_potential")),
                    self._fmt_raw(params.get("conditioning_time")),
                )
                text = iteration
            self._history_tree.insert(
                "",
                "end",
                iid=history_key,
                text=text,
                values=values,
            )
        if not self._history_rows:
            if hasattr(self, "_q_equation_text"):
                self._write_text(self._q_equation_text, "\n".join(self._q_equation_lines(self._bo_session.config)))
            self._render_raw_traces(None)
            self._render_corrected_traces(None)
        self._refresh_analysis_q_trend()
        self._refresh_surrogate_controls()

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
        if self._measurement_priority_active():
            self._render_raw_traces(obs)
            self._render_corrected_traces(obs)
            self._status_var.set(
                f"Selected BO iteration {obs.get('iteration')}; trace plots are deferred while a measurement is collecting data."
            )
            self._restore_tree_focus(self._history_tree)
            return
        self._render_raw_traces(obs)
        self._render_corrected_traces(obs)
        artifact_iterations = self._surrogate_artifact_iterations()
        try:
            selected_iteration = int(obs.get("iteration"))
        except Exception:
            selected_iteration = None
        if selected_iteration in artifact_iterations:
            self._surrogate_iteration_var.set(str(selected_iteration))
            self._refresh_surrogate_view()
        self._status_var.set(
            f"Viewing BO iteration {obs.get('iteration')}: Q_run={float(obs.get('Q_run', 0.0)):.3f}"
        )
        self._restore_tree_focus(self._history_tree)

    def _on_score_tree_select(self):
        if self._measurement_priority_active():
            self._render_raw_traces(self._selected_history_observation)
            self._render_corrected_traces(self._selected_history_observation)
            self._restore_tree_focus(self._score_tree)
            return
        self._render_raw_traces(self._selected_history_observation)
        self._render_corrected_traces(self._selected_history_observation)
        self._restore_tree_focus(self._score_tree)

    def _select_latest_history_iteration(self):
        if not self._history_rows:
            return
        latest = next(reversed(self._history_rows))
        self._history_tree.selection_set(latest)
        self._history_tree.focus(latest)
        self._history_tree.see(latest)
        self._select_history_iteration(latest)

    def _resolve_history_key(self, value):
        """Resolve either an opaque tree row ID or a legacy numeric iteration."""
        if value is None:
            return None
        key = str(value)
        if key in self._history_rows:
            return key
        matches = [
            row_key
            for row_key, observation in self._history_rows.items()
            if str(observation.get("iteration")) == key
        ]
        return matches[-1] if matches else None

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
        if self._measurement_priority_active():
            self._render_raw_traces(self._selected_history_observation)
            self._render_corrected_traces(self._selected_history_observation)
            return "break"
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

    def _history_metric_for_column(self, column_id):
        try:
            index = int(str(column_id).lstrip("#")) - 1
        except ValueError:
            return None
        labels = tuple(self._history_tree["columns"]) if hasattr(self, "_history_tree") else ()
        metrics = {
            "Q_run": "Q_run",
            "Mean": "Mean channel Q",
            "Std": "Std channel Q",
            "Failed": "Failed fraction",
            "Poor": "Poor fraction",
            "Paired Q": "Paired Q",
            "Buffer Q": "Buffer classic Q",
            "Target Q": "Target classic Q",
            "Classic Pair Q": "Classic pair Q",
            "Delta Peak": "Delta Peak",
            "Frac Delta": "Fractional delta peak",
            "Distance": "Distance",
            "Buffer Prominence": "Buffer peak prominence",
            "Target Prominence": "Target peak prominence",
            "Target Shape": "Target shape score",
            "Peak uA": "Mean peak uA",
            "Noise uA": "Mean noise uA",
            "Peak Prominence": "Mean peak prominence",
            "Prominence Score": "Mean prominence score",
            "Shape": "Mean shape score",
            "Baseline": "Mean baseline score",
            "Replicate": "Mean replicate score",
            "Success": "Mean success score",
            "BO Iter": "BO iteration",
            "Set": "Parameter set",
            "Buffer Trace": "Buffer trace",
            "Target Trace": "Target trace",
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
        if self._defer_results_render(
            self._raw_trace_frame,
            "Raw trace plotting is paused while a measurement is collecting data.\nAcquisition has priority.",
        ):
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
        selected_channels, selected_phases = self._selected_score_filters()
        if selected_channels:
            rows = [
                row for row in rows
                if str(row.get("channel") or "") in selected_channels
            ]
        if selected_phases:
            rows = [
                row for row in rows
                if str(row.get("phase") or "").strip().lower() in selected_phases
            ]
        if not rows:
            if selected_channels:
                message = (
                    "No raw SWV CSV paths were found for the selected channel(s) "
                    f"{', '.join(sorted(selected_channels, key=self._channel_sort_key))} in this iteration."
                )
                if selected_phases:
                    message += f" Requested phase(s): {', '.join(sorted(selected_phases))}."
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
            path = Path(row.get("path") or "")
            if row.get("voltage") is not None or row.get("current") is not None:
                volts = row.get("voltage") or []
                currents = row.get("current") or []
            else:
                try:
                    volts, currents = load_swv_csv(str(path))
                except Exception as exc:
                    errors.append(f"{path.name}: {exc}")
                    continue
            volts = self._to_float_list(volts)
            currents = self._to_float_list(currents)
            n = min(len(volts), len(currents))
            if n <= 0:
                continue
            if len(volts) != len(currents):
                errors.append(
                    f"{path.name}: trimmed mismatched raw trace lengths from "
                    f"({len(volts)}, {len(currents)}) to ({n}, {n})"
                )
            volts = volts[:n]
            currents = currents[:n]
            channel = row.get("channel")
            scan = row.get("scan")
            label = f"Ch {channel}" if channel not in (None, "") else (path.stem if str(path) not in ("", ".") else row.get("label", "trace"))
            if scan not in (None, ""):
                label = f"{label} scan {scan}"
            phase = row.get("phase")
            trace_number = row.get("trace_number")
            if phase:
                label = f"{phase} {label}"
            if trace_number not in (None, ""):
                label = f"Trace {trace_number} {label}"
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
        if selected_phases:
            channel_suffix += f" | {'/'.join(sorted(selected_phases))}"
        title_prefix = f"Cycle {observation.get('paired_cycle')} set {observation.get('paired_batch_index')}" if str(observation.get("objective") or "").lower() == "paired_response" else f"Iteration {observation.get('iteration')}"
        ax.set_title(f"{title_prefix} raw SWV traces{channel_suffix} | Q_run={self._fmt(observation.get('Q_run'))}")
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
                text="\n".join(errors[:8]),
                foreground="#8a5a44",
                justify="left",
                anchor="w",
            ).pack(fill="x")

    def _render_corrected_traces(self, observation):
        if not hasattr(self, "_corrected_trace_frame"):
            return
        if self._defer_results_render(
            self._corrected_trace_frame,
            "Corrected trace recomputation is paused while a measurement is collecting data.\nAcquisition has priority.",
        ):
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

        rows, diagnostics = self._corrected_trace_rows_for_observation(observation)
        selected_channels, selected_phases = self._selected_score_filters()
        if selected_channels:
            rows = [row for row in rows if str(row.get("channel") or "") in selected_channels]
            diagnostics = [d for d in diagnostics if str(d.get("channel") or "") in selected_channels]
        if selected_phases:
            rows = [row for row in rows if str(row.get("phase") or "").strip().lower() in selected_phases]
            diagnostics = [d for d in diagnostics if str(d.get("phase") or "").strip().lower() in selected_phases]
        if not rows:
            message = (
                "No corrected traces could be recomputed for this iteration.\n\n"
                "The corrected-trace panel now recomputes from archived raw SWV files "
                "using the current Setup-tab analysis settings."
            )
            if selected_channels:
                message = (
                    "No corrected traces could be recomputed for the selected channel(s) "
                    f"{', '.join(sorted(selected_channels, key=self._channel_sort_key))} in this iteration."
                )
                if selected_phases:
                    message += f" Requested phase(s): {', '.join(sorted(selected_phases))}."
            detail_lines = self._corrected_trace_diagnostic_lines(diagnostics)
            if detail_lines:
                message += "\n\nWhy recomputation failed:\n" + "\n".join(detail_lines[:12])
            ttk.Label(
                self._corrected_trace_frame,
                text=message,
                anchor="center",
                justify="left",
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
            phase = row.get("phase")
            trace_number = row.get("trace_number")
            if phase:
                label = f"{phase} {label}"
            if trace_number not in (None, ""):
                label = f"Trace {trace_number} {label}"
            ax.plot(
                row["voltage"],
                row["current"],
                linewidth=1.1,
                alpha=0.88,
                color=palette[idx % len(palette)],
                label=label,
            )
            marker_color = palette[idx % len(palette)]
            for key in ("left_min_idx", "right_min_idx"):
                point_idx = row.get(key)
                if point_idx is None:
                    continue
                try:
                    point_idx = int(point_idx)
                except (TypeError, ValueError):
                    continue
                if 0 <= point_idx < len(row["voltage"]) and 0 <= point_idx < len(row["current"]):
                    ax.scatter(
                        [row["voltage"][point_idx]],
                        [row["current"][point_idx]],
                        color=marker_color,
                        edgecolors="black",
                        linewidths=0.6,
                        s=34,
                        zorder=5,
                    )
            peak_idx = row.get("peak_idx_corr")
            try:
                peak_idx = int(peak_idx) if peak_idx is not None else None
            except (TypeError, ValueError):
                peak_idx = None
            if peak_idx is not None and 0 <= peak_idx < len(row["voltage"]) and 0 <= peak_idx < len(row["current"]):
                ax.scatter(
                    [row["voltage"][peak_idx]],
                    [row["current"][peak_idx]],
                    marker="x",
                    color="#c1121f",
                    linewidths=1.4,
                    s=52,
                    zorder=6,
                )

        channel_suffix = ""
        if selected_channels:
            channel_suffix = f" | Ch {', '.join(sorted(selected_channels, key=self._channel_sort_key))}"
        if selected_phases:
            channel_suffix += f" | {'/'.join(sorted(selected_phases))}"
        title_prefix = f"Cycle {observation.get('paired_cycle')} set {observation.get('paired_batch_index')}" if str(observation.get("objective") or "").lower() == "paired_response" else f"Iteration {observation.get('iteration')}"
        ax.set_title(f"{title_prefix} smoothed corrected traces{channel_suffix}", fontsize=9, pad=8, wrap=True)
        ax.set_xlabel("Voltage (V)", fontsize=9, labelpad=4)
        ax.set_ylabel("Corrected current (uA)", fontsize=9, labelpad=4)
        ax.tick_params(labelsize=8)
        ax.grid(alpha=0.25)
        if len(rows) <= 16:
            ax.legend(loc="best", fontsize=7)
        self._fit_embedded_figure(fig, top=0.84, bottom=0.22, left=0.17, right=0.96)
        canvas = FigureCanvasTkAgg(fig, master=self._corrected_trace_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        detail_lines = self._corrected_trace_diagnostic_lines(diagnostics)
        if detail_lines:
            ttk.Label(
                self._corrected_trace_frame,
                text="Some traces could not be fully recomputed:\n" + "\n".join(detail_lines[:8]),
                anchor="w",
                justify="left",
                foreground="#8a5a44",
            ).pack(fill="x", pady=(4, 0))

    def _corrected_trace_rows_for_observation(self, observation):
        rows = []
        diagnostics = []
        seen = set()
        selected_channels = self._observation_channel_filter(observation)
        external_results = self._external_analysis_results(observation)
        if external_results:
            for result in external_results:
                channel = self._normalize_observation_channel(result.get("channel"))
                if selected_channels and channel not in selected_channels:
                    continue
                scan = result.get("scan_number", result.get("scan_id_from_name"))
                label = str(result.get("file_name") or Path(str(result.get("file_path") or "trace")).name)
                voltage = self._to_float_list(result.get("voltage"))
                current = self._to_float_list(result.get("smoothed_corrected_current"))
                stage = "smoothed corrected"
                if not current:
                    current = self._to_float_list(result.get("corrected_current"))
                    stage = "corrected"
                if not voltage or not current:
                    diagnostics.append(
                        {
                            "label": label,
                            "channel": channel,
                            "scan": scan,
                            "reason": result.get("error") or result.get("partial_error") or "External analysis returned no corrected trace.",
                        }
                    )
                    continue
                n = min(len(voltage), len(current))
                rows.append(
                    {
                        "voltage": voltage[:n],
                        "current": current[:n],
                        "channel": channel,
                        "scan": scan,
                        "label": label,
                        "phase": result.get("_bo_phase"),
                        "trace_stage": stage,
                        "left_min_idx": result.get("left_min_idx"),
                        "right_min_idx": result.get("right_min_idx"),
                        "peak_idx_corr": result.get("peak_idx_corr"),
                    }
                )
            return self._sort_corrected_trace_rows(rows, diagnostics)
        analysis_cfg = self._current_analysis_settings()
        for raw_row in self._raw_trace_rows_for_observation(observation):
            file_path = self._resolve_observation_file_path(raw_row.get("path"), observation)
            if raw_row.get("voltage") is not None or raw_row.get("current") is not None:
                v_raw = raw_row.get("voltage") or []
                i_raw = raw_row.get("current") or []
                file_label = raw_row.get("label") or "embedded simulated trace"
            else:
                if not file_path.exists():
                    diagnostics.append(
                        {
                            "label": Path(raw_row.get("path") or "").name or "unknown trace",
                            "channel": raw_row.get("channel"),
                            "reason": "Raw SWV file could not be found.",
                        }
                    )
                    continue
                try:
                    v_raw, i_raw = load_swv_csv(str(file_path))
                except Exception as exc:
                    diagnostics.append(
                        {
                            "label": Path(file_path).name,
                            "channel": raw_row.get("channel"),
                            "reason": f"Raw SWV file could not be loaded: {exc}",
                        }
                    )
                    continue
                file_label = Path(file_path).name
            channel = self._normalize_observation_channel(
                raw_row.get("channel") or (self._infer_channel_from_path(file_path) if file_path else "")
            )
            scan = raw_row.get("scan")
            key = (str(file_path) if file_path else str(file_label), str(channel), str(scan or ""), str(raw_row.get("phase") or ""), str(raw_row.get("trace_number") or ""))
            if key in seen:
                continue
            seen.add(key)
            result = self._recompute_corrected_trace(v_raw, i_raw, analysis_cfg)
            voltage = self._to_float_list(result.get("voltage"))
            current = self._to_float_list(result.get("smoothed_corrected_current"))
            stage = "smoothed corrected"
            if not current:
                current = self._to_float_list(result.get("corrected_current"))
                if current:
                    stage = "corrected"
            if not current:
                current = self._to_float_list(result.get("smoothed_current"))
                if current:
                    stage = "smoothed raw"
            if not voltage or not current:
                diagnostics.append(
                    {
                        "label": Path(file_path).name,
                        "channel": channel,
                        "scan": scan,
                        "reason": result.get("partial_error") or "No plottable arrays were produced.",
                    }
                )
                continue
            n = min(len(voltage), len(current))
            if n <= 0:
                diagnostics.append(
                    {
                        "label": Path(file_path).name,
                        "channel": channel,
                        "scan": scan,
                        "reason": "Voltage/current arrays were empty after recomputation.",
                    }
                )
                continue
            partial_error = result.get("partial_error")
            if partial_error:
                diagnostics.append(
                    {
                        "label": Path(file_path).name,
                        "channel": channel,
                        "scan": scan,
                        "reason": partial_error,
                    }
                )
            rows.append(
                {
                    "voltage": voltage[:n],
                    "current": current[:n],
                    "channel": channel,
                    "scan": scan,
                    "label": str(raw_row.get("label") or Path(file_path).stem or file_label),
                    "phase": raw_row.get("phase"),
                    "trace_number": raw_row.get("trace_number"),
                    "trace_stage": stage,
                    "left_min_idx": result.get("left_min_idx"),
                    "right_min_idx": result.get("right_min_idx"),
                    "peak_idx_corr": result.get("peak_idx_corr"),
                }
            )
        return self._sort_corrected_trace_rows(rows, diagnostics)

    def _sort_corrected_trace_rows(self, rows, diagnostics):
        return sorted(
            rows,
            key=lambda row: (
                self._channel_sort_key(row.get("channel")),
                0 if str(row.get("phase") or "").lower() == "buffer" else 1 if str(row.get("phase") or "").lower() == "target" else 2,
                str(row.get("scan") or ""),
                str(row.get("trace_number") or ""),
                str(row.get("label") or ""),
            ),
        ), sorted(
            diagnostics,
            key=lambda row: (self._channel_sort_key(row.get("channel")), str(row.get("scan") or ""), str(row.get("label") or "")),
        )

    def _external_analysis_results(self, observation):
        paired = str((observation or {}).get("objective") or "").lower() == "paired_response"
        paired_sources = [
            (observation.get("buffer_analysis_results_json"), "buffer"),
            (observation.get("target_analysis_results_json"), "target"),
        ]
        sources = paired_sources if paired and any(path for path, _phase in paired_sources) else (
            (observation.get("analysis_results_json"), None),
            *paired_sources,
        )
        rows = []
        seen = set()
        for raw_path, phase in sources:
            if not raw_path:
                continue
            path = self._resolve_observation_file_path(raw_path, observation)
            key = str(path)
            if key in seen or not path.exists():
                continue
            seen.add(key)
            try:
                with open(path, "r", encoding="utf-8") as fh:
                    payload = json.load(fh)
            except Exception:
                continue
            if not isinstance(payload, list):
                continue
            for item in payload:
                if isinstance(item, dict):
                    row = dict(item)
                    row["_bo_phase"] = phase
                    rows.append(row)
        return rows

    def _scroll_history_horizontally(self, event):
        if not hasattr(self, "_history_tree"):
            return None
        delta = getattr(event, "delta", 0)
        if delta == 0:
            return "break"
        units = -1 if delta > 0 else 1
        self._history_tree.xview_scroll(units, "units")
        return "break"

    def _current_analysis_settings(self):
        analysis_cfg = {}
        self._update_analysis_config_from_vars(analysis_cfg)
        return analysis_cfg

    def _recompute_corrected_trace(self, v_raw, i_raw, analysis_cfg):
        kwargs = dict(
            v_raw=v_raw,
            i_raw=i_raw,
            crop_range=(
                float(analysis_cfg.get("crop_min_v", -0.61)),
                float(analysis_cfg.get("crop_max_v", -0.30)),
            ),
            smooth_window=int(analysis_cfg.get("smooth_window", 15)),
            smooth_polyorder=int(analysis_cfg.get("smooth_polyorder", 2)),
            minima_search_window_V=float(analysis_cfg.get("minima_search_window_v", 0.30)),
            use_prominent_minima=bool(analysis_cfg.get("use_prominent_minima", False)),
            require_local_minima_on_both_sides=bool(analysis_cfg.get("require_local_minima_on_both_sides", False)),
            use_double_correction=bool(analysis_cfg.get("use_double_correction", False)),
            min_peak_height_uA=analysis_cfg.get("min_peak_height_ua"),
            peak_voltage_min_V=analysis_cfg.get("peak_voltage_min_v"),
            peak_voltage_max_V=analysis_cfg.get("peak_voltage_max_v"),
            left_min_voltage_min_V=analysis_cfg.get("left_min_voltage_min_v"),
            left_min_voltage_max_V=analysis_cfg.get("left_min_voltage_max_v"),
            right_min_voltage_min_V=analysis_cfg.get("right_min_voltage_min_v"),
            right_min_voltage_max_V=analysis_cfg.get("right_min_voltage_max_v"),
            compute_skew=bool(analysis_cfg.get("compute_skew", True)),
            compute_wavelet_energy=bool(analysis_cfg.get("compute_wavelet_energy", True)),
            compute_wavelet_denoised_trace=bool(analysis_cfg.get("compute_wavelet_denoised_trace", False)),
            use_wavelet_for_correction=bool(analysis_cfg.get("use_wavelet_for_correction", False)),
        )
        try:
            result = analyze_swv_arrays(**kwargs)
        except Exception:
            result = partial_traces_for_failure_arrays(
                v_raw=v_raw,
                i_raw=i_raw,
                crop_range=kwargs["crop_range"],
                smooth_window=kwargs["smooth_window"],
                smooth_polyorder=kwargs["smooth_polyorder"],
                minima_search_window_V=kwargs["minima_search_window_V"],
                use_prominent_minima=kwargs["use_prominent_minima"],
                require_local_minima_on_both_sides=kwargs["require_local_minima_on_both_sides"],
                use_double_correction=kwargs["use_double_correction"],
                compute_wavelet_denoised_trace=kwargs["compute_wavelet_denoised_trace"],
                use_wavelet_for_correction=kwargs["use_wavelet_for_correction"],
            )
        self._normalize_trace_result_lengths(result)
        return result

    @staticmethod
    def _to_float_list(values):
        if values is None:
            return []
        out = []
        try:
            iterable = list(values)
        except Exception:
            return []
        for value in iterable:
            try:
                out.append(float(value))
            except (TypeError, ValueError):
                continue
        return out

    @staticmethod
    def _corrected_trace_diagnostic_lines(diagnostics):
        lines = []
        for item in diagnostics or []:
            label = str(item.get("label") or "trace")
            channel = item.get("channel")
            scan = item.get("scan")
            prefix = label
            if channel not in (None, ""):
                prefix = f"Ch {channel} | {prefix}"
            if scan not in (None, ""):
                prefix = f"{prefix} | scan {scan}"
            reason = str(item.get("reason") or "Unknown error")
            lines.append(f"- {prefix}: {reason}")
        return lines

    @staticmethod
    def _normalize_trace_result_lengths(result):
        if not isinstance(result, dict):
            return
        voltage = BayesianOptimizationTab._to_float_list(result.get("voltage"))
        if not voltage:
            return
        result["voltage"] = voltage
        for key in (
            "raw_current",
            "smoothed_current",
            "wavelet_denoised_current",
            "corrected_current",
            "smoothed_corrected_current",
            "local_baseline",
            "first_pass_corrected_current",
            "first_pass_smoothed_corrected_current",
            "first_pass_local_baseline",
            "second_pass_corrected_current",
            "second_pass_smoothed_corrected_current",
            "second_pass_local_baseline",
        ):
            values = BayesianOptimizationTab._to_float_list(result.get(key))
            if not values:
                continue
            n = min(len(voltage), len(values))
            if n <= 0:
                result[key] = []
                continue
            if n != len(voltage):
                result["voltage"] = voltage[:n]
                voltage = result["voltage"]
            result[key] = values[:n]

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

    @staticmethod
    def _minima_bracket_rms_noise_from_row(row):
        raw_current = BayesianOptimizationTab._parse_trace_array(row.get("raw_current"))
        if len(raw_current) < 2:
            return None, None, None
        try:
            left = int(float(row.get("left_min_idx")))
            right = int(float(row.get("right_min_idx")))
        except (TypeError, ValueError):
            return None, None, len(raw_current)
        if left < 0 or right < 0:
            return None, None, len(raw_current)
        lo = max(0, min(left, right))
        hi = min(len(raw_current) - 1, max(left, right))
        segment = raw_current[lo:hi + 1]
        if len(segment) < 2:
            return None, len(segment), len(raw_current)
        diffs = [segment[idx + 1] - segment[idx] for idx in range(len(segment) - 1)]
        noise = math.sqrt(sum(diff * diff for diff in diffs) / len(diffs)) / math.sqrt(2.0)
        return noise, len(segment), len(raw_current)

    def _rebuilt_channel_metrics_for_observation(self, observation):
        rows = []
        selected_channels = self._observation_channel_filter(observation)
        for path in self._analysis_results_paths_for_observation(observation):
            if not path.exists():
                continue
            try:
                with open(path, "r", encoding="utf-8-sig", newline="") as fh:
                    for row in csv.DictReader(fh):
                        row = dict(row)
                        if selected_channels:
                            channel = self._normalize_observation_channel(row.get("channel"))
                            if channel is None or channel not in selected_channels:
                                continue
                        noise, bracket_count, crop_count = self._minima_bracket_rms_noise_from_row(row)
                        if noise is not None:
                            row["background_current_rms"] = noise
                        if bracket_count is not None:
                            row["bracket_point_count"] = bracket_count
                        if crop_count is not None:
                            row["crop_point_count"] = crop_count
                        rows.append(row)
            except Exception:
                continue
        if not rows:
            return None
        return _build_channel_metrics(rows)

    def _analysis_results_paths_for_observation(self, observation):
        paths = []
        for analysis_record, _phase_hint in self._analysis_record_paths_with_phase(observation):
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
        return paths

    def _analysis_record_paths_for_observation(self, observation):
        return [path for path, _phase_hint in self._analysis_record_paths_with_phase(observation)]

    def _analysis_record_paths_with_phase(self, observation):
        paths = []
        seen = set()
        paired = str((observation or {}).get("objective") or "").lower() == "paired_response"
        paired_pairs = (
            ("buffer_analysis_record", "buffer"),
            ("target_analysis_record", "target"),
        )
        key_phase_pairs = paired_pairs if paired and any((observation or {}).get(key) for key, _phase in paired_pairs) else (
            *paired_pairs,
            ("analysis_record", ""),
        )
        for key, phase_hint in key_phase_pairs:
            raw_path = (observation or {}).get(key)
            if not raw_path:
                continue
            path = self._resolve_observation_file_path(raw_path, observation)
            if not path.exists():
                continue
            norm = str(path)
            if norm in seen:
                continue
            seen.add(norm)
            inferred_phase = phase_hint or self._infer_measurement_phase_from_path(path)
            paths.append((path, inferred_phase))
        return paths

    def _selected_score_filters(self):
        if not hasattr(self, "_score_tree"):
            return set(), set()
        selected_channels = set()
        selected_phases = set()
        for item in self._score_tree.selection():
            channel = self._normalize_observation_channel(self._score_tree.item(item, "text") or item)
            if channel:
                selected_channels.add(channel)
            values = self._score_tree.item(item, "values") or ()
            if values:
                phase = str(values[0] or "").strip().lower()
                if phase in ("buffer", "target"):
                    selected_phases.add(phase)
        return selected_channels, selected_phases

    def _raw_trace_rows_for_observation(self, observation):
        rows = []
        seen = set()
        paired = str((observation or {}).get("objective") or "").lower() == "paired_response"
        selected_channels = self._observation_channel_filter(observation)

        def add_embedded_traces(traces, phase=None, trace_number=None):
            if not isinstance(traces, dict):
                return
            for channel, trace in traces.items():
                if not isinstance(trace, dict):
                    continue
                normalized_channel = self._normalize_observation_channel(channel)
                if selected_channels and normalized_channel not in selected_channels:
                    continue
                voltage = self._to_float_list(trace.get("voltage_v") or trace.get("voltage") or [])
                current = self._to_float_list(trace.get("current_uA") or trace.get("current") or [])
                if not voltage or not current:
                    continue
                key = ("embedded", str(phase or ""), str(trace_number or ""), str(normalized_channel or channel))
                if key in seen:
                    continue
                seen.add(key)
                phase_label = str(phase or "simulated")
                trace_label = f"{phase_label} ch {normalized_channel or channel}"
                rows.append(
                    {
                        "path": "",
                        "voltage": voltage,
                        "current": current,
                        "channel": normalized_channel or str(channel),
                        "scan": "",
                        "phase": phase_label,
                        "trace_number": trace_number,
                        "label": trace_label,
                    }
                )

        def add(path, channel=None, scan=None, phase=None, trace_number=None):
            if not path:
                return
            p = self._resolve_observation_file_path(path, observation)
            phase_label = str(phase or self._infer_measurement_phase_from_path(p) or "")
            trace_no = trace_number
            if trace_no in (None, ""):
                if phase_label == "buffer":
                    trace_no = observation.get("buffer_trace_number")
                elif phase_label == "target":
                    trace_no = observation.get("target_trace_number")
            normalized_channel = self._normalize_observation_channel(
                channel if channel not in (None, "") else self._infer_channel_from_path(p)
            )
            if selected_channels and normalized_channel not in selected_channels:
                return
            key = (str(p), str(phase_label), str(trace_no or ""))
            if key in seen or not p.exists() or not p.is_file():
                return
            seen.add(key)
            rows.append(
                {
                    "path": str(p),
                    "channel": normalized_channel or "",
                    "scan": scan,
                    "phase": phase_label,
                    "trace_number": trace_no,
                    "measurement_id": self._infer_measurement_id_from_path(p),
                }
            )

        if str((observation or {}).get("objective") or "").lower() == "paired_response":
            add_embedded_traces(
                observation.get("buffer_swv_trace_preview") or {},
                phase="buffer",
                trace_number=observation.get("buffer_trace_number"),
            )
            add_embedded_traces(
                observation.get("target_swv_trace_preview") or observation.get("swv_trace_preview") or {},
                phase="target",
                trace_number=observation.get("target_trace_number"),
            )
        else:
            add_embedded_traces(observation.get("swv_trace_preview") or {}, phase="simulated", trace_number=observation.get("iteration"))

        archived_measurements = list(observation.get("archived_measurements") or [])
        for path in archived_measurements:
            add(path)

        for analysis_record, phase_hint in self._analysis_record_paths_with_phase(observation):
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

            for results_path in results_paths:
                if not results_path.exists():
                    continue
                try:
                    with open(results_path, "r", encoding="utf-8-sig", newline="") as fh:
                        for row in csv.DictReader(fh):
                            add(
                                row.get("file_path") or row.get("file_name"),
                                channel=row.get("channel"),
                                scan=row.get("scan_number"),
                                phase=phase_hint,
                            )
                except Exception:
                    pass

        if paired:
            rows = self._filter_paired_rows_to_matching_measurements(rows)

        return sorted(
            rows,
            key=lambda row: (
                self._channel_sort_key(row.get("channel")),
                0 if str(row.get("phase") or "").lower() == "buffer" else 1 if str(row.get("phase") or "").lower() == "target" else 2,
                str(row.get("scan") or ""),
                str(row.get("trace_number") or ""),
                Path(row.get("path") or "").name,
            ),
        )

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

            basename = Path(normalized).name
            if basename:
                for archived_path in (observation or {}).get("archived_measurements") or []:
                    candidate = Path(str(archived_path))
                    if candidate.name == basename and candidate.exists():
                        return candidate

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

    @staticmethod
    def _infer_channel_from_path(path):
        match = re.search(r"(?:^|[_\-\s])ch(?:annel)?\s*0*(\d+)(?:\D|$)", Path(path).stem, re.IGNORECASE)
        return match.group(1) if match else ""

    @staticmethod
    def _infer_measurement_phase_from_path(path):
        parts = [Path(path).name.lower()]
        try:
            parts.extend(part.lower() for part in Path(path).parts)
        except Exception:
            pass
        for part in parts:
            if "buffer" in part:
                return "buffer"
            if "target" in part:
                return "target"
        return ""

    @staticmethod
    def _infer_measurement_id_from_path(path):
        stem = Path(path).stem
        match = re.search(r"swv_ch\d+_([0-9a-f]{4,})_meas", stem, re.IGNORECASE)
        return match.group(1).lower() if match else ""

    def _filter_paired_rows_to_matching_measurements(self, rows):
        phase_groups = {"buffer": [], "target": []}
        for row in rows:
            phase = str(row.get("phase") or "").strip().lower()
            if phase in phase_groups:
                phase_groups[phase].append(row)
        non_empty = {phase: group for phase, group in phase_groups.items() if group}
        if len(non_empty) < 2:
            return rows
        reference_phase, reference_rows = min(non_empty.items(), key=lambda item: len(item[1]))
        preferred_ids = {
            str(row.get("measurement_id") or "").strip().lower()
            for row in reference_rows
            if str(row.get("measurement_id") or "").strip()
        }
        if not preferred_ids:
            return rows
        filtered = []
        for row in rows:
            phase = str(row.get("phase") or "").strip().lower()
            measurement_id = str(row.get("measurement_id") or "").strip().lower()
            if phase in non_empty and len(non_empty.get(phase, [])) > len(reference_rows):
                if measurement_id and measurement_id not in preferred_ids:
                    continue
            filtered.append(row)
        # Exact path duplicates are already removed while rows are assembled.
        # Do not collapse by measurement id here: repeats created from the same
        # BO method intentionally share that id and must remain independently
        # visible in Results & Records.
        return filtered

    def _collapse_paired_phase_duplicates(self, rows):
        grouped = {}
        for row in rows:
            phase = str(row.get("phase") or "").strip().lower()
            measurement_id = str(row.get("measurement_id") or "").strip().lower()
            if phase not in {"buffer", "target"} or not measurement_id:
                grouped.setdefault(None, []).append(row)
                continue
            grouped.setdefault((phase, measurement_id), []).append(row)

        archived_by_phase = {"buffer": set(), "target": set()}
        for row in rows:
            phase = str(row.get("phase") or "").strip().lower()
            path = str(row.get("path") or "").strip()
            if phase in archived_by_phase and path:
                archived_by_phase[phase].add(path.lower())

        collapsed = list(grouped.pop(None, []))
        for key, group in grouped.items():
            phase, _measurement_id = key
            if len(group) == 1:
                collapsed.append(group[0])
                continue
            archived_matches = [
                row for row in group
                if str(row.get("path") or "").strip().lower() in archived_by_phase.get(phase, set())
            ]
            if archived_matches:
                collapsed.append(archived_matches[0])
                continue
            if phase == "buffer":
                chosen = min(group, key=self._paired_row_scan_sort_key)
            else:
                chosen = max(group, key=self._paired_row_scan_sort_key)
            collapsed.append(chosen)
        return collapsed

    def _collapse_paired_rows_to_one_trace_per_phase_channel(self, rows):
        grouped = {}
        passthrough = []
        for index, row in enumerate(rows):
            phase = str(row.get("phase") or "").strip().lower()
            channel = self._normalize_observation_channel(row.get("channel"))
            if phase not in {"buffer", "target"} or not channel:
                passthrough.append(row)
                continue
            trace_number = str(row.get("trace_number") or "").strip()
            key = (phase, channel, trace_number)
            grouped.setdefault(key, []).append((index, row))
        collapsed = list(passthrough)
        for group in grouped.values():
            _index, row = min(group, key=lambda item: self._paired_trace_duplicate_sort_key(item[1], item[0]))
            collapsed.append(row)
        return collapsed

    @staticmethod
    def _paired_trace_duplicate_sort_key(row, index):
        has_path = 0 if str(row.get("path") or "").strip() else 1
        has_blank_scan = 0 if str(row.get("scan") or "").strip() == "" else 1
        scan_sort = BayesianOptimizationTab._paired_row_scan_sort_key(row)
        return (has_path, has_blank_scan, scan_sort, index)

    @staticmethod
    def _paired_row_scan_sort_key(row):
        scan = row.get("scan")
        try:
            return int(scan)
        except (TypeError, ValueError):
            return -1

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
        objective = str((observation or {}).get("objective") or quality.get("objective") or "").strip().lower()
        mode = str(scoring.get("mode", "classic") or "classic").strip().lower()
        channel_weights = dict(scoring.get("channel_weights") or {})
        paired_weights = dict(scoring.get("paired_response_weights") or {})
        run_weights = dict(scoring.get("run_weights") or {})
        direction = str(
            quality.get("optimization_direction")
            or (observation or {}).get("optimization_direction")
            or (source_config.get("acquisition") or {}).get("optimization_direction")
            or "maximize"
        ).strip().lower()
        direction = self._display_optimization_direction(direction)
        poor_expression = (
            "abs(Q_channel) <"
            if direction == "survey"
            else f"Q_channel {'>' if direction == 'minimize' else '<'}"
        )
        q_run = float(observation.get("Q_run", quality.get("Q_run", 0.0)) or 0.0)
        mean_q = float(quality.get("mean_Q_channel", 0.0) or 0.0)
        std_q = float(quality.get("std_Q_channel", 0.0) or 0.0)
        failed = float(quality.get("failed_channel_fraction", 0.0) or 0.0)
        low = float(quality.get("poor_channel_fraction", quality.get("low_channel_fraction", 0.0)) or 0.0)
        lambda_var = float(run_weights.get("lambda_variability", 0.20))
        lambda_repeat_std = float(
            paired_weights.get("lambda_repeat_std", 0.0)
            if objective == "paired_response"
            else run_weights.get("lambda_repeat_std", 0.0)
        )
        lambda_failed = float(run_weights.get("lambda_failed", 0.40))
        lambda_low = float(run_weights.get("lambda_low", 0.20))
        threshold = float(run_weights.get("low_channel_threshold", 0.50))
        noise_penalty = float(channel_weights.get("noise_penalty", 0.0))
        total = float(
            channel_weights.get("peak_prominence", channel_weights.get("snr", 0.35))
        ) + sum(
            float(channel_weights.get(key, default))
            for key, default in (
                ("repeat_scan_snr", 0.0),
                ("peak_height", 0.0),
                ("peak_shape", 0.20),
                ("baseline", 0.20),
                ("replicate_consistency", 0.15),
                ("success", 0.10),
            )
        )
        repeat_penalty_line = (
            f"  Paired run repeat relative-std penalty: {lambda_repeat_std:g} x mean repeat relative std "
            f"{float(quality.get('mean_repeat_relative_std', 0.0) or 0.0):.4f} = "
            f"{float(quality.get('repeat_std_penalty', 0.0) or 0.0):.4f}"
            if objective == "paired_response"
            else (
                f"  Classic run repeat relative-std penalty: {lambda_repeat_std:g} x mean repeat relative std "
                f"{float(quality.get('mean_repeat_relative_std', 0.0) or 0.0):.4f} = "
                f"{float(quality.get('repeat_std_penalty', 0.0) or 0.0):.4f}"
            )
        )
        lines = [
            "Q_run breakdown:",
            f"  optimization direction: {direction}",
            f"  mean channel Q: {mean_q:.4f}",
            f"  Run std penalty: {lambda_var:g} x std(Q_channel) {std_q:.4f} = {lambda_var * std_q:.4f}",
            repeat_penalty_line,
            f"  Run failed penalty: {lambda_failed:g} x failed_fraction {failed:.4f} = {lambda_failed * failed:.4f}",
            f"  Run poor-channel penalty: {lambda_low:g} x poor_channel_fraction {low:.4f} = {lambda_low * low:.4f} ({poor_expression} threshold {threshold:g})",
            f"  Directional penalty adjustment: {float(quality.get('run_penalty_adjustment', 0.0) or 0.0):+.4f} (subtract for maximize, add for minimize, move toward zero for survey)",
            f"  final Q_run: {q_run:.4f}",
            "",
        ]
        if objective == "paired_response":
            lines.extend(
                [
                    "Paired-response Q_channel terms:",
                    "  repeat_scan_SNR = delta_peak / (buffer peak STD + target peak STD)",
                    "  peak_prominence = delta_peak / (average buffer RMS + average target RMS)",
                    "  paired_Q_channel = repeat_SNR_weight*repeat_scan_SNR + prominence_weight*peak_prominence "
                    "+ sign(delta_peak)*(buffer_weight*buffer_classic_Q + target_weight*target_classic_Q)",
                    f"  delta_peak = target_peak_height_uA - buffer_peak_height_uA",
                    f"  Mean buffer channel noise: {float(quality.get('mean_buffer_channel_noise', 0.0) or 0.0):.4g} uA",
                    f"  Mean target channel noise: {float(quality.get('mean_target_channel_noise', 0.0) or 0.0):.4g} uA",
                    f"  Mean buffer classic Q: {float(quality.get('mean_buffer_classic_Q', 0.0) or 0.0):.4f}",
                    f"  Mean target classic Q: {float(quality.get('mean_target_classic_Q', 0.0) or 0.0):.4f}",
                    f"  Mean signed delta peak: {float(quality.get('mean_delta_peak_height_uA', 0.0) or 0.0):.4g} uA",
                    f"  Mean absolute delta peak: {float(quality.get('mean_abs_delta_peak_height_uA', 0.0) or 0.0):.4g} uA",
                ]
            )
            return lines
        if mode == "signal_priority_unbounded":
            lines.extend(
                [
                    "Q_channel terms:",
                    (
                        "  "
                        f"Peak prominence weight {float(channel_weights.get('peak_prominence', channel_weights.get('snr', 0.45))):g}, "
                        f"Repeat-scan SNR weight {float(channel_weights.get('repeat_scan_snr', 0.0)):g}, "
                        f"Channel peak weight {float(channel_weights.get('peak_height', 0.35)):g}, "
                        f"Channel baseline weight {float(channel_weights.get('baseline', 0.12)):g}, "
                        f"Channel shape weight {float(channel_weights.get('peak_shape', 0.05)):g}, "
                        f"Channel replicate weight {float(channel_weights.get('replicate_consistency', 0.03)):g}, "
                        f"Channel success weight {float(channel_weights.get('success', 0.0)):g}; total {total:g}"
                    ),
                ]
            )
        else:
            lines.extend(
                [
                    "Q_channel weights:",
                    (
                        "  "
                        f"Peak prominence weight {float(channel_weights.get('peak_prominence', channel_weights.get('snr', 0.35))):g}, "
                        f"Repeat-scan SNR weight {float(channel_weights.get('repeat_scan_snr', 0.0)):g}, "
                        f"Channel peak weight {float(channel_weights.get('peak_height', 0.0)):g}, "
                        f"Channel shape weight {float(channel_weights.get('peak_shape', 0.20)):g}, "
                        f"Channel baseline weight {float(channel_weights.get('baseline', 0.20)):g}, "
                        f"Channel replicate weight {float(channel_weights.get('replicate_consistency', 0.15)):g}, "
                        f"Channel success weight {float(channel_weights.get('success', 0.10)):g}; "
                        f"Channel noise penalty {noise_penalty:g}; no weight normalization"
                    ),
                ]
            )
        return lines

    def _q_equation_lines(self, config=None):
        source_config = config if config is not None else (self._bo_session.config if self._bo_session else self._config or {})
        scoring = dict((source_config or {}).get("scoring") or {})
        mode = str(scoring.get("mode", "classic") or "classic").strip().lower()
        channel_weights = dict(scoring.get("channel_weights") or {})
        paired_weights = dict(scoring.get("paired_response_weights") or {})
        run_weights = dict(scoring.get("run_weights") or {})
        objective = str((source_config or {}).get("objective") or "").strip().lower()
        direction = str(
            ((source_config or {}).get("acquisition") or {}).get(
                "optimization_direction", "maximize"
            )
        ).strip().lower()
        direction = self._display_optimization_direction(direction)
        penalty_operator = "+" if direction == "minimize" else "-"
        poor_expression = (
            "abs(Q_channel) <"
            if direction == "survey"
            else f"Q_channel {'>' if direction == 'minimize' else '<'}"
        )
        penalty_prefix = (
            "move mean(Q_channel) toward zero by"
            if direction == "survey"
            else f"mean(Q_channel) {penalty_operator}"
        )

        if objective == "paired_response":
            lambda_var = float(run_weights.get("lambda_variability", 0.20))
            lambda_repeat_std = float(paired_weights.get("lambda_repeat_std", 0.0))
            lambda_failed = float(run_weights.get("lambda_failed", 0.40))
            lambda_low = float(run_weights.get("lambda_low", 0.20))
            threshold = float(run_weights.get("low_channel_threshold", 0.50))
            return [
                "delta_peak = average target peak height - average buffer peak height",
                "repeat_scan_SNR = delta_peak / (buffer peak-height STD + target peak-height STD)",
                "peak_prominence = delta_peak / (average buffer RMS + average target RMS)",
                "Q_channel = paired_Q_channel = repeat_SNR_weight*repeat_scan_SNR + prominence_weight*peak_prominence "
                "+ sign(delta_peak)*(buffer_weight*buffer_classic_Q + target_weight*target_classic_Q)",
                (
                    f"Q_run = {penalty_prefix} [Run std penalty({lambda_var:g})*std(Q_channel) "
                    f"+ Paired run repeat relative-std penalty({lambda_repeat_std:g})*mean(repeat relative std) "
                    f"+ Run failed penalty({lambda_failed:g})*failed_fraction "
                    f"+ Run poor-channel penalty({lambda_low:g})*fraction({poor_expression} threshold {threshold:g})] ({direction})"
                ),
            ]

        if mode == "signal_priority_unbounded":
            terms = [
                ("Peak prominence weight", "log1p(Peak prominence)", float(channel_weights.get("peak_prominence", channel_weights.get("snr", 0.45)))),
                ("Repeat-scan SNR weight", "log1p(Repeat-scan SNR)", float(channel_weights.get("repeat_scan_snr", 0.0))),
                ("Channel peak weight", "log1p(Peak uA)", float(channel_weights.get("peak_height", 0.35))),
                ("Channel baseline weight", "Baseline", float(channel_weights.get("baseline", 0.12))),
                ("Channel shape weight", "Shape", float(channel_weights.get("peak_shape", 0.05))),
                ("Channel replicate weight", "Replicate", float(channel_weights.get("replicate_consistency", 0.03))),
                ("Channel success weight", "Success", float(channel_weights.get("success", 0.0))),
            ]
            noise_penalty = 0.0
        else:
            terms = [
                ("Peak prominence weight", "Peak prominence", float(channel_weights.get("peak_prominence", channel_weights.get("snr", 0.35)))),
                ("Repeat-scan SNR weight", "Repeat-scan SNR", float(channel_weights.get("repeat_scan_snr", 0.0))),
                ("Channel peak weight", "Peak uA", float(channel_weights.get("peak_height", 0.0))),
                ("Channel shape weight", "Shape", float(channel_weights.get("peak_shape", 0.20))),
                ("Channel baseline weight", "Baseline", float(channel_weights.get("baseline", 0.20))),
                ("Channel replicate weight", "Replicate", float(channel_weights.get("replicate_consistency", 0.15))),
                ("Channel success weight", "Success", float(channel_weights.get("success", 0.10))),
            ]
            noise_penalty = float(channel_weights.get("noise_penalty", 0.0))
        total = sum(weight for _weight_label, _metric_label, weight in terms)
        numerator = " + ".join(
            f"{weight_label}({weight:g})*{metric_label}"
            for weight_label, metric_label, weight in terms
            if weight
        )
        if not numerator:
            numerator = "0"
        if mode != "signal_priority_unbounded":
            channel_penalty_operator = "-" if direction == "survey" else penalty_operator
            numerator = f"{numerator} {channel_penalty_operator} Channel noise penalty({noise_penalty:g})*Noise uA"
        lambda_var = float(run_weights.get("lambda_variability", 0.20))
        lambda_repeat_std = float(run_weights.get("lambda_repeat_std", 0.0))
        lambda_failed = float(run_weights.get("lambda_failed", 0.40))
        lambda_low = float(run_weights.get("lambda_low", 0.20))
        threshold = float(run_weights.get("low_channel_threshold", 0.50))
        lines = [
            f"Q_channel = ({numerator}) / {max(total, 1e-12):g}" if mode == "signal_priority_unbounded" else f"Q_channel = {numerator}",
            (
                f"Q_run = {penalty_prefix} [Run std penalty({lambda_var:g})*std(Q_channel) "
                f"+ Repeat relative-std penalty({lambda_repeat_std:g})*mean(repeat relative std) "
                f"+ Run failed penalty({lambda_failed:g})*failed_fraction "
                f"+ Run poor-channel penalty({lambda_low:g})*fraction({poor_expression} threshold {threshold:g})] ({direction})"
            ),
        ]
        if mode != "signal_priority_unbounded":
            lines.append(f"Prominence Score display = clip(peak prominence / {float(channel_weights.get('peak_prominence_saturation', channel_weights.get('snr_saturation', 20.0))):g}, 0, 1); classic Q_channel uses peak prominence directly.")
            lines.append(
                "Classic Q_channel is floored at 0. Q_run clips only the undesired sign for maximize/minimize; survey keeps both signs."
            )
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
        ]
        idx = 1
        for kind, folder in roots:
            for path in sorted(folder.glob("*")):
                if path.is_file():
                    self._model_tree.insert("", "end", text=str(idx), values=(kind, path.name))
                    idx += 1
        self._refresh_surrogate_controls()

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

    def _refresh_surrogate_controls(self):
        if not hasattr(self, "_surrogate_iteration_combo"):
            return
        iterations = self._surrogate_artifact_iterations()
        values = [str(iteration) for iteration in iterations]
        self._surrogate_iteration_combo.configure(values=values)
        dims = self._surrogate_dimension_options()
        for combo in (
            getattr(self, "_surrogate_x_combo", None),
            getattr(self, "_surrogate_y_combo", None),
            getattr(self, "_surrogate_z_combo", None),
        ):
            if combo is not None:
                combo.configure(values=dims)
        if dims:
            if self._surrogate_x_var.get() not in dims:
                self._surrogate_x_var.set(dims[0])
            if self._surrogate_y_var.get() not in dims:
                self._surrogate_y_var.set(dims[1] if len(dims) > 1 else dims[0])
            if self._surrogate_z_var.get() not in dims:
                self._surrogate_z_var.set(dims[2] if len(dims) > 2 else dims[-1])
        else:
            self._surrogate_x_var.set("")
            self._surrogate_y_var.set("")
            self._surrogate_z_var.set("")
        if values and self._surrogate_iteration_var.get() not in values:
            selected = None
            obs = self._selected_history_observation or {}
            try:
                selected = int(obs.get("iteration"))
            except Exception:
                selected = None
            self._surrogate_iteration_var.set(str(selected) if selected in iterations else values[-1])
        elif not values:
            self._surrogate_iteration_var.set("")

    def _surrogate_group_id(self):
        selected = self._selected_history_observation
        if selected is not None:
            try:
                return int(selected.get("group_id", 1))
            except Exception:
                return None
        return None

    def _surrogate_artifact_iterations(self):
        if self._bo_session is None:
            return []
        selected_group_id = self._surrogate_group_id()
        iterations = set()
        for folder in (self._bo_session.surrogate_dir, self._bo_session.acquisition_dir):
            for path in Path(folder).glob("*_candidate_predictions.csv"):
                group_id, iteration = self._parse_surrogate_artifact_name(path.name)
                if iteration is None:
                    continue
                if selected_group_id is not None and group_id != selected_group_id:
                    continue
                iterations.add(iteration)
            for path in Path(folder).glob("*_acquisition_values.csv"):
                group_id, iteration = self._parse_surrogate_artifact_name(path.name)
                if iteration is None:
                    continue
                if selected_group_id is not None and group_id != selected_group_id:
                    continue
                iterations.add(iteration)
        return sorted(iterations)

    @staticmethod
    def _parse_surrogate_artifact_name(name):
        text = str(name or "")
        grouped = re.search(r"group_(\d+)_iter_(\d+)_", text)
        if grouped:
            return int(grouped.group(1)), int(grouped.group(2))
        plain = re.search(r"iter_(\d+)_", text)
        if plain:
            return None, int(plain.group(1))
        return None, None

    def _surrogate_dimension_options(self):
        cfg = self._bo_session.config if self._bo_session is not None else self._config
        dims = []
        if cfg:
            try:
                dims = list(active_parameters(cfg))
            except Exception:
                dims = []
            if not dims:
                params = dict((cfg or {}).get("parameters") or {})
                dims = [
                    name for name in PARAMETER_ORDER
                    if str((params.get(name) or {}).get("mode", "")).lower() == "active"
                ]
        if not dims:
            rows = self._read_surrogate_rows()
            if rows:
                dims = [name for name in PARAMETER_ORDER if name in rows[0]]
        return [name for name in PARAMETER_ORDER if name in dims]

    def _surrogate_artifact_path(self, iteration=None):
        if self._bo_session is None:
            return None
        if iteration is None:
            raw = self._surrogate_iteration_var.get()
            if not raw:
                return None
            iteration = int(raw)
        group_id = self._surrogate_group_id()
        stem = self._bo_session._group_iteration_stem(iteration, group_id=group_id) if group_id is not None else f"iter_{int(iteration):03d}"
        candidates = (
            self._bo_session.surrogate_dir / f"{stem}_candidate_predictions.csv",
            self._bo_session.acquisition_dir / f"{stem}_acquisition_values.csv",
            self._bo_session.surrogate_dir / f"iter_{int(iteration):03d}_candidate_predictions.csv",
            self._bo_session.acquisition_dir / f"iter_{int(iteration):03d}_acquisition_values.csv",
        )
        for path in candidates:
            if path.exists():
                return path
        return None

    def _read_surrogate_rows(self, iteration=None):
        path = self._surrogate_artifact_path(iteration)
        if path is None:
            return []
        rows = []
        with open(path, "r", newline="", encoding="utf-8") as fh:
            for row in csv.DictReader(fh):
                parsed = {}
                for key, value in row.items():
                    if value in (None, ""):
                        parsed[key] = None
                        continue
                    text = str(value)
                    if text.lower() in ("true", "false"):
                        parsed[key] = text.lower() == "true"
                        continue
                    try:
                        parsed[key] = float(text)
                    except ValueError:
                        parsed[key] = value
                rows.append(parsed)
        return rows

    def _refresh_surrogate_view(self):
        if not hasattr(self, "_surrogate_plot_frame"):
            return
        if self._defer_results_render(
            self._surrogate_plot_frame,
            "Surrogate/acquisition plot rendering is paused while a measurement is collecting data.\nAcquisition has priority.",
        ):
            return
        self._refresh_surrogate_controls()
        for child in self._surrogate_plot_frame.winfo_children():
            child.destroy()
        rows = self._read_surrogate_rows()
        if not rows:
            ttk.Label(
                self._surrogate_plot_frame,
                text="No surrogate/acquisition CSV artifacts found for the selected iteration.",
            ).pack(fill="both", expand=True)
            return
        value_key = self._surrogate_value_var.get() or "predicted_mean_Q"
        x_name = self._surrogate_x_var.get()
        y_name = self._surrogate_y_var.get()
        z_name = self._surrogate_z_var.get()
        if not x_name or value_key not in rows[0]:
            ttk.Label(self._surrogate_plot_frame, text="Choose a valid value and X parameter.").pack(fill="both", expand=True)
            return
        try:
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            from matplotlib.figure import Figure
        except Exception as exc:
            ttk.Label(self._surrogate_plot_frame, text=f"Matplotlib plot unavailable: {exc}").pack(fill="both", expand=True)
            return
        fig = Figure(figsize=(7.2, 4.0), dpi=100)
        view = self._surrogate_view_var.get() or "1D slice"
        color_limits = self._surrogate_manual_color_range()
        if view == "Correlation falloff":
            self._plot_surrogate_correlation_falloff(fig, rows, x_name)
        elif view == "3D tensor" and y_name and z_name:
            self._plot_surrogate_3d(fig, rows, value_key, x_name, y_name, z_name, color_limits)
        elif view == "2D map" and y_name:
            self._plot_surrogate_2d(fig, rows, value_key, x_name, y_name, color_limits)
        else:
            self._plot_surrogate_1d(fig, rows, value_key, x_name)
        if view == "3D tensor":
            self._fit_embedded_figure(fig, top=0.84, bottom=0.14, left=0.06, right=0.86)
        else:
            self._fit_embedded_figure(fig, top=0.84, bottom=0.20, left=0.16, right=0.93)
        canvas = FigureCanvasTkAgg(fig, master=self._surrogate_plot_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)
        self._surrogate_plot_canvas = canvas

    def _surrogate_gp_model_path(self, iteration=None):
        if self._bo_session is None:
            return None
        if iteration is None:
            raw = self._surrogate_iteration_var.get()
            if not raw:
                return None
            iteration = int(raw)
        group_id = self._surrogate_group_id()
        candidates = []
        if group_id is not None:
            candidates.append(self._bo_session.surrogate_dir / f"{self._bo_session._group_iteration_stem(iteration, group_id=group_id)}_gp_model.pkl")
        candidates.append(self._bo_session.surrogate_dir / f"iter_{int(iteration):03d}_gp_model.pkl")
        for path in candidates:
            if path.exists():
                return path
        return None

    def _load_surrogate_gp_model(self):
        path = self._surrogate_gp_model_path()
        if path is not None:
            try:
                with open(path, "rb") as fh:
                    return pickle.load(fh)
            except Exception:
                pass
        # Older/incomplete real-session artifacts may lack a usable pickle.
        # Refit from only the observations available at the selected iteration
        # so historical views do not leak later measurements.
        if self._bo_session is None:
            return None
        try:
            historical_session = copy.copy(self._bo_session)
            historical_session.observations = self._surrogate_observations_so_far()
            gp, _train = historical_session._fit_gp_surrogate()
            return gp
        except Exception:
            return None

    def _extract_matern_kernel(self, kernel):
        if kernel is None:
            return None
        if kernel.__class__.__name__ == "Matern":
            return kernel
        for attr in ("k1", "k2"):
            child = getattr(kernel, attr, None)
            found = self._extract_matern_kernel(child)
            if found is not None:
                return found
        return None

    def _extract_white_noise_level(self, kernel):
        if kernel is None:
            return None
        if kernel.__class__.__name__ == "WhiteKernel":
            try:
                return float(kernel.noise_level)
            except Exception:
                return None
        for attr in ("k1", "k2"):
            child = getattr(kernel, attr, None)
            found = self._extract_white_noise_level(child)
            if found is not None:
                return found
        return None

    def _gp_length_scale_by_parameter(self, gp):
        matern = self._extract_matern_kernel(getattr(gp, "kernel_", None))
        if matern is None:
            return {}, None
        raw = getattr(matern, "length_scale", None)
        try:
            import numpy as np
            values = np.asarray(raw, dtype=float).ravel().tolist()
        except Exception:
            try:
                values = [float(raw)]
            except Exception:
                return {}, matern
        if len(values) == 1:
            values = values * len(OPTIMIZER_ORDER)
        order = list(OPTIMIZER_ORDER)
        if len(values) != len(order):
            try:
                active = active_parameters(self._bo_session.config if self._bo_session is not None else self._config)
            except Exception:
                active = []
            if len(values) == len(active):
                order = active
        return {
            name: values[idx]
            for idx, name in enumerate(order)
            if idx < len(values)
        }, matern

    def _plot_surrogate_correlation_falloff(self, fig, rows, x_name):
        ax = fig.add_subplot(111)
        gp = self._load_surrogate_gp_model()
        if gp is None:
            ax.text(
                0.5,
                0.5,
                "No saved GP model for this artifact iteration.\n"
                "Correlation falloff is available only when iter_###_gp_model.pkl exists.",
                ha="center",
                va="center",
            )
            ax.set_axis_off()
            return
        length_scales, matern = self._gp_length_scale_by_parameter(gp)
        length_scale = length_scales.get(x_name)
        if length_scale is None or length_scale <= 0:
            ax.text(0.5, 0.5, f"No GP length scale found for {x_name}.", ha="center", va="center")
            ax.set_axis_off()
            return
        center = self._selected_surrogate_observation() or self._selected_history_observation
        if center is None and self._bo_session is not None:
            center = self._bo_session.best_observation()
        center_params = dict((center or {}).get("params") or {})
        if x_name not in center_params and rows:
            values = [float(row[x_name]) for row in rows if row.get(x_name) is not None]
            if values:
                center_params[x_name] = sum(values) / len(values)
        if not center_params or x_name not in center_params:
            ax.text(0.5, 0.5, "No center point available for falloff plot.", ha="center", va="center")
            ax.set_axis_off()
            return
        base = dict(center_params)
        cfg = self._bo_session.config if self._bo_session is not None else self._config
        for name in PARAMETER_ORDER:
            if name not in base:
                try:
                    base[name] = float((cfg or {}).get("initial_parameters", {}).get(name, 0.0))
                except Exception:
                    base[name] = 0.0
        bounds = self._parameter_plot_bounds(cfg, x_name)
        if bounds is None:
            ax.text(0.5, 0.5, f"No valid bounds found for {x_name}.", ha="center", va="center")
            ax.set_axis_off()
            return
        x_min, x_max, is_log = bounds
        base_encoded = encode_candidate(base, cfg)
        x_index = OPTIMIZER_ORDER.index(x_name)
        center_encoded = max(0.0, min(1.0, float(base_encoded[x_index])))
        full_half_width = 0.5
        local_half_width = max(0.01, min(full_half_width, 4.0 * float(length_scale)))
        zoomed = local_half_width < 0.20
        if zoomed:
            enc_min = max(0.0, center_encoded - local_half_width)
            enc_max = min(1.0, center_encoded + local_half_width)
            if enc_max <= enc_min:
                enc_min, enc_max = 0.0, 1.0
        else:
            enc_min, enc_max = 0.0, 1.0
        sample_count = 240
        encoded_values = [enc_min + (enc_max - enc_min) * idx / (sample_count - 1) for idx in range(sample_count)]
        x_values = [self._raw_from_encoded_parameter(value, x_min, x_max, is_log) for value in encoded_values]
        correlations = []
        for value, encoded_value in zip(x_values, encoded_values):
            candidate = dict(base)
            candidate[x_name] = value
            try:
                distance = abs(float(encoded_value) - center_encoded) / max(float(length_scale), 1e-12)
            except Exception:
                distance = 0.0
            correlations.append(self._matern_correlation(distance, getattr(matern, "nu", 2.5)))
        ax.plot(x_values, correlations, color=self.ACCENT_DARK, linewidth=2.0)
        ax.axvline(float(base[x_name]), color="#d67b32", linestyle="--", linewidth=1.4, label="center")
        for frac, label in ((0.5, "50%"), (0.1, "10%")):
            crossing = self._falloff_crossing(x_values, correlations, frac, float(base[x_name]))
            if crossing is not None:
                ax.axvline(crossing, color="#5a6b84", linestyle=":", linewidth=1.0)
                ax.text(crossing, frac, label, fontsize=8, ha="left", va="bottom")
        noise = self._extract_white_noise_level(getattr(gp, "kernel_", None))
        noise_text = f", noise {noise:.3g}" if noise is not None else ""
        zoom_text = " | local zoom" if zoomed else ""
        ax.set_title(
            f"GP correlation falloff | {x_name} | length scale {length_scale:.3g}{noise_text}{zoom_text}",
            fontsize=9,
            pad=8,
            wrap=True,
        )
        ax.set_xlabel(x_name, fontsize=9, labelpad=4)
        ax.set_ylabel("Correlation to center", fontsize=9, labelpad=4)
        if is_log and x_min > 0:
            ax.set_xscale("log")
        ax.set_ylim(-0.02, 1.02)
        ax.grid(alpha=0.25)
        ax.legend(loc="best", fontsize=8)
        ax.tick_params(labelsize=8)

    def _parameter_plot_bounds(self, config, name):
        try:
            cfg = normalize_bo_config(config or {})
            p_cfg = dict(cfg["parameters"].get(name) or {})
            lo = float(p_cfg.get("min"))
            hi = float(p_cfg.get("max"))
            is_log = str(p_cfg.get("scale", "")).lower() in ("log", "log10")
        except Exception:
            values = []
            for row in self._read_surrogate_rows():
                try:
                    values.append(float(row.get(name)))
                except Exception:
                    pass
            if not values:
                return None
            lo, hi = min(values), max(values)
            is_log = False
        if hi < lo:
            lo, hi = hi, lo
        if hi <= lo:
            pad = max(abs(lo) * 0.1, 1.0)
            lo -= pad
            hi += pad
        if is_log:
            lo = max(lo, 1e-12)
            hi = max(hi, lo * 1.0001)
        return lo, hi, is_log

    @staticmethod
    def _raw_from_encoded_parameter(encoded_value, lo, hi, is_log):
        encoded_value = max(0.0, min(1.0, float(encoded_value)))
        if is_log:
            log_lo = math.log10(max(float(lo), 1e-12))
            log_hi = math.log10(max(float(hi), 1e-12))
            return 10 ** (log_lo + encoded_value * (log_hi - log_lo))
        return float(lo) + encoded_value * (float(hi) - float(lo))

    @staticmethod
    def _matern_correlation(distance, nu):
        d = max(0.0, float(distance))
        if abs(float(nu) - 0.5) < 1e-9:
            return math.exp(-d)
        if abs(float(nu) - 1.5) < 1e-9:
            r = math.sqrt(3.0) * d
            return (1.0 + r) * math.exp(-r)
        if abs(float(nu) - 2.5) < 1e-9:
            r = math.sqrt(5.0) * d
            return (1.0 + r + (r * r) / 3.0) * math.exp(-r)
        return math.exp(-0.5 * d * d)

    @staticmethod
    def _falloff_crossing(x_values, correlations, threshold, center):
        candidates = []
        for idx in range(1, len(x_values)):
            y0 = correlations[idx - 1]
            y1 = correlations[idx]
            if (y0 - threshold) == 0:
                candidates.append(x_values[idx - 1])
            elif (y0 - threshold) * (y1 - threshold) < 0:
                t = (threshold - y0) / max(y1 - y0, 1e-12)
                candidates.append(x_values[idx - 1] + t * (x_values[idx] - x_values[idx - 1]))
        if not candidates:
            return None
        return min(candidates, key=lambda value: abs(value - center))

    def _surrogate_manual_color_range(self):
        min_text = (self._surrogate_color_min_var.get() or "").strip()
        max_text = (self._surrogate_color_max_var.get() or "").strip()
        if not min_text and not max_text:
            return None
        try:
            vmin = float(min_text)
            vmax = float(max_text)
        except Exception:
            self._status_var.set("Surrogate color range ignored: enter numeric min and max, or clear both for auto color.")
            return None
        if vmax <= vmin:
            self._status_var.set("Surrogate color range ignored: max must be greater than min.")
            return None
        return vmin, vmax

    def _plot_surrogate_1d(self, fig, rows, value_key, x_name):
        ax = fig.add_subplot(111)
        all_rows = [row for row in rows if row.get(x_name) is not None and row.get(value_key) is not None]
        if not all_rows:
            ax.text(0.5, 0.5, "No plottable surrogate rows", ha="center", va="center")
            ax.set_axis_off()
            return
        x_all = [float(row[x_name]) for row in all_rows]
        y_all = [float(row[value_key]) for row in all_rows]
        ax.scatter(x_all, y_all, color=self.ACCENT_DARK, s=12, alpha=0.35, label="all candidate predictions")
        grouped = {}
        for x_value, y_value in zip(x_all, y_all):
            grouped.setdefault(x_value, []).append(y_value)
        if 1 < len(grouped) < len(x_all):
            trend_x = []
            trend_y = []
            for x_value in sorted(grouped):
                values = sorted(grouped[x_value])
                mid = len(values) // 2
                median = values[mid] if len(values) % 2 else (values[mid - 1] + values[mid]) / 2.0
                trend_x.append(x_value)
                trend_y.append(median)
            ax.plot(trend_x, trend_y, color="#d67b32", linewidth=1.4, alpha=0.85, label="median at X")
        self._overlay_observed_points(ax, x_name, None)
        ax.set_xlabel(x_name, fontsize=9, labelpad=4)
        ax.set_ylabel(value_key, fontsize=9, labelpad=4)
        ax.set_title(self._surrogate_plot_title(value_key, "1D all-candidate view"), fontsize=9, pad=8, wrap=True)
        ax.tick_params(labelsize=8)
        ax.grid(alpha=0.25)
        ax.legend(loc="best", fontsize=8)

    def _plot_surrogate_2d(self, fig, rows, value_key, x_name, y_name, color_limits=None):
        ax = fig.add_subplot(111)
        color_kwargs = {}
        if color_limits is not None:
            color_kwargs = {"vmin": color_limits[0], "vmax": color_limits[1]}
        grid = self._surrogate_2d_prediction_grid(
            rows,
            value_key,
            x_name,
            y_name,
        )
        if grid is not None:
            x_values, y_values, z_values, x_is_log, y_is_log = grid
            mesh = ax.pcolormesh(
                x_values,
                y_values,
                z_values,
                cmap="viridis",
                shading="auto",
                **color_kwargs,
            )
            if x_is_log:
                ax.set_xscale("log")
            if y_is_log:
                ax.set_yscale("log")
        else:
            # Old sessions may not have a saved GP model. A scatter plot is an
            # honest fallback; triangulating a high-dimensional candidate
            # projection invents jagged connections and holes.
            points = [
                row for row in rows
                if row.get(x_name) is not None
                and row.get(y_name) is not None
                and row.get(value_key) is not None
            ]
            if not points:
                ax.text(0.5, 0.5, "No plottable surrogate rows", ha="center", va="center")
                ax.set_axis_off()
                return
            mesh = ax.scatter(
                [float(row[x_name]) for row in points],
                [float(row[y_name]) for row in points],
                c=[float(row[value_key]) for row in points],
                cmap="viridis",
                s=14,
                alpha=0.8,
                **color_kwargs,
            )
        cbar = fig.colorbar(mesh, ax=ax, label=value_key)
        cbar.ax.tick_params(labelsize=8)
        cbar.set_label(value_key, fontsize=8)
        self._overlay_observed_points(ax, x_name, y_name)
        ax.set_xlabel(x_name, fontsize=9, labelpad=4)
        ax.set_ylabel(y_name, fontsize=9, labelpad=4)
        ax.set_title(self._surrogate_plot_title(value_key, "2D surrogate slice"), fontsize=9, pad=8, wrap=True)
        ax.tick_params(labelsize=8)
        ax.grid(alpha=0.2)

    def _surrogate_2d_prediction_grid(self, rows, value_key, x_name, y_name, grid_size=75):
        gp = self._load_surrogate_gp_model()
        cfg = self._bo_session.config if self._bo_session is not None else self._config
        if not cfg or x_name == y_name:
            return None
        if x_name not in OPTIMIZER_ORDER or y_name not in OPTIMIZER_ORDER:
            return None
        x_bounds = self._parameter_plot_bounds(cfg, x_name)
        y_bounds = self._parameter_plot_bounds(cfg, y_name)
        if x_bounds is None or y_bounds is None:
            return None
        try:
            import numpy as np

            observations = self._surrogate_observations_so_far()
            center = self._selected_surrogate_observation()
            if center is None and observations:
                center = max(
                    observations,
                    key=lambda obs: self._optimization_objective_value(obs.get("Q_run", 0.0), cfg),
                )
            base = dict((center or {}).get("params") or {})
            initial = dict((cfg or {}).get("initial_parameters") or {})
            for name in PARAMETER_ORDER:
                if base.get(name) is None:
                    base[name] = float(initial.get(name, 0.0))
            base_encoded = np.asarray(encode_candidate(base, cfg), dtype=float)

            size = max(20, min(200, int(grid_size)))
            x_encoded = np.linspace(0.0, 1.0, size)
            y_encoded = np.linspace(0.0, 1.0, size)
            x_mesh_encoded, y_mesh_encoded = np.meshgrid(x_encoded, y_encoded)
            encoded_points = np.repeat(base_encoded[None, :], size * size, axis=0)
            encoded_points[:, OPTIMIZER_ORDER.index(x_name)] = x_mesh_encoded.ravel()
            encoded_points[:, OPTIMIZER_ORDER.index(y_name)] = y_mesh_encoded.ravel()
            if gp is not None:
                means, stds = gp.predict(encoded_points, return_std=True)
            else:
                observed_points = [
                    (encode_candidate(obs["params"], cfg), float(obs["Q_run"]))
                    for obs in observations
                    if obs.get("params") is not None and obs.get("Q_run") is not None
                ]
                if not observed_points:
                    return None
                train_x = np.asarray([point for point, _q in observed_points], dtype=float)
                train_y = np.asarray([q for _point, q in observed_points], dtype=float)
                distances = np.linalg.norm(
                    encoded_points[:, None, :] - train_x[None, :, :],
                    axis=2,
                )
                weights = 1.0 / (distances + 0.05)
                means = (weights @ train_y) / np.maximum(weights.sum(axis=1), 1e-12)
                stds = distances.min(axis=1)

            if value_key == "predicted_mean_Q":
                values = means
            elif value_key == "predicted_std_Q":
                values = stds
            elif value_key == "acquisition_value":
                best_objective = max(
                    (self._optimization_objective_value(obs.get("Q_run", 0.0), cfg) for obs in observations),
                    default=self._optimization_objective_value(rows[0].get("best_observed_Q", 0.0), cfg) if rows else 0.0,
                )
                exploration = float((cfg.get("acquisition") or {}).get("exploration", 0.35))
                values = np.asarray(
                    [
                        _acquisition_score(
                            self._optimization_objective_value(float(mean), cfg),
                            float(std),
                            best_objective,
                            exploration,
                        )
                        for mean, std in zip(means, stds)
                    ],
                    dtype=float,
                )
            else:
                return None

            x_lo, x_hi, x_is_log = x_bounds
            y_lo, y_hi, y_is_log = y_bounds
            x_values = np.asarray(
                [
                    self._raw_from_encoded_parameter(value, x_lo, x_hi, x_is_log)
                    for value in x_encoded
                ],
                dtype=float,
            )
            y_values = np.asarray(
                [
                    self._raw_from_encoded_parameter(value, y_lo, y_hi, y_is_log)
                    for value in y_encoded
                ],
                dtype=float,
            )
            return x_values, y_values, values.reshape(size, size), x_is_log, y_is_log
        except Exception:
            return None

    def _plot_surrogate_3d(self, fig, rows, value_key, x_name, y_name, z_name, color_limits=None):
        ax = fig.add_subplot(111, projection="3d")
        points = [
            row for row in rows
            if row.get(x_name) is not None
            and row.get(y_name) is not None
            and row.get(z_name) is not None
            and row.get(value_key) is not None
        ]
        if not points:
            ax.text2D(0.5, 0.5, "No plottable surrogate rows", ha="center", va="center", transform=ax.transAxes)
            return
        scatter = ax.scatter(
            [float(row[x_name]) for row in points],
            [float(row[y_name]) for row in points],
            [float(row[z_name]) for row in points],
            c=[float(row[value_key]) for row in points],
            cmap="viridis",
            s=7,
            alpha=0.35,
            vmin=color_limits[0] if color_limits is not None else None,
            vmax=color_limits[1] if color_limits is not None else None,
        )
        observations = self._surrogate_observations_so_far()
        path = [
            obs for obs in observations
            if all((obs.get("params") or {}).get(name) is not None for name in (x_name, y_name, z_name))
        ]
        if path:
            xs = [float(obs["params"][x_name]) for obs in path]
            ys = [float(obs["params"][y_name]) for obs in path]
            zs = [float(obs["params"][z_name]) for obs in path]
            ax.plot(xs, ys, zs, color="#d67b32", linewidth=1.4, label="observed path")
            ax.scatter(xs, ys, zs, color="#d67b32", s=18, depthshade=False)
            selected = self._selected_surrogate_observation()
            if selected is not None and all((selected.get("params") or {}).get(name) is not None for name in (x_name, y_name, z_name)):
                params = selected["params"]
                ax.scatter(
                    [float(params[x_name])],
                    [float(params[y_name])],
                    [float(params[z_name])],
                    color="#ffd166",
                    edgecolors="black",
                    linewidths=0.8,
                    s=90,
                    depthshade=False,
                )
        cbar = fig.colorbar(scatter, ax=ax, label=value_key, shrink=0.75, pad=0.08)
        cbar.ax.tick_params(labelsize=8)
        cbar.set_label(value_key, fontsize=8)
        ax.set_xlabel(x_name, fontsize=8, labelpad=3)
        ax.set_ylabel(y_name, fontsize=8, labelpad=3)
        ax.set_zlabel(z_name, fontsize=8, labelpad=3)
        ax.set_title(self._surrogate_plot_title(value_key, "3D surrogate tensor"), fontsize=9, pad=8, wrap=True)
        ax.tick_params(labelsize=7)
        if path:
            ax.legend(loc="best", fontsize=8)

    def _surrogate_iteration_limit(self):
        try:
            return int(self._surrogate_iteration_var.get())
        except Exception:
            return None

    def _surrogate_observations_so_far(self):
        if self._bo_session is None:
            return []
        limit = self._surrogate_iteration_limit()
        selected_group_id = self._surrogate_group_id()
        observations = []
        for obs in self._bo_session.observations:
            try:
                iteration = int(obs.get("iteration"))
            except Exception:
                continue
            if selected_group_id is not None:
                try:
                    if int(obs.get("group_id", 1)) != selected_group_id:
                        continue
                except Exception:
                    continue
            if limit is None or iteration <= limit:
                observations.append(obs)
        return observations

    def _selected_surrogate_observation(self):
        selected = self._selected_history_observation
        if selected is None:
            return None
        limit = self._surrogate_iteration_limit()
        try:
            iteration = int(selected.get("iteration"))
        except Exception:
            return None
        if limit is not None and iteration > limit:
            return None
        return selected

    def _overlay_observed_points(self, ax, x_name, y_name=None):
        observations = self._surrogate_observations_so_far()
        points = []
        for obs in observations:
            params = obs.get("params") or {}
            if params.get(x_name) is None:
                continue
            if y_name and params.get(y_name) is None:
                continue
            points.append(obs)
        if not points:
            return
        if y_name:
            xs = [float(obs["params"][x_name]) for obs in points]
            ys = [float(obs["params"][y_name]) for obs in points]
            ax.plot(xs, ys, color="#d67b32", linewidth=1.4, alpha=0.95, label="observed path")
            ax.scatter(xs, ys, color="#d67b32", s=18, zorder=4)
            selected = self._selected_surrogate_observation()
            if selected is not None:
                params = selected.get("params") or {}
                if params.get(x_name) is not None and params.get(y_name) is not None:
                    ax.scatter(
                        [float(params[x_name])],
                        [float(params[y_name])],
                        color="#ffd166",
                        edgecolors="black",
                        linewidths=1.0,
                        s=85,
                        zorder=5,
                        label="selected iteration",
                    )
        else:
            xs = [float(obs["params"][x_name]) for obs in points]
            ys = [float(obs.get("Q_run", 0.0)) for obs in points]
            ax.scatter(xs, ys, color="#d67b32", s=24, zorder=4, label="observed Q_run")
            selected = self._selected_surrogate_observation()
            if selected is not None and (selected.get("params") or {}).get(x_name) is not None:
                ax.scatter(
                    [float(selected["params"][x_name])],
                    [float(selected.get("Q_run", 0.0))],
                    color="#ffd166",
                    edgecolors="black",
                    linewidths=1.0,
                    s=85,
                    zorder=5,
                    label="selected iteration",
                )

    def _surrogate_plot_title(self, value_key, prefix):
        iteration = self._surrogate_iteration_var.get() or "?"
        backend = ""
        if self._bo_session is not None:
            metadata_candidates = []
            if str(iteration).isdigit():
                group_id = self._surrogate_group_id()
                if group_id is not None:
                    metadata_candidates.append(
                        self._bo_session.surrogate_dir / f"{self._bo_session._group_iteration_stem(int(iteration), group_id=group_id)}_surrogate_metadata.json"
                    )
                metadata_candidates.append(
                    self._bo_session.surrogate_dir / f"iter_{int(iteration):03d}_surrogate_metadata.json"
                )
            for metadata in metadata_candidates:
                if metadata.exists():
                    try:
                        with open(metadata, "r", encoding="utf-8") as fh:
                            backend = str((json.load(fh) or {}).get("backend") or "")
                        break
                    except Exception:
                        backend = ""
        suffix = f" ({backend})" if backend else ""
        group_text = ""
        group_id = self._surrogate_group_id()
        if group_id is not None:
            group_text = f" | group {group_id}"
        return f"{prefix} | {value_key}{group_text} | iter {iteration}{suffix}"

    @staticmethod
    def _analysis_trend_metric_options():
        return (
            "Q_run",
            "Paired Q",
            "Buffer classic Q",
            "Target classic Q",
            "Classic pair Q",
            "Delta Peak",
            "Fractional delta peak",
            "Distance",
            "Paired peak prominence",
            "Paired repeat-scan SNR",
            "Buffer peak prominence",
            "Target peak prominence",
            "Buffer channel noise",
            "Target channel noise",
            "Combined channel noise",
            "Buffer prominence score",
            "Target prominence score",
            "Target shape score",
            "Mean channel Q",
            "Std channel Q",
            "Failed fraction",
            "Poor fraction",
            "Mean peak uA",
            "Mean peak prominence",
            "Mean repeat-scan SNR",
            "Mean prominence score",
            "Mean shape score",
            "Mean baseline score",
            "Mean replicate score",
            "Mean success score",
            "Mean noise uA",
            "Step potential",
            "Amplitude",
            "Frequency",
            "Begin potential",
            "End potential",
            "Conditioning potential",
            "Conditioning time",
            "BO iteration",
            "Parameter set",
            "Buffer trace",
            "Target trace",
        )

    def _analysis_trend_value(self, observation, metric):
        quality = dict((observation or {}).get("quality") or {})
        params = dict((observation or {}).get("params") or {})
        truth = dict((observation or {}).get("simulation_truth") or {})
        if metric == "Q_run":
            return (observation or {}).get("Q_run")
        if metric == "Paired Q":
            return truth.get("paired_Q_score", quality.get("mean_paired_Q_channel", quality.get("mean_Q_channel")))
        if metric == "Buffer classic Q":
            return quality.get("mean_buffer_classic_Q")
        if metric == "Target classic Q":
            return quality.get("mean_target_classic_Q")
        if metric == "Classic pair Q":
            return quality.get("mean_classic_pair_Q")
        if metric == "Delta Peak":
            return quality.get("mean_abs_delta_peak_height_uA", truth.get("expected_delta_peak_uA"))
        if metric == "Fractional delta peak":
            return quality.get("mean_fractional_delta_peak")
        if metric == "Distance":
            return truth.get("normalized_distance")
        if metric == "Paired peak prominence":
            return quality.get("mean_peak_prominence", quality.get("mean_delta_peak_score"))
        if metric == "Paired repeat-scan SNR":
            return quality.get("mean_repeat_scan_snr")
        if metric == "Buffer peak prominence":
            return quality.get("mean_buffer_peak_prominence", quality.get("mean_buffer_snr_raw"))
        if metric == "Target peak prominence":
            return quality.get("mean_target_peak_prominence", quality.get("mean_target_snr_raw"))
        if metric == "Buffer channel noise":
            return quality.get("mean_buffer_channel_noise")
        if metric == "Target channel noise":
            return quality.get("mean_target_channel_noise")
        if metric == "Combined channel noise":
            return quality.get("mean_combined_channel_noise")
        if metric == "Buffer prominence score":
            return quality.get("mean_buffer_snr_score")
        if metric == "Target prominence score":
            return quality.get("mean_target_snr_score")
        if metric == "Target shape score":
            return quality.get("mean_target_shape_score")
        if metric == "BO iteration":
            return (observation or {}).get("iteration")
        if metric == "Parameter set":
            return (observation or {}).get("paired_batch_index")
        if metric == "Buffer trace":
            return (observation or {}).get("buffer_trace_number")
        if metric == "Target trace":
            return (observation or {}).get("target_trace_number")
        if metric == "Mean channel Q":
            return quality.get("mean_Q_channel")
        if metric == "Std channel Q":
            return quality.get("std_Q_channel")
        if metric == "Failed fraction":
            return quality.get("failed_channel_fraction")
        if metric == "Poor fraction":
            return quality.get("poor_channel_fraction", quality.get("low_channel_fraction"))
        if metric == "Mean peak uA":
            peak, _rms = self._observation_peak_rms(observation)
            return peak
        if metric == "Mean noise uA":
            _peak, rms = self._observation_peak_rms(observation)
            return rms
        if metric == "Mean peak prominence":
            return self._observation_component_mean(observation, "peak_prominence_raw")
        if metric == "Mean repeat-scan SNR":
            return self._observation_component_mean(observation, "repeat_scan_snr_raw")
        if metric == "Mean prominence score":
            return self._observation_component_mean(observation, "normalized_peak_prominence")
        if metric == "Mean shape score":
            return self._observation_component_mean(observation, "peak_shape_score")
        if metric == "Mean baseline score":
            return self._observation_component_mean(observation, "baseline_stability_score")
        if metric == "Mean replicate score":
            return self._observation_component_mean(observation, "replicate_consistency_score")
        if metric == "Mean success score":
            return self._observation_component_mean(observation, "success_score")
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

    @staticmethod
    def _normalize_observation_channel(channel):
        if channel in (None, ""):
            return None
        try:
            return str(int(float(channel)))
        except (TypeError, ValueError):
            text = str(channel).strip()
            return text or None

    @classmethod
    def _observation_channel_filter(cls, observation):
        channels = set()
        for channel in (observation or {}).get("channels", []):
            normalized = cls._normalize_observation_channel(channel)
            if normalized is not None:
                channels.add(normalized)
        return channels

    @staticmethod
    def _observation_component_mean(observation, key):
        components = dict(((observation or {}).get("quality") or {}).get("channel_components") or {})
        if not components:
            return None
        values = []
        for data in components.values():
            if not isinstance(data, dict):
                continue
            value = data.get(key)
            if value is None:
                continue
            try:
                values.append(float(value))
            except (TypeError, ValueError):
                continue
        return sum(values) / len(values) if values else None

    def _observation_peak_rms(self, observation):
        row_peaks = []
        row_rms_values = []
        selected_channels = self._observation_channel_filter(observation)
        try:
            analysis_result_paths = self._analysis_results_paths_for_observation(observation)
        except Exception:
            analysis_result_paths = []
        for results_path in analysis_result_paths:
            if not results_path.exists():
                continue
            try:
                with open(results_path, "r", encoding="utf-8-sig", newline="") as fh:
                    for row in csv.DictReader(fh):
                        if selected_channels:
                            channel = self._normalize_observation_channel(row.get("channel"))
                            if channel is None or channel not in selected_channels:
                                continue
                        if str(row.get("status", "")).upper() != "OK":
                            continue
                        peak_text = row.get("peak_current")
                        noise, _bracket_count, _crop_count = self._minima_bracket_rms_noise_from_row(row)
                        rms_text = noise if noise is not None else row.get("background_current_rms")
                        try:
                            if peak_text not in (None, ""):
                                row_peaks.append(float(peak_text))
                        except (TypeError, ValueError):
                            pass
                        try:
                            if rms_text not in (None, ""):
                                row_rms_values.append(float(rms_text))
                        except (TypeError, ValueError):
                            pass
            except Exception:
                continue
        if row_peaks or row_rms_values:
            peak_avg = sum(row_peaks) / len(row_peaks) if row_peaks else None
            rms_avg = sum(row_rms_values) / len(row_rms_values) if row_rms_values else None
            return peak_avg, rms_avg

        metrics = (observation or {}).get("channel_metrics", {})
        if not isinstance(metrics, dict) or not metrics:
            return None, None
        peaks = []
        rms_values = []
        for data in metrics.values():
            peak = self._channel_peak_height(data)
            rms = self._channel_background_rms(data)
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
