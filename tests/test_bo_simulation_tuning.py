import tkinter as tk

from core.bo_session import load_bo_config
from gui.tab_bayesian_optimization import BayesianOptimizationTab


def _var(master, value):
    return tk.StringVar(master=master, value=str(value))


def test_simulation_tuning_builds_independent_paired_bo_config():
    master = tk.Tcl()
    tab = BayesianOptimizationTab.__new__(BayesianOptimizationTab)
    tab._config = load_bo_config("optimizer/bo_configs/default_swv_bo.json")
    original_exploration = tab._config["acquisition"]["exploration"]
    original_scoring = dict(tab._config["scoring"]["paired_response_weights"])

    tab._engine_exploration_var = tk.DoubleVar(master=master, value=0.45)
    tab._engine_candidate_pool_var = _var(master, 1500)
    tab._engine_local_pool_var = _var(master, 25)
    tab._engine_initial_point_mode_var = _var(master, "random")
    tab._engine_warmup_iterations_var = _var(master, 2)
    tab._engine_gp_length_scale_vars = {
        name: _var(master, 0.35) for name in tab._config["parameters"]
    }
    tab._engine_channel_group_vars = [
        _var(master, "1, 2, 3"),
        _var(master, "4, 5"),
    ]
    tab._engine_channel_group_settings = [
        {
            "exploration_var": _var(master, 0.3),
            "warmup_var": _var(master, 2),
            "candidate_pool_var": _var(master, 700),
            "local_pool_var": _var(master, 70),
            "start_mode_var": _var(master, "specific"),
            "initial_parameters": dict(tab._config["initial_parameters"]),
        },
        {
            "exploration_var": _var(master, 0.5),
            "warmup_var": _var(master, 4),
            "candidate_pool_var": _var(master, 900),
            "local_pool_var": _var(master, 90),
            "start_mode_var": _var(master, "random"),
            "initial_parameters": dict(tab._config["initial_parameters"]),
        },
    ]
    tab._engine_score_vars = {
        "mode": _var(master, "classic"),
        "snr": _var(master, 0.5),
        "peak_height": _var(master, 1.0),
        "peak_shape": _var(master, 0.0),
        "baseline": _var(master, 0.0),
        "replicate_consistency": _var(master, 0.0),
        "success": _var(master, 0.0),
        "noise_penalty": _var(master, 5.0),
        "snr_saturation": _var(master, 20.0),
        "lambda_variability": _var(master, 0.0),
        "lambda_failed": _var(master, 0.0),
        "lambda_low": _var(master, 0.0),
        "low_channel_threshold": _var(master, 0.0),
    }
    tab._engine_paired_score_vars = {
        "buffer_classic_Q": _var(master, 0.1),
        "target_classic_Q": _var(master, 0.2),
        "delta_peak": _var(master, 2.0),
        "delta_scale_uA": _var(master, 0.5),
    }
    tab._engine_analysis_vars = {
        "crop_min_v": _var(master, -0.55),
        "crop_max_v": _var(master, -0.20),
        "smooth_window": _var(master, 11),
        "smooth_polyorder": _var(master, 3),
        "minima_search_window_v": _var(master, 0.18),
        "min_peak_height_ua": _var(master, 0.02),
        "peak_voltage_min_v": _var(master, -0.50),
        "peak_voltage_max_v": _var(master, -0.25),
        "min_start_voltage_v": _var(master, -0.70),
        "scan_windows": _var(master, ""),
        "use_prominent_minima": tk.BooleanVar(master=master, value=True),
        "require_local_minima_on_both_sides": tk.BooleanVar(master=master, value=True),
        "use_double_correction": tk.BooleanVar(master=master, value=False),
        "compute_skew": tk.BooleanVar(master=master, value=True),
        "compute_wavelet_energy": tk.BooleanVar(master=master, value=False),
        "compute_wavelet_denoised_trace": tk.BooleanVar(master=master, value=False),
        "use_wavelet_for_correction": tk.BooleanVar(master=master, value=False),
    }

    tuned = tab._engine_bo_config(
        {"paired_response": True, "paired_batch_size": 5}
    )

    assert tuned["acquisition"]["exploration"] == 0.45
    assert tuned["acquisition"]["candidate_pool_size"] == 1500
    assert tuned["acquisition"]["local_candidate_pool_size"] == 25
    assert tuned["acquisition"]["gp_falloff_fractions"]["frequency"] == 0.35
    assert tuned["paired_warmup_cycles"] == 2
    assert tuned["n_initial_points"] == 10
    assert tuned["scoring"]["channel_weights"]["noise_penalty"] == 5.0
    assert tuned["scoring"]["paired_response_weights"]["delta_peak"] == 2.0
    assert tuned["analysis"]["crop_min_v"] == -0.55
    assert tuned["analysis"]["smooth_window"] == 11
    assert tuned["analysis"]["require_local_minima_on_both_sides"] is True
    assert tuned["channel_groups"][0]["exploration"] == 0.3
    assert tuned["channel_groups"][0]["n_initial_points"] == 2
    assert tuned["channel_groups"][1]["exploration"] == 0.5
    assert tuned["channel_groups"][1]["n_initial_points"] == 4

    assert tab._config["acquisition"]["exploration"] == original_exploration
    assert tab._config["scoring"]["paired_response_weights"] == original_scoring


def test_main_setup_persists_per_group_optimizer_settings():
    master = tk.Tcl()
    tab = BayesianOptimizationTab.__new__(BayesianOptimizationTab)
    tab._config = load_bo_config("optimizer/bo_configs/default_swv_bo.json")
    tab._channels_var = _var(master, "")
    tab._channel_group_vars = [_var(master, "1, 2"), _var(master, "3, 4")]
    first_initial = dict(tab._config["initial_parameters"])
    second_initial = dict(tab._config["initial_parameters"])
    first_initial["amplitude"] = 0.02
    second_initial["amplitude"] = 0.05
    tab._channel_group_settings = [
        {
            "exploration_var": _var(master, 0.3),
            "warmup_var": _var(master, 2),
            "candidate_pool_var": _var(master, 700),
            "local_pool_var": _var(master, 70),
            "start_mode_var": _var(master, "specific"),
            "initial_parameters": first_initial,
        },
        {
            "exploration_var": _var(master, 0.5),
            "warmup_var": _var(master, 6),
            "candidate_pool_var": _var(master, 900),
            "local_pool_var": _var(master, 90),
            "start_mode_var": _var(master, "random"),
            "initial_parameters": second_initial,
        },
    ]

    tab._sync_channel_groups(show_error=False)

    assert tab._config["channel_groups"][0]["exploration"] == 0.3
    assert tab._config["channel_groups"][0]["n_initial_points"] == 2
    assert tab._config["channel_groups"][0]["initial_parameters"]["amplitude"] == 0.02
    assert tab._config["channel_groups"][1]["exploration"] == 0.5
    assert tab._config["channel_groups"][1]["n_initial_points"] == 6
    assert tab._config["channel_groups"][1]["initial_parameters"]["amplitude"] == 0.05
