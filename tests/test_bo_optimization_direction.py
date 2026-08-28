import copy
import json
import random
from pathlib import Path

from core.bo_session import (
    BOIntegrationSession,
    candidate_key,
    compute_channel_quality,
    compute_paired_response_quality,
    compute_run_quality,
    load_bo_config,
    load_bo_setup_metadata,
    normalize_bo_config,
    save_bo_config,
    save_bo_setup_metadata,
)
from gui.tab_bayesian_optimization import BayesianOptimizationTab
from gui.tab_queue import QueueTab


def _directional_scoring():
    return {
        "mode": "classic",
        "channel_weights": {
            "peak_prominence": 1.0,
            "repeat_scan_snr": 0.0,
            "peak_height": 0.0,
            "peak_shape": 0.0,
            "baseline": 0.0,
            "replicate_consistency": 0.0,
            "success": 0.0,
            "noise_penalty": 1.0,
        },
        "run_weights": {
            "lambda_variability": 1.0,
            "lambda_failed": 0.0,
            "lambda_low": 2.0,
            "low_channel_threshold": 5.0,
            "lambda_repeat_std": 3.0,
        },
        "paired_response_weights": {
            "buffer_classic_Q": 0.0,
            "target_classic_Q": 0.0,
            "peak_prominence": 1.0,
            "repeat_scan_snr": 0.0,
            "lambda_repeat_std": 3.0,
        },
    }


def _config(direction="maximize"):
    return normalize_bo_config(
        {
            "n_initial_points": 0,
            "initial_parameters": {
                "begin_potential": -0.7,
                "end_potential": -0.1,
                "step_potential": 0.002,
                "amplitude": 0.04,
                "frequency": 100.0,
                "conditioning_potential": -0.7,
                "conditioning_time": 0.0,
            },
            "parameters": {
                "amplitude": {
                    "mode": "active",
                    "space": "continuous",
                    "min": 0.01,
                    "max": 0.08,
                    "step": None,
                    "value": 0.04,
                }
            },
            "acquisition": {
                "use_gp": False,
                "exploration": 0.0,
                "optimization_direction": direction,
            },
        }
    )


def _session(direction="maximize"):
    config = _config(direction)
    session = BOIntegrationSession.__new__(BOIntegrationSession)
    session.config = config
    session._start_candidate = dict(config["initial_parameters"])
    session._rng = random.Random(42)

    def candidate(amplitude):
        params = copy.deepcopy(config["initial_parameters"])
        params["amplitude"] = amplitude
        return params

    session.observations = [
        {"params": candidate(0.02), "Q_run": -2.0},
        {"params": candidate(0.07), "Q_run": 1.0},
    ]
    session.candidates = [candidate(0.021), candidate(0.069)]
    return session


def test_minimize_direction_treats_more_negative_q_as_better():
    session = _session("minimize")

    best = session.best_observation()
    choice = session._choose_candidate_current(session.candidates)

    assert best["Q_run"] == -2.0
    assert candidate_key(choice) == candidate_key(session.candidates[0])


def test_maximize_direction_keeps_more_positive_q_as_better():
    session = _session("maximize")

    best = session.best_observation()
    choice = session._choose_candidate_current(session.candidates)

    assert best["Q_run"] == 1.0
    assert candidate_key(choice) == candidate_key(session.candidates[1])


def test_survey_direction_keeps_sign_and_prefers_larger_magnitude():
    session = _session("survey")

    best = session.best_observation()
    choice = session._choose_candidate_current(session.candidates)

    assert best["Q_run"] == -2.0
    assert candidate_key(choice) == candidate_key(session.candidates[0])


def test_maximize_and_minimize_normalizes_from_display_label():
    config = _config("maximize and minimize")

    assert config["acquisition"]["optimization_direction"] == "maximize_and_minimize"


def test_maximize_and_minimize_asks_two_isolated_optimizer_streams(tmp_path):
    config = _config("maximize_and_minimize")
    session = BOIntegrationSession(config, tmp_path)

    first = session.ask_batch(1)

    assert [suggestion.optimization_direction for suggestion in first] == [
        "maximize",
        "minimize",
    ]
    assert first[0].iteration == first[1].iteration == 1
    assert first[0].method_id != first[1].method_id

    session.observations = [
        {
            "iteration": 1,
            "group_id": 1,
            "params": dict(first[0].params),
            "Q_run": 10.0,
            "optimization_direction": "maximize",
        },
        {
            "iteration": 1,
            "group_id": 1,
            "params": dict(first[1].params),
            "Q_run": -10.0,
            "optimization_direction": "minimize",
        },
    ]
    session.pending_batch = []

    captured = []
    original_choose = session._choose_candidate

    def capture(available, pending_params=None, observations=None, config=None):
        captured.append(
            (
                config["acquisition"]["optimization_direction"],
                [obs["optimization_direction"] for obs in observations],
            )
        )
        return original_choose(available, pending_params, observations, config)

    session._choose_candidate = capture
    second = session.ask_batch(1)

    assert [suggestion.optimization_direction for suggestion in second] == [
        "maximize",
        "minimize",
    ]
    assert captured == [
        ("maximize", ["maximize"]),
        ("minimize", ["minimize"]),
    ]


def test_dual_direction_queue_items_use_virtual_analysis_channels(tmp_path):
    class Registry:
        def save_script(self, *_args, **_kwargs):
            return tmp_path / "method.psmethod", "method.psmethod"

        def hash_key_for(self, path):
            return Path(path).stem

    session = BOIntegrationSession(_config("maximize_and_minimize"), tmp_path)
    maximize, minimize = session.ask_batch(1)

    maximize_item = session.build_queue_items(Registry(), maximize)[0]
    minimize_item = session.build_queue_items(Registry(), minimize)[0]

    assert maximize_item["method_ref"]["mux_channel"] == 1
    assert maximize_item["method_ref"]["channel_label"] == "1_max"
    assert maximize_item["bo_ref"]["channel_label"] == "1_max"
    assert minimize_item["method_ref"]["mux_channel"] == 1
    assert minimize_item["method_ref"]["channel_label"] == "1_min"


def test_dual_direction_metrics_are_saved_as_distinct_virtual_channels():
    metrics = {
        "1": {"peak_prominence": 2.0},
        "1_max": {"peak_prominence": 5.0},
        "1_min": {"peak_prominence": 3.0},
    }

    maximize = BOIntegrationSession._filter_metrics(metrics, [1], "maximize")
    minimize = BOIntegrationSession._filter_metrics(metrics, [1], "minimize")

    assert maximize == {"1_max": {"peak_prominence": 5.0}}
    assert minimize == {"1_min": {"peak_prominence": 3.0}}


def test_dual_direction_measurement_tag_keeps_physical_mux_and_virtual_channel():
    class Session:
        def next_meas_tag(self):
            return "meas_20260824_1200_001"

        def next_meas_tag_with_mux(self, mux_channel):
            return f"ordinary_ch{mux_channel}"

    tab = QueueTab.__new__(QueueTab)
    tab._session = Session()
    item = {
        "method_ref": {"mux_channel": 1, "channel_label": "1_max"},
        "bo_ref": {"optimization_direction": "maximize"},
    }

    assert tab._next_measurement_tag(item, 1) == "meas_20260824_1200_001_ch1_max"


def test_results_label_displays_virtual_direction_channel():
    observation = {
        "group_id": 1,
        "group_name": "Group 1",
        "analysis_channels": ["1_min"],
    }

    assert BayesianOptimizationTab._observation_result_label(observation) == "ch1 min"


def test_mixed_group_directions_survive_normalization_and_persistence(tmp_path):
    config = _config("maximize")
    config["channel_groups"] = [
        {"name": "Positive", "channels": [1], "optimization_direction": "maximize"},
        {"name": "Negative", "channels": [2], "optimization_direction": "minimize"},
        {"name": "Either", "channels": [3], "optimization_direction": "survey"},
    ]

    normalized = normalize_bo_config(config)
    assert [group["optimization_direction"] for group in normalized["channel_groups"]] == [
        "maximize",
        "minimize",
        "survey",
    ]

    config_path = save_bo_config(normalized, tmp_path / "mixed-directions.json")
    session = BOIntegrationSession(load_bo_config(config_path), tmp_path)

    assert session._config_for_group(1)["acquisition"]["optimization_direction"] == "maximize"
    assert session._config_for_group(2)["acquisition"]["optimization_direction"] == "minimize"
    assert session._config_for_group(3)["acquisition"]["optimization_direction"] == "survey"


def test_last_bo_setup_metadata_round_trip_preserves_config_and_ui_settings(tmp_path):
    config = _config("minimize")
    config["channel_groups"] = [
        {"channels": [1, 2], "optimization_direction": "minimize"},
    ]
    ui_settings = {
        "config_path": str(tmp_path / "custom.json"),
        "target_iterations": "37",
        "paired_target_exchange_block": "target.json",
    }
    metadata_path = tmp_path / "last_bo_setup_metadata.json"

    saved_path = save_bo_setup_metadata(config, ui_settings, metadata_path)
    loaded = load_bo_setup_metadata(saved_path)

    assert loaded is not None
    assert loaded["bo_config"]["acquisition"]["optimization_direction"] == "minimize"
    assert loaded["bo_config"]["channel_groups"][0]["optimization_direction"] == "minimize"
    assert loaded["ui_settings"] == ui_settings
    assert loaded["saved_at"]


def test_invalid_last_bo_setup_metadata_is_ignored(tmp_path):
    metadata_path = tmp_path / "last_bo_setup_metadata.json"
    metadata_path.write_text("not JSON", encoding="utf-8")

    assert load_bo_setup_metadata(metadata_path) is None


def test_classic_channel_noise_penalty_moves_opposite_for_each_direction():
    metrics = {
        "peak_prominence": 10.0,
        "mean_background_rms_uA": 2.0,
        "success_score": 1.0,
    }
    scoring = _directional_scoring()

    maximize = compute_channel_quality(metrics, scoring, "maximize")
    minimize = compute_channel_quality(metrics, scoring, "minimize")

    assert maximize["Q_channel"] == 8.0
    assert maximize["noise_penalty_adjustment"] == -2.0
    assert minimize["Q_channel"] == 12.0
    assert minimize["noise_penalty_adjustment"] == 2.0


def test_classic_run_penalties_worsen_both_maximize_and_minimize_scores():
    metrics = {
        "1": {
            "peak_prominence": 4.0,
            "repeat_relative_std": 0.5,
            "success_score": 1.0,
        },
        "2": {
            "peak_prominence": 8.0,
            "repeat_relative_std": 0.5,
            "success_score": 1.0,
        },
    }
    scoring = _directional_scoring()
    scoring["channel_weights"]["noise_penalty"] = 0.0

    maximize = compute_run_quality(metrics, scoring, "maximize")
    minimize = compute_run_quality(metrics, scoring, "minimize")

    assert maximize["optimization_direction"] == "maximize"
    assert maximize["run_penalty_adjustment"] < 0.0
    assert maximize["Q_run"] < maximize["mean_Q_channel"]
    assert minimize["optimization_direction"] == "minimize"
    assert minimize["run_penalty_adjustment"] > 0.0
    assert minimize["Q_run"] == 0.0
    assert maximize["poor_channel_fraction"] == 0.5  # 4 is poor when maximizing.
    assert minimize["poor_channel_fraction"] == 0.5  # 8 is poor when minimizing.


def test_paired_signal_sign_and_run_penalties_match_optimization_direction():
    scoring = _directional_scoring()
    scoring["channel_weights"]["noise_penalty"] = 0.0
    buffer = {
        "1": {
            "mean_peak_current_uA": 8.0,
            "peak_prominence": 8.0,
            "mean_background_rms_uA": 1.0,
            "repeat_relative_std": 0.5,
            "success_score": 1.0,
        }
    }
    target = {
        "1": {
            "mean_peak_current_uA": 2.0,
            "peak_prominence": 2.0,
            "mean_background_rms_uA": 1.0,
            "repeat_relative_std": 0.5,
            "success_score": 1.0,
        }
    }

    maximize = compute_paired_response_quality(buffer, target, scoring, "maximize")
    minimize = compute_paired_response_quality(buffer, target, scoring, "minimize")

    assert maximize["mean_Q_channel"] == -3.0
    assert minimize["mean_Q_channel"] == -3.0
    assert maximize["run_penalty_adjustment"] < 0.0
    assert minimize["run_penalty_adjustment"] > 0.0
    assert maximize["Q_run"] == 0.0
    assert minimize["Q_run"] > -3.0


def test_paired_run_clipping_and_survey_signed_scores():
    scoring = _directional_scoring()
    scoring["channel_weights"]["noise_penalty"] = 0.0
    scoring["run_weights"] = {
        "lambda_variability": 0.0,
        "lambda_failed": 0.0,
        "lambda_low": 0.0,
        "low_channel_threshold": 0.5,
        "lambda_repeat_std": 0.0,
    }
    scoring["paired_response_weights"]["lambda_repeat_std"] = 0.0

    negative_buffer = {"1": {"mean_peak_current_uA": 8.0, "peak_prominence": 8.0, "mean_background_rms_uA": 1.0}}
    negative_target = {"1": {"mean_peak_current_uA": 2.0, "peak_prominence": 2.0, "mean_background_rms_uA": 1.0}}
    positive_buffer = {"1": {"mean_peak_current_uA": 2.0, "peak_prominence": 2.0, "mean_background_rms_uA": 1.0}}
    positive_target = {"1": {"mean_peak_current_uA": 8.0, "peak_prominence": 8.0, "mean_background_rms_uA": 1.0}}

    assert compute_paired_response_quality(negative_buffer, negative_target, scoring, "maximize")["Q_run"] == 0.0
    assert compute_paired_response_quality(positive_buffer, positive_target, scoring, "minimize")["Q_run"] == 0.0
    assert compute_paired_response_quality(negative_buffer, negative_target, scoring, "survey")["Q_run"] == -3.0
    assert compute_paired_response_quality(positive_buffer, positive_target, scoring, "survey")["Q_run"] == 3.0


def test_group_import_uses_each_groups_direction_for_penalty_sign(tmp_path):
    config = _config("maximize")
    config["channel_groups"] = [
        {"name": "Positive", "channels": [1], "optimization_direction": "maximize"},
        {"name": "Negative", "channels": [2], "optimization_direction": "minimize"},
    ]
    config["scoring"] = _directional_scoring()
    config["scoring"]["channel_weights"]["noise_penalty"] = 0.0
    session = BOIntegrationSession(normalize_bo_config(config), tmp_path)
    suggestions = session.ask_next_groups()

    observations = []
    for suggestion in suggestions:
        payload = tmp_path / f"group_{suggestion.group_id}.json"
        payload.write_text(
            json.dumps(
                {
                    "channel_metrics": {
                        str(suggestion.channels[0]): {
                            "peak_prominence": 6.0,
                            "repeat_relative_std": 0.5,
                            "success_score": 1.0,
                        }
                    }
                }
            ),
            encoding="utf-8",
        )
        observations.append(session.import_analysis(payload, suggestion=suggestion))

    maximize, minimize = observations
    assert maximize["optimization_direction"] == "maximize"
    assert minimize["optimization_direction"] == "minimize"
    assert maximize["Q_run"] < maximize["quality"]["mean_Q_channel"]
    assert minimize["Q_run"] == 0.0
