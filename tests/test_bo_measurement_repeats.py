from pathlib import Path

import pytest

from analysis_worker.bo_headless import _build_channel_metrics as external_metrics
from core.bo_analysis import _build_channel_metrics as built_in_metrics
from core.bo_session import (
    BOIntegrationSession,
    BOSuggestion,
    compute_paired_response_quality,
    compute_run_quality,
    normalize_bo_config,
)


class _Registry:
    def __init__(self, root):
        self.root = Path(root)
        self.calls = []

    def save_script(self, _technique, _script, **kwargs):
        self.calls.append(kwargs)
        path = self.root / f"method_{len(self.calls)}.psmethod"
        return path, path.name

    def hash_key_for(self, path):
        return Path(path).stem


def _config(measurements=1):
    return normalize_bo_config(
        {
            "measurements_per_channel": measurements,
            "channels": [1, 2, 3],
            "channel_groups": [{"name": "All", "channels": [1, 2, 3]}],
            "acquisition": {"use_gp": False},
        }
    )


def _suggestion():
    config = _config()
    return BOSuggestion(
        iteration=1,
        method_id="g01_i001_test",
        params=dict(config["initial_parameters"]),
        created_at="2026-01-01T00:00:00",
        group_name="All",
        channels=[1, 2, 3],
    )


def _rows():
    base = {
        "channel": 1,
        "status": "OK",
        "peak_offset_norm": 0.0,
        "bracket_width_V": 0.2,
        "bracket_point_count": 10,
        "crop_point_count": 20,
    }
    return [
        {**base, "peak_current": 1.0, "background_current_rms": 0.1},
        {**base, "peak_current": 1.4, "background_current_rms": 0.3},
    ]


def _scoring(repeat_penalty):
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
        },
        "run_weights": {
            "lambda_variability": 0.0,
            "lambda_failed": 0.0,
            "lambda_low": 0.0,
            "low_channel_threshold": 0.0,
            "lambda_repeat_std": repeat_penalty,
        },
        "paired_response_weights": {
            "buffer_classic_Q": 0.0,
            "target_classic_Q": 0.0,
            "peak_prominence": 1.0,
            "repeat_scan_snr": 0.0,
            "lambda_repeat_std": repeat_penalty,
        },
    }


def test_queue_repeats_each_channel_contiguously_and_records_repeat_metadata(tmp_path):
    session = BOIntegrationSession(_config(measurements=2), tmp_path)
    items = session.build_queue_items(_Registry(tmp_path), _suggestion(), phase="buffer")

    assert [item["method_ref"]["mux_channel"] for item in items] == [1, 1, 2, 2, 3, 3]
    assert [item["bo_ref"]["measurement_repeat_index"] for item in items] == [1, 2, 1, 2, 1, 2]
    assert all(item["bo_ref"]["measurement_repeat_count"] == 2 for item in items)
    assert all(item["bo_ref"]["phase"] == "buffer" for item in items)


@pytest.mark.parametrize("builder", [built_in_metrics, external_metrics])
def test_repeat_analysis_averages_signal_and_noise_and_reports_std(builder):
    metrics = builder(_rows())["1"]

    assert metrics["mean_peak_current_uA"] == pytest.approx(1.2)
    assert metrics["mean_background_rms_uA"] == pytest.approx(0.2)
    assert metrics["snr"] == pytest.approx(6.0)
    assert metrics["peak_prominence"] == pytest.approx(6.0)
    assert metrics["repeat_scan_snr"] == pytest.approx(
        1.2 / (0.4 / (2 ** 0.5))
    )
    assert metrics["std_peak_current_uA"] == pytest.approx(0.4 / (2 ** 0.5))
    assert metrics["std_background_rms_uA"] == pytest.approx(0.2 / (2 ** 0.5))
    assert metrics["repeat_relative_std"] > 0.0


@pytest.mark.parametrize("builder", [built_in_metrics, external_metrics])
def test_single_measurement_has_zero_repeat_scan_snr(builder):
    metrics = builder(_rows()[:1])["1"]

    assert metrics["peak_prominence"] == pytest.approx(10.0)
    assert metrics["repeat_scan_snr"] == 0.0


def test_classic_q_subtracts_configured_repeat_std_penalty():
    metrics = built_in_metrics(_rows())
    without_penalty = compute_run_quality(metrics, _scoring(0.0))
    with_penalty = compute_run_quality(metrics, _scoring(2.0))

    expected = 2.0 * metrics["1"]["repeat_relative_std"]
    assert with_penalty["repeat_std_penalty"] == pytest.approx(expected)
    assert without_penalty["Q_run"] - with_penalty["Q_run"] == pytest.approx(expected)


def test_classic_q_weights_peak_prominence_and_repeat_scan_snr_separately():
    metrics = built_in_metrics(_rows())
    scoring = _scoring(0.0)
    scoring["channel_weights"]["peak_prominence"] = 2.0
    scoring["channel_weights"]["repeat_scan_snr"] = 3.0

    quality = compute_run_quality(metrics, scoring)
    channel = quality["channel_components"]["1"]

    expected = 2.0 * metrics["1"]["peak_prominence"] + 3.0 * metrics["1"]["repeat_scan_snr"]
    assert channel["peak_prominence_raw"] == pytest.approx(metrics["1"]["peak_prominence"])
    assert channel["repeat_scan_snr_raw"] == pytest.approx(metrics["1"]["repeat_scan_snr"])
    assert channel["Q_channel"] == pytest.approx(expected)
    assert channel["peak_prominence_contribution"] == pytest.approx(
        2.0 * metrics["1"]["peak_prominence"]
    )
    assert channel["repeat_scan_snr_contribution"] == pytest.approx(
        3.0 * metrics["1"]["repeat_scan_snr"]
    )
    assert sum(
        channel[key]
        for key in (
            "peak_prominence_contribution",
            "repeat_scan_snr_contribution",
            "peak_height_contribution",
            "peak_shape_contribution",
            "baseline_contribution",
            "replicate_consistency_contribution",
            "success_contribution",
            "noise_penalty_adjustment",
            "clip_adjustment",
        )
    ) == pytest.approx(channel["Q_channel"])


def test_paired_q_uses_averages_and_subtracts_repeat_std_penalty():
    buffer_metrics = built_in_metrics(_rows())
    target_rows = [
        {**row, "peak_current": row["peak_current"] + 1.0}
        for row in _rows()
    ]
    target_metrics = built_in_metrics(target_rows)
    without_penalty = compute_paired_response_quality(
        buffer_metrics, target_metrics, _scoring(0.0)
    )
    with_penalty = compute_paired_response_quality(
        buffer_metrics, target_metrics, _scoring(2.0)
    )

    channel = with_penalty["channel_components"]["1"]
    unpenalized_channel = without_penalty["channel_components"]["1"]
    assert channel["buffer_peak_height_raw"] == pytest.approx(1.2)
    assert channel["target_peak_height_raw"] == pytest.approx(2.2)
    assert channel["delta_peak_height_uA"] == pytest.approx(1.0)
    assert channel["paired_Q_channel"] == pytest.approx(
        unpenalized_channel["paired_Q_channel"]
    )
    assert without_penalty["Q_run"] - with_penalty["Q_run"] == pytest.approx(
        with_penalty["repeat_std_penalty"]
    )


def test_paired_repeat_penalty_is_applied_at_run_not_channel_level():
    buffer_metrics = built_in_metrics(_rows())
    target_metrics = built_in_metrics(
        [{**row, "peak_current": 0.5 * row["peak_current"]} for row in _rows()]
    )

    without_penalty = compute_paired_response_quality(
        buffer_metrics, target_metrics, _scoring(0.0), "survey"
    )
    with_penalty = compute_paired_response_quality(
        buffer_metrics, target_metrics, _scoring(2.0), "survey"
    )

    raw_channel_q = without_penalty["channel_components"]["1"]["paired_Q_channel"]
    penalized_channel_q = with_penalty["channel_components"]["1"]["paired_Q_channel"]
    assert raw_channel_q < 0.0
    assert penalized_channel_q == pytest.approx(raw_channel_q)
    assert abs(with_penalty["Q_run"]) < abs(without_penalty["Q_run"])


def test_classic_and_paired_repeat_penalty_weights_are_independent():
    metrics = built_in_metrics(_rows())
    scoring = _scoring(1.0)
    scoring["paired_response_weights"]["lambda_repeat_std"] = 3.0

    classic = compute_run_quality(metrics, scoring)
    paired = compute_paired_response_quality(metrics, metrics, scoring)

    repeat_std = metrics["1"]["repeat_relative_std"]
    assert classic["repeat_std_penalty"] == pytest.approx(repeat_std)
    assert paired["repeat_std_penalty"] == pytest.approx(3.0 * repeat_std)


def test_repeat_count_is_normalized_to_at_least_one():
    assert _config(measurements=0)["measurements_per_channel"] == 1
    assert _config(measurements=3)["measurements_per_channel"] == 3


def test_legacy_snr_and_delta_weights_migrate_to_new_metric_names():
    config = normalize_bo_config(
        {
            "scoring": {
                "channel_weights": {"snr": 0.7, "snr_saturation": 12.0},
                "paired_response_weights": {"delta_peak": 2.5},
            }
        }
    )

    channel = config["scoring"]["channel_weights"]
    paired = config["scoring"]["paired_response_weights"]
    assert channel["peak_prominence"] == 0.7
    assert channel["peak_prominence_saturation"] == 12.0
    assert channel["repeat_scan_snr"] == 0.0
    assert "snr" not in channel
    assert paired["peak_prominence"] == 2.5
    assert paired["repeat_scan_snr"] == 0.0
    assert "delta_peak" not in paired
