import csv
import json

import pytest

from core.bo_session import BOIntegrationSession, compute_paired_response_quality
from gui.tab_bayesian_optimization import BayesianOptimizationTab


def _pairwise_observation():
    scoring = {
        "channel_weights": {"success": 1.0, "peak_prominence": 0.0},
        "paired_response_weights": {
            "buffer_classic_Q": 0.0,
            "target_classic_Q": 0.0,
            "peak_prominence": 0.0,
            "repeat_scan_snr": 1.0,
            "repeat_scan_snr_definition": "pairwise",
            "pairwise_std_floor_uA": 0.25,
        },
        "run_weights": {
            "lambda_variability": 0.0,
            "lambda_failed": 0.0,
            "lambda_low": 0.0,
        },
    }
    buffer_metrics = {
        "1": {
            "peak_currents_uA": [1.0, 2.0],
            "mean_peak_current_uA": 1.5,
            "std_peak_current_uA": 0.5,
            "ok_scan_count": 2,
            "success_score": 1.0,
        }
    }
    target_metrics = {
        "1": {
            "peak_currents_uA": [4.0, 6.0],
            "mean_peak_current_uA": 5.0,
            "std_peak_current_uA": 1.0,
            "ok_scan_count": 2,
            "success_score": 1.0,
        }
    }
    quality = compute_paired_response_quality(
        buffer_metrics,
        target_metrics,
        scoring,
    )
    quality["paired_batch_size"] = 1
    return {
        "iteration": 1,
        "group_id": 1,
        "group_name": "Group 1",
        "objective": "paired_response",
        "Q_run": quality["Q_run"],
        "quality": quality,
        "buffer_channel_metrics": buffer_metrics,
        "target_channel_metrics": target_metrics,
        "params": {},
    }


def test_pairwise_history_row_contains_every_numeric_snr_input():
    tab = BayesianOptimizationTab.__new__(BayesianOptimizationTab)
    tab._bo_session = None
    tab._config = {}
    observation = _pairwise_observation()

    row = dict(
        zip(
            tab._paired_history_columns(),
            tab._paired_history_values(observation),
        )
    )

    assert row["Pair Δ Mean"] == "3.5"
    assert row["Pair Count"] == "4"
    assert row["Pair Δ STD"] == "1.29099"
    assert row["Pair STD Floor"] == "0.25"
    assert row["Pair Reg STD"] == "1.31498"
    assert row["Pair Raw SNR"] == "2.71109"
    assert row["Paired Repeat SNR"] == "2.662"
    assert row["Buffer Peaks uA"] == "ch1:[1.0,2.0]"
    assert row["Target Peaks uA"] == "ch1:[4.0,6.0]"
    assert row["Pair Δ Values uA"] == "ch1:[3.0,5.0,2.0,4.0]"


@pytest.mark.parametrize(
    ("metric", "expected"),
    [
        ("Pairwise mean difference", 3.5),
        ("Pairwise difference count", 4.0),
        ("Pairwise sample STD", 1.2909944487358056),
        ("Pairwise regularized STD", 1.3149778198382918),
        ("Pairwise STD floor", 0.25),
        ("Pairwise unregularized SNR", 2.711088342345192),
        ("Paired repeat-scan SNR", 2.6616427790506595),
    ],
)
def test_pairwise_snr_inputs_are_available_as_trend_metrics(metric, expected):
    tab = BayesianOptimizationTab.__new__(BayesianOptimizationTab)
    observation = _pairwise_observation()

    assert tab._analysis_trend_value(observation, metric) == pytest.approx(expected)


def test_pairwise_snr_inputs_are_written_to_history_csv(tmp_path):
    observation = _pairwise_observation()
    observation.update(
        {
            "method_id": "method-1",
            "channels": [1],
            "completed_at": "2026-08-15T12:00:00",
            "analysis_record": "analysis.json",
        }
    )
    session = BOIntegrationSession.__new__(BOIntegrationSession)
    session.record_dir = tmp_path
    session.config = {"paired_batch_size": 1}
    session.observations = [observation]

    session._write_history_csv()

    with open(tmp_path / "history.csv", newline="", encoding="utf-8") as handle:
        row = next(csv.DictReader(handle))
    assert row["repeat_scan_snr_definition"] == "pairwise"
    assert float(row["mean_pairwise_mean_peak_difference_uA"]) == pytest.approx(3.5)
    assert float(row["mean_pairwise_peak_difference_count"]) == pytest.approx(4)
    assert float(row["mean_pairwise_peak_difference_std_uA"]) == pytest.approx(
        1.2909944487358056
    )
    assert float(row["mean_pairwise_regularized_std_uA"]) == pytest.approx(
        1.3149778198382918
    )
    assert float(row["mean_pairwise_std_floor_uA"]) == pytest.approx(0.25)
    assert float(row["mean_unregularized_repeat_scan_snr"]) == pytest.approx(
        2.711088342345192
    )
    assert json.loads(row["ch1_pairwise_peak_differences_uA"]) == [
        3.0,
        5.0,
        2.0,
        4.0,
    ]
    assert json.loads(row["ch1_buffer_peak_currents_uA"]) == [1.0, 2.0]
    assert json.loads(row["ch1_target_peak_currents_uA"]) == [4.0, 6.0]


def test_pairwise_history_reconstructs_mean_and_count_for_older_records():
    tab = BayesianOptimizationTab.__new__(BayesianOptimizationTab)
    tab._bo_session = None
    tab._config = {}
    observation = _pairwise_observation()
    quality = observation["quality"]
    quality.pop("mean_pairwise_mean_peak_difference_uA")
    quality.pop("mean_pairwise_peak_difference_count")
    component = quality["channel_components"]["1"]
    component.pop("pairwise_mean_peak_difference_uA")
    component.pop("pairwise_peak_difference_count")

    row = dict(
        zip(
            tab._paired_history_columns(),
            tab._paired_history_values(observation),
        )
    )

    assert row["Pair Δ Mean"] == "3.5"
    assert row["Pair Count"] == "4"
