import csv
import json

import pytest

from core.bo_session import (
    BOIntegrationSession,
    channel_groups,
    normalize_bo_config,
    validate_bo_config,
)
from gui.tab_bayesian_optimization import BayesianOptimizationTab


def _config():
    return normalize_bo_config({
        "name": "group-test",
        "channels": [1, 2, 3, 4],
        "channel_groups": [
            {"name": "Left", "channels": [1, 2]},
            {"name": "Right", "channels": [3, 4]},
        ],
        "n_initial_points": 1,
        "acquisition": {"use_gp": False},
        "parameters": {
            "amplitude": {
                "mode": "active",
                "space": "discrete",
                "value": 0.036,
                "values": [0.02, 0.036, 0.05],
            },
        },
    })


def test_channel_group_validation_rejects_overlap():
    config = _config()
    config["channel_groups"][1]["channels"] = [2, 3, 4]
    assert any("only belong to one group" in error for error in validate_bo_config(config))


def test_group_suggestions_have_independent_histories(tmp_path):
    session = BOIntegrationSession(_config(), tmp_path)
    first = session.ask_next_groups()
    assert [(item.group_name, item.channels) for item in first] == [
        ("Left", [1, 2]),
        ("Right", [3, 4]),
    ]

    for suggestion, q in zip(first, (0.1, 0.9)):
        payload = tmp_path / f"{suggestion.group_id}.json"
        payload.write_text(json.dumps({
            "channel_metrics": {
                str(channel): {
                    "snr": q * 20,
                    "peak_shape_score": q,
                    "baseline_stability_score": q,
                    "replicate_consistency_score": q,
                    "success_score": 1,
                }
                for channel in range(1, 5)
            }
        }))
        session.import_analysis(payload, suggestion=suggestion)

    second = session.ask_next_groups()
    assert all(item.iteration == 2 for item in second)
    assert {item.group_id for item in second} == {1, 2}
    assert all(obs["channels"] == ([1, 2] if obs["group_id"] == 1 else [3, 4]) for obs in session.observations)
    assert all(set(obs["channel_metrics"]) == {str(ch) for ch in obs["channels"]} for obs in session.observations)

    with open(session.record_dir / "history.csv", newline="", encoding="utf-8") as handle:
        rows = list(csv.DictReader(handle))
    assert {(row["group_name"], row["channels"]) for row in rows} == {
        ("Left", "1,2"),
        ("Right", "3,4"),
    }


def test_legacy_channels_become_one_group():
    config = normalize_bo_config({"channels": [2, 5]})
    assert channel_groups(config) == [{"id": 1, "name": "Group 1", "channels": [2, 5]}]


def test_grouped_history_ids_are_resolved_without_integer_conversion():
    tab = BayesianOptimizationTab.__new__(BayesianOptimizationTab)
    tab._history_rows = {
        "g1:i1": {"group_id": 1, "iteration": 1},
        "g2:i1": {"group_id": 2, "iteration": 1},
    }

    assert tab._resolve_history_key("g1:i1") == "g1:i1"
    assert tab._resolve_history_key(1) == "g2:i1"


def test_trend_series_are_split_by_group_and_iteration():
    tab = BayesianOptimizationTab.__new__(BayesianOptimizationTab)
    rows = [
        {"group_id": 1, "group_name": "Group 1", "iteration": 2, "Q_run": 0.2},
        {"group_id": 2, "group_name": "Group 2", "iteration": 1, "Q_run": 0.8},
        {"group_id": 1, "group_name": "Group 1", "iteration": 1, "Q_run": 0.1},
    ]

    assert tab._grouped_trend_series(rows, "Q_run") == [
        ("Group 1", [("g1:i1", 1, 0.1), ("g1:i2", 2, 0.2)]),
        ("Group 2", [("g2:i1", 1, 0.8)]),
    ]
