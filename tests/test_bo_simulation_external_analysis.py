import json

from core import bo_simulation
from core.bo_analysis import run_request
from core.bo_session import load_bo_config


def _external_worker_stub(calls):
    def run(request_path, request):
        assert request_path.exists()
        assert request["source"] == "experiment_automation_simulation"
        calls.append(dict(request))
        summary = run_request(request)
        summary_path = summary["summary_path"]
        with open(summary_path, "r", encoding="utf-8") as fh:
            payload = json.load(fh)
        payload["analysis_engine"] = "external-test-worker"
        with open(summary_path, "w", encoding="utf-8") as fh:
            json.dump(payload, fh)
        summary["analysis_engine"] = "external-test-worker"
        return summary

    return run


def _simulation_config(config):
    return {
        "dimensions": bo_simulation.default_dimensions(config),
        "grid_size": 5,
        "trace_points": 61,
        "measurement_noise": 0.01,
    }


def _with_groups(config):
    config = dict(config)
    config["channel_groups"] = [
        {"name": "First", "channels": [1, 2]},
        {"name": "Second", "channels": [3, 4]},
    ]
    config["channels"] = [1, 2, 3, 4]
    return config


def test_classic_simulation_uses_analysis_worker(monkeypatch, tmp_path):
    calls = []
    monkeypatch.setattr(bo_simulation, "run_analysis", _external_worker_stub(calls))
    config = load_bo_config("optimizer/bo_configs/default_swv_bo.json")

    result = bo_simulation.run_optimizer_simulation(
        config,
        _simulation_config(config),
        tmp_path,
        iterations=1,
    )

    assert len(calls) == 1
    assert result["session"].observations[0]["analysis_engine"] == "external-test-worker"


def test_paired_simulation_analyzes_buffer_and_target_externally(monkeypatch, tmp_path):
    calls = []
    monkeypatch.setattr(bo_simulation, "run_analysis", _external_worker_stub(calls))
    config = load_bo_config("optimizer/bo_configs/default_swv_bo.json")

    result = bo_simulation.run_paired_response_optimizer_simulation(
        config,
        {**_simulation_config(config), "paired_response": True},
        tmp_path,
        cycles=1,
        batch_size=1,
    )

    assert len(calls) == 2
    assert {call["output_stem"].rsplit("_", 1)[-1] for call in calls} == {"buffer", "target"}
    assert result["session"].observations[0]["analysis_engine"] == "external-test-worker"


def test_classic_simulation_runs_each_channel_group(monkeypatch, tmp_path):
    calls = []
    monkeypatch.setattr(bo_simulation, "run_analysis", _external_worker_stub(calls))
    config = _with_groups(load_bo_config("optimizer/bo_configs/default_swv_bo.json"))

    result = bo_simulation.run_optimizer_simulation(
        config,
        _simulation_config(config),
        tmp_path,
        iterations=1,
    )

    assert len(calls) == 2
    assert [row["group_name"] for row in result["rows"]] == ["First", "Second"]
    assert [set(obs["channel_metrics"]) for obs in result["session"].observations] == [
        {"1", "2"},
        {"3", "4"},
    ]


def test_paired_simulation_runs_each_channel_group(monkeypatch, tmp_path):
    calls = []
    monkeypatch.setattr(bo_simulation, "run_analysis", _external_worker_stub(calls))
    config = _with_groups(load_bo_config("optimizer/bo_configs/default_swv_bo.json"))

    result = bo_simulation.run_paired_response_optimizer_simulation(
        config,
        {**_simulation_config(config), "paired_response": True},
        tmp_path,
        cycles=1,
        batch_size=1,
    )

    assert len(calls) == 4
    assert result["group_count"] == 2
    assert result["total_swv_traces"] == 4
    assert {row["group_name"] for row in result["rows"]} == {"First", "Second"}
