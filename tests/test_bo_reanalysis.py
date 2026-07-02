import json
from pathlib import Path

import pytest

from core.bo_session import BOIntegrationSession, load_bo_config


def _worker(tmp_path, calls):
    def run(_request_path, request):
        calls.append(request)
        phase = "target" if request["output_stem"].endswith("_target") else (
            "buffer" if request["output_stem"].endswith("_buffer") else "classic"
        )
        peak = {"buffer": 2.0, "target": 5.0, "classic": 4.0}[phase]
        noise = {"buffer": 0.5, "target": 1.0, "classic": 0.5}[phase]
        metrics = {
            "1": {
                "mean_peak_current_uA": peak,
                "mean_background_rms_uA": noise,
                "snr": peak / noise,
                "success_score": 1.0,
            }
        }
        results = tmp_path / f"{request['output_stem']}_results.json"
        results.write_text(json.dumps([{"channel": 1, "peak_current": peak}]), encoding="utf-8")
        summary_path = tmp_path / f"{request['output_stem']}.json"
        summary = {
            "result_count": 1,
            "channel_metrics": metrics,
            "results_json": str(results),
            "results_csv": "",
            "analysis_engine": "external-test-worker",
            "summary_path": str(summary_path),
        }
        summary_path.write_text(json.dumps(summary), encoding="utf-8")
        return summary

    return run


def _raw(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("Potential (V),Current (uA)\n-0.5,1\n-0.4,2\n", encoding="utf-8")
    return str(path)


def test_reanalyze_classic_observation(monkeypatch, tmp_path):
    calls = []
    monkeypatch.setattr("core.analysis_worker.run_analysis", _worker(tmp_path, calls))
    config = load_bo_config("optimizer/bo_configs/default_swv_bo.json")
    session = BOIntegrationSession(config, tmp_path / "experiment")
    observation = {
        "iteration": 1,
        "method_id": "method-1",
        "params": dict(config["initial_parameters"]),
        "archived_measurements": [_raw(tmp_path / "experiment" / "legacy" / "iter_001" / "ch001.csv")],
    }

    rebuilt = session.reanalyze_observation(
        observation,
        analysis=dict(config["analysis"]),
        scoring=dict(config["scoring"]),
        output_dir=tmp_path / "reanalyzed",
    )

    assert len(calls) == 1
    assert rebuilt["analysis_engine"] == "external-test-worker"
    assert rebuilt["channel_metrics"]["1"]["mean_peak_current_uA"] == 4.0
    assert rebuilt["Q_run"] > 0.0


def test_reanalyze_discovers_raw_files_for_legacy_observation(monkeypatch, tmp_path):
    calls = []
    monkeypatch.setattr("core.analysis_worker.run_analysis", _worker(tmp_path, calls))
    config = load_bo_config("optimizer/bo_configs/default_swv_bo.json")
    experiment = tmp_path / "experiment"
    session = BOIntegrationSession(config, experiment)
    raw_path = Path(_raw(experiment / "legacy" / "iter_003" / "ch001.csv"))
    _raw(experiment / "legacy" / "iter_003" / "analysis_results.csv")

    rebuilt = session.reanalyze_observation(
        {"iteration": 3, "method_id": "old-method"},
        analysis=dict(config["analysis"]),
        scoring=dict(config["scoring"]),
        output_dir=tmp_path / "reanalyzed",
    )

    assert rebuilt["Q_run"] > 0.0
    assert calls[0]["folders"] == [str(raw_path.resolve())]


def test_reanalyze_does_not_replace_q_when_all_peak_analyses_failed(monkeypatch, tmp_path):
    config = load_bo_config("optimizer/bo_configs/default_swv_bo.json")
    session = BOIntegrationSession(config, tmp_path / "experiment")
    raw_path = _raw(tmp_path / "experiment" / "legacy" / "iter_001" / "ch001.csv")
    results_path = tmp_path / "failed_results.json"
    results_path.write_text(
        json.dumps([{"channel": 1, "status": "FAILED", "error": "peak below cutoff"}]),
        encoding="utf-8",
    )
    summary_path = tmp_path / "failed_summary.json"

    def failed_worker(_request_path, _request):
        summary = {
            "result_count": 1,
            "channel_metrics": {
                "1": {
                    "success_score": 0.0,
                    "ok_scan_count": 0,
                    "total_scan_count": 1,
                }
            },
            "results_json": str(results_path),
            "summary_path": str(summary_path),
        }
        summary_path.write_text(json.dumps(summary), encoding="utf-8")
        return summary

    monkeypatch.setattr("core.analysis_worker.run_analysis", failed_worker)

    with pytest.raises(ValueError, match="peak below cutoff"):
        session.reanalyze_observation(
            {
                "iteration": 1,
                "method_id": "method-1",
                "archived_measurements": [raw_path],
            },
            analysis=dict(config["analysis"]),
            scoring=dict(config["scoring"]),
            output_dir=tmp_path / "reanalyzed",
        )


def test_reanalyze_paired_observation_routes_phases_and_rebuilds_paired_q(monkeypatch, tmp_path):
    calls = []
    monkeypatch.setattr("core.analysis_worker.run_analysis", _worker(tmp_path, calls))
    config = load_bo_config("optimizer/bo_configs/default_swv_bo.json")
    config["objective"] = "paired_response"
    session = BOIntegrationSession(config, tmp_path / "experiment")
    observation = {
        "iteration": 1,
        "method_id": "method-1",
        "objective": "paired_response",
        "params": dict(config["initial_parameters"]),
        "archived_measurements": [
            _raw(tmp_path / "experiment" / "legacy" / "iter_001_buffer" / "ch001.csv"),
            _raw(tmp_path / "experiment" / "legacy" / "iter_001_target" / "ch001.csv"),
        ],
    }

    rebuilt = session.reanalyze_observation(
        observation,
        analysis=dict(config["analysis"]),
        scoring=dict(config["scoring"]),
        output_dir=tmp_path / "reanalyzed",
    )

    assert len(calls) == 2
    assert {call["output_stem"].rsplit("_", 1)[-1] for call in calls} == {"buffer", "target"}
    assert rebuilt["quality"]["channel_components"]["1"]["delta_peak_height_uA"] == 3.0
    assert rebuilt["Q_run"] == pytest.approx(2.0)
