import csv
import json
from pathlib import Path

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
    assert not list(session.plots_dir.glob("*.png"))
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
    plot_names = [path.name for path in session.plots_dir.glob("*.png")]
    assert plot_names
    assert all(name.startswith(("group_01_", "group_02_")) for name in plot_names)


def test_groups_use_distinct_optimizer_settings_and_starting_parameters(tmp_path):
    config = _config()
    config["channel_groups"][0].update({
        "exploration": 0.3,
        "n_initial_points": 1,
        "candidate_pool_size": 700,
        "local_candidate_pool_size": 70,
        "initial_point_mode": "specific",
        "gp_falloff_fractions": {name: 0.15 for name in config["parameters"]},
        "initial_parameters": {"amplitude": 0.02},
    })
    config["channel_groups"][1].update({
        "exploration": 0.5,
        "n_initial_points": 4,
        "candidate_pool_size": 900,
        "local_candidate_pool_size": 90,
        "initial_point_mode": "specific",
        "gp_falloff_fractions": {name: 0.35 for name in config["parameters"]},
        "initial_parameters": {"amplitude": 0.05},
    })
    session = BOIntegrationSession(config, tmp_path)

    suggestions = session.ask_next_groups()

    assert suggestions[0].params["amplitude"] == pytest.approx(0.02)
    assert suggestions[1].params["amplitude"] == pytest.approx(0.05)
    assert session._config_for_group(1)["acquisition"]["exploration"] == 0.3
    assert session._config_for_group(2)["acquisition"]["exploration"] == 0.5
    assert session._config_for_group(1)["n_initial_points"] == 1
    assert session._config_for_group(2)["n_initial_points"] == 4
    assert session._config_for_group(1)["acquisition"]["candidate_pool_size"] == 700
    assert session._config_for_group(2)["acquisition"]["local_candidate_pool_size"] == 90
    assert session._config_for_group(1)["acquisition"]["gp_falloff_fractions"]["frequency"] == 0.15
    assert session._config_for_group(2)["acquisition"]["gp_falloff_fractions"]["frequency"] == 0.35


def test_random_group_starts_do_not_require_group_initial_parameters(tmp_path):
    config = _config()
    for group in config["channel_groups"]:
        group["initial_point_mode"] = "random"
        group.pop("initial_parameters", None)

    assert validate_bo_config(config) == []

    session = BOIntegrationSession(config, tmp_path)
    suggestions = session.ask_next_groups()

    assert len(suggestions) == 2
    assert {suggestion.group_id for suggestion in suggestions} == {1, 2}
    assert all(
        suggestion.params["amplitude"] in {0.02, 0.036, 0.05}
        for suggestion in suggestions
    )


def test_random_start_mode_uses_acquisition_after_observations(tmp_path):
    config = _config()
    config["n_initial_points"] = 0
    config["acquisition"]["use_gp"] = True
    config["channel_groups"] = [
        {
            "id": 1,
            "name": "Only",
            "channels": [1, 2],
            "initial_point_mode": "random",
            "n_initial_points": 0,
        }
    ]
    session = BOIntegrationSession(config, tmp_path)
    group_config = session._config_for_group(1)
    observations = [
        {"params": dict(session.candidates[0]), "Q_run": 0.1},
        {"params": dict(session.candidates[1]), "Q_run": 0.2},
    ]
    tried = {tuple(sorted(obs["params"].items())) for obs in observations}
    available = [
        candidate for candidate in session.candidates
        if tuple(sorted(candidate.items())) not in tried
    ]
    random_start_candidate = dict(available[0])
    acquisition_candidate = dict(available[-1])

    session._resolve_start_candidate = lambda _config=None: random_start_candidate
    session._gp_expected_improvement_candidate = (
        lambda _available, pending_params=None: acquisition_candidate
    )

    selected = session._choose_candidate(
        available,
        observations=observations,
        config=group_config,
    )

    assert selected == acquisition_candidate


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


def test_grouped_peak_and_noise_use_only_observation_channels(tmp_path):
    tab = BayesianOptimizationTab.__new__(BayesianOptimizationTab)
    results_csv = tmp_path / "results.csv"
    results_csv.write_text(
        "\n".join([
            "channel,status,peak_current,background_current_rms",
            "1,OK,10,1",
            "2,OK,12,3",
            "3,OK,100,50",
            "4,OK,120,60",
        ]),
        encoding="utf-8",
    )
    tab._analysis_results_paths_for_observation = lambda observation: [results_csv]

    peak, noise = tab._observation_peak_rms({"channels": [1, 2]})

    assert peak == pytest.approx(11.0)
    assert noise == pytest.approx(2.0)


def test_rebuilt_group_metrics_ignore_other_groups_in_results_csv(tmp_path):
    tab = BayesianOptimizationTab.__new__(BayesianOptimizationTab)
    results_csv = tmp_path / "results.csv"
    results_csv.write_text(
        "\n".join([
            "channel,status,peak_current,background_current_rms,peak_offset_norm,bracket_width_V",
            "1,OK,10,1,0.1,0.05",
            "2,OK,12,3,0.2,0.06",
            "3,OK,100,50,0.9,0.30",
            "4,OK,120,60,0.8,0.35",
        ]),
        encoding="utf-8",
    )
    tab._analysis_results_paths_for_observation = lambda observation: [results_csv]

    metrics = tab._rebuilt_channel_metrics_for_observation({"channels": [1, 2]})

    assert set(metrics) == {"1", "2"}


def test_raw_trace_rows_are_filtered_to_observation_channels(tmp_path):
    tab = BayesianOptimizationTab.__new__(BayesianOptimizationTab)
    kept = tmp_path / "ch003_trace.csv"
    dropped = tmp_path / "ch009_trace.csv"
    kept.write_text("voltage,current\n0,0\n", encoding="utf-8")
    dropped.write_text("voltage,current\n0,0\n", encoding="utf-8")

    analysis_record = tmp_path / "analysis.json"
    results_csv = tmp_path / "results.csv"
    results_csv.write_text(
        "\n".join([
            "channel,file_path,scan_number",
            f"3,{kept},1",
            f"9,{dropped},1",
        ]),
        encoding="utf-8",
    )
    analysis_record.write_text(json.dumps({"results_csv": str(results_csv)}), encoding="utf-8")

    tab._analysis_record_paths_with_phase = lambda observation: [(analysis_record, "")]
    tab._resolve_observation_file_path = lambda raw_path, observation=None: Path(raw_path)
    tab._infer_measurement_phase_from_path = lambda path: ""
    tab._infer_channel_from_path = lambda path: Path(path).stem
    tab._infer_measurement_id_from_path = lambda path: ""

    rows = tab._raw_trace_rows_for_observation({"channels": [3, 4], "archived_measurements": []})

    assert [row["channel"] for row in rows] == ["3"]


def test_external_corrected_rows_are_filtered_to_observation_channels():
    tab = BayesianOptimizationTab.__new__(BayesianOptimizationTab)
    tab._external_analysis_results = lambda observation: [
        {"channel": 3, "voltage": [0.0, 1.0], "smoothed_corrected_current": [0.1, 0.2], "file_name": "ch3.csv"},
        {"channel": 9, "voltage": [0.0, 1.0], "smoothed_corrected_current": [0.3, 0.4], "file_name": "ch9.csv"},
    ]

    rows, diagnostics = tab._corrected_trace_rows_for_observation({"channels": [3, 4]})

    assert diagnostics == []
    assert [row["channel"] for row in rows] == ["3"]


def test_paired_external_results_skip_duplicate_analysis_results_json(tmp_path):
    tab = BayesianOptimizationTab.__new__(BayesianOptimizationTab)
    buffer_results = tmp_path / "buffer_results.json"
    target_results = tmp_path / "target_results.json"
    buffer_results.write_text(
        json.dumps([{"channel": 2, "voltage": [0, 1], "smoothed_corrected_current": [0.1, 0.2]}]),
        encoding="utf-8",
    )
    target_results.write_text(
        json.dumps([{"channel": 2, "voltage": [0, 1], "smoothed_corrected_current": [0.3, 0.4]}]),
        encoding="utf-8",
    )
    tab._resolve_observation_file_path = lambda raw_path, observation=None: Path(raw_path)
    observation = {
        "objective": "paired_response",
        "analysis_results_json": str(target_results),
        "buffer_analysis_results_json": str(buffer_results),
        "target_analysis_results_json": str(target_results),
    }

    rows = tab._external_analysis_results(observation)

    assert [row["_bo_phase"] for row in rows] == ["buffer", "target"]


def test_paired_analysis_records_skip_duplicate_analysis_record(tmp_path):
    tab = BayesianOptimizationTab.__new__(BayesianOptimizationTab)
    buffer_record = tmp_path / "buffer_summary.json"
    target_record = tmp_path / "target_summary.json"
    buffer_record.write_text("{}", encoding="utf-8")
    target_record.write_text("{}", encoding="utf-8")
    tab._resolve_observation_file_path = lambda raw_path, observation=None: Path(raw_path)
    tab._infer_measurement_phase_from_path = lambda path: ""
    observation = {
        "objective": "paired_response",
        "analysis_record": str(target_record),
        "buffer_analysis_record": str(buffer_record),
        "target_analysis_record": str(target_record),
    }

    rows = tab._analysis_record_paths_with_phase(observation)

    assert rows == [(buffer_record, "buffer"), (target_record, "target")]


def test_paired_raw_trace_rows_deduplicate_same_paths_without_collapsing_repeats(tmp_path):
    tab = BayesianOptimizationTab.__new__(BayesianOptimizationTab)
    tab._bo_session = None
    buffer_csv = tmp_path / "buffer_ch4.csv"
    buffer_repeat_csv = tmp_path / "buffer_ch4_repeat_02.csv"
    target_csv = tmp_path / "target_ch4.csv"
    target_repeat_csv = tmp_path / "target_ch4_repeat_02.csv"
    buffer_csv.write_text("Voltage,Current\n0,1\n", encoding="utf-8")
    buffer_repeat_csv.write_text("Voltage,Current\n0,1.1\n", encoding="utf-8")
    target_csv.write_text("Voltage,Current\n0,2\n", encoding="utf-8")
    target_repeat_csv.write_text("Voltage,Current\n0,2.1\n", encoding="utf-8")
    buffer_results = tmp_path / "buffer_results.csv"
    target_results = tmp_path / "target_results.csv"
    buffer_results.write_text(
        f"file_path,channel,scan_number\n{buffer_csv},4,1\n{buffer_repeat_csv},4,2\n",
        encoding="utf-8",
    )
    target_results.write_text(
        f"file_path,channel,scan_number\n{target_csv},4,1\n{target_repeat_csv},4,2\n",
        encoding="utf-8",
    )
    buffer_record = tmp_path / "buffer_summary.json"
    target_record = tmp_path / "target_summary.json"
    buffer_record.write_text(json.dumps({"results_csv": str(buffer_results)}), encoding="utf-8")
    target_record.write_text(json.dumps({"results_csv": str(target_results)}), encoding="utf-8")

    rows = tab._raw_trace_rows_for_observation(
        {
            "objective": "paired_response",
            "channels": [4],
            "archived_measurements": [
                str(buffer_csv),
                str(buffer_repeat_csv),
                str(target_csv),
                str(target_repeat_csv),
            ],
            "buffer_analysis_record": str(buffer_record),
            "target_analysis_record": str(target_record),
        }
    )

    assert [(row["phase"], row["channel"], Path(row["path"]).name) for row in rows] == [
        ("buffer", "4", "buffer_ch4.csv"),
        ("buffer", "4", "buffer_ch4_repeat_02.csv"),
        ("target", "4", "target_ch4.csv"),
        ("target", "4", "target_ch4_repeat_02.csv"),
    ]


def test_paired_corrected_trace_rows_preserve_external_repeat_scans():
    tab = BayesianOptimizationTab.__new__(BayesianOptimizationTab)
    tab._external_analysis_results = lambda observation: [
        {
            "_bo_phase": "buffer",
            "channel": 4,
            "scan_number": 1,
            "voltage": [0.0, 1.0],
            "smoothed_corrected_current": [0.1, 0.2],
        },
        {
            "_bo_phase": "buffer",
            "channel": 4,
            "scan_number": 2,
            "voltage": [0.0, 1.0],
            "smoothed_corrected_current": [0.11, 0.21],
        },
        {
            "_bo_phase": "target",
            "channel": 4,
            "scan_number": 1,
            "voltage": [0.0, 1.0],
            "smoothed_corrected_current": [0.3, 0.4],
        },
        {
            "_bo_phase": "target",
            "channel": 4,
            "scan_number": 2,
            "voltage": [0.0, 1.0],
            "smoothed_corrected_current": [0.31, 0.41],
        },
    ]

    rows, diagnostics = tab._corrected_trace_rows_for_observation(
        {"objective": "paired_response", "channels": [4]}
    )

    assert diagnostics == []
    assert [(row["phase"], row["channel"], row.get("scan")) for row in rows] == [
        ("buffer", "4", 1),
        ("buffer", "4", 2),
        ("target", "4", 1),
        ("target", "4", 2),
    ]


def test_paired_score_table_builds_one_row_per_analyzed_trace():
    tab = BayesianOptimizationTab.__new__(BayesianOptimizationTab)
    tab._bo_session = None
    tab._config = _config()
    tab._external_analysis_results = lambda observation: [
        {
            "_bo_phase": phase,
            "channel": 4,
            "scan_number": scan,
            "status": "OK",
            "peak_current": peak,
            "background_current_rms": 0.1,
            "peak_offset_norm": 0.0,
            "bracket_width_V": 0.2,
            "bracket_point_count": 10,
            "crop_point_count": 20,
        }
        for phase, scan, peak in (
            ("buffer", 1, 1.0),
            ("buffer", 2, 1.2),
            ("target", 1, 2.0),
            ("target", 2, 2.2),
        )
    ]

    rows = tab._individual_trace_score_rows(
        {"objective": "paired_response", "channels": [4]}
    )

    assert [(row["phase"], row["channel"], row["trace"]) for row in rows] == [
        ("buffer", "4", 1),
        ("buffer", "4", 2),
        ("target", "4", 1),
        ("target", "4", 2),
    ]
    assert [row["metrics"]["mean_peak_current_uA"] for row in rows] == [
        1.0,
        1.2,
        2.0,
        2.2,
    ]


def test_grouped_analysis_records_are_namespaced(tmp_path):
    session = BOIntegrationSession(_config(), tmp_path)
    first = session.ask_next_groups()

    source = tmp_path / "analysis_summary.json"
    source.write_text(json.dumps({"channel_metrics": {"1": {"snr": 1, "success_score": 1}}}), encoding="utf-8")

    left = session.import_analysis(source, suggestion=first[0])
    right = session.import_analysis(source, suggestion=first[1])

    assert Path(left["analysis_record"]).name.startswith("group_01_iter_001_")
    assert Path(right["analysis_record"]).name.startswith("group_02_iter_001_")
    assert left["analysis_record"] != right["analysis_record"]


def test_run_pending_analysis_uses_group_namespaced_output_stem(tmp_path):
    session = BOIntegrationSession(_config(), tmp_path)
    suggestion = session.ask_next_groups()[1]

    import core.analysis_worker as analysis_worker

    captured = {}
    summary_path = tmp_path / "summary.json"
    results_json = tmp_path / "results.json"
    results_json.write_text("[]", encoding="utf-8")
    summary_path.write_text(json.dumps({"ok": True}), encoding="utf-8")
    original = analysis_worker.run_analysis
    try:
        def fake_run_analysis(request_path, request):
            captured["request_path"] = str(request_path)
            captured["request"] = dict(request)
            return {
                "summary_path": str(summary_path),
                "channel_metrics": {"3": {"snr": 1, "success_score": 1}},
                "result_count": 1,
                "results_json": str(results_json),
            }

        analysis_worker.run_analysis = fake_run_analysis
        session.run_pending_analysis(folders=[tmp_path], suggestion=suggestion)
    finally:
        analysis_worker.run_analysis = original

    assert captured["request"]["output_stem"] == "bo_group_02_iter_001"
    assert Path(captured["request_path"]).name == "group_02_iter_001_analysis_request.json"


def test_reanalysis_filters_mixed_archived_measurements_by_group(tmp_path):
    session = BOIntegrationSession(_config(), tmp_path)

    keep = tmp_path / "ch003_keep.csv"
    drop = tmp_path / "ch009_drop.csv"
    keep.write_text("voltage,current\n0,0\n", encoding="utf-8")
    drop.write_text("voltage,current\n0,0\n", encoding="utf-8")

    import core.analysis_worker as analysis_worker

    captured = {}
    summary_path = tmp_path / "reanalyzed_summary.json"
    results_json = tmp_path / "reanalyzed_results.json"
    results_json.write_text("[]", encoding="utf-8")
    summary_path.write_text(json.dumps({"ok": True}), encoding="utf-8")
    original = analysis_worker.run_analysis
    try:
        def fake_run_analysis(request_path, request):
            captured["request"] = dict(request)
            return {
                "summary_path": str(summary_path),
                "channel_metrics": {"3": {"snr": 1, "success_score": 1, "ok_scan_count": 1}},
                "result_count": 1,
                "results_json": str(results_json),
                "analysis_engine": "test",
            }

        analysis_worker.run_analysis = fake_run_analysis
        session.reanalyze_observation(
            {
                "iteration": 1,
                "group_id": 1,
                "channels": [3, 4],
                "archived_measurements": [str(keep), str(drop)],
            },
            analysis={},
            scoring=session.config["scoring"],
            output_dir=tmp_path / "reanalysis",
        )
    finally:
        analysis_worker.run_analysis = original

    assert captured["request"]["folders"] == [str(keep.resolve())]


def test_candidate_prediction_rows_are_group_specific(tmp_path):
    session = BOIntegrationSession(_config(), tmp_path)
    suggestions = session.ask_next_groups()

    for suggestion, q in zip(suggestions, (0.1, 0.9)):
        payload = tmp_path / f"group_{suggestion.group_id}.json"
        payload.write_text(json.dumps({
            "channel_metrics": {
                str(channel): {
                    "snr": q * 20,
                    "peak_shape_score": q,
                    "baseline_stability_score": q,
                    "replicate_consistency_score": q,
                    "success_score": 1,
                }
                for channel in suggestion.channels
            }
        }), encoding="utf-8")
        session.import_analysis(payload, suggestion=suggestion)

    left_rows, left_meta, _gp = session._candidate_prediction_rows(group_id=1)
    right_rows, right_meta, _gp = session._candidate_prediction_rows(group_id=2)

    assert left_rows
    assert right_rows
    assert left_meta["observation_count"] == 1
    assert right_meta["observation_count"] == 1
    assert left_meta["group_id"] == 1
    assert right_meta["group_id"] == 2
    assert left_meta["best_Q_run"] != right_meta["best_Q_run"]


def test_surrogate_history_filters_to_selected_group():
    tab = BayesianOptimizationTab.__new__(BayesianOptimizationTab)
    tab._bo_session = type("Session", (), {
        "observations": [
            {"group_id": 1, "iteration": 1, "params": {"amplitude": 0.02}, "Q_run": 0.1},
            {"group_id": 2, "iteration": 1, "params": {"amplitude": 0.05}, "Q_run": 0.9},
            {"group_id": 1, "iteration": 2, "params": {"amplitude": 0.036}, "Q_run": 0.2},
        ]
    })()
    tab._selected_history_observation = {"group_id": 1, "iteration": 2}
    tab._surrogate_iteration_var = type("Var", (), {"get": lambda self: "2"})()

    observations = tab._surrogate_observations_so_far()

    assert [(obs["group_id"], obs["iteration"]) for obs in observations] == [(1, 1), (1, 2)]


def test_classic_auto_loop_appends_queue_history_and_runs_only_new_rows():
    config = normalize_bo_config(
        {
            "channels": [1],
            "channel_groups": [{"name": "Group 1", "channels": [1]}],
        }
    )
    suggestion = type(
        "Suggestion",
        (),
        {"iteration": 1, "group_id": 1, "group_name": "Group 1"},
    )()

    class FakeBOSession:
        session_id = "bo-session"
        observations = []

        def __init__(self):
            self.config = config

        def ask_next_for_group(self, _group_id):
            return suggestion

        def build_queue_items(self, _registry, _suggestion):
            return [
                {
                    "type": "SWV",
                    "status": "pending",
                    "bo_ref": {"session_id": self.session_id},
                }
            ]

        def record_queued(self, _suggestion, _items):
            return None

    old_item = {
        "type": "SWV",
        "status": "completed",
        "bo_ref": {"session_id": "bo-session"},
    }
    measurement_session = type(
        "MeasurementSession",
        (),
        {
            "is_running": False,
            "measurement_queue": [old_item],
            "registry": object(),
        },
    )()
    started_at = []
    tab = BayesianOptimizationTab.__new__(BayesianOptimizationTab)
    tab._auto_running = True
    tab._auto_target_var = type("Var", (), {"get": lambda self: "2"})()
    tab._auto_status_var = type("Var", (), {"set": lambda self, _value: None})()
    tab._bo_session = FakeBOSession()
    tab._session = measurement_session
    tab._suggestion = None
    tab._render_suggestion = lambda: None
    tab._add_to_queue = measurement_session.measurement_queue.append
    tab._refresh_queue = lambda: None
    tab._refresh_record_files = lambda: None
    tab._run_queue = lambda: pytest.fail("Completed queue rows must not be replayed")
    tab._run_queue_from_index = started_at.append

    tab._auto_submit_next()

    assert measurement_session.measurement_queue[0] is old_item
    assert len(measurement_session.measurement_queue) == 2
    assert started_at == [1]
