from analysis_worker.bo_headless import _apply_result_constraints, _build_channel_metrics


def _row(channel=1):
    return {
        "channel": channel,
        "status": "OK",
        "peak_voltage": -0.25,
        "voltage": [-0.45, -0.35, -0.25, -0.15, -0.05],
        "left_min_idx": 1,
        "right_min_idx": 3,
        "peak_current": 1.0,
        "background_current_rms": 0.1,
        "peak_offset_norm": 0.0,
        "bracket_width_V": 0.2,
    }


def test_bo_voltage_constraints_accept_peak_and_minima_in_range():
    results = [_row()]

    _apply_result_constraints(
        results,
        {
            "peak_voltage_min_v": -0.3,
            "peak_voltage_max_v": -0.2,
            "left_min_voltage_min_v": -0.4,
            "left_min_voltage_max_v": -0.3,
            "right_min_voltage_min_v": -0.2,
            "right_min_voltage_max_v": -0.1,
        },
    )

    assert results[0]["status"] == "OK"
    metrics = _build_channel_metrics(results)
    assert metrics["1"]["success_score"] == 1.0


def test_bo_voltage_constraints_fail_and_score_zero_for_out_of_range_minima():
    results = [_row()]

    _apply_result_constraints(
        results,
        {
            "peak_voltage_min_v": -0.3,
            "peak_voltage_max_v": -0.2,
            "left_min_voltage_min_v": -0.2,
            "right_min_voltage_max_v": -0.2,
        },
    )

    assert results[0]["status"] == "FAILED"
    assert "left minimum voltage" in results[0]["error"]
    assert "right minimum voltage" in results[0]["error"]
    metrics = _build_channel_metrics(results)
    assert metrics["1"]["success_score"] == 0.0
    assert metrics["1"]["peak_shape_score"] == 0.0
