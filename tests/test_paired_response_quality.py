from core.bo_session import compute_paired_response_quality


def test_zero_buffer_classic_q_invalidates_paired_q():
    scoring = {
        "mode": "classic",
        "channel_weights": {
            "snr": 0.0,
            "peak_height": 0.0,
            "peak_shape": 0.0,
            "baseline": 0.0,
            "replicate_consistency": 0.0,
            "success": 1.0,
            "noise_penalty": 0.0,
            "snr_saturation": 20.0,
        },
        "paired_response_weights": {
            "buffer_classic_Q": 0.25,
            "target_classic_Q": 0.25,
            "delta_peak": 10.0,
            "delta_scale_uA": 1.0,
        },
        "run_weights": {
            "low_channel_threshold": 0.5,
            "lambda_variability": 0.0,
            "lambda_failed": 0.0,
            "lambda_low": 0.0,
        },
    }
    buffer_metrics = {
        "1": {
            "mean_peak_current_uA": 0.0,
            "success_score": 0.0,
        }
    }
    target_metrics = {
        "1": {
            "mean_peak_current_uA": 100.0,
            "success_score": 1.0,
        }
    }

    quality = compute_paired_response_quality(buffer_metrics, target_metrics, scoring)
    channel = quality["channel_components"]["1"]

    assert channel["buffer_classic_Q"] == 0.0
    assert channel["target_classic_Q"] > 0.0
    assert channel["delta_peak_score"] > 0.0
    assert channel["valid_classic_pair"] is False
    assert channel["buffer_classic_Q_contribution"] == 0.0
    assert channel["target_classic_Q_contribution"] == 0.0
    assert channel["delta_peak_contribution"] == 0.0
    assert channel["paired_Q_channel"] == 0.0
    assert quality["Q_run"] == 0.0


def test_paired_q_contribution_terms_sum_to_paired_q():
    scoring = {
        "mode": "classic",
        "channel_weights": {
            "snr": 0.0,
            "peak_height": 0.0,
            "peak_shape": 0.0,
            "baseline": 0.0,
            "replicate_consistency": 0.0,
            "success": 1.0,
            "noise_penalty": 0.0,
            "snr_saturation": 20.0,
        },
        "paired_response_weights": {
            "buffer_classic_Q": 0.25,
            "target_classic_Q": 0.25,
            "delta_peak": 1.0,
            "delta_scale_uA": 1.0,
        },
        "run_weights": {
            "low_channel_threshold": 0.5,
            "lambda_variability": 0.0,
            "lambda_failed": 0.0,
            "lambda_low": 0.0,
        },
    }
    buffer_metrics = {
        "1": {
            "mean_peak_current_uA": 1.0,
            "success_score": 1.0,
        }
    }
    target_metrics = {
        "1": {
            "mean_peak_current_uA": 4.0,
            "success_score": 1.0,
        }
    }

    quality = compute_paired_response_quality(buffer_metrics, target_metrics, scoring)
    channel = quality["channel_components"]["1"]
    contribution_sum = (
        channel["buffer_classic_Q_contribution"]
        + channel["target_classic_Q_contribution"]
        + channel["delta_peak_contribution"]
    )

    assert channel["valid_classic_pair"] is True
    assert contribution_sum == channel["paired_Q_channel"]
    assert quality["mean_buffer_classic_Q_contribution"] == channel["buffer_classic_Q_contribution"]
    assert quality["mean_target_classic_Q_contribution"] == channel["target_classic_Q_contribution"]
    assert quality["mean_delta_peak_contribution"] == channel["delta_peak_contribution"]
