import pytest

from core.bo_session import compute_paired_response_quality


def test_paired_q_is_delta_peak_over_combined_noise():
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
            "buffer_classic_Q": 0.0,
            "target_classic_Q": 0.0,
            "peak_prominence": 1.0,
            "repeat_scan_snr": 0.0,
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
            "mean_peak_current_uA": 2.0,
            "mean_background_rms_uA": 0.5,
            "success_score": 1.0,
        }
    }
    target_metrics = {
        "1": {
            "mean_peak_current_uA": 8.0,
            "mean_background_rms_uA": 1.5,
            "success_score": 1.0,
        }
    }

    quality = compute_paired_response_quality(buffer_metrics, target_metrics, scoring)
    channel = quality["channel_components"]["1"]

    assert channel["delta_peak_height_uA"] == 6.0
    assert channel["buffer_channel_noise"] == 0.5
    assert channel["target_channel_noise"] == 1.5
    assert channel["combined_channel_noise"] == 2.0
    assert channel["repeat_scan_snr"] == 0.0
    assert channel["buffer_classic_Q_contribution"] == 0.0
    assert channel["target_classic_Q_contribution"] == 0.0
    assert channel["delta_peak_contribution"] == 3.0
    assert channel["paired_Q_channel"] == 3.0
    assert quality["Q_run"] == 3.0


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
            "peak_prominence": 1.0,
            "repeat_scan_snr": 0.0,
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
            "mean_background_rms_uA": 0.25,
            "success_score": 1.0,
        }
    }
    target_metrics = {
        "1": {
            "mean_peak_current_uA": 4.0,
            "mean_background_rms_uA": 0.75,
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
    assert channel["buffer_classic_Q"] == 1.0
    assert channel["target_classic_Q"] == 1.0
    assert channel["buffer_classic_Q_contribution"] == 0.25
    assert channel["target_classic_Q_contribution"] == 0.25
    assert channel["delta_peak_contribution"] == 3.0
    assert channel["paired_Q_channel"] == 3.5
    assert contribution_sum == channel["paired_Q_channel"]
    assert quality["mean_buffer_classic_Q_contribution"] == channel["buffer_classic_Q_contribution"]
    assert quality["mean_target_classic_Q_contribution"] == channel["target_classic_Q_contribution"]
    assert quality["mean_delta_peak_contribution"] == channel["delta_peak_contribution"]
    assert quality["Q_run"] == 3.5


def test_paired_q_weights_repeat_scan_snr_and_peak_prominence_separately():
    scoring = {
        "mode": "classic",
        "channel_weights": {"success": 1.0, "peak_prominence": 0.0},
        "paired_response_weights": {
            "buffer_classic_Q": 0.0,
            "target_classic_Q": 0.0,
            "peak_prominence": 2.0,
            "repeat_scan_snr": 3.0,
            "repeat_scan_snr_definition": "pairwise",
        },
        "run_weights": {
            "low_channel_threshold": -100.0,
            "lambda_variability": 0.0,
            "lambda_failed": 0.0,
            "lambda_low": 0.0,
        },
    }
    buffer_metrics = {
        "1": {
            "mean_peak_current_uA": 2.0,
            "std_peak_current_uA": 0.25,
            "ok_scan_count": 2,
            "mean_background_rms_uA": 0.5,
            "success_score": 1.0,
        }
    }
    target_metrics = {
        "1": {
            "mean_peak_current_uA": 8.0,
            "std_peak_current_uA": 0.75,
            "ok_scan_count": 2,
            "mean_background_rms_uA": 1.5,
            "success_score": 1.0,
        }
    }

    quality = compute_paired_response_quality(buffer_metrics, target_metrics, scoring)
    channel = quality["channel_components"]["1"]

    assert channel["delta_peak_height_uA"] == 6.0
    assert channel["peak_prominence"] == 3.0
    pairwise_std = (1.25 / 3.0) ** 0.5
    assert channel["combined_peak_std_uA"] == pytest.approx(pairwise_std)
    assert channel["pairwise_peak_difference_std_uA"] == pytest.approx(
        pairwise_std
    )
    assert channel["repeat_scan_snr"] == pytest.approx(6.0 / pairwise_std)
    assert channel["peak_prominence_contribution"] == 6.0
    assert channel["repeat_scan_snr_contribution"] == pytest.approx(
        18.0 / pairwise_std
    )
    assert channel["paired_Q_channel"] == pytest.approx(
        6.0 + 18.0 / pairwise_std
    )


def test_pairwise_rescore_saves_every_target_minus_buffer_difference():
    scoring = {
        "channel_weights": {"success": 1.0, "peak_prominence": 0.0},
        "paired_response_weights": {
            "buffer_classic_Q": 0.0,
            "target_classic_Q": 0.0,
            "peak_prominence": 0.0,
            "repeat_scan_snr": 1.0,
            "repeat_scan_snr_definition": "pairwise",
        },
        "run_weights": {
            "low_channel_threshold": -100.0,
            "lambda_variability": 0.0,
            "lambda_failed": 0.0,
            "lambda_low": 0.0,
        },
    }
    buffer_metrics = {"1": {
        "peak_currents_uA": [1.0, 2.0],
        "mean_peak_current_uA": 1.5,
        "std_peak_current_uA": 0.5,
        "ok_scan_count": 2,
        "success_score": 1.0,
    }}
    target_metrics = {"1": {
        "peak_currents_uA": [4.0, 6.0],
        "mean_peak_current_uA": 5.0,
        "std_peak_current_uA": 1.0,
        "ok_scan_count": 2,
        "success_score": 1.0,
    }}

    quality = compute_paired_response_quality(buffer_metrics, target_metrics, scoring)
    channel = quality["channel_components"]["1"]
    differences = [3.0, 5.0, 2.0, 4.0]
    pairwise_std = 1.2909944487358056

    assert channel["pairwise_peak_differences_uA"] == differences
    assert channel["pairwise_mean_peak_difference_uA"] == pytest.approx(3.5)
    assert channel["pairwise_peak_difference_count"] == 4
    assert channel["pairwise_peak_difference_std_uA"] == pytest.approx(pairwise_std)
    assert channel["repeat_scan_snr"] == pytest.approx(3.5 / pairwise_std)
    assert quality["repeat_scan_snr_definition"] == "pairwise"
    assert quality["mean_pairwise_mean_peak_difference_uA"] == pytest.approx(3.5)
    assert quality["mean_pairwise_peak_difference_count"] == pytest.approx(4)
    assert quality["mean_pairwise_peak_difference_std_uA"] == pytest.approx(
        pairwise_std
    )
    assert quality["mean_pairwise_regularized_std_uA"] == pytest.approx(
        pairwise_std
    )
    assert quality["mean_pairwise_std_floor_uA"] == pytest.approx(0.0)
    assert quality["mean_unregularized_repeat_scan_snr"] == pytest.approx(
        3.5 / pairwise_std
    )


def test_paired_repeat_scan_snr_defaults_to_original_definition():
    scoring = {
        "channel_weights": {"success": 1.0, "peak_prominence": 0.0},
        "paired_response_weights": {
            "buffer_classic_Q": 0.0,
            "target_classic_Q": 0.0,
            "peak_prominence": 0.0,
            "repeat_scan_snr": 1.0,
        },
        "run_weights": {
            "lambda_variability": 0.0,
            "lambda_failed": 0.0,
            "lambda_low": 0.0,
        },
    }
    buffer_metrics = {
        "1": {
            "mean_peak_current_uA": 2.0,
            "std_peak_current_uA": 0.25,
            "ok_scan_count": 2,
            "success_score": 1.0,
        }
    }
    target_metrics = {
        "1": {
            "mean_peak_current_uA": 8.0,
            "std_peak_current_uA": 0.75,
            "ok_scan_count": 2,
            "success_score": 1.0,
        }
    }

    quality = compute_paired_response_quality(
        buffer_metrics,
        target_metrics,
        scoring,
        "survey",
    )
    channel = quality["channel_components"]["1"]

    assert channel["repeat_scan_snr_definition"] == "original"
    assert channel["combined_peak_std_uA"] == pytest.approx(1.0)
    assert channel["repeat_scan_snr"] == pytest.approx(6.0)
    assert "pairwise_peak_difference_std_uA" not in channel


def test_paired_q_is_zero_when_either_phase_failed():
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
        "run_weights": {
            "low_channel_threshold": 0.5,
            "lambda_variability": 0.0,
            "lambda_failed": 0.0,
            "lambda_low": 0.0,
        },
    }
    buffer_metrics = {
        "1": {
            "snr": 0.0,
            "success_score": 0.0,
            "ok_scan_count": 0,
            "total_scan_count": 1,
        }
    }
    target_metrics = {
        "1": {
            "mean_peak_current_uA": 8.0,
            "mean_background_rms_uA": 1.5,
            "success_score": 1.0,
        }
    }

    quality = compute_paired_response_quality(buffer_metrics, target_metrics, scoring)
    channel = quality["channel_components"]["1"]

    assert channel["success_score"] == 0.0
    assert channel["delta_peak_contribution"] == 0.0
    assert channel["paired_Q_channel"] == 0.0
    assert quality["Q_run"] == 0.0


def test_paired_q_preserves_negative_delta_score():
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
        "run_weights": {
            "low_channel_threshold": -10.0,
            "lambda_variability": 0.0,
            "lambda_failed": 0.0,
            "lambda_low": 0.0,
        },
    }
    buffer_metrics = {
        "1": {
            "mean_peak_current_uA": 8.0,
            "mean_background_rms_uA": 1.5,
            "success_score": 1.0,
        }
    }
    target_metrics = {
        "1": {
            "mean_peak_current_uA": 2.0,
            "mean_background_rms_uA": 0.5,
            "success_score": 1.0,
        }
    }

    quality = compute_paired_response_quality(
        buffer_metrics, target_metrics, scoring, "survey"
    )
    channel = quality["channel_components"]["1"]

    assert channel["delta_peak_height_uA"] == -6.0
    assert channel["paired_Q_channel"] == -3.0
    assert quality["Q_run"] == -3.0


def test_negative_paired_delta_subtracts_weighted_classic_q_terms():
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
        },
        "run_weights": {
            "low_channel_threshold": -10.0,
            "lambda_variability": 0.0,
            "lambda_failed": 0.0,
            "lambda_low": 0.0,
        },
    }
    buffer_metrics = {
        "1": {
            "mean_peak_current_uA": 8.0,
            "mean_background_rms_uA": 1.5,
            "success_score": 1.0,
        }
    }
    target_metrics = {
        "1": {
            "mean_peak_current_uA": 2.0,
            "mean_background_rms_uA": 0.5,
            "success_score": 1.0,
        }
    }

    quality = compute_paired_response_quality(
        buffer_metrics, target_metrics, scoring, "survey"
    )
    channel = quality["channel_components"]["1"]

    assert channel["delta_peak_contribution"] == -3.0
    assert channel["buffer_classic_Q_contribution"] == -0.25
    assert channel["target_classic_Q_contribution"] == -0.25
    assert channel["paired_Q_channel"] == -3.5
    assert quality["Q_run"] == -3.5
