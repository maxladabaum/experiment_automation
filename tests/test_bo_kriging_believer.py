import copy

import pytest

from core.bo_session import (
    BOIntegrationSession,
    encode_candidate,
    normalize_bo_config,
)


pytest.importorskip("numpy")
pytest.importorskip("sklearn")


def _session_with_amplitude_observations():
    config = normalize_bo_config(
        {
            "n_initial_points": 2,
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
                "gp_falloff_fractions": {
                    name: 0.2
                    for name in (
                        "begin_potential",
                        "end_potential",
                        "step_potential",
                        "amplitude",
                        "frequency",
                        "conditioning_potential",
                        "conditioning_time",
                    )
                }
            },
        }
    )
    session = BOIntegrationSession.__new__(BOIntegrationSession)
    session.config = config

    def candidate(amplitude):
        params = copy.deepcopy(config["initial_parameters"])
        params["amplitude"] = amplitude
        return params

    session.observations = [
        {"params": candidate(0.02), "Q_run": 0.2},
        {"params": candidate(0.04), "Q_run": 0.9},
        {"params": candidate(0.07), "Q_run": 0.3},
    ]
    return session, candidate


def test_kriging_believer_adds_pending_points_as_mean_fantasies():
    session, candidate = _session_with_amplitude_observations()
    pending = candidate(0.05)

    base_gp, base_train = session._fit_gp_surrogate()
    fantasy_gp, fantasy_train = session._fit_gp_surrogate([pending])

    assert base_gp is not None
    assert fantasy_gp is not None
    assert len(fantasy_train["y_train"]) == len(base_train["y_train"]) + 1

    encoded_pending = [encode_candidate(pending, session.config)]
    expected_fantasy = float(base_gp.predict(encoded_pending)[0])
    assert fantasy_train["y_train"][-1] == pytest.approx(expected_fantasy)


def test_kriging_believer_reduces_uncertainty_near_pending_point():
    session, candidate = _session_with_amplitude_observations()
    pending = candidate(0.05)
    nearby = candidate(0.051)

    base_gp, _ = session._fit_gp_surrogate()
    fantasy_gp, _ = session._fit_gp_surrogate([pending])
    encoded_nearby = [encode_candidate(nearby, session.config)]

    _, base_std = base_gp.predict(encoded_nearby, return_std=True)
    _, fantasy_std = fantasy_gp.predict(encoded_nearby, return_std=True)

    assert fantasy_std[0] < base_std[0]
