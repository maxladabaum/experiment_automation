import copy
import random

from core.bo_session import BOIntegrationSession, candidate_key, normalize_bo_config


def _config(direction="maximize"):
    return normalize_bo_config(
        {
            "n_initial_points": 0,
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
                "use_gp": False,
                "exploration": 0.0,
                "optimization_direction": direction,
            },
        }
    )


def _session(direction="maximize"):
    config = _config(direction)
    session = BOIntegrationSession.__new__(BOIntegrationSession)
    session.config = config
    session._start_candidate = dict(config["initial_parameters"])
    session._rng = random.Random(42)

    def candidate(amplitude):
        params = copy.deepcopy(config["initial_parameters"])
        params["amplitude"] = amplitude
        return params

    session.observations = [
        {"params": candidate(0.02), "Q_run": -2.0},
        {"params": candidate(0.07), "Q_run": 1.0},
    ]
    session.candidates = [candidate(0.021), candidate(0.069)]
    return session


def test_minimize_direction_treats_more_negative_q_as_better():
    session = _session("minimize")

    best = session.best_observation()
    choice = session._choose_candidate_current(session.candidates)

    assert best["Q_run"] == -2.0
    assert candidate_key(choice) == candidate_key(session.candidates[0])


def test_maximize_direction_keeps_more_positive_q_as_better():
    session = _session("maximize")

    best = session.best_observation()
    choice = session._choose_candidate_current(session.candidates)

    assert best["Q_run"] == 1.0
    assert candidate_key(choice) == candidate_key(session.candidates[1])
