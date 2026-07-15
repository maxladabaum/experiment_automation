import random

from core.bo_session import BOIntegrationSession, candidate_key, normalize_bo_config


def test_choose_candidate_with_zero_warmup_and_no_observations_does_not_crash():
    config = normalize_bo_config(
        {
            "objective": "paired_response",
            "n_initial_points": 0,
            "paired_warmup_cycles": 0,
            "paired_batch_size": 1,
            "channels": [1, 2, 3],
            "channel_groups": [
                {
                    "name": "Group 1",
                    "channels": [1, 2, 3],
                    "n_initial_points": 0,
                    "initial_point_mode": "random",
                }
            ],
            "initial_parameters": {
                "begin_potential": -0.6,
                "end_potential": 0.0,
                "step_potential": 0.002,
                "amplitude": 0.036,
                "frequency": 200.0,
                "conditioning_potential": -0.6,
                "conditioning_time": 0.0,
            },
            "parameters": {
                "amplitude": {
                    "mode": "active",
                    "space": "continuous",
                    "min": 0.01,
                    "max": 0.08,
                    "step": None,
                    "value": 0.036,
                }
            },
            "acquisition": {
                "use_gp": False,
                "exploration": 0.5,
                "initial_point_mode": "random",
            },
        }
    )
    session = BOIntegrationSession.__new__(BOIntegrationSession)
    session.config = config
    session._start_candidate = dict(config["initial_parameters"])
    session._rng = random.Random(42)
    session.observations = []

    candidates = [
        dict(config["initial_parameters"], amplitude=0.02),
        dict(config["initial_parameters"], amplitude=0.07),
    ]

    choice = session._choose_candidate_current(candidates)

    assert candidate_key(choice) in {candidate_key(candidate) for candidate in candidates}
