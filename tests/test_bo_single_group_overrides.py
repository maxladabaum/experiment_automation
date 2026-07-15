from core.bo_session import BOIntegrationSession, normalize_bo_config


def test_single_group_batch_honors_group_warmup_override(tmp_path):
    config = normalize_bo_config(
        {
            "n_initial_points": 0,
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
                    "step": 0.01,
                    "value": 0.036,
                }
            },
            "acquisition": {
                "use_gp": False,
                "initial_point_mode": "specific",
            },
                "channel_groups": [
                    {
                        "name": "Group 1",
                        "channels": [1, 2],
                        "n_initial_points": 3,
                        "initial_point_mode": "random",
                        "candidate_pool_size": 50,
                    }
                ],
        }
    )
    session = BOIntegrationSession(config, tmp_path)

    captured = {}
    original_choose = session._choose_candidate

    def wrapped_choose(available, pending_params=None, observations=None, config=None):
        captured["n_initial_points"] = int((config or {}).get("n_initial_points", -1))
        captured["initial_point_mode"] = str((config or {}).get("acquisition", {}).get("initial_point_mode", ""))
        return original_choose(
            available,
            pending_params=pending_params,
            observations=observations,
            config=config,
        )

    session._choose_candidate = wrapped_choose

    suggestions = session.ask_batch(3)

    assert len(suggestions) == 3
    assert all(suggestion.group_id == 1 for suggestion in suggestions)
    assert captured["n_initial_points"] == 3
    assert captured["initial_point_mode"] == "random"


def test_single_group_available_candidates_use_group_candidate_pool_size(tmp_path):
    config = normalize_bo_config(
        {
            "n_initial_points": 0,
            "acquisition": {
                "candidate_pool_size": 1000,
            },
                "channel_groups": [
                    {
                        "name": "Group 1",
                        "channels": [1],
                        "candidate_pool_size": 50,
                        "local_candidate_pool_size": 0,
                    }
                ],
            "parameters": {
                "amplitude": {
                    "mode": "active",
                    "space": "continuous",
                    "min": 0.01,
                    "max": 0.08,
                    "step": 0.001,
                    "value": 0.036,
                }
            },
        }
    )
    session = BOIntegrationSession(config, tmp_path)

    group_config = session._config_for_group(1)
    available = session._available_candidates(set(), config=group_config)

    assert len(available) == 50
