from core.bo_session import load_bo_config, generate_candidates


def test_swv_candidate_pool_does_not_strongly_bias_low_step_values():
    config = load_bo_config("optimizer/bo_configs/default_swv_bo.json")
    # Exercise both halves of a fixed search interval, independently of the
    # editable experiment defaults shipped with the app.
    config["parameters"]["step_potential"].update(
        mode="active", space="continuous", min=0.001, max=0.020, step=0.001,
        scale="linear",
    )
    candidates = generate_candidates(config)

    low_step = sum(1 for candidate in candidates if candidate["step_potential"] < 0.0105)
    high_step = len(candidates) - low_step

    assert low_step > 0
    assert high_step > 0
    ratio = low_step / max(high_step, 1)
    assert ratio < 1.35
