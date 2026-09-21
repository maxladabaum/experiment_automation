from core.bo_session import load_bo_config, normalize_bo_config


def test_missing_bo_current_range_uses_safe_autorange_default():
    config = normalize_bo_config({})

    assert config["method_options"]["ba_range"] == {
        "mode": "auto",
        "fixed": "100 nA",
        "auto_min": "100 nA",
        "auto_max": "100 uA",
    }


def test_partial_bo_current_range_only_fills_missing_defaults():
    config = normalize_bo_config(
        {
            "method_options": {
                "ba_range": {"mode": "fixed", "fixed": "25 uA"}
            }
        }
    )

    assert config["method_options"]["ba_range"] == {
        "mode": "fixed",
        "fixed": "25 uA",
        "auto_min": "100 nA",
        "auto_max": "100 uA",
    }


def test_explicit_saved_bo_current_range_is_preserved():
    saved_range = {
        "mode": "auto",
        "fixed": "1 uA",
        "auto_min": "1 uA",
        "auto_max": "25 uA",
    }

    config = normalize_bo_config(
        {"method_options": {"ba_range": dict(saved_range)}}
    )

    assert config["method_options"]["ba_range"] == saved_range


def test_bundled_bo_config_uses_safe_autorange_default():
    config = load_bo_config("optimizer/bo_configs/default_swv_bo.json")

    assert config["method_options"]["ba_range"]["mode"] == "auto"
    assert config["method_options"]["ba_range"]["auto_min"] == "100 nA"
    assert config["method_options"]["ba_range"]["auto_max"] == "100 uA"
