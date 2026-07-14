from core.swv_method import build_swv_methodscript


def test_build_swv_methodscript_rounds_frequency_to_integer_hz():
    script = build_swv_methodscript(
        {
            "begin_potential": -0.6,
            "end_potential": 0.0,
            "step_potential": 0.013,
            "amplitude": 0.057,
            "frequency": 384.615384615,
            "conditioning_potential": -0.6,
            "conditioning_time": 0.0,
        },
        {
            "bandwidth": "4k",
            "ba_range": {
                "mode": "auto",
                "fixed": "100 nA",
                "auto_min": "100 nA",
                "auto_max": "25 uA",
            },
        },
    )

    assert "meas_loop_swv p c f r -600m 0 13m 57m 385" in script
    assert "384.615" not in script
