import pytest
from pathlib import Path

from gui.tab_automated_titration import AutomatedTitrationTab
from core.titration import calculate_titration_plan


def test_air_assisted_stock_delivery_order_and_volumes():
    tab = AutomatedTitrationTab.__new__(AutomatedTitrationTab)
    recipe = []
    settings = {
        "air_port": 9,
        "stock_port": 5,
        "mix_port": 4,
        "waste_port": 2,
        "speed": 20,
        "syringe_capacity": 250.0,
        "stock_air_spacer": 100.0,
        "mix_line_air_push": 250.0,
        "mix_line_volume": 75.0,
    }

    tab._append_air_assisted_stock_delivery(
        recipe,
        stock_volume_ul=10.0,
        settings=settings,
        point="10 µM",
    )

    assert [item["type"] for item in recipe] == [
        "PUMP_VALVE",
        "PUMP_ASPIRATE",
        "PUMP_VALVE",
        "PUMP_ASPIRATE",
        "PUMP_VALVE",
        "PUMP_DISPENSE",
        "PUMP_VALVE",
        "PUMP_ASPIRATE",
        "PUMP_VALVE",
        "PUMP_DISPENSE",
        "PUMP_VALVE",
        "PUMP_ASPIRATE",
        "PUMP_VALVE",
        "PUMP_DISPENSE",
    ]
    assert recipe[0]["pump_action"]["params"]["port"] == 9
    assert recipe[1]["pump_action"]["params"]["volume"] == pytest.approx(100)
    assert recipe[2]["pump_action"]["params"]["port"] == 5
    assert recipe[3]["pump_action"]["params"]["volume"] == pytest.approx(10)
    assert recipe[4]["pump_action"]["params"]["port"] == 4
    assert recipe[5]["pump_action"]["params"]["volume"] == pytest.approx(110)

    assert recipe[6]["pump_action"]["params"]["port"] == 9
    assert recipe[7]["pump_action"]["params"]["volume"] == pytest.approx(250)
    assert recipe[8]["pump_action"]["params"]["port"] == 4
    assert recipe[9]["pump_action"]["params"]["volume"] == pytest.approx(250)

    assert recipe[10]["pump_action"]["params"]["port"] == 4
    assert recipe[11]["pump_action"]["params"]["volume"] == pytest.approx(75)
    assert recipe[12]["pump_action"]["params"]["port"] == 2
    assert recipe[13]["pump_action"]["params"]["volume"] == pytest.approx(75)


def test_port4_clear_includes_configured_bubble_volume():
    tab = AutomatedTitrationTab.__new__(AutomatedTitrationTab)
    recipe = []
    settings = {
        "air_port": 9,
        "mix_port": 4,
        "waste_port": 2,
        "speed": 20,
        "syringe_capacity": 250.0,
        "mix_line_air_push": 250.0,
        "mix_line_volume": 110.0,
        "mix_line_bubble_volume": 50.0,
    }

    tab._append_port4_air_flush_and_clear(
        recipe,
        settings=settings,
        point="test",
        label="Bubble clear",
    )

    assert recipe[-4]["pump_action"]["params"]["port"] == 4
    assert recipe[-3]["pump_action"]["params"]["volume"] == pytest.approx(160)
    assert recipe[-2]["pump_action"]["params"]["port"] == 2
    assert recipe[-1]["pump_action"]["params"]["volume"] == pytest.approx(160)


def test_stock_at_or_above_syringe_capacity_omits_air_spacer():
    tab = AutomatedTitrationTab.__new__(AutomatedTitrationTab)
    recipe = []
    settings = {
        "air_port": 9,
        "stock_port": 5,
        "mix_port": 4,
        "waste_port": 2,
        "speed": 20,
        "syringe_capacity": 250.0,
        "stock_air_spacer": 100.0,
        "mix_line_air_push": 250.0,
        "mix_line_volume": 75.0,
    }

    tab._append_air_assisted_stock_delivery(
        recipe,
        stock_volume_ul=300.0,
        settings=settings,
        point="large transfer",
    )

    # Two direct stock strokes occur before the one final air push.
    assert [item["pump_action"]["params"].get("port") for item in recipe[:8:2]] == [
        5,
        4,
        5,
        4,
    ]
    assert recipe[1]["pump_action"]["params"]["volume"] == pytest.approx(250)
    assert recipe[5]["pump_action"]["params"]["volume"] == pytest.approx(50)
    assert recipe[8]["pump_action"]["params"]["port"] == 9
    assert sum(
        1
        for item in recipe
        if item["pump_action"]["params"].get("port") == 9
    ) == 1


def test_spacer_is_limited_to_remaining_syringe_capacity():
    tab = AutomatedTitrationTab.__new__(AutomatedTitrationTab)
    recipe = []
    settings = {
        "air_port": 9,
        "stock_port": 5,
        "mix_port": 4,
        "waste_port": 2,
        "speed": 20,
        "syringe_capacity": 250.0,
        "stock_air_spacer": 100.0,
        "mix_line_air_push": 250.0,
        "mix_line_volume": 75.0,
    }

    tab._append_air_assisted_stock_delivery(
        recipe,
        stock_volume_ul=200.0,
        settings=settings,
        point="near capacity",
    )

    assert recipe[1]["pump_action"]["params"]["volume"] == pytest.approx(50)
    assert recipe[3]["pump_action"]["params"]["volume"] == pytest.approx(200)
    assert recipe[5]["pump_action"]["params"]["volume"] == pytest.approx(250)


def test_every_mix_cycle_is_followed_by_port4_air_flush_and_waste_clear():
    tab = AutomatedTitrationTab.__new__(AutomatedTitrationTab)
    tab._parameter_groups = []
    settings = {
        "air_port": 9,
        "stock_port": 5,
        "buffer_port": 6,
        "mix_port": 4,
        "flow_port": 1,
        "waste_port": 2,
        "speed": 20,
        "syringe_capacity": 250.0,
        "stock_air_spacer": 100.0,
        "mix_line_air_push": 250.0,
        "mix_line_volume": 75.0,
        "initial_buffer_volume": 1000.0,
        "aliquot_volume": 100.0,
        "mix_cycles": 2,
        "mix_volume": 200.0,
        "equilibration": 0.0,
        "replicates": 1,
    }
    plan = calculate_titration_plan(
        [10],
        stock_concentration_um=10_000,
        initial_buffer_volume_ul=settings["initial_buffer_volume"],
        aliquot_volume_ul=settings["aliquot_volume"],
    )

    recipe = tab._build_recipe(settings, plan)
    details = [item["details"] for item in recipe]

    for cycle in (1, 2):
        mix_dispense = details.index(f"Mix cycle {cycle}: dispense 200.00 µL")
        following = recipe[mix_dispense + 1:mix_dispense + 9]
        assert [item["type"] for item in following] == [
            "PUMP_VALVE",
            "PUMP_ASPIRATE",
            "PUMP_VALVE",
            "PUMP_DISPENSE",
            "PUMP_VALVE",
            "PUMP_ASPIRATE",
            "PUMP_VALVE",
            "PUMP_DISPENSE",
        ]
        assert following[0]["pump_action"]["params"]["port"] == 9
        assert following[1]["pump_action"]["params"]["volume"] == pytest.approx(250)
        assert following[2]["pump_action"]["params"]["port"] == 4
        assert following[4]["pump_action"]["params"]["port"] == 4
        assert following[5]["pump_action"]["params"]["volume"] == pytest.approx(75)
        assert following[6]["pump_action"]["params"]["port"] == 2


def test_recipe_ends_by_draining_calculated_remaining_mix_volume_to_waste():
    tab = AutomatedTitrationTab.__new__(AutomatedTitrationTab)
    tab._parameter_groups = []
    settings = {
        "air_port": 9,
        "stock_port": 5,
        "buffer_port": 6,
        "mix_port": 4,
        "flow_port": 1,
        "waste_port": 2,
        "speed": 20,
        "syringe_capacity": 250.0,
        "stock_air_spacer": 100.0,
        "mix_line_air_push": 250.0,
        "mix_line_volume": 75.0,
        "initial_buffer_volume": 1000.0,
        "aliquot_volume": 100.0,
        "mix_cycles": 1,
        "mix_volume": 200.0,
        "equilibration": 0.0,
        "replicates": 1,
    }
    plan = calculate_titration_plan(
        [10],
        stock_concentration_um=10_000,
        initial_buffer_volume_ul=settings["initial_buffer_volume"],
        aliquot_volume_ul=settings["aliquot_volume"],
    )

    recipe = tab._build_recipe(settings, plan)
    cleanup = [
        item for item in recipe
        if item.get("_point") == "Final cleanup"
    ]

    assert cleanup
    assert cleanup[0]["pump_action"]["params"]["port"] == 4
    assert cleanup[-2]["pump_action"]["params"]["port"] == 2
    assert cleanup[-1]["type"] == "PUMP_DISPENSE"
    aspirated = sum(
        item["pump_action"]["params"]["volume"]
        for item in cleanup
        if item["type"] == "PUMP_ASPIRATE"
    )
    dispensed = sum(
        item["pump_action"]["params"]["volume"]
        for item in cleanup
        if item["type"] == "PUMP_DISPENSE"
    )
    assert aspirated == pytest.approx(plan[-1].volume_remaining_ul + 250)
    assert dispensed == pytest.approx(plan[-1].volume_remaining_ul + 250)
    assert cleanup[-3]["pump_action"]["params"]["volume"] == pytest.approx(250)
    assert cleanup[-1]["pump_action"]["params"]["volume"] == pytest.approx(250)


def test_swv_replicates_cycle_across_channels():
    tab = AutomatedTitrationTab.__new__(AutomatedTitrationTab)
    tab._parameter_groups = [
        {"name": "Group 1", "channels": [1, 2]},
        {"name": "Group 2", "channels": [3]},
    ]
    settings = {
        "air_port": 9,
        "stock_port": 5,
        "buffer_port": 6,
        "mix_port": 4,
        "flow_port": 1,
        "waste_port": 2,
        "speed": 20,
        "syringe_capacity": 250.0,
        "stock_air_spacer": 100.0,
        "mix_line_air_push": 250.0,
        "mix_line_volume": 75.0,
        "initial_buffer_volume": 1000.0,
        "aliquot_volume": 100.0,
        "mix_cycles": 0,
        "mix_volume": 200.0,
        "equilibration": 0.0,
        "replicates": 3,
    }
    plan = calculate_titration_plan(
        [10],
        stock_concentration_um=10_000,
        initial_buffer_volume_ul=settings["initial_buffer_volume"],
        aliquot_volume_ul=settings["aliquot_volume"],
    )

    recipe = tab._build_recipe(settings, plan)
    swv_items = [
        item for item in recipe
        if item["type"] == "SWV"
        and item["_point"] == "Initial buffer"
    ]

    assert [item["_mux_channel"] for item in swv_items] == [
        1, 2, 3,
        1, 2, 3,
        1, 2, 3,
    ]
    assert [
        item["details"].rsplit("rep ", 1)[1]
        for item in swv_items
    ] == [
        "1/3", "1/3", "1/3",
        "2/3", "2/3", "2/3",
        "3/3", "3/3", "3/3",
    ]


def test_manual_swv_pass_follows_optimized_pass_each_replicate():
    tab = AutomatedTitrationTab.__new__(AutomatedTitrationTab)
    optimized = {
        "begin_potential": -0.7,
        "end_potential": -0.1,
        "step_potential": 0.002,
        "amplitude": 0.036,
        "frequency": 200.0,
        "conditioning_potential": -0.7,
        "conditioning_time": 0.0,
    }
    tab._parameter_groups = [
        {"name": "Group 1", "channels": [1, 2, 3], "params": optimized}
    ]
    tab._manual_channel_params = {
        channel: {**optimized, "frequency": 100.0 + channel}
        for channel in (1, 2, 3)
    }
    settings = {
        "air_port": 9, "stock_port": 5, "buffer_port": 6, "mix_port": 4,
        "flow_port": 1, "waste_port": 2, "speed": 20,
        "syringe_capacity": 250.0, "stock_air_spacer": 100.0,
        "mix_line_air_push": 250.0, "mix_line_volume": 75.0,
        "initial_buffer_volume": 1000.0, "aliquot_volume": 100.0,
        "mix_cycles": 0, "mix_volume": 200.0, "equilibration": 0.0,
        "replicates": 2,
    }
    plan = calculate_titration_plan(
        [10], stock_concentration_um=10_000,
        initial_buffer_volume_ul=1000, aliquot_volume_ul=100,
    )

    swv_items = [
        item for item in tab._build_recipe(settings, plan)
        if item["type"] == "SWV"
        and item["_point"] == "Initial buffer"
    ]

    assert [
        (item["_mux_channel"], item["_swv_source"])
        for item in swv_items
    ] == [
        (1, "optimized"), (2, "optimized"), (3, "optimized"),
        (1, "manual"), (2, "manual"), (3, "manual"),
        (1, "optimized"), (2, "optimized"), (3, "optimized"),
        (1, "manual"), (2, "manual"), (3, "manual"),
    ]
    assert swv_items[3]["_titration_group"]["params"]["frequency"] == 101.0


def test_first_stock_mixes_before_buffer_measurement_and_first_flow_load():
    tab = AutomatedTitrationTab.__new__(AutomatedTitrationTab)
    tab._parameter_groups = [
        {"name": "Group 1", "channels": [1, 2, 3]}
    ]
    tab._manual_channel_params = {}
    settings = {
        "air_port": 9, "stock_port": 5, "buffer_port": 6, "mix_port": 4,
        "flow_port": 1, "waste_port": 2, "speed": 20,
        "syringe_capacity": 250.0, "stock_air_spacer": 100.0,
        "mix_line_air_push": 250.0, "mix_line_volume": 110.0,
        "initial_buffer_volume": 1000.0, "aliquot_volume": 100.0,
        "mix_cycles": 0, "mix_volume": 200.0, "equilibration": 0.0,
        "replicates": 3,
    }
    plan = calculate_titration_plan(
        [10], stock_concentration_um=10_000,
        initial_buffer_volume_ul=1000, aliquot_volume_ul=100,
    )

    recipe = tab._build_recipe(settings, plan)
    initial_swv_indices = [
        index for index, item in enumerate(recipe)
        if item["type"] == "SWV"
        and item["_point"] == "Initial buffer"
    ]
    first_stock_index = next(
        index for index, item in enumerate(recipe)
        if "Stock delivery" in item.get("details", "")
    )
    first_titrated_flow_load = next(
        index for index, item in enumerate(recipe)
        if item["type"] == "PUMP_DISPENSE"
        and "Mixing tube" in item.get("details", "")
        and "flow cell" in item.get("details", "")
    )

    assert len(initial_swv_indices) == 9
    assert first_stock_index < min(initial_swv_indices)
    assert max(initial_swv_indices) < first_titrated_flow_load


def test_multi_point_recipe_pipelines_next_mix_before_current_measurement():
    tab = AutomatedTitrationTab.__new__(AutomatedTitrationTab)
    tab._parameter_groups = [{"name": "Group 1", "channels": [1]}]
    tab._manual_channel_params = {}
    settings = {
        "air_port": 9, "stock_port": 5, "buffer_port": 6, "mix_port": 4,
        "flow_port": 1, "waste_port": 2, "speed": 20,
        "syringe_capacity": 250.0, "stock_air_spacer": 100.0,
        "mix_line_air_push": 250.0, "mix_line_volume": 110.0,
        "mix_line_bubble_volume": 50.0,
        "initial_buffer_volume": 1000.0, "aliquot_volume": 100.0,
        "mix_cycles": 1, "mix_volume": 200.0, "equilibration": 0.0,
        "replicates": 1,
    }
    plan = calculate_titration_plan(
        [10, 25, 50], stock_concentration_um=10_000,
        initial_buffer_volume_ul=1000, aliquot_volume_ul=100,
    )
    recipe = tab._build_recipe(settings, plan)

    def first_index(predicate):
        return next(i for i, item in enumerate(recipe) if predicate(item))

    def stock_index(point):
        return first_index(
            lambda item: item.get("_point") == point
            and "Stock delivery" in item.get("details", "")
        )

    def measure_index(point):
        return first_index(
            lambda item: item.get("_point") == point and item["type"] == "SWV"
        )

    def load_index(point):
        return first_index(
            lambda item: item.get("_point") == point
            and item["type"] == "PUMP_DISPENSE"
            and "Mixing tube" in item.get("details", "")
            and "flow cell" in item.get("details", "")
        )

    assert (
        stock_index("10 µM")
        < measure_index("Initial buffer")
        < load_index("10 µM")
        < stock_index("25 µM")
        < measure_index("10 µM")
        < load_index("25 µM")
        < stock_index("50 µM")
        < measure_index("25 µM")
        < load_index("50 µM")
        < measure_index("50 µM")
    )


def test_optional_plain_buffer_measurements_are_inserted_between_concentrations():
    tab = AutomatedTitrationTab.__new__(AutomatedTitrationTab)
    tab._parameter_groups = [{"name": "Group 1", "channels": [1]}]
    tab._manual_channel_params = {}
    settings = {
        "air_port": 9, "stock_port": 5, "buffer_port": 6, "mix_port": 4,
        "flow_port": 1, "waste_port": 2, "speed": 20,
        "initial_buffer_speed": 12,
        "syringe_capacity": 250.0, "stock_air_spacer": 100.0,
        "mix_line_air_push": 250.0, "mix_line_volume": 110.0,
        "mix_line_bubble_volume": 50.0,
        "initial_buffer_volume": 1000.0, "aliquot_volume": 100.0,
        "plain_buffer_volume": 175.0,
        "mix_cycles": 1, "mix_volume": 200.0, "equilibration": 0.0,
        "replicates": 1, "measure_buffer_between": True,
    }
    plan = calculate_titration_plan(
        [40, 80, 160], stock_concentration_um=10_000,
        initial_buffer_volume_ul=1000, aliquot_volume_ul=100,
    )

    recipe = tab._build_recipe(settings, plan)
    measured_points = [
        item["_point"] for item in recipe if item["type"] == "SWV"
    ]

    assert measured_points == [
        "Initial buffer",
        "40 µM",
        "0 µM buffer between 40 µM and 80 µM",
        "80 µM",
        "0 µM buffer between 80 µM and 160 µM",
        "160 µM",
    ]
    buffer_loads = [
        item for item in recipe
        if item["type"] == "PUMP_DISPENSE"
        and "Plain buffer" in item.get("details", "")
    ]
    assert len(buffer_loads) == 2
    assert all(
        item["pump_action"]["params"]["volume"] == pytest.approx(175)
        for item in buffer_loads
    )
    titrated_loads = [
        item for item in recipe
        if item["type"] == "PUMP_DISPENSE"
        and "Mixing tube → flow cell" in item.get("details", "")
    ]
    assert all(
        item["pump_action"]["params"]["volume"] == pytest.approx(100)
        for item in titrated_loads
    )


def test_previous_flow_cell_aliquot_is_not_moved_to_waste():
    tab = AutomatedTitrationTab.__new__(AutomatedTitrationTab)
    tab._parameter_groups = []
    settings = {
        "air_port": 9, "stock_port": 5, "buffer_port": 6, "mix_port": 4,
        "flow_port": 1, "waste_port": 2, "speed": 20,
        "syringe_capacity": 250.0, "stock_air_spacer": 100.0,
        "mix_line_air_push": 250.0, "mix_line_volume": 110.0,
        "initial_buffer_volume": 1000.0, "aliquot_volume": 100.0,
        "mix_cycles": 0, "mix_volume": 200.0, "equilibration": 0.0,
        "replicates": 1,
    }
    plan = calculate_titration_plan(
        [10, 25], stock_concentration_um=10_000,
        initial_buffer_volume_ul=1000, aliquot_volume_ul=100,
    )

    recipe = tab._build_recipe(settings, plan)

    assert not any(
        "Previous flow-cell aliquot" in item.get("details", "")
        for item in recipe
    )
    flow_cell_loads = [
        item for item in recipe
        if "Mixing tube" in item.get("details", "")
        and "flow cell" in item.get("details", "")
        and item["type"] == "PUMP_DISPENSE"
    ]
    assert len(flow_cell_loads) == 2


def test_new_manual_channels_use_requested_swv_defaults():
    tab = AutomatedTitrationTab.__new__(AutomatedTitrationTab)
    optimized = {
        "begin_potential": -0.65,
        "end_potential": -0.15,
        "step_potential": 0.01,
        "amplitude": 0.1,
        "frequency": 50.0,
        "conditioning_potential": -0.7,
        "conditioning_time": 2.0,
    }
    manual = tab._default_manual_params(optimized)

    assert manual["amplitude"] == pytest.approx(0.036)
    assert manual["step_potential"] == pytest.approx(0.002)
    assert manual["frequency"] == pytest.approx(200.0)
    assert manual["begin_potential"] == pytest.approx(-0.65)
    assert manual["conditioning_time"] == pytest.approx(2.0)


def test_initial_buffer_transfer_can_be_bypassed():
    tab = AutomatedTitrationTab.__new__(AutomatedTitrationTab)
    tab._parameter_groups = []
    settings = {
        "air_port": 9, "stock_port": 5, "buffer_port": 6, "mix_port": 4,
        "flow_port": 1, "waste_port": 2, "speed": 20,
        "syringe_capacity": 250.0, "stock_air_spacer": 100.0,
        "mix_line_air_push": 250.0, "mix_line_volume": 75.0,
        "initial_buffer_volume": 1000.0, "aliquot_volume": 100.0,
        "mix_cycles": 0, "mix_volume": 200.0, "equilibration": 0.0,
        "replicates": 1, "skip_initial_buffer": True,
    }
    plan = calculate_titration_plan(
        [10], stock_concentration_um=10_000,
        initial_buffer_volume_ul=1000, aliquot_volume_ul=100,
    )

    recipe = tab._build_recipe(settings, plan)

    assert recipe[0]["type"] == "PUMP_INIT"
    assert not any(item.get("_point") == "Setup" for item in recipe)


def test_locked_post_bo_titration_materializes_queues_and_starts():
    class Registry:
        def save_script(self, *_args, **_kwargs):
            return Path("manual-test.ms"), "manual-test"

        def hash_key_for(self, _path):
            return "test-hash"

    class Session:
        registry = Registry()
        measurement_queue = [
            {"type": "BO_AUTO_LOOP", "status": "completed", "details": "Finished BO"}
        ]

    tab = AutomatedTitrationTab.__new__(AutomatedTitrationTab)
    tab._session = Session()
    tab._manual_channel_params = {}
    tab._bo_locked_settings = {
        "air_port": 9, "stock_port": 5, "buffer_port": 6, "mix_port": 4,
        "flow_port": 1, "waste_port": 2, "speed": 20,
        "initial_buffer_speed": 7, "final_cleanup_speed": 11,
        "syringe_capacity": 250.0, "stock_air_spacer": 100.0,
        "mix_line_air_push": 250.0, "mix_line_volume": 75.0,
        "initial_buffer_volume": 1000.0, "aliquot_volume": 100.0,
        "mix_cycles": 0, "mix_volume": 200.0, "equilibration": 0.0,
        "replicates": 1, "skip_initial_buffer": True,
    }
    tab._bo_locked_plan = calculate_titration_plan(
        [10], stock_concentration_um=10_000,
        initial_buffer_volume_ul=1000, aliquot_volume_ul=100,
    )
    queued = []
    run_calls = []
    tab._send_queue_item = queued.append
    tab._run_queue = lambda start_index: run_calls.append(start_index)
    tab._status_var = type("Status", (), {"set": lambda self, value: None})()
    optimized = {
        "begin_potential": -0.7, "end_potential": -0.1,
        "step_potential": 0.002, "amplitude": 0.036, "frequency": 200.0,
        "conditioning_potential": -0.7, "conditioning_time": 0.0,
    }

    tab.run_locked_after_bo([
        {"name": "Group 1", "channels": [1], "params": optimized}
    ])

    assert queued
    assert any(item["type"] == "SWV" for item in queued)
    assert run_calls == [1]
