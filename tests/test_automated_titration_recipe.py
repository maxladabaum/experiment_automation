import pytest

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
    assert aspirated == pytest.approx(plan[-1].volume_remaining_ul)
    assert dispensed == pytest.approx(plan[-1].volume_remaining_ul)
