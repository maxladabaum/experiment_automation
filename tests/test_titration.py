import pytest

from core.titration import (
    calculate_titration_plan,
    parse_concentrations,
    split_transfer,
)


def test_exact_serial_titration_accounts_for_added_and_removed_volume():
    points = calculate_titration_plan(
        [10, 20],
        stock_concentration_um=10_000,
        initial_buffer_volume_ul=10_000,
        aliquot_volume_ul=500,
    )

    assert points[0].stock_added_ul == pytest.approx(10.01001001001)
    assert points[0].volume_remaining_ul == pytest.approx(9510.01001001001)

    expected_second_addition = (
        points[0].volume_remaining_ul * (20 - 10) / (10_000 - 20)
    )
    assert points[1].stock_added_ul == pytest.approx(expected_second_addition)
    assert points[1].volume_before_stock_ul == pytest.approx(
        points[0].volume_remaining_ul
    )


def test_zero_first_point_requires_no_stock():
    point = calculate_titration_plan(
        [0],
        stock_concentration_um=10_000,
        initial_buffer_volume_ul=10_000,
        aliquot_volume_ul=500,
    )[0]
    assert point.stock_added_ul == 0
    assert point.volume_remaining_ul == 9500


def test_descending_targets_are_rejected():
    with pytest.raises(ValueError, match="nondecreasing"):
        calculate_titration_plan(
            [10, 5],
            stock_concentration_um=10_000,
            initial_buffer_volume_ul=10_000,
            aliquot_volume_ul=500,
        )


def test_target_must_be_below_stock():
    with pytest.raises(ValueError, match="below the stock concentration"):
        calculate_titration_plan(
            [10_000],
            stock_concentration_um=10_000,
            initial_buffer_volume_ul=10_000,
            aliquot_volume_ul=500,
        )


def test_parse_concentrations_accepts_common_separators():
    assert parse_concentrations("0, 1  10;100\n500") == [0, 1, 10, 100, 500]


def test_transfer_is_split_to_syringe_capacity():
    assert split_transfer(625, 250) == [250, 250, 125]
