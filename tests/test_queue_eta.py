from core.queue_eta import estimate_queue_eta, estimate_running_queue_eta


def test_estimate_queue_eta_applies_step_delay_between_items():
    queue = [
        {"type": "PAUSE", "pause_seconds": 10},
        {"type": "PAUSE", "pause_seconds": 5},
    ]

    eta = estimate_queue_eta(queue, step_delay_seconds=2.0)

    assert eta.total_seconds == 17.0
    assert eta.known_seconds == 17.0
    assert eta.unknown_item_count == 0


def test_running_eta_splits_current_step_and_remaining_queue():
    queue = [
        {"type": "PAUSE", "pause_seconds": 20},
        {"type": "PAUSE", "pause_seconds": 10},
    ]

    eta = estimate_running_queue_eta(
        queue,
        next_index=1,
        current_step_elapsed_seconds=5.0,
        current_step_estimated_seconds=20.0,
        step_delay_seconds=3.0,
        include_next_step_delay=True,
    )

    assert eta.current_step_remaining_seconds == 15.0
    assert eta.remaining_after_current_seconds == 13.0
    assert eta.total_remaining_seconds == 28.0


def test_running_eta_marks_manual_alert_as_unpredictable():
    queue = [
        {"type": "ALERT", "alert_message": "Continue?"},
        {"type": "PAUSE", "pause_seconds": 10},
    ]

    eta = estimate_running_queue_eta(
        queue,
        next_index=1,
        current_step_elapsed_seconds=0.0,
        current_step_estimated_seconds=None,
        step_delay_seconds=2.0,
        include_next_step_delay=True,
    )

    assert eta.current_step_predictable is False
    assert eta.current_step_remaining_seconds is None
    assert eta.total_remaining_seconds is None
    assert eta.remaining_after_current_seconds == 12.0
    assert eta.unknown_item_count == 1


def test_running_eta_does_not_double_count_current_step_delay():
    queue = [
        {"type": "PAUSE", "pause_seconds": 10},
        {"type": "PAUSE", "pause_seconds": 6},
    ]

    eta = estimate_running_queue_eta(
        queue,
        next_index=1,
        current_step_elapsed_seconds=2.0,
        current_step_estimated_seconds=4.0,
        step_delay_seconds=4.0,
        include_next_step_delay=False,
    )

    assert eta.current_step_remaining_seconds == 2.0
    assert eta.remaining_after_current_seconds == 6.0
    assert eta.total_remaining_seconds == 8.0
