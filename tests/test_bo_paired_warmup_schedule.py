from gui.tab_queue import QueueTab


class _Suggestion:
    def __init__(self, iteration, group_id):
        self.iteration = iteration
        self.group_id = group_id


def test_paired_warmup_uses_common_per_group_warmup_prefix():
    config = {
        "n_initial_points": 0,
        "channel_groups": [
            {"name": "Group 1", "channels": [1], "n_initial_points": 2},
            {"name": "Group 2", "channels": [2], "n_initial_points": 2},
            {"name": "Group 3", "channels": [3], "n_initial_points": 2},
        ],
    }

    warmup = QueueTab._paired_bo_warmup_parameter_sets(config)
    count, cycle_span = QueueTab._paired_bo_batch_span(
        completed_observations=0,
        target_observations=4,
        batch_size=1,
        warmup_observations=warmup,
    )

    assert warmup == 2
    assert count == 2
    assert cycle_span == 2


def test_paired_warmup_only_consolidates_shared_group_warmup_prefix():
    config = {
        "n_initial_points": 8,
        "channel_groups": [
            {"name": "Group 1", "channels": [1], "n_initial_points": 2},
            {"name": "Group 2", "channels": [2], "n_initial_points": 4},
            {"name": "Group 3", "channels": [3], "n_initial_points": 3},
        ],
    }

    assert QueueTab._paired_bo_warmup_parameter_sets(config) == 2


def test_paired_execution_order_runs_all_groups_for_each_iteration():
    suggestions = [
        _Suggestion(iteration=1, group_id=1),
        _Suggestion(iteration=2, group_id=1),
        _Suggestion(iteration=1, group_id=2),
        _Suggestion(iteration=2, group_id=2),
        _Suggestion(iteration=1, group_id=3),
        _Suggestion(iteration=2, group_id=3),
    ]

    ordered = QueueTab._paired_bo_execution_order(suggestions)

    assert [(item.group_id, item.iteration) for item in ordered] == [
        (1, 1),
        (2, 1),
        (3, 1),
        (1, 2),
        (2, 2),
        (3, 2),
    ]


def test_paired_warmup_cycles_are_consolidated_into_one_batch():
    count, cycle_span = QueueTab._paired_bo_batch_span(
        completed_observations=0,
        target_observations=25,
        batch_size=5,
        warmup_observations=10,
    )

    assert count == 10
    assert cycle_span == 2


def test_paired_gp_batches_return_to_configured_batch_size_after_warmup():
    count, cycle_span = QueueTab._paired_bo_batch_span(
        completed_observations=10,
        target_observations=25,
        batch_size=5,
        warmup_observations=10,
    )

    assert count == 5
    assert cycle_span == 1


def test_paired_warmup_is_capped_by_total_requested_observations():
    count, cycle_span = QueueTab._paired_bo_batch_span(
        completed_observations=0,
        target_observations=5,
        batch_size=5,
        warmup_observations=10,
    )

    assert count == 5
    assert cycle_span == 1
