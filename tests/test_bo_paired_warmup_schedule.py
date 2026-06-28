from gui.tab_queue import QueueTab


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
