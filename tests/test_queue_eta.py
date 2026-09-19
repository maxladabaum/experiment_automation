from core.queue_eta import estimate_queue_eta, estimate_running_queue_eta
import json
from pathlib import Path
from types import SimpleNamespace
import pytest
from core import queue_eta


@pytest.fixture(autouse=True)
def reset_pump_samples():
    queue_eta._PUMP_SAMPLES.clear()
    queue_eta._PUMP_CONTEXT = None
    yield
    queue_eta._PUMP_SAMPLES.clear()
    queue_eta._PUMP_CONTEXT = None


def pump(volume, speed=15):
    return {'type': 'PUMP_ASPIRATE', 'pump_action': {
        'name': 'ASPIRATE', 'params': {'volume': volume, 'speed': speed}}}


def test_pump_eta_learns_speed_and_volume_with_overhead():
    queue_eta.record_pump_seconds(pump(100), 12)
    queue_eta.record_pump_seconds(pump(200), 22)
    queue_eta.record_pump_seconds(pump(200, 20), 42)
    assert queue_eta.estimate_item_seconds(pump(150)) == pytest.approx(17)
    assert queue_eta.estimate_item_seconds(pump(200, 20)) == 42
    assert queue_eta.estimate_item_seconds(pump(200)) == 22


def test_pump_calibration_change_discards_old_timings():
    controller = SimpleNamespace(syringe_ul=250, use_sim=False)
    queue_eta.set_pump_timing_context(controller)
    queue_eta.record_pump_seconds(pump(250), 30)
    assert queue_eta.estimate_item_seconds(pump(250)) == 30
    controller.syringe_ul = 1000
    queue_eta.set_pump_timing_context(controller)
    assert not queue_eta._PUMP_SAMPLES


def bo_item(tmp_path):
    config = {
        'channels': [1, 2],
        'channel_groups': [
            {'id': 1, 'channels': [1], 'optimization_direction': 'maximize_and_minimize'},
            {'id': 2, 'channels': [2], 'optimization_direction': 'maximize'},
        ],
        'measurements_per_channel': 3,
    }
    path = tmp_path / 'bo.json'
    path.write_text(json.dumps(config))
    return {'type': 'BO_AUTO_LOOP', 'bo_block': {
        'bo_config_path': str(path), 'objective': 'paired_response',
        'target_iterations': 5, 'batch_size': 2, 'warmup_iterations': 2,
        'warmup_batch_size': 1, 'target_equilibration_seconds': 60,
        'buffer_equilibration_seconds': 60,
    }}


def test_bo_eta_counts_warmup_batches_and_exchange_repetitions(tmp_path):
    item = bo_item(tmp_path)
    initial = queue_eta.estimate_item_seconds(item)
    assert initial is not None and initial > 480
    exchange = tmp_path / 'exchange.json'
    exchange.write_text(json.dumps({'items': [{'type': 'PAUSE', 'pause_seconds': 17}]}))
    item['bo_block']['target_exchange_block_path'] = str(exchange)
    # Two warmup batches + ceil(3/2) regular batches = four exchanges.
    assert queue_eta.estimate_item_seconds(item) - initial == pytest.approx(4*17)
    item['bo_block']['warmup_single_batch'] = True
    assert queue_eta.estimate_item_seconds(item) == pytest.approx(initial + 4*17 - 137)


def test_bo_eta_counts_directions_channels_and_repeats(tmp_path):
    item = bo_item(tmp_path)
    item['bo_block'].update(target_equilibration_seconds=0, buffer_equilibration_seconds=0)
    original = queue_eta.estimate_item_seconds(item)
    path = Path(item['bo_block']['bo_config_path'])
    cfg = json.loads(path.read_text())
    cfg['channel_groups'][0]['optimization_direction'] = 'maximize'
    path.write_text(json.dumps(cfg))
    assert queue_eta.estimate_item_seconds(item) == pytest.approx(original*2/3)


def test_bo_missing_exchange_is_unknown_and_never_created(tmp_path):
    item = bo_item(tmp_path)
    missing = tmp_path / 'missing.json'
    item['bo_block']['target_exchange_block_path'] = str(missing)
    assert queue_eta.estimate_item_seconds(item) is None
    assert not missing.exists()


def test_live_bo_eta_includes_loop_not_just_nested_step():
    from gui.tab_queue import QueueTab
    tab = QueueTab.__new__(QueueTab)
    tab._session = SimpleNamespace(
        measurement_queue=[{'type': 'BO_AUTO_LOOP'}, {'type': 'PAUSE', 'pause_seconds': 10}],
        step_delay=0,
        get_queue_status=lambda: {
            'active_queue_index': 0, 'next_queue_index': 1,
            'active_step_estimated_seconds': 2, 'active_step_type': 'SWV',
            'bo_eta_seconds': 100,
        },
    )
    tab._elapsed_since = lambda timestamp: 0
    lines = tab._build_live_eta_lines()
    assert 'Total remaining: 1m 50s' in lines


def test_failed_pump_action_does_not_train_eta(monkeypatch):
    from gui.tab_queue import QueueTab
    ticks = iter([0.0, 10.0, 15.0])
    monkeypatch.setattr('gui.tab_queue.time.monotonic', lambda: next(ticks))
    tab = QueueTab.__new__(QueueTab)
    tab._pump_ctrl = SimpleNamespace(syringe_ul=250, use_sim=False)
    tab._session = SimpleNamespace(is_running=True)
    tab._exec_pump_action = lambda item: False
    assert tab._exec_pump(pump(250)) is False
    assert not queue_eta._PUMP_SAMPLES
    tab._exec_pump_action = lambda item: True
    assert tab._exec_pump(pump(250)) is True
    assert len(queue_eta._PUMP_SAMPLES[('PUMP_ASPIRATE', 15)]) == 1


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
