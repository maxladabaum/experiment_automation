from types import SimpleNamespace
from unittest.mock import Mock

from gui.tab_automated_titration import AutomatedTitrationTab
from gui.tab_bayesian_optimization import BayesianOptimizationTab


class Value:
    def __init__(self, value=None):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


def make_bo_handoff(start_result):
    tab = BayesianOptimizationTab.__new__(BayesianOptimizationTab)
    tab._run_auto_titration_var = Value(True)
    tab._post_bo_titration_started = False
    tab._auto_status_var = Value()
    tab._session = SimpleNamespace(session_manager=SimpleNamespace(log=Mock()))
    tab._on_bo_finished = Mock(return_value=start_result)
    return tab


def test_bo_started_state_is_set_only_after_actual_queue_start():
    tab = make_bo_handoff(False)
    summary = {"run_id": 4, "superseded": False}

    assert tab._start_post_bo_titration(summary) is False
    assert tab._post_bo_titration_started is False
    assert "rejected" in tab._auto_status_var.get()

    tab._on_bo_finished.return_value = True
    assert tab._start_post_bo_titration(summary) is True
    assert tab._post_bo_titration_started is True
    assert "actually started" in tab._auto_status_var.get()


def test_duplicate_bo_completion_callback_does_not_execute_twice():
    tab = make_bo_handoff(True)
    summary = {"run_id": 8, "superseded": False}

    assert tab._start_post_bo_titration(summary) is True
    assert tab._start_post_bo_titration(summary) is False
    tab._on_bo_finished.assert_called_once_with(summary)


def test_superseded_bo_completion_cancels_pending_automatic_start():
    tab = make_bo_handoff(True)

    assert tab._start_post_bo_titration(
        {"run_id": 2, "superseded": True}
    ) is False
    assert tab._post_bo_titration_started is False
    tab._on_bo_finished.assert_not_called()
    assert "canceled" in tab._auto_status_var.get()


def test_rejected_handoff_retries_without_duplicate_recipe_insertion():
    tab = AutomatedTitrationTab.__new__(AutomatedTitrationTab)
    tab._session = SimpleNamespace(
        measurement_queue=[{"type": "BO_AUTO_LOOP", "status": "completed"}]
    )
    tab._bo_locked_settings = {"locked": True}
    tab._bo_locked_plan = [{"point": 1}]
    tab._bo_locked_manual_channel_params = {}
    tab._manual_channel_params = {}
    tab._bo_queue_handoff = None
    tab._status_var = Value()
    tab._build_recipe = Mock(
        return_value=[
            {"type": "PUMP_INIT", "status": "pending", "details": "one"},
            {"type": "PUMP_VALVE", "status": "pending", "details": "two"},
        ]
    )
    queued = []
    starts = []
    start_results = iter([False, True])
    tab._send_queue_item = queued.append

    def run_queue(start_index, **kwargs):
        starts.append((start_index, kwargs))
        return next(start_results)

    tab._run_queue = run_queue
    groups = [{"name": "Group 1", "channels": [1], "params": {}}]

    assert tab.run_locked_after_bo(
        groups, handoff_id=12, expected_run_id=12
    ) is False
    assert tab.run_locked_after_bo(
        groups, handoff_id=12, expected_run_id=12
    ) is True

    assert [item["details"] for item in queued] == ["one", "two"]
    assert starts == [
        (1, {"expected_run_id": 12, "automatic": True}),
        (1, {"expected_run_id": 12, "automatic": True}),
    ]


def test_nonpaired_completion_passes_specific_run_to_continuation():
    tab = BayesianOptimizationTab.__new__(BayesianOptimizationTab)
    summary = {
        "run_id": 6,
        "superseded": False,
        "failed": 0,
        "stopped": 0,
    }
    tab._paired_queue_running = False
    tab._bo_session = Mock()
    tab._refresh_record_files = Mock()
    tab._auto_running = True
    tab._auto_status_var = Value()
    tab._run_analysis_for_pending = Mock(return_value={"iteration": 2})
    tab._auto_submit_next = Mock()

    tab.on_queue_complete(summary)

    tab._bo_session.record_queue_completion.assert_called_once_with(summary)
    tab._auto_submit_next.assert_called_once_with(summary)


def test_paired_completion_passes_specific_run_to_titration_handoff():
    tab = BayesianOptimizationTab.__new__(BayesianOptimizationTab)
    summary = {
        "run_id": 9,
        "superseded": False,
        "failed": 0,
        "stopped": 0,
    }
    loaded = SimpleNamespace(config={}, record_dir="record", observations=[{}])
    tab._paired_queue_running = True
    tab._load_latest_paired_queue_session = Mock(return_value=loaded)
    tab._sync_suggestion_from_session = Mock()
    tab._set_rescore_vars_from_config = Mock()
    tab._record_dir_var = Value()
    tab._flush_deferred_results_render = Mock()
    tab._auto_status_var = Value()
    tab._tabs = SimpleNamespace(select=Mock())
    tab._start_post_bo_titration = Mock()

    tab.on_queue_complete(summary)

    tab._start_post_bo_titration.assert_called_once_with(summary)
