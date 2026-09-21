import threading
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from gui.app import ElectrochemGUI
from gui.tab_queue import QueueTab


@pytest.mark.parametrize('start_method', ['run_queue', 'run_from_selected', 'run_from_index', 'run_now'])
def test_stop_cannot_restart_a_still_live_worker(monkeypatch, start_method):
    release = threading.Event()
    entered = threading.Event()
    tab = QueueTab.__new__(QueueTab)
    tab._session = SimpleNamespace(is_running=True, measurement_queue=[{'type': 'PUMP_ASPIRATE'}],
                                   stop_current_runner=Mock(), update_queue_status=Mock())
    tab.log = Mock()
    tab.set_status = Mock()
    tab._reset_reorder = Mock()
    warning = Mock()
    monkeypatch.setattr('gui.tab_queue.messagebox.showwarning', warning)

    def hardware_call():
        entered.set()
        release.wait(timeout=5)

    worker = threading.Thread(target=hardware_call)
    tab._queue_thread = worker
    worker.start()
    try:
        assert entered.wait(timeout=2)
        tab.stop_queue()
        assert not tab._session.is_running
        assert worker.is_alive()
        if start_method == 'run_now':
            app = ElectrochemGUI.__new__(ElectrochemGUI)
            app._session = tab._session
            app._queue_tab = tab
            app._run_now('CV', 'e\n')
        elif start_method == 'run_from_index':
            tab.run_from_index(0)
        else:
            getattr(tab, start_method)()
        warning.assert_called_once()
        assert not tab._session.is_running
        assert tab._queue_thread is worker
    finally:
        release.set()
        worker.join(timeout=2)
    assert not tab.worker_is_active()
    assert not tab._queue_start_is_blocked()


@pytest.mark.parametrize('operation', [
    'paste_after_selected', 'duplicate_selected', 'delete_selected',
    'clear_queue', 'load_queue', 'confirm_reorder',
])
def test_stopping_worker_keeps_queue_protected_from_edits(monkeypatch, operation):
    original = [{'type': 'PUMP_ASPIRATE', 'status': 'running'}]
    tab = QueueTab.__new__(QueueTab)
    tab._session = SimpleNamespace(is_running=False, measurement_queue=original.copy())
    tab._queue_thread = Mock(is_alive=Mock(return_value=True))
    tab._reset_reorder = Mock()
    warning = Mock()
    monkeypatch.setattr('gui.tab_queue.messagebox.showwarning', warning)
    getattr(tab, operation)()
    warning.assert_called_once()
    assert tab._session.measurement_queue == original


def test_completion_callback_waits_for_completed_worker_exit():
    callback_ran = threading.Event()
    worker_was_alive = []
    tab = QueueTab.__new__(QueueTab)
    tab._session = SimpleNamespace(measurement_queue=[])
    tab._root = SimpleNamespace(after=lambda _ms, callback, *args: callback(*args))
    tab.log = Mock()

    def on_complete(_summary):
        worker_was_alive.append(tab.worker_is_active())
        callback_ran.set()

    tab._completion_callbacks = [on_complete]
    worker = threading.Thread(target=tab._notify_completion_callbacks, args=(0,))
    tab._queue_thread = worker

    worker.start()
    worker.join(timeout=2)

    assert callback_ran.wait(timeout=2)
    assert worker_was_alive == [False]


def test_completion_callback_waits_through_delayed_worker_exit():
    release_worker = threading.Event()
    notification_prepared = threading.Event()
    callback_ran = threading.Event()
    tab = QueueTab.__new__(QueueTab)
    tab._session = SimpleNamespace(measurement_queue=[])
    tab._root = SimpleNamespace(after=lambda _ms, callback, *args: callback(*args))
    tab._completion_callbacks = [lambda _summary: callback_ran.set()]
    tab._queue_run_id = 7
    tab.log = Mock()

    def worker_target():
        tab._notify_completion_callbacks(0, run_id=7)
        notification_prepared.set()
        release_worker.wait(timeout=2)

    worker = threading.Thread(target=worker_target)
    tab._queue_thread = worker
    worker.start()

    assert notification_prepared.wait(timeout=2)
    assert not callback_ran.wait(timeout=0.05)
    release_worker.set()
    worker.join(timeout=2)
    assert callback_ran.wait(timeout=2)


def test_stop_invalidates_completion_waiting_for_worker_exit():
    release_worker = threading.Event()
    notification_prepared = threading.Event()
    callback_ran = threading.Event()
    callback_summary = []
    tab = QueueTab.__new__(QueueTab)
    tab._session = SimpleNamespace(is_running=False, measurement_queue=[])
    tab._root = SimpleNamespace(after=lambda _ms, callback, *args: callback(*args))

    def on_complete(summary):
        callback_summary.append(summary)
        callback_ran.set()

    tab._completion_callbacks = [on_complete]
    tab._queue_run_id = 3
    tab.log = Mock()

    def worker_target():
        tab._notify_completion_callbacks(0, run_id=3)
        notification_prepared.set()
        release_worker.wait(timeout=2)

    worker = threading.Thread(target=worker_target)
    tab._queue_thread = worker
    worker.start()

    assert notification_prepared.wait(timeout=2)
    tab.stop_queue()
    release_worker.set()
    worker.join(timeout=2)

    assert callback_ran.wait(timeout=2)
    assert callback_summary[0]["superseded"] is True
    assert tab._queue_run_id == 4


def test_stop_invalidates_callback_already_posted_to_gui():
    posted = []
    callback_summary = []
    callback_posted = threading.Event()
    tab = QueueTab.__new__(QueueTab)
    tab._session = SimpleNamespace(is_running=False, measurement_queue=[])

    def after(_ms, callback, *args):
        posted.append((callback, args))
        callback_posted.set()

    tab._root = SimpleNamespace(after=after)
    tab._completion_callbacks = [callback_summary.append]
    tab._queue_run_id = 20
    tab.log = Mock()
    worker = threading.Thread(
        target=tab._notify_completion_callbacks, args=(0, 20)
    )
    tab._queue_thread = worker
    worker.start()
    worker.join(timeout=2)

    assert callback_posted.wait(timeout=2)
    tab.stop_queue()
    callback, args = posted.pop()
    callback(*args)

    assert callback_summary[0]["superseded"] is True


def test_all_completion_callbacks_share_pre_handoff_summary():
    received = []
    tab = QueueTab.__new__(QueueTab)
    tab._session = SimpleNamespace(measurement_queue=[])
    tab._root = SimpleNamespace(after=lambda _ms, callback, *args: callback(*args))
    tab._queue_run_id = 11
    tab.log = Mock()

    def bo_callback(summary):
        received.append(("bo", summary))
        # Simulate BO starting the next queue before the archive callback runs.
        tab._queue_run_id = 12

    def archive_callback(summary):
        received.append(("archive", summary))

    tab._completion_callbacks = [bo_callback, archive_callback]
    tab._notify_completion_callbacks(0, run_id=11)

    assert [name for name, _summary in received] == ["bo", "archive"]
    assert [summary["superseded"] for _name, summary in received] == [False, False]


def test_automatic_start_reports_superseded_rejection_without_busy_dialog(monkeypatch):
    tab = QueueTab.__new__(QueueTab)
    tab._session = SimpleNamespace(
        is_running=False,
        measurement_queue=[{"type": "PUMP_INIT", "status": "pending"}],
    )
    tab._queue_run_id = 5
    tab._queue_thread = None
    tab._reset_reorder = Mock()
    tab._validate_queue_scripts = Mock(return_value=True)
    tab.log = Mock()
    warning = Mock()
    monkeypatch.setattr("gui.tab_queue.messagebox.showwarning", warning)

    started = tab.run_from_index(0, expected_run_id=4, automatic=True)

    assert started is False
    warning.assert_not_called()
    tab._validate_queue_scripts.assert_not_called()
    assert "superseded" in tab.log.call_args.args[0]


def test_automatic_start_reports_live_worker_rejection_without_busy_dialog(monkeypatch):
    tab = QueueTab.__new__(QueueTab)
    tab._session = SimpleNamespace(
        is_running=False,
        measurement_queue=[{"type": "PUMP_INIT", "status": "pending"}],
    )
    tab._queue_run_id = 5
    tab._queue_thread = Mock(is_alive=Mock(return_value=True))
    tab._reset_reorder = Mock()
    tab._validate_queue_scripts = Mock(return_value=True)
    tab.log = Mock()
    warning = Mock()
    monkeypatch.setattr("gui.tab_queue.messagebox.showwarning", warning)

    started = tab.run_from_index(0, expected_run_id=5, automatic=True)

    assert started is False
    warning.assert_not_called()
    tab._validate_queue_scripts.assert_not_called()
    assert "worker is still active" in tab.log.call_args.args[0]
