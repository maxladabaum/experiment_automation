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
