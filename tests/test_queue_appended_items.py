from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from gui.tab_queue import QueueTab
from core.runner import SerialMeasurementRunner
from core.session import SessionState


def make_queue(monkeypatch, items, delay=0):
    monkeypatch.setattr('gui.tab_queue.estimate_item_seconds', lambda item: 0)
    tab = QueueTab.__new__(QueueTab)
    tab._session = SimpleNamespace(measurement_queue=items, is_running=True,
                                   step_delay=delay, update_queue_status=Mock())
    tab._root = SimpleNamespace(after=lambda ms, callback, *args: callback(*args))
    tab.refresh = Mock()
    tab.set_status = Mock()
    tab.log = Mock()
    tab._copy_queue_file = Mock()
    tab._announce_queue_end = Mock()
    tab._completion_callbacks = [Mock()]
    return tab


def pump(name):
    return {'type': 'PUMP_VALVE', 'details': name, 'status': 'pending'}


def test_items_added_during_last_step_execute_in_order_and_are_counted(monkeypatch):
    tab = make_queue(monkeypatch, [pump('first')])
    executed = []

    def execute(item):
        executed.append(item['details'])
        if item['details'] == 'first':
            tab.add_item(pump('second'))
            tab.add_item(pump('third'))
        return True

    tab._exec_pump = execute
    tab._execute_queue()
    assert executed == ['first', 'second', 'third']
    assert all(item['status'] == 'completed' for item in tab._session.measurement_queue)
    summary = tab._completion_callbacks[0].call_args.args[0]
    assert summary['total'] == summary['completed'] == 3
    tab._copy_queue_file.assert_called_once_with('queue_updated')


def test_appended_steps_stay_pending_when_user_stops(monkeypatch):
    tab = make_queue(monkeypatch, [pump('first')])

    def execute(item):
        tab.add_item(pump('later'))
        tab._session.is_running = False
        return True

    tab._exec_pump = Mock(side_effect=execute)
    tab._execute_queue()
    tab._exec_pump.assert_called_once()
    assert tab._session.measurement_queue[1]['status'] == 'pending'


def test_resume_and_append_during_inter_step_delay(monkeypatch):
    items = [pump('already run'), pump('second'), pump('third')]
    items[0]['status'] = 'completed'
    tab = make_queue(monkeypatch, items, delay=1)
    tab._exec_pump = Mock(return_value=True)

    def pause(seconds):
        if len(items) == 3:
            tab.add_item(pump('added during delay'))
        return True

    tab._exec_pause = pause
    tab._execute_queue(1)
    assert [c.args[0]['details'] for c in tab._exec_pump.call_args_list] == [
        'second', 'third', 'added during delay']
    assert tab._completion_callbacks[0].call_args.args[0]['completed'] == 3


def test_raw_packets_default_on_and_runner_can_opt_out(monkeypatch, tmp_path):
    monkeypatch.setattr('core.session.MethodRegistry', Mock())
    assert SessionState().save_raw_packets is True
    runner = SerialMeasurementRunner(tmp_path / 'method.ms', data_folder=tmp_path)
    assert runner.save_raw_packets is True
    runner = SerialMeasurementRunner(tmp_path / 'method.ms', data_folder=tmp_path,
                                     save_raw_packets=False)
    assert runner.save_raw_packets is False


def test_pump_failure_stops_queue_before_any_later_action(monkeypatch):
    tab = make_queue(monkeypatch, [pump('failed dispense'), pump('later valve')])
    tab._exec_pump = Mock(return_value=False)
    tab._execute_queue()
    tab._exec_pump.assert_called_once()
    assert [item['status'] for item in tab._session.measurement_queue] == ['failed', 'pending']
    assert not tab._session.is_running
    assert not any(call.args[0]=='Queue completed.' for call in tab.log.call_args_list)
