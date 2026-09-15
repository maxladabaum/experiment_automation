"""The withdrawn bipot workaround must not alter normal queued measurements."""
from types import SimpleNamespace
from unittest.mock import Mock

from core.runner import SerialMeasurementRunner
from core.swv_method import build_swv_methodscript
from gui.tab_queue import QueueTab


def test_200hz_swv_sent_and_saved_without_routing_conversion(tmp_path):
    script = build_swv_methodscript(dict(begin_potential=-.6, end_potential=0,
        step_potential=.002, amplitude=.036, frequency=200,
        conditioning_potential=-.6, conditioning_time=.2))
    source = tmp_path / 'swv.ms'
    source.write_text(script)
    runner = SerialMeasurementRunner(source, data_folder=tmp_path)
    runner.connect = Mock(return_value=True)
    runner.run_script = Mock(return_value=False)
    runner.disconnect = Mock()
    runner.execute('unchanged')
    runner.run_script.assert_called_once_with(script)
    used = list((tmp_path/'methods_used').glob('*.ms'))
    assert len(used) == 1 and used[0].read_text() == script
    assert 'set_max_bandwidth 4k' in script
    assert 'poly_we' not in script


def test_queue_does_not_forward_stale_experimental_selection(tmp_path, monkeypatch):
    tab = QueueTab.__new__(QueueTab)
    tab._session = SimpleNamespace(queue_electrode_route='we0_re1_ce1',
        session_manager=None, save_raw_packets=True, simulate_measurements=False,
        device_port=None, current_runner=None)
    tab._root = Mock()
    tab._plotter = Mock()
    tab._pump_ctrl = None
    tab.log = Mock()
    tab._ensure_mux_script_for_item = Mock()
    tab._extract_mux_channel = Mock(return_value=1)
    tab._next_measurement_tag = Mock(return_value='unchanged')
    factory = Mock()
    factory.return_value.execute.return_value = (True, tmp_path/'result.csv')
    monkeypatch.setattr('gui.tab_queue.SerialMeasurementRunner', factory)
    tab._execute_measurement_item({'type':'SWV','script_path':str(tmp_path/'swv.ms')})
    assert 'electrode_route' not in factory.call_args.kwargs
    assert factory.call_args.kwargs['invert_current'] is True
