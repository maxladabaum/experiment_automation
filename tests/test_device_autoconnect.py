from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from core.runner import SerialMeasurementRunner, get_pump_com_port
from gui.tab_method import MethodTab
from gui.tab_queue import QueueTab


def port(device):
    return SimpleNamespace(device=device, description='USB Serial Port', manufacturer='FTDI')


@pytest.mark.parametrize('pump_port', [3, '3', 'COM3', ' com3 '])
def test_autodetect_and_preview_deprioritize_pump(monkeypatch, tmp_path, pump_port):
    ports = [port('COM3'), port('COM8')]
    monkeypatch.setattr('core.runner.serial.tools.list_ports.comports', lambda **kwargs: ports)
    runner = SerialMeasurementRunner('test.ms', data_folder=tmp_path, pump_com_port=pump_port)
    assert runner.find_device_port() == 'COM8'
    assert MethodTab._auto_detect_port(ports, pump_port) == 'COM8'


@pytest.mark.parametrize('devices, pump_port, expected', [
    (['COM3', 'COM8'], None, 'COM3'),
    (['COM3'], 3, 'COM3'),
    ([], 3, None),
])
def test_existing_autodetect_fallbacks(monkeypatch, tmp_path, devices, pump_port, expected):
    ports = [port(device) for device in devices]
    monkeypatch.setattr('core.runner.serial.tools.list_ports.comports', lambda **kwargs: ports)
    runner = SerialMeasurementRunner('test.ms', data_folder=tmp_path, pump_com_port=pump_port)
    assert runner.find_device_port() == expected
    assert MethodTab._auto_detect_port(ports, pump_port) == expected


def test_current_real_pump_port_is_available_during_connect():
    pump = SimpleNamespace(use_sim=False, com_port=3, connected=False)
    assert get_pump_com_port(pump) == 3
    pump.com_port = 8
    pump.connected = True
    assert get_pump_com_port(pump) == 8
    pump.use_sim = True
    assert get_pump_com_port(pump) is None
    assert get_pump_com_port(None) is None


def test_queue_measurement_uses_live_pump_port(monkeypatch, tmp_path):
    tab = QueueTab.__new__(QueueTab)
    tab._pump_ctrl = SimpleNamespace(use_sim=False, com_port=3, connected=True)
    tab._session = SimpleNamespace(session_manager=None, save_raw_packets=False,
        simulate_measurements=False, device_port=None, current_runner=None)
    tab._root = Mock()
    tab._plotter = Mock()
    tab._ensure_mux_script_for_item = Mock()
    tab._extract_mux_channel = Mock(return_value=None)
    tab._next_measurement_tag = Mock(return_value='test')
    tab.log = Mock()
    tab.refresh_labels = Mock()
    selected_ports = []

    def make_runner(*args, **kwargs):
        kwargs['data_folder'] = tmp_path
        runner = SerialMeasurementRunner(*args, **kwargs)
        def execute(**kwargs):
            selected_ports.append(runner.find_device_port())
            return True, None
        runner.execute = execute
        return runner

    monkeypatch.setattr('gui.tab_queue.SerialMeasurementRunner', make_runner)
    monkeypatch.setattr('core.runner.serial.tools.list_ports.comports',
                        lambda **kwargs: [port('COM3'), port('COM8')])
    assert tab._execute_measurement_item({'type': 'CV', 'script_path': 'test.ms'}) == (True, None)
    tab._pump_ctrl.com_port = 8
    assert tab._execute_measurement_item({'type': 'CV', 'script_path': 'test.ms'}) == (True, None)
    assert selected_ports == ['COM8', 'COM3']
