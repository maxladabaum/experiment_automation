import json
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from gui.tab_pump import PumpTab


class PumpAutoconnectTests(unittest.TestCase):
    def run_search(self, responding, saved=8):
        tab = PumpTab.__new__(PumpTab)
        tab._ctrl = Mock()
        def connect(port, baud, dev):
            if port not in responding:
                raise RuntimeError('No response')
        tab._ctrl.connect.side_effect = connect
        tab._root = Mock()
        tab._root.after.side_effect = lambda delay, callback: callback()
        tab._var_com = Mock()
        tab.log = Mock()
        tab._remember_pump_port = Mock()
        ports = [SimpleNamespace(device='COM4', vid=1027),
                 SimpleNamespace(device='COM5', vid=1027),
                 SimpleNamespace(device='COM1', vid=None)]
        with patch('gui.tab_pump.HAS_SERIAL_PORTS', True), \
             patch('gui.tab_pump.serial.tools.list_ports.comports', return_value=ports) as scan, \
             patch('gui.tab_pump.messagebox.showerror') as error:
            tab._do_autoconnect(False, saved, 9600, 1)
        return tab, scan, error

    def test_saved_port_success_does_not_scan(self):
        tab, scan, error = self.run_search({4}, saved=4)
        scan.assert_not_called()
        error.assert_not_called()
        tab._ctrl.connect.assert_called_once_with(4, 9600, 1)

    def test_discovers_usb_port_and_remembers_it(self):
        tab, scan, error = self.run_search({4})
        self.assertEqual([c.args[0] for c in tab._ctrl.connect.call_args_list], [8, 4, 5, 4])
        tab._remember_pump_port.assert_called_once_with(4, 9600, 1)
        tab._var_com.set.assert_called_once_with('COM4')
        error.assert_not_called()

    def test_multiple_pumps_require_manual_selection(self):
        tab, scan, error = self.run_search({4, 5})
        self.assertIn('Multiple pumps', error.call_args.args[1])
        tab._remember_pump_port.assert_not_called()
        self.assertEqual(tab._ctrl.disconnect.call_count, 2)

    def test_no_pump_reports_failure(self):
        tab, scan, error = self.run_search(set())
        self.assertIn('No pump responded', error.call_args.args[1])
        tab._remember_pump_port.assert_not_called()

    def test_save_preserves_other_settings(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'local_config.json'
            path.write_text(json.dumps({'data_dir': 'existing'}), encoding='utf-8')
            with patch('gui.tab_pump.LOCAL_CONFIG_PATH', path):
                PumpTab._remember_pump_port(4, 9600, 1)
            self.assertEqual(json.loads(path.read_text()), {
                'data_dir': 'existing', 'pump_com_port': 4,
                'pump_baud': 9600, 'pump_dev': 1})

    def test_real_connection_cannot_fall_back_to_simulation(self):
        from pump_gui import PumpCtrl
        with patch('pump_gui.HAS_COM', False):
            with self.assertRaisesRegex(RuntimeError, 'COM driver unavailable'):
                PumpCtrl(use_sim=False).connect(4, 9600, 1)


if __name__ == '__main__':
    unittest.main()
