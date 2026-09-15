"""No serial hardware is accessed by these tests."""
import tempfile
from pathlib import Path
import unittest
from unittest.mock import patch

from core.pico_routing_test import build_test_script, RoutingTestConnection, RESTORE


class FakeSerial:
    def __init__(self, response):
        self.response = response
        self.sent = []
        self.closed = False

    @property
    def in_waiting(self):
        return len(self.response)

    def read(self, n):
        data, self.response = self.response[:n], self.response[n:]
        return data

    def write(self, data):
        self.sent.append(data)

    def close(self):
        self.closed = True


class RoutingTests(unittest.TestCase):
    def test_default_preserves_mux_and_uses_existing_route(self):
        script=build_test_script()
        self.assertNotIn('set_gpio',script)
        self.assertNotIn('poly_we(',script)
        self.assertNotIn('set_pgstat_mode 5',script)
        self.assertIn('cell_on',script)

    def test_candidate_records_secondary_after_main(self):
        script=build_test_script('RE1/CE1 candidate')
        self.assertIn('poly_we(0 b)',script)
        self.assertIn('pck_add c\npck_add b',script)
        self.assertIn('if b > 50u',script)
        self.assertNotIn('cell_on',build_test_script('RE1/CE1 candidate',kind='configure'))
        self.assertNotIn('cell_off', '\n'.join(RESTORE))

    def test_invalid_requests_fail_before_serial(self):
        for route in ['RE1/CE0','RE0/CE1','unknown']:
            with self.assertRaises(ValueError):build_test_script(route)
        for center in [float('nan'),float('inf'),.6]:
            with self.assertRaises(ValueError):build_test_script(center=center)
        for mux in [0,17,True]:
            with self.assertRaises(ValueError):build_test_script(mux_position=mux)

    def test_observed_overload_is_aborted_and_saved(self):
        packet=b'Pda73B41B3n;ba8000000 ,14,201;ba7DA796Bp,12,201\n'
        with tempfile.TemporaryDirectory() as tmp:
            device=RoutingTestConnection('unused',tmp,log=lambda s:None)
            device.serial=FakeSerial(packet)
            with self.assertRaisesRegex(RuntimeError,'field 2'):
                device.run('candidate','e\n\n')
            self.assertIn(b'Z\n',device.serial.sent)
            self.assertIn('ba7DA796Bp',(Path(tmp)/'01_candidate_raw.txt').read_text())

    def test_normal_completion_needs_marker(self):
        with tempfile.TemporaryDirectory() as tmp:
            device=RoutingTestConnection('unused',tmp,log=lambda s:None)
            device.serial=FakeSerial(b'e\nTROUTING_TEST_DONE\n')
            self.assertEqual(device.run('configuration','e\n\n'),[])

    def test_cancel_sends_abort(self):
        with tempfile.TemporaryDirectory() as tmp:
            device=RoutingTestConnection('unused',tmp,log=lambda s:None,cancelled=lambda:True)
            device.serial=FakeSerial(b'')
            with self.assertRaisesRegex(RuntimeError,'cancelled'):
                device.run('candidate','e\n\n')
            self.assertEqual(device.serial.sent[-1],b'Z\n')

    def test_restore_failure_preserves_original_error(self):
        with tempfile.TemporaryDirectory() as tmp:
            device=RoutingTestConnection('unused',tmp,log=lambda s:None)
            device.serial=FakeSerial(b'')
            with patch.object(device,'run',side_effect=RuntimeError('restore failure')):
                device.__exit__(RuntimeError,RuntimeError('original'),None)
            self.assertFalse(device.restore_confirmed)
            self.assertEqual(device.restore_error,'restore failure')
            self.assertTrue(device.serial.closed)


if __name__ == '__main__':
    unittest.main()
