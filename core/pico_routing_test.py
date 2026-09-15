"""Bounded, opt-in live routing investigation using legacy Pico commands.

Retains the MUX unless a position is explicitly supplied. Does not change NVM,
calibration, or existing method files.
Requested terminal routing is a hypothesis until independently identified.
"""
import csv
import json
import math
from pathlib import Path
import time

import serial
from core.mscript_parser import parse_mscript_data_package

RESTORE = ['set_pgstat_chan 1', 'set_pgstat_mode 0',
           'set_pgstat_chan 0', 'set_pgstat_mode 0', 'set_pgstat_mode 3']


def build_test_script(route='RE0/CE0', center=0.0, kind='ca', mux_position=None):
    if route not in ('RE0/CE0', 'RE1/CE1 candidate'):
        raise ValueError('Mixed physical RE/CE routing is not established.')
    if kind not in ('ca', 'ocp', 'configure'):
        raise ValueError('Unknown test kind')
    if not math.isfinite(center) or abs(center) > .5:
        raise ValueError('Test center must be finite and within +/-0.5 V')
    reverse = route != 'RE0/CE0'
    if reverse and kind == 'ocp':
        raise ValueError('OCP is not supported in bipot mode')
    main, secondary = (1, 0) if reverse else (0, 1)
    lines = ['e', 'var p', 'var c', 'var b'] + RESTORE
    if mux_position is not None:
        if isinstance(mux_position, bool) or mux_position not in range(1,17):
            raise ValueError('MUX position must be 1-16')
        idx = mux_position-1
        lines += ['set_gpio_cfg 0x3FFi 1', f'set_gpio {(idx << 4) | idx}i']
    if reverse:
        lines += ['set_pgstat_chan 0', 'set_pgstat_mode 5',
                  'set_poly_we_mode 1', 'set_e 0', 'set_range ba 1u']
    lines += [f'set_pgstat_chan {main}', 'set_pgstat_mode 2',
              'set_range ba 1u', 'set_max_bandwidth 40',
              'set_range_minmax da -550m 550m']
    if kind == 'configure':
        pass  # Validate configuration with the cell off before any candidate scan.
    elif kind == 'ocp':
        lines += ['cell_off', 'meas_loop_ocp p 100m 2', 'pck_start',
                  'pck_add p', 'pck_end', 'endloop']
    else:
        start = round(center*1000)
        lines += [f'set_e {start}m', 'cell_on']
        for delta in (0, 10, -10, 0):
            extra = f' poly_we({secondary} b)' if reverse else ''
            lines += [f'meas_loop_ca p c {start+delta}m 100m 500m{extra}',
                      'pck_start', 'pck_add p', 'pck_add c']
            if reverse:
                lines += ['pck_add b']
            lines += ['pck_end']
            for var in (('c', 'b') if reverse else ('c',)):
                for op, threshold in [('>', '50u'), ('<', '-50u')]:
                    lines += [f'if {var} {op} {threshold}', 'cell_off', 'abort', 'endif']
            lines += ['endloop']
        lines += ['cell_off']
    lines += RESTORE + ['send_string "ROUTING_TEST_DONE"']
    return '\n'.join(lines)+'\n\n'


class RoutingTestConnection:
    def __init__(self, port, output_dir, log=print, cancelled=lambda: False):
        self.port, self.output = port, Path(output_dir)
        self.log, self.cancelled = log, cancelled
        self.serial = None
        self.sequence = 0
        self.restore_confirmed = False
        self.restore_error = None
        self.stage = None

    def __enter__(self):
        self.output.mkdir(parents=True, exist_ok=False)
        self.serial = serial.Serial(self.port, 230400, timeout=.1, write_timeout=1)
        try:
            time.sleep(.25)
            pending = self.serial.read(self.serial.in_waiting)
            if pending:
                raise RuntimeError('Device already emitting data; test not started')
            self.serial.write(b't\n')
            deadline = time.monotonic()+2
            identity = b''
            while time.monotonic() < deadline:
                identity += self.serial.read(self.serial.in_waiting or 1)
            (self.output/'identity.txt').write_bytes(identity)
            if b'tespico' not in identity.lower():
                raise RuntimeError('Port did not identify as EmStat Pico')
            self.log(identity.decode(errors='replace').strip())
            return self
        except BaseException:
            self.serial.close()
            raise

    def __exit__(self, *args):
        try:
            # Any fault or cancellation must finish with both cells off.
            self.serial.write(b'Z\n')
            time.sleep(.15)
            self.serial.read(self.serial.in_waiting)
            self.run('restore', 'e\n'+'\n'.join(RESTORE)+
                     '\nsend_string "ROUTING_TEST_DONE"\n\n', allow_cancel=False)
            self.restore_confirmed = True
        except Exception as exc:
            self.restore_error = str(exc)
            self.log('RESTORATION FAILED: '+str(exc))
            if args[0] is None:
                raise
        finally:
            self.serial.close()

    def run(self, label, script, allow_cancel=True):
        self.sequence += 1
        if label != 'restore':
            self.stage = label
        stem = f'{self.sequence:02d}_{label}'
        (self.output/(stem+'.ms')).write_text(script, encoding='utf8')
        packets, lines, buffer = [], [], b''
        self.log('Running '+label)
        self.serial.write(script.encode('ascii'))
        deadline = time.monotonic()+12
        completed = False
        try:
            while time.monotonic() < deadline:
                if allow_cancel and self.cancelled():
                    raise RuntimeError('Test cancelled')
                buffer += self.serial.read(self.serial.in_waiting or 1)
                while b'\n' in buffer:
                    raw, buffer = buffer.split(b'\n', 1)
                    line = raw.decode('ascii', errors='replace').rstrip('\r')
                    lines.append(line)
                    if line.startswith('!'):
                        raise RuntimeError('Pico rejected test: '+line)
                    if line.startswith('P'):
                        if 'nan' in line:
                            raise RuntimeError('Invalid measurement packet; stopped')
                        packet = parse_mscript_data_package(line+'\n')
                        if packet:
                            values = [v.value for v in packet]
                            if not all(math.isfinite(v) for v in values):
                                raise RuntimeError('Non-finite measurement; stopped')
                            packets.append([(v.id, v.value, v.raw_metadata) for v in packet])
                            for field,v in enumerate(packet):
                                for m in v.raw_metadata:
                                    if m.startswith('1') and int(m[1:],16) & 3:
                                        raise RuntimeError(f'{label}: field {field} ({v.id}) status {m}: overload/timing fault; stopped')
                            if any(abs(v.value) > 50e-6 for v in packet if v.id in ('ba','bb')):
                                raise RuntimeError('50 uA test limit exceeded; stopped')
                    if line == 'TROUTING_TEST_DONE':
                        completed = True
                if completed:
                    break
            if not completed:
                raise RuntimeError('Test timed out without completion marker')
        finally:
            if not completed:
                self.serial.write(b'Z\n')
            (self.output/(stem+'_raw.txt')).write_text('\n'.join(lines)+'\n'+buffer.decode(errors='replace'), encoding='utf8')
            (self.output/(stem+'.json')).write_text(json.dumps(packets,indent=2),encoding='utf8')
            with (self.output/(stem+'.csv')).open('w',newline='',encoding='utf8') as fh:
                w=csv.writer(fh)
                w.writerow(['sample','field_index','vartype','value_SI','metadata'])
                for n,packet in enumerate(packets):
                    w.writerows((n,i,tag,value,json.dumps(meta)) for i,(tag,value,meta) in enumerate(packet))
        self.log(f'{label}: {len(packets)} packets; completed')
        return packets


def run_flowcell_test(port, output, log=print, cancelled=lambda: False, candidate=False, mux_position=None):
    result = {'port':port,'mux':mux_position or 'unchanged','candidate_requested':candidate,
              'physical_routing_verified':False}
    device = RoutingTestConnection(port,output,log,cancelled)
    # Do not overwrite results from another run on an output-path collision.
    if Path(output).exists():
        raise FileExistsError(output)
    try:
        with device:
            ocp = device.run('baseline_ocp',build_test_script(kind='ocp',mux_position=mux_position))
            if len(ocp) < 5:
                raise RuntimeError('Insufficient OCP data')
            center = sum(p[0][1] for p in ocp[-5:])/5
            result['baseline_ocp_V'] = center
            device.run('baseline',build_test_script(center=center))
            if candidate:
                device.run('reverse_configuration',build_test_script('RE1/CE1 candidate',kind='configure'))
                result['candidate_configuration_accepted'] = True
                device.run('reverse',build_test_script('RE1/CE1 candidate',center=center))
            device.run('baseline_return',build_test_script(center=center))
            result['completed'] = True
    except Exception as exc:
        result.update(completed=False,error=str(exc))
        log(str(exc))
    finally:
        result['last_stage'] = device.stage
        result['restore_confirmed'] = device.restore_confirmed
        result['restore_error'] = device.restore_error
        if Path(output).is_dir():
            (Path(output)/'result.json').write_text(json.dumps(result,indent=2),encoding='utf8')
    return result
