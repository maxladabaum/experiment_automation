#!/usr/bin/env python3
# electrochemistry_automation_gui.py
# For bash use -> source venv_gui/Scripts/activate
# For powershell use -> venv_gui\Scripts\Activate.ps1
# To run -> python -m gui_script

import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import json
import os
from datetime import datetime
from pathlib import Path
import threading
import io
import time
import sys
import serial
import serial.tools.list_ports
import csv
import math
import collections
import queue
import warnings
import traceback
import itertools
import hashlib
from typing import Dict, List, Optional

# ── Canonical version of this file ──────────────────────────────────────────
APP_VERSION = "2.0.0"          # Phase 0 consolidation + Phase 1 bug-fixes
# ────────────────────────────────────────────────────────────────────────────


# --- Matplotlib and Pandas imports for plotting ---
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk


try:
    from pump_gui import (
        PumpCtrl,
        HAS_COM as PUMP_HAS_COM,
        SPEED_MIN as PUMP_SPEED_MIN,
        SPEED_MAX as PUMP_SPEED_MAX,
        DEFAULT_COM_PORT as PUMP_DEFAULT_COM_PORT,
        DEFAULT_BAUD as PUMP_DEFAULT_BAUD,
        DEFAULT_DEV as PUMP_DEFAULT_DEV,
        DEFAULT_STEPS as PUMP_DEFAULT_STEPS,
        DEFAULT_SYRINGE as PUMP_DEFAULT_SYRINGE,
    )
    PUMP_AVAILABLE = True
except ImportError:
    PumpCtrl = None
    PUMP_HAS_COM = False
    PUMP_AVAILABLE = False
    PUMP_DEFAULT_COM_PORT = 1
    PUMP_DEFAULT_BAUD = 9600
    PUMP_DEFAULT_DEV = 1
    PUMP_DEFAULT_STEPS = 100000
    PUMP_DEFAULT_SYRINGE = 1250.0
    PUMP_SPEED_MIN = 1
    PUMP_SPEED_MAX = 40
    print("Warning: pump_gui backend not found. Pump features disabled.")

try:
    import pythoncom  # type: ignore
except Exception:
    pythoncom = None

PREFERRED_SYRINGE_UL = 250.0
PREFERRED_STEPS_PER_STROKE = 181490

PUMP_DEFAULT_STEPS = PREFERRED_STEPS_PER_STROKE
PUMP_DEFAULT_SYRINGE = PREFERRED_SYRINGE_UL

# --- PalmSens MethodSCRIPT Parser Integration ---
# The following code is adapted from the provided mscript.py file
# to correctly parse data packages from the device.

# Custom types
VarType = collections.namedtuple('VarType', ['id', 'name', 'unit'])

# Dictionary for the conversion of the SI prefixes.
SI_PREFIX_FACTOR = {
    'a': 1e-18, 'f': 1e-15, 'p': 1e-12, 'n': 1e-9, 'u': 1e-6,
    'm': 1e-3, ' ': 1e0, 'k': 1e3, 'M': 1e6, 'G': 1e9,
    'T': 1e12, 'P': 1e15, 'E': 1e18, 'i': 1e0,
}

# List of MethodSCRIPT variable types.
MSCRIPT_VAR_TYPES_LIST = [
    VarType('aa', 'unknown', ''),
    VarType('ab', 'WE vs RE potential', 'V'),
    VarType('ac', 'CE vs GND potential', 'V'),
    VarType('ad', 'SE vs GND potential', 'V'),
    VarType('ae', 'RE vs GND potential', 'V'),
    VarType('af', 'WE vs GND potential', 'V'),
    VarType('ag', 'WE vs CE potential', 'V'),
    VarType('as', 'AIN0 potential', 'V'),
    VarType('at', 'AIN1 potential', 'V'),
    VarType('au', 'AIN2 potential', 'V'),
    VarType('av', 'AIN3 potential', 'V'),
    VarType('aw', 'AIN4 potential', 'V'),
    VarType('ax', 'AIN5 potential', 'V'),
    VarType('ay', 'AIN6 potential', 'V'),
    VarType('az', 'AIN7 potential', 'V'),
    VarType('ba', 'WE current', 'A'),
    VarType('ca', 'Phase', 'degrees'),
    VarType('cb', 'Impedance', '\u2126'),
    VarType('cc', 'Z_real', '\u2126'),
    VarType('cd', 'Z_imag', '\u2126'),
    VarType('ce', 'EIS E TDD', 'V'),
    VarType('cf', 'EIS I TDD', 'A'),
    VarType('cg', 'EIS sampling frequency', 'Hz'),
    VarType('ch', 'EIS E AC', 'Vrms'),
    VarType('ci', 'EIS E DC', 'V'),
    VarType('cj', 'EIS I AC', 'Arms'),
    VarType('ck', 'EIS I DC', 'A'),
    VarType('da', 'Applied potential', 'V'),
    VarType('db', 'Applied current', 'A'),
    VarType('dc', 'Applied frequency', 'Hz'),
    VarType('dd', 'Applied AC amplitude', 'Vrms'),
    VarType('ea', 'Channel', ''),
    VarType('eb', 'Time', 's'),
    VarType('ec', 'Pin mask', ''),
    VarType('ed', 'Temperature', '\u00B0 Celsius'),
    VarType('ee', 'Count', ''),
    VarType('ha', 'Generic current 1', 'A'),
    VarType('hb', 'Generic current 2', 'A'),
    VarType('hc', 'Generic current 3', 'A'),
    VarType('hd', 'Generic current 4', 'A'),
    VarType('ia', 'Generic potential 1', 'V'),
    VarType('ib', 'Generic potential 2', 'V'),
    VarType('ic', 'Generic potential 3', 'V'),
    VarType('id', 'Generic potential 4', 'V'),
    VarType('ja', 'Misc. generic 1', ''),
    VarType('jb', 'Misc. generic 2', ''),
    VarType('jc', 'Misc. generic 3', ''),
    VarType('jd', 'Misc. generic 4', ''),
]

MSCRIPT_VAR_TYPES_DICT = {x.id: x for x in MSCRIPT_VAR_TYPES_LIST}


def get_variable_type(var_id: str) -> VarType:
    """Get the variable type with the specified id."""
    if var_id in MSCRIPT_VAR_TYPES_DICT:
        return MSCRIPT_VAR_TYPES_DICT[var_id]
    warnings.warn(f'Unsupported VarType id "{var_id}"!')
    return VarType(var_id, 'unknown', '')


class MScriptVar:
    """Class to store and parse a received MethodSCRIPT variable."""
    def __init__(self, data: str):
        assert len(data) >= 10
        self.data = data[:]
        self.id = data[0:2]
        if data[2:10] == '     nan':
            self.raw_value = math.nan
            self.si_prefix = ' '
        else:
            self.raw_value = self.decode_value(data[2:9])
            self.si_prefix = data[9]
        self.raw_metadata = data.split(',')[1:]
        self.metadata = self.parse_metadata(self.raw_metadata)

    @property
    def type(self) -> VarType:
        return get_variable_type(self.id)

    @property
    def si_prefix_factor(self) -> float:
        return SI_PREFIX_FACTOR[self.si_prefix]

    @property
    def value(self) -> float:
        return self.raw_value * self.si_prefix_factor

    @staticmethod
    def decode_value(var: str):
        """Decode the raw value of a MethodSCRIPT variable."""
        assert len(var) == 7
        return int(var, 16) - (2 ** 27)

    @staticmethod
    def parse_metadata(tokens: List[str]) -> Dict[str, int]:
        """Parse the (optional) metadata."""
        metadata = {}
        for token in tokens:
            if (len(token) == 2) and (token[0] == '1'):
                metadata['status'] = int(token[1], 16)
            if (len(token) == 3) and (token[0] == '2'):
                metadata['cr'] = int(token[1:], 16)
        return metadata


def parse_mscript_data_package(line: str) -> Optional[List[MScriptVar]]:
    """Parse a MethodSCRIPT data package."""
    if line.startswith('P') and line.endswith('\n'):
        vars_out = []
        for var in line[1:-1].split(';'):
            if len(var) < 10:
                continue
            value_token = var[2:9]
            if any(ch not in '0123456789ABCDEFabcdef' for ch in value_token):
                continue
            if var[9] not in SI_PREFIX_FACTOR:
                continue
            try:
                vars_out.append(MScriptVar(var))
            except Exception:
                continue
        return vars_out
    return None

# --- End of PalmSens MethodSCRIPT Parser Integration ---


# --- Helper function to convert float to SI string ---
def to_si_string(value_str, unit='V'):
    """Converts a string float value to an SI unit string for the device."""
    try:
        val = float(value_str)
    except (ValueError, TypeError):
        return value_str  # Return original if conversion fails

    if unit in ['V', 'V/s']:
        if val == 0:
            return "0"
        milli_value = val * 1000.0
        formatted = f"{milli_value:.12f}".rstrip('0').rstrip('.')
        if formatted in ('', '-0', '+0'):
            formatted = '0'
        return f"{formatted}m"
    if unit == 'Hz':
        if val.is_integer():
            return f"{int(val)}"
        return f"{val:g}"
    return value_str  # Fallback for unsupported units

# --- Integrated SerialMeasurementRunner Class (Unchanged) ---
class SerialMeasurementRunner:
    def __init__(self, script_path, log_callback=print, data_callback=None):
        self.script_path = Path(script_path)
        self.data_points = []
        self.connection = None
        self.log = log_callback # Callback to log messages to the GUI
        self.data_callback = data_callback
        self.is_running = True
        self.partial_packet = ""

        self.data_base_path = Path("measurement_data") #TODO: change this to the DataDrive Path
        self.data_base_path.mkdir(exist_ok=True)
        date_folder = datetime.now().strftime('%Y-%m-%d')
        self.data_folder = self.data_base_path / date_folder
        self.data_folder.mkdir(exist_ok=True)

    def find_device_port(self):
        self.log("Scanning for devices...")
        ports = serial.tools.list_ports.comports(include_links=False)
        candidates = []
        for port in ports:
            self.log(f"  Found port: {port.description} ({port.device})")
            if any(name in port.description for name in ['ESPicoDev', 'EmStat', 'USB Serial Port', 'FTDI']):
                candidates.append(port.device)
        if not candidates:
            self.log("ERROR: No measurement device found")
            return None

        pump_port_upper = None
        if PUMP_AVAILABLE and PUMP_DEFAULT_COM_PORT:
            try:
                pump_port_upper = f"COM{int(PUMP_DEFAULT_COM_PORT)}".upper()
            except (TypeError, ValueError):
                pump_port_upper = str(PUMP_DEFAULT_COM_PORT).upper()

        def candidate_key(dev: str):
            return (pump_port_upper is not None and dev.upper() == pump_port_upper, dev)

        candidates.sort(key=candidate_key)

        if len(candidates) > 1:
            self.log(f"Multiple devices found: {candidates}")
            selected = candidates[0]
            if pump_port_upper and selected.upper() != pump_port_upper and any(dev.upper() == pump_port_upper for dev in candidates):
                self.log(f"Using first device: {selected} (pump port {pump_port_upper} deprioritized)")
            else:
                self.log(f"Using first device: {selected}")
        return candidates[0]

    def connect(self, port=None):
        if port is None:
            port = self.find_device_port()
            if port is None: return False
        try:
            self.log(f"Connecting to {port}...")
            self.connection = serial.Serial(port=port, baudrate=230400, timeout=1, write_timeout=1)
            time.sleep(2)
            self.connection.reset_input_buffer()
            self.connection.reset_output_buffer()
            self.connection.write(b't\n')
            response = self.connection.readline()
            if response:
                self.log(f"Device responded: {response.decode('utf-8', errors='ignore').strip()}")
                return True
            else:
                self.log("No response from device")
                return False
        except Exception as e:
            self.log(f"Connection failed: {e}")
            return False

    def stop(self):
        self.is_running = False

    def run_script(self, script):
        if not self.connection:
            self.log("ERROR: Not connected to device")
            return False
        try:
            self.log("Sending script to device...")
            lines = script.strip().split('\n')
            for line in lines:
                self.connection.write((line + '\n').encode('utf-8'))
                time.sleep(0.01)
            self.connection.write(b'\n')
            self.log("Script sent. Collecting data...")
            self.log("-" * 40)
            
            empty_reads = 0
            while self.is_running:
                try:
                    line = self.connection.readline()
                    if not line:
                        if self.partial_packet:
                            empty_reads += 1
                            if empty_reads >= 3:
                                self.log("Warning: incomplete data packet timed out")
                                break
                        continue
                    empty_reads = 0
                    text = line.decode('utf-8', errors='ignore').rstrip('\r\n')
                    if not text: continue
                    if self.partial_packet:
                        if text.startswith('P'):
                            self.log("Warning: dropped incomplete data packet")
                            self.partial_packet = ""
                        else:
                            self.log("Warning: dropped incomplete data packet")
                            self.partial_packet = ""
                    self.log(text)
                    if text.startswith('P'):
                        if not self._is_complete_packet(text):
                            self.partial_packet = text
                            continue
                        self.parse_data_line(text)
                    if text in ['*', 'Measurement completed', 'Script completed']:
                        self.log("\nMeasurement completed")
                        break
                    if text.startswith('!'):
                        self.log(f"Device error: {text}")
                        if "abort" in text.lower(): break
                except serial.SerialException as e:
                    self.log(f"Serial Error: {e}")
                    break
            
            if not self.is_running: self.log("Measurement stopped by user.")
            return True
        except Exception as e:
            self.log(f"Error running script: {type(e).__name__}: {e}")
            self.log(traceback.format_exc())
            return False

    def parse_data_line(self, line):
        """
        Parses a MethodSCRIPT data package line (e.g., 'Pab...') using the
        integrated mscript parsing logic.
        """
        package = parse_mscript_data_package(line + '\n')
        
        if not package:
            return

        try:
            data_point = {}
            for var in package:
                if var.id in ['ab', 'da']:
                    data_point['potential'] = var.value
                elif var.id == 'ba':
                    data_point['current'] = var.value * 1e6
            
            if 'potential' in data_point and 'current' in data_point:
                self.data_points.append(data_point)
                if self.data_callback:
                    try:
                        self.data_callback(data_point)
                    except Exception as exc:
                        self.log(f"Live plot callback error: {type(exc).__name__}: {exc}")
        except Exception as e:
            self.log(f"Error parsing data package: {line} -> {e}")

    @staticmethod
    def _is_complete_packet(line: str) -> bool:
        if not line.startswith('P'):
            return False
        parts = line[1:].split(';')
        if not parts:
            return False
        for part in parts:
            if len(part) < 10:
                return False
        return True

    def save_data_to_csv(self):
        if not self.data_points:
            self.log("No data to save")
            return None
        base_name = self.script_path.stem
        timestamp = datetime.now().strftime('%H%M%S')
        csv_filename = self.data_folder / f"{base_name}_{timestamp}.csv"
        with open(csv_filename, 'w', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=['potential', 'current'])
            writer.writerow({'potential': 'Potential (V)', 'current': 'Current (uA)'})
            writer.writerows(self.data_points)
        self.log(f"\nData saved to: {csv_filename}")
        return csv_filename

    def disconnect(self):
        if self.connection and self.connection.is_open:
            try:
                self.connection.close()
                self.log("Disconnected from device")
            except Exception as e: self.log(f"Error on disconnect: {e}")

    def execute(self):
        self.log("=" * 60)
        self.log(f"Starting measurement for: {self.script_path.name}")
        self.log("=" * 60)
        csv_path = None
        try:
            with open(self.script_path, 'r') as f: script = f.read()
        except Exception as e:
            self.log(f"ERROR: Failed to read script: {e}")
            return False, None
        
        if not self.connect():
            self.log("ERROR: Failed to connect to device")
            return False, None
        
        success = False
        try:
            if self.run_script(script):
                if self.data_points:
                    csv_path = self.save_data_to_csv()
                self.log(f"Total data points: {len(self.data_points)}")
                success = True
        finally:
            self.disconnect()
        return success, csv_path


class ElectrochemGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Electrochemistry Automation System")
        self.root.geometry("1400x900")
        
        self.script_dir = Path(__file__).parent.absolute()
        os.chdir(self.script_dir)
        
        self.base_path = Path("methods")
        self.base_path.mkdir(exist_ok=True)
        
        self.measurement_queue = []
        self.is_running = False
        self.current_script = ""
        self.current_runner = None
        
        self.pump_ctrl = None
        self.pump_busy = False
        self.pump_disable_widgets = []
        self.pump_log_text = None
        self.pump_early_logs = []

        self.live_plot_queue = queue.Queue(maxsize=10000)
        self.live_plot_x = []
        self.live_plot_y = []
        self.live_plot_active = False
        self.live_plot_job = None
        self.plot_line = None
        self.last_live_plot_color = None
        self.last_live_plot_label = None
        self.plot_colors = (
            plt.rcParams.get('axes.prop_cycle', plt.cycler(color=['#1f77b4']))
            .by_key()
            .get('color', ['#1f77b4'])
        )
        self.plot_color_cycle = itertools.cycle(self.plot_colors)

        self.queue_drag_item = None
        self.queue_reorder_pending = False
        self.queue_reorder_snapshot = None
        
        self.setup_gui()
        
    def setup_gui(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        if PUMP_AVAILABLE:
            self.pump_frame = ttk.Frame(self.notebook)
            self.notebook.add(self.pump_frame, text="Pump Control")
            self.setup_pump_tab()
        else:
            self.pump_frame = None

        self.method_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.method_frame, text="Method Creation")
        self.setup_method_tab()

        self.script_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.script_frame, text="Script Preview")
        self.setup_script_tab()

        self.queue_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.queue_frame, text="Queue & Execution")
        self.setup_queue_tab()

        # --- NEW: Add the plotter tab ---
        self.plotter_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.plotter_frame, text="Plotter")
        self.setup_plotter_tab()

    def create_cv_methodscript(self):
        """Create MethodSCRIPT for CV with correct SI unit formatting"""
        begin = to_si_string(self.cv_params['begin_potential'].get(), 'V')
        v1 = to_si_string(self.cv_params['vertex1'].get(), 'V')
        v2 = to_si_string(self.cv_params['vertex2'].get(), 'V')
        step = to_si_string(self.cv_params['step_potential'].get(), 'V')
        scan_rate = to_si_string(self.cv_params['scan_rate'].get(), 'V/s')
        n_scans = self.cv_params['n_scans'].get()
        cond_pot = to_si_string(self.cv_params['cond_potential'].get(), 'V')
        cond_time = self.cv_params['cond_time'].get()

        script_parts = [
            "e", "var c", "var p", "set_pgstat_mode 2", "set_max_bandwidth 40",
            "set_range ba 100u", "set_autoranging ba 1n 100u"
        ]
        
        if float(cond_time) > 0:
            script_parts.extend([
                f"set_e {cond_pot}", "cell_on",
                f"# Condition for {cond_time}s",
                f"wait {cond_time}"
            ])
        else:
            script_parts.extend([f"set_e {begin}", "cell_on"])
        
        cv_command = f"meas_loop_cv p c {begin} {v1} {v2} {step} {scan_rate}"
        if int(n_scans) > 1:
            cv_command += f" nscans({n_scans})"
        
        script_parts.extend([
            "# CV measurement loop",
            cv_command, "\tpck_start", "\tpck_add p", "\tpck_add c", "\tpck_end",
            "endloop", "on_finished:", "cell_off"
        ])
        
        return "\n".join(script_parts)

    def create_swv_methodscript(self):
        """Create MethodSCRIPT for SWV with correct SI unit formatting"""
        begin_v = float(self.swv_params['begin_potential'].get())
        end_v = float(self.swv_params['end_potential'].get())
        amp_v = float(self.swv_params['amplitude'].get())

        begin = to_si_string(self.swv_params['begin_potential'].get(), 'V')
        end = to_si_string(self.swv_params['end_potential'].get(), 'V')
        step = to_si_string(self.swv_params['step_potential'].get(), 'V')
        amplitude = to_si_string(self.swv_params['amplitude'].get(), 'V')
        frequency = to_si_string(self.swv_params['frequency'].get(), 'Hz')
        cond_pot = to_si_string(self.swv_params['cond_potential'].get(), 'V')
        cond_time = self.swv_params['cond_time'].get()
        min_pot = min(begin_v, end_v) - amp_v
        max_pot = max(begin_v, end_v) + amp_v
        min_pot_mv, max_pot_mv = int(min_pot * 1000), int(max_pot * 1000)

        script_parts = [
            "e", "var c", "var p", "var f", "var r","set_pgstat_mode 2", "set_max_bandwidth 1600",
            f"set_range_minmax da {min_pot_mv}m {max_pot_mv}m",
            "set_range ba 5m", "set_autoranging ba 100n 5m", "cell_on"
        ]
        
        if float(cond_time) > 0:
            script_parts.extend([
                f"# Equilibrate at {cond_pot} for {cond_time}s",
                f"set_e {cond_pot}", f"wait {cond_time}"
            ])
            
        script_parts.extend([
            f"meas_loop_swv p c f r {begin} {end} {step} {amplitude} {frequency}",
            "\tpck_start", "\t\tpck_add p", "\t\tpck_add c", "\t\tpck_add f", "\t\tpck_add r",
            "\tpck_end", "endloop",
        ])

        script_parts.extend([
            "on_finished:", "cell_off"
        ])

        return "\n".join(script_parts)

    def _get_mux_channels(self, params):
        entry = params.get('mux_channel')
        if entry is None:
            return []
        raw = entry.get().strip()
        if raw in ("", "0"):
            return []

        channels = []
        seen = set()
        parts = [p.strip() for p in raw.split(",") if p.strip()]
        if not parts:
            messagebox.showerror("Invalid MUX Channel", "Enter channels like 1-3,7-9 or 0 to disable.")
            return None

        for part in parts:
            if "-" in part:
                start_s, end_s = part.split("-", 1)
                if not start_s.strip() or not end_s.strip():
                    messagebox.showerror("Invalid MUX Channel", f"Invalid range: '{part}'")
                    return None
                try:
                    start = int(start_s)
                    end = int(end_s)
                except Exception:
                    messagebox.showerror("Invalid MUX Channel", f"Invalid range: '{part}'")
                    return None
                if start > end:
                    messagebox.showerror("Invalid MUX Channel", f"Invalid range: '{part}'")
                    return None
                for ch in range(start, end + 1):
                    if not (1 <= ch <= 16):
                        messagebox.showerror("Invalid MUX Channel", "MUX16 channels must be between 1 and 16.")
                        return None
                    if ch not in seen:
                        seen.add(ch)
                        channels.append(ch)
            else:
                try:
                    ch = int(part)
                except Exception:
                    messagebox.showerror("Invalid MUX Channel", f"Invalid channel: '{part}'")
                    return None
                if not (1 <= ch <= 16):
                    messagebox.showerror("Invalid MUX Channel", "MUX16 channels must be between 1 and 16.")
                    return None
                if ch not in seen:
                    seen.add(ch)
                    channels.append(ch)

        return channels

    @staticmethod
    def _mux_channel_address(channel: int) -> int:
        idx = channel - 1
        return (idx << 4) | idx

    def _build_mux_script(self, base_script: str, channel: int) -> str:
        lines = base_script.splitlines()
        header = "e"
        rest = lines
        if lines and lines[0].strip() in ("e", "l"):
            header = lines[0].strip()
            rest = lines[1:]

        addr = self._mux_channel_address(channel)
        prefix = [
            header,
            "# MUX16 channel select",
            "set_gpio_cfg 0x3FFi 1",
            f"set_gpio {addr}i",
        ]
        merged = prefix + rest
        return "\n".join(merged)

    def _get_swv_cycles_and_delay(self):
        try:
            n_scans = int(self.swv_params['n_scans'].get())
        except Exception:
            messagebox.showerror("Invalid SWV Scans", "Number of scans must be an integer.")
            return None, None
        try:
            delay = float(self.swv_params.get('cycle_delay').get())
        except Exception:
            messagebox.showerror("Invalid SWV Delay", "Delay between scans must be a number.")
            return None, None
        if n_scans < 1:
            messagebox.showerror("Invalid SWV Scans", "Number of scans must be at least 1.")
            return None, None
        if delay < 0:
            messagebox.showerror("Invalid SWV Delay", "Delay between scans must be non-negative.")
            return None, None
        return n_scans, delay
    
    def setup_method_tab(self):
        left_frame = ttk.Frame(self.method_frame)
        left_frame.pack(side='left', fill='both', expand=True, padx=5)
        ttk.Label(left_frame, text="Select Technique:", font=('Arial', 12, 'bold')).pack(pady=5)
        technique_frame = ttk.Frame(left_frame)
        technique_frame.pack(pady=10)
        ttk.Button(technique_frame, text="Cyclic Voltammetry (CV)", command=self.show_cv_params, width=25).pack(pady=5)
        ttk.Button(technique_frame, text="Square Wave Voltammetry (SWV)", command=self.show_swv_params, width=25).pack(pady=5)
        ttk.Separator(technique_frame, orient='horizontal').pack(fill='x', pady=6)
        ttk.Button(technique_frame, text="Pause", command=self.show_pause_params, width=25).pack(pady=5)
        self.device_status = ttk.Label(left_frame, text="", foreground="blue")
        self.device_status.pack(pady=10)
        ttk.Button(left_frame, text="Check Device Connection", command=self.check_device).pack(pady=5)
        self.params_frame = ttk.LabelFrame(self.method_frame, text="Parameters", padding=10)
        self.params_frame.pack(side='right', fill='both', expand=True, padx=5)
        self.show_cv_params()

    def check_device(self):
        ports = list(serial.tools.list_ports.comports())
        if ports:
            self.device_status.config(text="Devices found (check console)", foreground="green")
            print("Available serial devices:\n" + "\n".join([f"{p.device}: {p.description}" for p in ports]))
        else: self.device_status.config(text="No devices found", foreground="red")
            
    def show_cv_params(self):
        self.clear_params_frame()
        self.current_technique = "CV"
        self.cv_params = {}
        params = [("Begin Potential (V):", "begin_potential", "0"), ("Vertex 1 (V):", "vertex1", "-0.5"), ("Vertex 2 (V):", "vertex2", "0.5"), ("Step Potential (V):", "step_potential", "0.002"), ("Scan Rate (V/s):", "scan_rate", "0.1"), ("Number of Scans:", "n_scans", "1"), ("Conditioning Potential (V):", "cond_potential", "0"), ("Conditioning Time (s):", "cond_time", "0"), ("MUX16 Channels (1-16, 0=off, e.g. 1-3,7-9):", "mux_channel", "0")]
        for i, (label, key, default) in enumerate(params):
            ttk.Label(self.params_frame, text=label).grid(row=i, column=0, sticky='w', pady=2)
            entry = ttk.Entry(self.params_frame, width=15); entry.insert(0, default); entry.grid(row=i, column=1, pady=2)
            self.cv_params[key] = entry
        button_frame = ttk.Frame(self.params_frame); button_frame.grid(row=len(params), column=0, columnspan=2, pady=20)
        ttk.Button(button_frame, text="Generate Script", command=self.generate_cv_script).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Run Now", command=self.run_cv_immediately).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Add to Queue", command=self.add_cv_to_queue).pack(side='left', padx=5)
        
    def show_swv_params(self):
        self.clear_params_frame()
        self.current_technique = "SWV"
        self.swv_params = {}
        params = [("Begin Potential (V):", "begin_potential", "-0.5"), ("End Potential (V):", "end_potential", "0.5"), ("Step Potential (V):", "step_potential", "0.002"), ("Amplitude (V):", "amplitude", "0.02"), ("Frequency (Hz):", "frequency", "15"), ("Number of Scans:", "n_scans", "1"), ("Delay Between Scans (s):", "cycle_delay", "0"), ("Conditioning Potential (V):", "cond_potential", "0"), ("Conditioning Time (s):", "cond_time", "0"), ("MUX16 Channels (1-16, 0=off, e.g. 1-3,7-9):", "mux_channel", "0")]
        for i, (label, key, default) in enumerate(params):
            ttk.Label(self.params_frame, text=label).grid(row=i, column=0, sticky='w', pady=2)
            entry = ttk.Entry(self.params_frame, width=15); entry.insert(0, default); entry.grid(row=i, column=1, pady=2)
            self.swv_params[key] = entry
        button_frame = ttk.Frame(self.params_frame); button_frame.grid(row=len(params), column=0, columnspan=2, pady=20)
        ttk.Button(button_frame, text="Generate Script", command=self.generate_swv_script).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Run Now", command=self.run_swv_immediately).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Add to Queue", command=self.add_swv_to_queue).pack(side='left', padx=5)
        
    def show_pause_params(self):
        self.clear_params_frame()
        self.current_technique = "PAUSE"
        self.pause_params = {}
        ttk.Label(self.params_frame, text="Pause Time (sec):").grid(row=0, column=0, sticky='w', pady=2)
        entry = ttk.Entry(self.params_frame, width=15)
        entry.insert(0, "10")
        entry.grid(row=0, column=1, pady=2)
        self.pause_params['pause_time'] = entry
        ttk.Label(self.params_frame, text="Alert Message:").grid(row=1, column=0, sticky='w', pady=2)
        alert_entry = ttk.Entry(self.params_frame, width=30)
        alert_entry.insert(0, "Paused - click OK to continue.")
        alert_entry.grid(row=1, column=1, pady=2, sticky='w')
        self.pause_params['alert_message'] = alert_entry

        button_frame = ttk.Frame(self.params_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=20)
        ttk.Button(button_frame, text="Add Pause to Queue", command=self.add_pause_to_queue).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Run Pause Now", command=self.run_pause_immediately).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Add Alert Pause", command=self.add_alert_pause_to_queue).pack(side='left', padx=5)

    def setup_pump_tab(self): #TODO: pump autoconnect on start
        if not PUMP_AVAILABLE or PumpCtrl is None:
            ttk.Label(self.pump_frame, text="Pump controls unavailable.").pack(pady=20)
            return

        pad = {"padx": 6, "pady": 4}

        container = ttk.Frame(self.pump_frame)
        container.pack(fill='both', expand=True, padx=10, pady=10)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(3, weight=1)

        self.pump_ctrl = PumpCtrl(use_sim=(not PUMP_HAS_COM), log_cb=self.pump_log)
        self.pump_ctrl.configure_calibration(PREFERRED_STEPS_PER_STROKE, PREFERRED_SYRINGE_UL)

        # Connection section
        conn_frame = ttk.LabelFrame(container, text="Connection")
        conn_frame.grid(row=0, column=0, sticky='ew')

        self.pump_var_sim = tk.BooleanVar(value=self.pump_ctrl.use_sim)
        self.pump_chk_sim = ttk.Checkbutton(conn_frame, text="Simulate (no hardware)", variable=self.pump_var_sim)
        self.pump_chk_sim.grid(row=0, column=0, columnspan=2, **pad, sticky='w')
        if not PUMP_HAS_COM:
            self.pump_chk_sim.configure(state='disabled')
            self.pump_log("COM not available; defaulting to SIM mode.")

        ttk.Label(conn_frame, text="COM port:").grid(row=1, column=0, **pad, sticky='e')
        self.pump_var_com = tk.IntVar(value=int(PUMP_DEFAULT_COM_PORT))
        self.pump_spin_com = ttk.Spinbox(conn_frame, from_=1, to=60, width=6, textvariable=self.pump_var_com)
        self.pump_spin_com.grid(row=1, column=1, **pad)

        ttk.Label(conn_frame, text="Baud:").grid(row=1, column=2, **pad, sticky='e')
        self.pump_var_baud = tk.StringVar(value=str(PUMP_DEFAULT_BAUD))
        self.pump_combo_baud = ttk.Combobox(conn_frame, values=['9600', '38400'], width=8, textvariable=self.pump_var_baud)
        self.pump_combo_baud.grid(row=1, column=3, **pad)
        self.pump_combo_baud.set(str(PUMP_DEFAULT_BAUD))

        ttk.Label(conn_frame, text="Device #:").grid(row=1, column=4, **pad, sticky='e')
        self.pump_var_dev = tk.IntVar(value=int(PUMP_DEFAULT_DEV))
        self.pump_spin_dev = ttk.Spinbox(conn_frame, from_=0, to=30, width=6, textvariable=self.pump_var_dev)
        self.pump_spin_dev.grid(row=1, column=5, **pad)

        self.pump_btn_connect = ttk.Button(
            conn_frame,
            text="Connect",
            command=lambda: self._pump_launch_with_validation(
                self.pump_on_connect,
                lambda: bool(self.pump_var_sim.get()),
                lambda: int(self.pump_var_com.get()),
                lambda: int(self.pump_var_baud.get()),
                lambda: int(self.pump_var_dev.get()),
            ),
        )
        self.pump_btn_connect.grid(row=2, column=0, columnspan=2, **pad)

        self.pump_btn_disconnect = ttk.Button(
            conn_frame,
            text="Disconnect",
            command=lambda: self.pump_threaded(self.pump_on_disconnect),
        )
        self.pump_btn_disconnect.grid(row=2, column=2, columnspan=2, **pad)

        # Calibration section
        cal_frame = ttk.LabelFrame(container, text="Calibration (\u00B5L \u2194 steps)")
        cal_frame.grid(row=1, column=0, sticky='ew', pady=(10, 0))

        ttk.Label(cal_frame, text="Steps/stroke:").grid(row=0, column=0, **pad, sticky='e')
        self.pump_var_steps = tk.IntVar(value=int(PUMP_DEFAULT_STEPS))
        self.pump_entry_steps = ttk.Entry(cal_frame, width=10, textvariable=self.pump_var_steps)
        self.pump_entry_steps.grid(row=0, column=1, **pad)

        ttk.Label(cal_frame, text="Syringe (\u00B5L):").grid(row=0, column=2, **pad, sticky='e')
        self.pump_var_syringe = tk.DoubleVar(value=float(PUMP_DEFAULT_SYRINGE))
        self.pump_entry_syringe = ttk.Entry(cal_frame, width=10, textvariable=self.pump_var_syringe)
        self.pump_entry_syringe.grid(row=0, column=3, **pad)

        self.pump_btn_apply_cal = ttk.Button(
            cal_frame,
            text="Apply",
            command=lambda: self._pump_launch_with_validation(
                self.pump_on_apply_cal,
                lambda: int(self.pump_var_steps.get()),
                lambda: float(self.pump_var_syringe.get()),
            ),
        )
        self.pump_btn_apply_cal.grid(row=0, column=4, **pad)

        # Actions section
        actions_frame = ttk.LabelFrame(container, text="Actions")
        actions_frame.grid(row=2, column=0, sticky='ew', pady=(10, 0))

        self.pump_btn_init = ttk.Button(
            actions_frame,
            text="Initialize (ZR)",
            command=lambda: self.pump_threaded(self.pump_do_init),
        )
        self.pump_btn_init.grid(row=0, column=0, **pad)

        self.pump_btn_queue_init = ttk.Button(
            actions_frame,
            text="Queue Init",
            command=self.queue_pump_init,
        )
        self.pump_btn_queue_init.grid(row=0, column=1, **pad)

        ttk.Label(actions_frame, text="Volume (\u00B5L):").grid(row=0, column=2, **pad, sticky='e')
        self.pump_var_volume = tk.DoubleVar(value=50.0)
        self.pump_entry_volume = ttk.Entry(actions_frame, width=10, textvariable=self.pump_var_volume)
        self.pump_entry_volume.grid(row=0, column=3, **pad)

        ttk.Label(actions_frame, text="Plunger speed (SnnR):").grid(row=0, column=4, **pad, sticky='e')
        self.pump_var_speed = tk.IntVar(value=20)
        self.pump_spin_speed = ttk.Spinbox(
            actions_frame,
            from_=PUMP_SPEED_MIN if PUMP_AVAILABLE else 1,
            to=PUMP_SPEED_MAX if PUMP_AVAILABLE else 40,
            width=6,
            textvariable=self.pump_var_speed,
        )
        self.pump_spin_speed.grid(row=0, column=5, **pad)

        self.pump_btn_set_speed = ttk.Button(
            actions_frame,
            text="Set Speed",
            command=lambda: self._pump_launch_with_validation(
                self.pump_do_set_speed,
                lambda: int(self.pump_var_speed.get()),
            ),
        )
        self.pump_btn_set_speed.grid(row=0, column=6, **pad)

        self.pump_btn_queue_set_speed = ttk.Button(
            actions_frame,
            text="Queue Speed",
            command=self.queue_pump_set_speed,
        )
        self.pump_btn_queue_set_speed.grid(row=0, column=7, **pad)

        ttk.Label(actions_frame, text="Valve port:").grid(row=1, column=0, **pad, sticky='e')
        self.pump_var_valve = tk.IntVar(value=1)
        self.pump_spin_valve = ttk.Spinbox(actions_frame, from_=1, to=9, width=6, textvariable=self.pump_var_valve)
        self.pump_spin_valve.grid(row=1, column=1, **pad)

        self.pump_btn_valve = ttk.Button(
            actions_frame,
            text="Move Valve (I#R)",
            command=lambda: self._pump_launch_with_validation(
                self.pump_do_valve,
                lambda: int(self.pump_var_valve.get()),
            ),
        )
        self.pump_btn_valve.grid(row=1, column=2, **pad)

        self.pump_btn_queue_valve = ttk.Button(
            actions_frame,
            text="Queue Valve",
            command=self.queue_pump_valve,
        )
        self.pump_btn_queue_valve.grid(row=1, column=3, **pad)

        valve_quick = ttk.LabelFrame(actions_frame, text="Valve quick")
        valve_quick.grid(row=2, column=0, columnspan=8, padx=6, pady=(6, 2))

        self.pump_valve_buttons = []
        for i in range(1, 10):
            btn = ttk.Button(
                valve_quick,
                text=str(i),
                width=3,
                command=lambda p=i: self.pump_threaded(self.pump_do_valve_num, p),
            )
            btn.grid(row=(i - 1) // 5, column=(i - 1) % 5, padx=3, pady=3)
            self.pump_valve_buttons.append(btn)

        self.pump_btn_aspirate = ttk.Button(
            actions_frame,
            text="Aspirate",
            command=lambda: self._pump_launch_with_validation(
                self.pump_do_aspirate,
                lambda: float(self.pump_var_volume.get()),
                lambda: int(self.pump_var_speed.get()),
            ),
        )
        self.pump_btn_aspirate.grid(row=3, column=2, **pad)

        self.pump_btn_queue_aspirate = ttk.Button(
            actions_frame,
            text="Queue Aspirate",
            command=self.queue_pump_aspirate,
        )
        self.pump_btn_queue_aspirate.grid(row=3, column=3, **pad)

        self.pump_btn_dispense = ttk.Button(
            actions_frame,
            text="Dispense",
            command=lambda: self._pump_launch_with_validation(
                self.pump_do_dispense,
                lambda: float(self.pump_var_volume.get()),
                lambda: int(self.pump_var_speed.get()),
            ),
        )
        self.pump_btn_dispense.grid(row=3, column=4, **pad)

        self.pump_btn_queue_dispense = ttk.Button(
            actions_frame,
            text="Queue Dispense",
            command=self.queue_pump_dispense,
        )
        self.pump_btn_queue_dispense.grid(row=3, column=5, **pad)

        # Log section
        log_frame = ttk.LabelFrame(container, text="Log")
        log_frame.grid(row=3, column=0, sticky='nsew', pady=(10, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.pump_log_text = tk.Text(log_frame, height=10, state='disabled')
        self.pump_log_text.grid(row=0, column=0, sticky='nsew', padx=6, pady=6)

        self._pump_flush_early_logs()

        self.pump_disable_widgets = [
            self.pump_btn_connect,
            self.pump_btn_disconnect,
            self.pump_btn_init,
            self.pump_btn_aspirate,
            self.pump_btn_dispense,
            self.pump_btn_valve,
            self.pump_btn_apply_cal,
            self.pump_spin_com,
            self.pump_combo_baud,
            self.pump_spin_dev,
            self.pump_entry_steps,
            self.pump_entry_syringe,
            self.pump_entry_volume,
            self.pump_spin_valve,
            self.pump_btn_set_speed,
            self.pump_spin_speed,
            *self.pump_valve_buttons,
        ]
        self.root.after(200, self._pump_auto_connect)

    def add_pump_action_to_queue(self, action_name: str, *, params=None, details: str):
        if not PUMP_AVAILABLE or self.pump_ctrl is None:
            messagebox.showerror("Pump Error", "Pump backend unavailable.")
            return
        item = {
            'type': f'PUMP_{action_name}',
            'status': 'pending',
            'details': details,
            'pump_action': {
                'name': action_name,
                'params': params or {},
            },
        }
        self.measurement_queue.append(item)
        self.refresh_queue_display()
        messagebox.showinfo("Added to Queue", details)

    def queue_pump_init(self):
        self.add_pump_action_to_queue('INIT', details='Pump: Initialize (ZR)')

    def queue_pump_set_speed(self):
        try:
            speed = int(self.pump_var_speed.get())
        except (ValueError, tk.TclError) as exc:
            messagebox.showerror("Invalid speed", str(exc))
            return
        details = f'Pump: Set speed S{speed}R'
        self.add_pump_action_to_queue('SET_SPEED', params={'speed': speed}, details=details)

    def queue_pump_valve(self):
        try:
            port = int(self.pump_var_valve.get())
        except (ValueError, tk.TclError) as exc:
            messagebox.showerror("Invalid valve port", str(exc))
            return
        details = f'Pump: Valve -> {port}'
        self.add_pump_action_to_queue('VALVE', params={'port': port}, details=details)

    def queue_pump_aspirate(self):
        try:
            volume = float(self.pump_var_volume.get())
            speed = int(self.pump_var_speed.get())
        except (ValueError, tk.TclError) as exc:
            messagebox.showerror("Invalid aspirate parameters", str(exc))
            return
        if PUMP_AVAILABLE and self.pump_ctrl is not None:
            pending_steps = self._pump_pending_plunger_steps()
            projected_steps = pending_steps + self.pump_ctrl.steps_for_volume(volume)
            if projected_steps > self.pump_ctrl.steps_per_stroke:
                remaining_ul = self._pump_remaining_capacity_for_queue()
                messagebox.showerror(
                    "Pump Capacity Exceeded",
                    f"Queued aspirations would exceed the syringe capacity. Remaining capacity: {remaining_ul:.2f} µL."
                )
                return
        details = f'Pump: Aspirate {volume:.2f} µL @ S{speed}R'
        self.add_pump_action_to_queue('ASPIRATE', params={'volume': volume, 'speed': speed}, details=details)

    def queue_pump_dispense(self):
        try:
            volume = float(self.pump_var_volume.get())
            speed = int(self.pump_var_speed.get())
        except (ValueError, tk.TclError) as exc:
            messagebox.showerror("Invalid dispense parameters", str(exc))
            return
        details = f'Pump: Dispense {volume:.2f} µL @ S{speed}R'
        self.add_pump_action_to_queue('DISPENSE', params={'volume': volume, 'speed': speed}, details=details)

    def _pump_launch_with_validation(self, target, *value_factories):
        if not PUMP_AVAILABLE or self.pump_ctrl is None:
            messagebox.showerror("Pump Error", "Pump backend unavailable.")
            return
        try:
            values = [factory() for factory in value_factories]
        except (ValueError, tk.TclError) as exc:
            messagebox.showerror("Invalid value", str(exc))
            return
        self.pump_threaded(target, *values)

    def pump_log(self, message):
        if not PUMP_AVAILABLE:
            return
        if self.pump_log_text is None:
            self.pump_early_logs.append(message)
            return

        def append():
            self.pump_log_text.configure(state='normal')
            self.pump_log_text.insert('end', message + '\n')
            self.pump_log_text.see('end')
            self.pump_log_text.configure(state='disabled')

        self.root.after(0, append)

    def _pump_flush_early_logs(self):
        if self.pump_log_text is None or not self.pump_early_logs:
            return
        pending = self.pump_early_logs[:]
        self.pump_early_logs.clear()

        def flush():
            self.pump_log_text.configure(state='normal')
            for msg in pending:
                self.pump_log_text.insert('end', msg + '\n')
            self.pump_log_text.see('end')
            self.pump_log_text.configure(state='disabled')

        self.root.after(0, flush)

    def set_pump_busy(self, busy: bool):
        self.pump_busy = busy
        if not self.pump_disable_widgets:
            return
        state = 'disabled' if busy else 'normal'

        def apply_state():
            for widget in self.pump_disable_widgets:
                try:
                    widget.configure(state=state)
                except tk.TclError:
                    pass

        self.root.after(0, apply_state)

    def _pump_pending_plunger_steps(self) -> int:
        if not PUMP_AVAILABLE or self.pump_ctrl is None:
            return 0
        steps = self.pump_ctrl.plunger_steps
        for item in self.measurement_queue:
            status = item.get('status')
            if status not in (None, 'pending', 'in_progress'):
                continue
            action_info = item.get('pump_action') or {}
            action = action_info.get('name')
            params = action_info.get('params') or {}
            if not action:
                continue
            if action == 'INIT':
                steps = 0
                continue
            if action not in ('ASPIRATE', 'DISPENSE'):
                continue
            try:
                volume = float(params.get('volume', 0.0))
            except (TypeError, ValueError):
                volume = 0.0
            delta_steps = self.pump_ctrl.steps_for_volume(volume)
            if action == 'ASPIRATE':
                steps = min(self.pump_ctrl.steps_per_stroke, steps + delta_steps)
            else:
                steps = max(0, steps - delta_steps)
        return steps

    def _pump_remaining_capacity_for_queue(self) -> float:
        if not PUMP_AVAILABLE or self.pump_ctrl is None:
            return 0.0
        pending_steps = self._pump_pending_plunger_steps()
        remaining_steps = max(0, self.pump_ctrl.steps_per_stroke - pending_steps)
        return self.pump_ctrl.volume_for_steps(remaining_steps)

    def _pump_auto_connect(self):
        if not PUMP_AVAILABLE or self.pump_ctrl is None:
            return
        if getattr(self, '_pump_auto_connect_attempted', False):
            return
        self._pump_auto_connect_attempted = True
        try:
            sim_mode = bool(self.pump_var_sim.get())
            com_port = int(self.pump_var_com.get())
            baud = int(self.pump_var_baud.get())
            dev = int(self.pump_var_dev.get())
        except (ValueError, tk.TclError):
            self.pump_log('Auto-connect skipped: invalid connection parameters.')
            return
        if self.pump_ctrl.connected:
            return
        self.pump_log('Auto-connecting to pump...')
        self.pump_threaded(self.pump_on_connect, sim_mode, com_port, baud, dev)

    def pump_threaded(self, fn, *args):
        if not PUMP_AVAILABLE or self.pump_ctrl is None:
            messagebox.showerror("Pump Error", "Pump backend unavailable.")
            return
        if self.pump_busy:
            return

        sim_mode = bool(self.pump_var_sim.get()) if hasattr(self, 'pump_var_sim') else True

        def run():
            com_required = bool(PUMP_HAS_COM and pythoncom and not sim_mode)
            if com_required:
                try:
                    pythoncom.CoInitialize()
                except Exception:
                    com_required = False
            try:
                self.set_pump_busy(True)
                fn(*args)
            except Exception as exc:
                self.pump_log(f"ERROR: {exc}")
                self.root.after(0, lambda: messagebox.showerror("Pump Error", str(exc)))
            finally:
                self.set_pump_busy(False)
                if com_required:
                    try:
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass

        threading.Thread(target=run, daemon=True).start()

    def pump_on_connect(self, sim_mode: bool, com_port: int, baud: int, dev: int):
        if self.pump_ctrl is None:
            return
        try:
            self.pump_ctrl.use_sim = bool(sim_mode) or (not PUMP_HAS_COM)
            mode = "[SIM]" if self.pump_ctrl.use_sim else "[REAL]"
            self.pump_log(f"{mode} Connecting…")
            self.pump_ctrl.connect(com_port, baud, dev)
            if self.pump_ctrl.use_sim:
                self.pump_log("Sim mode ready: speed/valve/A-D behave with realistic timing.")
            else:
                self.pump_log("Real mode ready. Tip: set plunger speed (SnnR) before A/D.")
        except Exception as exc:
            self.root.after(0, lambda: messagebox.showerror("Connect failed", str(exc)))
            self.pump_log(f"Connect failed: {exc}")

    def pump_on_disconnect(self):
        if self.pump_ctrl is None:
            return
        try:
            self.pump_ctrl.disconnect()
        except Exception as exc:
            self.pump_log(f"Disconnect error: {exc}")

    def pump_on_apply_cal(self, steps: int, syringe_ul: float):
        if self.pump_ctrl is None:
            return
        try:
            self.pump_ctrl.configure_calibration(int(steps), float(syringe_ul))
            self.pump_log(
                f"Applied: steps/stroke={self.pump_ctrl.steps_per_stroke}, syringe={self.pump_ctrl.syringe_ul:.0f} µL"
            )
        except Exception as exc:
            self.root.after(0, lambda: messagebox.showerror("Invalid calibration", str(exc)))

    def _pump_require_connection(self) -> bool:
        if not self.pump_ctrl or not self.pump_ctrl.connected:
            self.pump_log("Not connected.")
            return False
        return True

    def pump_do_init(self):
        if not self._pump_require_connection():
            return
        self.pump_log("Initialize (ZR)…")
        self.pump_ctrl.initialize()
        self.pump_log("Init done.")

    def pump_do_set_speed(self, speed: int):
        if not self._pump_require_connection():
            return
        speed = int(speed)
        self.pump_log(f"Set plunger speed: S{speed}R")
        self.pump_ctrl.set_speed(speed)

    def pump_do_valve(self, port: int):
        if not self._pump_require_connection():
            return
        port = int(port)
        self.pump_log(f"Valve -> {port} (I{port}R)")
        self.pump_ctrl.valve_to(port)
        self.pump_log("Valve move done.")

    def pump_do_valve_num(self, port: int):
        self.pump_do_valve(port)

    def pump_do_aspirate(self, volume_ul: float, speed: int):
        if not self._pump_require_connection():
            return
        volume = max(0.0, float(volume_ul))
        speed_val = int(speed)
        self.pump_log(f"Aspirate {volume:.2f} \u00B5L @ S{speed_val}R")
        self.pump_ctrl.set_speed(speed_val)
        self.pump_ctrl.aspirate_ul(volume)
        self.pump_log("Aspirate done.")

    def pump_do_dispense(self, volume_ul: float, speed: int):
        if not self._pump_require_connection():
            return
        volume = max(0.0, float(volume_ul))
        speed_val = int(speed)
        self.pump_log(f"Dispense  {volume:.2f} \u00B5L @ S{speed_val}R")
        self.pump_ctrl.set_speed(speed_val)
        self.pump_ctrl.dispense_ul(volume)
        self.pump_log("Dispense done.")

    def setup_queue_tab(self):
        main_pane = ttk.PanedWindow(self.queue_frame, orient=tk.VERTICAL); main_pane.pack(fill='both', expand=True)
        top_frame = ttk.Frame(main_pane); main_pane.add(top_frame, weight=1)
        bottom_frame = ttk.Frame(main_pane); main_pane.add(bottom_frame, weight=1)
        control_frame = ttk.Frame(top_frame); control_frame.pack(pady=10, fill='x', padx=10)
        ttk.Button(control_frame, text="Run Queue", command=self.run_queue).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Run From Selected", command=self.run_queue_from_selected).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Stop", command=self.stop_queue).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Save Queue", command=self.save_queue).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Load Queue", command=self.load_queue).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Clear Queue", command=self.clear_queue).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Delete Selected", command=self.delete_selected_queue_item).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Confirm Move", command=self.confirm_queue_reorder).pack(side='left', padx=5)
        self.queue_tree = ttk.Treeview(top_frame, columns=('Type', 'Status', 'Details'), show='tree headings', height=8)
        self.queue_tree.heading('#0', text='#'); self.queue_tree.heading('Type', text='Type'); self.queue_tree.heading('Status', text='Status'); self.queue_tree.heading('Details', text='Details')
        self.queue_tree.column('#0', width=50); self.queue_tree.column('Type', width=150); self.queue_tree.column('Status', width=100); self.queue_tree.column('Details', width=400)
        self.queue_tree.pack(fill='both', expand=True, padx=10, pady=5)
        self.queue_tree.bind("<ButtonPress-1>", self._on_queue_drag_start)
        self.queue_tree.bind("<B1-Motion>", self._on_queue_drag_motion)
        self.queue_tree.bind("<ButtonRelease-1>", self._on_queue_drag_release)
        log_frame = ttk.LabelFrame(bottom_frame, text="Live Output Log"); log_frame.pack(fill='both', expand=True, padx=10, pady=5)
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10); self.log_text.pack(fill='both', expand=True)
        self.log_text.config(state='disabled')
        self.status_label = ttk.Label(self.queue_frame, text="Status: Ready", relief='sunken'); self.status_label.pack(side='bottom', fill='x', padx=10, pady=5)

    def clear_log(self):
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')

    def log_message(self, message):
        def append_message():
            self.log_text.config(state='normal')
            self.log_text.insert(tk.END, message + '\n')
            self.log_text.see(tk.END)
            self.log_text.config(state='disabled')
        self.root.after(0, append_message)
        print(message)

    def setup_script_tab(self):
        text_frame = ttk.Frame(self.script_frame); text_frame.pack(fill='both', expand=True, padx=10, pady=5)
        self.script_text = tk.Text(text_frame, wrap='none', font=('Courier', 11)); self.script_text.pack(fill='both', expand=True)
    
    def setup_plotter_tab(self):
        """Sets up the new tab for plotting data."""
        plot_controls = ttk.Frame(self.plotter_frame)
        plot_controls.pack(side='top', fill='x', pady=5, padx=5)
        
        ttk.Button(plot_controls, text="Load and Plot CSV", command=self.load_and_plot_csv).pack(side='left')
        ttk.Button(plot_controls, text="Clear Plot", command=self.clear_plot).pack(side='left', padx=5)

        # Create a Matplotlib figure and axis
        self.fig = Figure(figsize=(8, 6), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.ax.set_title('Voltammogram')
        self.ax.set_xlabel('Potential (V)')
        self.ax.set_ylabel('Current (uA)')
        self.ax.grid(True)

        # Create a canvas to embed the plot in Tkinter
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.plotter_frame)
        self.canvas.draw()
        toolbar_frame = ttk.Frame(self.plotter_frame)
        toolbar_frame.pack(side='top', fill='x')
        self.toolbar = NavigationToolbar2Tk(self.canvas, toolbar_frame)
        self.toolbar.update()

        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

    def load_and_plot_csv(self):
        """Opens a file dialog to select a CSV and plots it."""
        filepath = filedialog.askopenfilename(
            title="Select a measurement CSV",
            filetypes=(("CSV files", "*.csv"), ("All files", "*.*"))
        )
        if filepath:
            self.plot_data(filepath)

    def _read_csv_with_fallback(self, csv_path):
        encodings_to_try = ("utf-8-sig", "utf-8", "cp1252", "latin-1")
        last_error = None
        for encoding in encodings_to_try:
            try:
                return pd.read_csv(csv_path, encoding=encoding)
            except UnicodeDecodeError as exc:
                last_error = exc
        with open(csv_path, "r", encoding="utf-8", errors="replace") as handle:
            data = handle.read()
        try:
            return pd.read_csv(io.StringIO(data))
        except Exception as exc:
            if last_error is not None:
                raise last_error from exc
            raise

    @staticmethod
    def _normalize_header(header: str) -> str:
        normalized = header.strip().lower()
        normalized = normalized.replace("\u03BC", "\u00B5")
        normalized = normalized.replace("\u00B5", "mu")
        normalized = normalized.replace("\uFFFD", "mu")
        return normalized

    def _find_column(self, df, candidates):
        for candidate in candidates:
            if candidate in df.columns:
                return candidate
        normalized_map = {self._normalize_header(col): col for col in df.columns}
        for candidate in candidates:
            normalized_candidate = self._normalize_header(candidate)
            if normalized_candidate in normalized_map:
                return normalized_map[normalized_candidate]
        return None

    def plot_data(self, csv_path, color=None, label=None):
        """Reads a CSV file and plots the voltammogram."""
        try:
            df = self._read_csv_with_fallback(csv_path)
        except Exception as exc:
            self.log_message(f"Plot error: failed to read {csv_path}: {exc}")
            messagebox.showerror("Plot Error", f"Failed to read data: {exc}")
            self.update_status("Plot failed: see log for details")
            return

        potential_col = self._find_column(df, ("Potential (V)",))
        current_col = self._find_column(df, ("Current (uA)", "Current (µA)", "Current (�A)"))

        if not potential_col or not current_col:
            message = "Plot error: CSV file must contain 'Potential (V)' and 'Current (uA)' columns."
            self.log_message(message)
            messagebox.showerror("Plot Error", message)
            self.update_status("Plot failed: missing required columns")
            return

        try:
            if color is None:
                color = next(self.plot_color_cycle)
            if label is None:
                label = Path(csv_path).name
            if label:
                for line in list(self.ax.lines):
                    if line.get_label() == label:
                        line.remove()
            self.ax.plot(df[potential_col], df[current_col], color=color, label=label)
            self.ax.set_title('Voltammogram')
            self.ax.set_xlabel('Potential (V)')
            self.ax.set_ylabel('Current (uA)')
            self.ax.grid(visible=True, which='major', linestyle='-')
            self.ax.grid(visible=True, which='minor', linestyle='--', alpha=0.2)
            self.ax.minorticks_on()
            self.ax.legend(loc='best')
            self.canvas.draw()
            self.notebook.select(self.plotter_frame)
        except Exception as exc:
            self.log_message(f"Plot error: failed to render {csv_path}: {exc}")
            messagebox.showerror("Plot Error", f"Failed to render plot: {exc}")
            self.update_status("Plot failed: see log for details")

    def clear_plot(self):
        self.ax.clear()
        self.ax.set_title('Voltammogram')
        self.ax.set_xlabel('Potential (V)')
        self.ax.set_ylabel('Current (uA)')
        self.ax.grid(visible=True, which='major', linestyle='-')
        self.ax.grid(visible=True, which='minor', linestyle='--', alpha=0.2)
        self.ax.minorticks_on()
        self.plot_color_cycle = itertools.cycle(self.plot_colors)
        self.plot_line = None
        self.live_plot_x = []
        self.live_plot_y = []
        self.last_live_plot_color = None
        self.last_live_plot_label = None
        self.canvas.draw()

    def start_live_plot(self, title=None, color=None, label=None):
        self.live_plot_queue = queue.Queue(maxsize=10000)
        self.live_plot_x = []
        self.live_plot_y = []
        self.live_plot_active = True

        if color is None:
            color = next(self.plot_color_cycle)
        self.last_live_plot_color = color
        if label is None:
            label = title or "Live"
        self.last_live_plot_label = label
        self.ax.set_title(title or "Live Voltammogram")
        self.ax.set_xlabel('Potential (V)')
        self.ax.set_ylabel('Current (uA)')
        self.ax.grid(visible=True, which='major', linestyle='-')
        self.ax.grid(visible=True, which='minor', linestyle='--', alpha=0.2)
        self.ax.minorticks_on()
        (self.plot_line,) = self.ax.plot([], [], lw=1, color=color, label=label)
        self.canvas.draw()
        self.notebook.select(self.plotter_frame)

        if self.live_plot_job is None:
            self.live_plot_job = self.root.after(100, self._poll_live_plot_queue)

    def stop_live_plot(self):
        self.live_plot_active = False
        if self.live_plot_job is not None:
            self.root.after_cancel(self.live_plot_job)
            self.live_plot_job = None

    def queue_live_point(self, data_point):
        if not self.live_plot_active:
            return
        try:
            potential = data_point['potential']
            current = data_point['current']
        except KeyError:
            return
        try:
            self.live_plot_queue.put_nowait((potential, current))
        except queue.Full:
            return

    def _poll_live_plot_queue(self):
        if not self.live_plot_active:
            self.live_plot_job = None
            return

        updated = False
        while True:
            try:
                potential, current = self.live_plot_queue.get_nowait()
            except queue.Empty:
                break
            self.live_plot_x.append(potential)
            self.live_plot_y.append(current)
            updated = True

        if updated:
            if self.plot_line is None:
                (self.plot_line,) = self.ax.plot(self.live_plot_x, self.live_plot_y, lw=1)
            else:
                self.plot_line.set_data(self.live_plot_x, self.live_plot_y)
            self.ax.relim()
            self.ax.autoscale_view()
            if self.last_live_plot_label:
                self.ax.legend(loc='best')
            self.canvas.draw_idle()

        self.live_plot_job = self.root.after(100, self._poll_live_plot_queue)
    def clear_params_frame(self):
        for widget in self.params_frame.winfo_children(): widget.destroy()

    def generate_cv_script(self):
        try:
            base_script = self.create_cv_methodscript()
            mux_channels = self._get_mux_channels(self.cv_params)
            if mux_channels is None:
                return None
            script = base_script
            if mux_channels:
                script = self._build_mux_script(base_script, mux_channels[0])
                if len(mux_channels) > 1:
                    note = f"# NOTE: Multiple MUX16 channels selected ({', '.join(map(str, mux_channels))}). Preview shows channel {mux_channels[0]}."
                    script = f"{note}\n{script}"
            self.current_script = script
            self.update_script_preview(script)
            return script
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate script: {str(e)}")
            return None

    def generate_swv_script(self):
        try:
            base_script = self.create_swv_methodscript()
            mux_channels = self._get_mux_channels(self.swv_params)
            if mux_channels is None:
                return None
            script = base_script
            if mux_channels:
                script = self._build_mux_script(base_script, mux_channels[0])
                if len(mux_channels) > 1:
                    note = f"# NOTE: Multiple MUX16 channels selected ({', '.join(map(str, mux_channels))}). Preview shows channel {mux_channels[0]}."
                    script = f"{note}\n{script}"
            self.current_script = script
            self.update_script_preview(script)
            return script
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate script: {str(e)}")
            return None

    def update_script_preview(self, script):
        self.script_text.delete(1.0, tk.END)
        self.script_text.insert(1.0, script)
    
    def add_cv_to_queue(self):
        base_script = self.create_cv_methodscript()
        mux_channels = self._get_mux_channels(self.cv_params)
        if mux_channels is None:
            return
        if mux_channels:
            self.add_mux_scripts_to_queue("CV", base_script, mux_channels)
        else:
            self.add_to_queue("CV", base_script)
    
    def add_swv_to_queue(self):
        base_script = self.create_swv_methodscript()
        mux_channels = self._get_mux_channels(self.swv_params)
        if mux_channels is None:
            return
        n_scans, delay = self._get_swv_cycles_and_delay()
        if n_scans is None:
            return
        if n_scans <= 1:
            if mux_channels:
                self.add_mux_scripts_to_queue("SWV", base_script, mux_channels)
            else:
                self.add_to_queue("SWV", base_script)
            return

        added = []
        for cycle in range(1, n_scans + 1):
            if mux_channels:
                for ch in mux_channels:
                    mux_script = self._build_mux_script(base_script, ch)
                    filepath, filename = self.save_script_file("SWV", mux_script, mux_channel=ch)
                    queue_item = {
                        'type': "SWV",
                        'script_path': str(filepath),
                        'status': 'pending',
                        'details': f"{filename} (MUX ch {ch})",
                    }
                    self.measurement_queue.append(queue_item)
                    added.append(f"{filename} (ch {ch})")
            else:
                filepath, filename = self.save_script_file("SWV", base_script)
                queue_item = {'type': "SWV", 'script_path': str(filepath), 'status': 'pending', 'details': filename}
                self.measurement_queue.append(queue_item)
                added.append(filename)

            if delay > 0 and cycle < n_scans:
                pause_item = {
                    'type': 'PAUSE',
                    'status': 'pending',
                    'details': f'Pause for {delay:.1f} sec',
                    'pause_seconds': delay,
                }
                self.measurement_queue.append(pause_item)

        self.refresh_queue_display()
        messagebox.showinfo(
            "Success",
            f"SWV added for {n_scans} scan(s)\nSaved as: {', '.join(added)}",
        )

    def add_pause_to_queue(self):
        try:
            seconds = float(self.pause_params['pause_time'].get())
            if seconds < 0:
                raise ValueError("Pause time must be non-negative")
        except (KeyError, ValueError, tk.TclError) as exc:
            messagebox.showerror("Invalid Pause", str(exc))
            return

        queue_item = {
            'type': 'PAUSE',
            'status': 'pending',
            'details': f'Pause for {seconds:.1f} sec',
            'pause_seconds': seconds,
        }
        self.measurement_queue.append(queue_item)
        self.refresh_queue_display()
        messagebox.showinfo("Success", f"Pause ({seconds:.1f} sec) added to queue")

    def add_alert_pause_to_queue(self):
        message = (self.pause_params.get('alert_message').get() or "").strip()
        if not message:
            messagebox.showerror("Invalid Alert", "Alert message cannot be empty.")
            return
        queue_item = {
            'type': 'ALERT',
            'status': 'pending',
            'details': "Alert pause",
            'alert_message': message,
        }
        self.measurement_queue.append(queue_item)
        self.refresh_queue_display()
        messagebox.showinfo("Success", "Alert pause added to queue")

    def run_cv_immediately(self):
        base_script = self.create_cv_methodscript()
        mux_channels = self._get_mux_channels(self.cv_params)
        if mux_channels is None:
            return
        if mux_channels:
            if len(mux_channels) == 1:
                mux_script = self._build_mux_script(base_script, mux_channels[0])
                self.run_script_immediately("CV", mux_script, mux_channel=mux_channels[0])
            else:
                self.run_mux_script_sequence("CV", base_script, mux_channels)
        else:
            self.run_script_immediately("CV", base_script)

    def run_swv_immediately(self):
        base_script = self.create_swv_methodscript()
        mux_channels = self._get_mux_channels(self.swv_params)
        if mux_channels is None:
            return
        n_scans, delay = self._get_swv_cycles_and_delay()
        if n_scans is None:
            return
        if n_scans <= 1:
            if mux_channels:
                if len(mux_channels) == 1:
                    mux_script = self._build_mux_script(base_script, mux_channels[0])
                    self.run_script_immediately("SWV", mux_script, mux_channel=mux_channels[0])
                else:
                    self.run_mux_script_sequence("SWV", base_script, mux_channels)
            else:
                self.run_script_immediately("SWV", base_script)
            return

        if mux_channels:
            self.run_mux_swv_cycles(base_script, mux_channels, n_scans, delay)
        else:
            self.run_swv_cycles(base_script, n_scans, delay)

    def run_pause_immediately(self):
        try:
            seconds = float(self.pause_params['pause_time'].get())
            if seconds < 0:
                raise ValueError("Pause time must be non-negative")
        except (KeyError, ValueError, tk.TclError) as exc:
            messagebox.showerror("Invalid Pause", str(exc))
            return

        self.log_message(f"Pausing for {seconds:.1f} seconds...")

        def perform_pause():
            time.sleep(seconds)
            self.root.after(0, lambda: self.log_message(f"Pause completed ({seconds:.1f} sec)"))

        threading.Thread(target=perform_pause, daemon=True).start()

    def run_script_immediately(self, technique, script, mux_channel=None):
        if self.is_running:
            messagebox.showwarning("Busy", "Another measurement is currently running. Stop it or wait for it to finish before starting a new run.")
            return

        try:
            filepath, filename = self.save_script_file(technique, script, mux_channel=mux_channel)
        except Exception as e:
            messagebox.showerror("File Error", f"Failed to save {technique} script: {e}")
            return

        self.clear_log()
        self.is_running = True
        self.update_status(f"Running: {technique} - {filename}")
        self.log_message(f"Starting immediate {technique} run ({filename})")
        self.start_live_plot(f"{technique} (live)", label=technique)

        def worker():
            success = False
            csv_path = None
            stopped_by_user = False
            runner = None
            try:
                runner = SerialMeasurementRunner(
                    Path(filepath),
                    log_callback=self.log_message,
                    data_callback=self.queue_live_point,
                )
                self.current_runner = runner
                success, csv_path = runner.execute()
                stopped_by_user = not runner.is_running
            except Exception as exc:
                self.log_message(f"CRITICAL ERROR executing {technique}: {exc}")
            finally:
                self.current_runner = None

                def finalize():
                    self.is_running = False
                    self.stop_live_plot()
                    if stopped_by_user:
                        self.update_status("Ready (stopped)")
                        if csv_path:
                            self.plot_data(csv_path, color=self.last_live_plot_color, label=self.last_live_plot_label)
                        detail = f"{technique} run was stopped. Script: {filename}"
                        if csv_path:
                            detail += f"\nData saved to: {csv_path}"
                        self.log_message(f"{technique} run stopped by user.")
                        messagebox.showinfo("Run Stopped", detail)
                    elif success:
                        self.update_status("Ready")
                        if csv_path:
                            self.plot_data(csv_path, color=self.last_live_plot_color, label=self.last_live_plot_label)
                        detail = f"{technique} run completed. Script: {filename}"
                        if csv_path:
                            detail += f"\nData saved to: {csv_path}"
                        self.log_message(f"{technique} run completed successfully.")
                        messagebox.showinfo("Run Complete", detail)
                    else:
                        self.update_status("Ready (last run failed)")
                        if csv_path:
                            self.plot_data(csv_path, color=self.last_live_plot_color, label=self.last_live_plot_label)
                        self.log_message(f"{technique} run failed.")
                        messagebox.showerror("Run Failed", f"{technique} run failed. Check the log for details.")

                self.root.after(0, finalize)

        threading.Thread(target=worker, daemon=True).start()

    def save_script_file(self, technique, script, mux_channel=None):
        date_folder = self.base_path / datetime.now().strftime('%Y-%m-%d')
        date_folder.mkdir(exist_ok=True)
        slug = technique.lower().replace(' ', '_')
        if mux_channel is not None:
            slug = f"{slug}_ch{mux_channel}"
        filename = f"{len(list(date_folder.glob('*.ms'))) + 1:03d}_{slug}.ms"
        filepath = date_folder / filename
        with open(filepath, 'w') as f:
            f.write(script)
        return filepath, filename

    def add_to_queue(self, technique, script):
        filepath, filename = self.save_script_file(technique, script)
        queue_item = {'type': technique, 'script_path': str(filepath), 'status': 'pending', 'details': filename}
        self.measurement_queue.append(queue_item)
        self.refresh_queue_display()
        messagebox.showinfo("Success", f"{technique} added to queue\nSaved as: {filename}")

    def add_mux_scripts_to_queue(self, technique, base_script, channels):
        saved = []
        for ch in channels:
            mux_script = self._build_mux_script(base_script, ch)
            filepath, filename = self.save_script_file(technique, mux_script, mux_channel=ch)
            queue_item = {
                'type': technique,
                'script_path': str(filepath),
                'status': 'pending',
                'details': f"{filename} (MUX ch {ch})",
            }
            self.measurement_queue.append(queue_item)
            saved.append(f"{filename} (ch {ch})")
        self.refresh_queue_display()
        messagebox.showinfo(
            "Success",
            f"{technique} added for MUX channels: {', '.join(map(str, channels))}\nSaved as: {', '.join(saved)}",
        )

    def run_mux_script_sequence(self, technique, base_script, channels):
        if self.is_running:
            messagebox.showwarning("Busy", "Another measurement is currently running. Stop it or wait for it to finish before starting a new run.")
            return

        if not channels:
            return

        self.clear_log()
        self.is_running = True
        self.update_status(f"Running: {technique} (MUX ch {channels[0]})")
        self.log_message(f"Starting {technique} run for MUX channels: {', '.join(map(str, channels))}")

        def worker():
            stopped_by_user = False
            success = True
            last_csv_path = None

            for idx, ch in enumerate(channels, start=1):
                if not self.is_running:
                    stopped_by_user = True
                    success = False
                    break

                mux_script = self._build_mux_script(base_script, ch)
                try:
                    filepath, filename = self.save_script_file(technique, mux_script, mux_channel=ch)
                except Exception as exc:
                    self.log_message(f"File error: failed to save {technique} (MUX ch {ch}) script: {exc}")
                    success = False
                    break

                self.root.after(0, self.update_status, f"Running: {technique} (MUX ch {ch}) [{idx}/{len(channels)}]")
                self.log_message(f"Starting {technique} run (MUX ch {ch}) ({filename})")
                color = next(self.plot_color_cycle)
                label = f"MUX ch {ch}"
                self.last_live_plot_color = color
                self.last_live_plot_label = label
                self.root.after(0, self.start_live_plot, f"{technique} (MUX ch {ch} live)", color, label)

                runner = None
                csv_path = None
                try:
                    runner = SerialMeasurementRunner(
                        Path(filepath),
                        log_callback=self.log_message,
                        data_callback=self.queue_live_point,
                    )
                    self.current_runner = runner
                    ok, csv_path = runner.execute()
                    if not ok:
                        success = False
                        if not runner.is_running:
                            stopped_by_user = True
                        break
                except Exception as exc:
                    self.log_message(f"CRITICAL ERROR executing {technique} (MUX ch {ch}): {exc}")
                    success = False
                    break
                finally:
                    self.current_runner = None

                if csv_path:
                    last_csv_path = csv_path
                    self.root.after(0, self.plot_data, csv_path, self.last_live_plot_color, self.last_live_plot_label)

            def finalize():
                self.is_running = False
                self.stop_live_plot()
                if stopped_by_user:
                    self.update_status("Ready (stopped)")
                    detail = f"{technique} MUX run was stopped."
                    if last_csv_path:
                        detail += f"\nLast data saved to: {last_csv_path}"
                    self.log_message(f"{technique} MUX run stopped by user.")
                    messagebox.showinfo("Run Stopped", detail)
                elif success:
                    self.update_status("Ready")
                    detail = f"{technique} MUX run completed. Channels: {', '.join(map(str, channels))}"
                    if last_csv_path:
                        detail += f"\nLast data saved to: {last_csv_path}"
                    self.log_message(f"{technique} MUX run completed successfully.")
                    messagebox.showinfo("Run Complete", detail)
                else:
                    self.update_status("Ready (last run failed)")
                    self.log_message(f"{technique} MUX run failed.")
                    messagebox.showerror("Run Failed", f"{technique} MUX run failed. Check the log for details.")

            self.root.after(0, finalize)

        threading.Thread(target=worker, daemon=True).start()

    def run_swv_cycles(self, base_script, n_scans, delay_seconds):
        if self.is_running:
            messagebox.showwarning("Busy", "Another measurement is currently running. Stop it or wait for it to finish before starting a new run.")
            return

        self.clear_log()
        self.is_running = True
        self.update_status("Running: SWV (multi-scan)")
        self.log_message(f"Starting SWV run for {n_scans} scan(s)")

        def worker():
            stopped_by_user = False
            success = True
            last_csv_path = None

            for scan in range(1, n_scans + 1):
                if not self.is_running:
                    stopped_by_user = True
                    success = False
                    break

                try:
                    filepath, filename = self.save_script_file("SWV", base_script)
                except Exception as exc:
                    self.log_message(f"File error: failed to save SWV script: {exc}")
                    success = False
                    break

                label = f"SWV scan {scan}"
                color = next(self.plot_color_cycle)
                self.last_live_plot_color = color
                self.last_live_plot_label = label
                self.root.after(0, self.start_live_plot, f"SWV (scan {scan}/{n_scans} live)", color, label)
                self.root.after(0, self.update_status, f"Running: SWV scan {scan}/{n_scans}")
                self.log_message(f"Starting SWV scan {scan}/{n_scans} ({filename})")

                runner = None
                csv_path = None
                try:
                    runner = SerialMeasurementRunner(
                        Path(filepath),
                        log_callback=self.log_message,
                        data_callback=self.queue_live_point,
                    )
                    self.current_runner = runner
                    ok, csv_path = runner.execute()
                    if not ok:
                        success = False
                        if not runner.is_running:
                            stopped_by_user = True
                        break
                except Exception as exc:
                    self.log_message(f"CRITICAL ERROR executing SWV scan {scan}: {exc}")
                    success = False
                    break
                finally:
                    self.current_runner = None
                    self.root.after(0, self.stop_live_plot)

                if csv_path:
                    last_csv_path = csv_path
                    self.root.after(0, self.plot_data, csv_path, self.last_live_plot_color, self.last_live_plot_label)

                if delay_seconds > 0 and scan < n_scans:
                    self.log_message(f"Waiting {delay_seconds:.1f} sec before next SWV scan...")
                    waited = 0.0
                    while waited < delay_seconds and self.is_running:
                        time.sleep(min(0.5, delay_seconds - waited))
                        waited += min(0.5, delay_seconds - waited)

            def finalize():
                self.is_running = False
                if stopped_by_user:
                    self.update_status("Ready (stopped)")
                    detail = "SWV run was stopped."
                    if last_csv_path:
                        detail += f"\nLast data saved to: {last_csv_path}"
                    self.log_message("SWV run stopped by user.")
                    messagebox.showinfo("Run Stopped", detail)
                elif success:
                    self.update_status("Ready")
                    detail = f"SWV run completed. Scans: {n_scans}"
                    if last_csv_path:
                        detail += f"\nLast data saved to: {last_csv_path}"
                    self.log_message("SWV run completed successfully.")
                    messagebox.showinfo("Run Complete", detail)
                else:
                    self.update_status("Ready (last run failed)")
                    self.log_message("SWV run failed.")
                    messagebox.showerror("Run Failed", "SWV run failed. Check the log for details.")

            self.root.after(0, finalize)

        threading.Thread(target=worker, daemon=True).start()

    def run_mux_swv_cycles(self, base_script, channels, n_scans, delay_seconds):
        if self.is_running:
            messagebox.showwarning("Busy", "Another measurement is currently running. Stop it or wait for it to finish before starting a new run.")
            return

        if not channels:
            return

        self.clear_log()
        self.is_running = True
        self.update_status(f"Running: SWV (MUX ch {channels[0]})")
        self.log_message(
            f"Starting SWV run for MUX channels: {', '.join(map(str, channels))} ({n_scans} scan(s))"
        )

        def worker():
            stopped_by_user = False
            success = True
            last_csv_path = None

            for scan in range(1, n_scans + 1):
                for idx, ch in enumerate(channels, start=1):
                    if not self.is_running:
                        stopped_by_user = True
                        success = False
                        break

                    mux_script = self._build_mux_script(base_script, ch)
                    try:
                        filepath, filename = self.save_script_file("SWV", mux_script, mux_channel=ch)
                    except Exception as exc:
                        self.log_message(f"File error: failed to save SWV (MUX ch {ch}) script: {exc}")
                        success = False
                        break

                    label = f"MUX ch {ch} (scan {scan})"
                    color = next(self.plot_color_cycle)
                    self.last_live_plot_color = color
                    self.last_live_plot_label = label
                    self.root.after(0, self.start_live_plot, f"SWV (MUX ch {ch} scan {scan}/{n_scans})", color, label)
                    self.root.after(
                        0,
                        self.update_status,
                        f"Running: SWV MUX ch {ch} scan {scan}/{n_scans} [{idx}/{len(channels)}]",
                    )
                    self.log_message(f"Starting SWV MUX ch {ch} scan {scan}/{n_scans} ({filename})")

                    runner = None
                    csv_path = None
                    try:
                        runner = SerialMeasurementRunner(
                            Path(filepath),
                            log_callback=self.log_message,
                            data_callback=self.queue_live_point,
                        )
                        self.current_runner = runner
                        ok, csv_path = runner.execute()
                        if not ok:
                            success = False
                            if not runner.is_running:
                                stopped_by_user = True
                            break
                    except Exception as exc:
                        self.log_message(f"CRITICAL ERROR executing SWV MUX ch {ch} scan {scan}: {exc}")
                        success = False
                        break
                    finally:
                        self.current_runner = None
                        self.root.after(0, self.stop_live_plot)

                    if csv_path:
                        last_csv_path = csv_path
                        self.root.after(0, self.plot_data, csv_path, self.last_live_plot_color, self.last_live_plot_label)

                if not success or stopped_by_user:
                    break

                if delay_seconds > 0 and scan < n_scans:
                    self.log_message(f"Waiting {delay_seconds:.1f} sec before next SWV cycle...")
                    waited = 0.0
                    while waited < delay_seconds and self.is_running:
                        time.sleep(min(0.5, delay_seconds - waited))
                        waited += min(0.5, delay_seconds - waited)

            def finalize():
                self.is_running = False
                if stopped_by_user:
                    self.update_status("Ready (stopped)")
                    detail = "SWV MUX run was stopped."
                    if last_csv_path:
                        detail += f"\nLast data saved to: {last_csv_path}"
                    self.log_message("SWV MUX run stopped by user.")
                    messagebox.showinfo("Run Stopped", detail)
                elif success:
                    self.update_status("Ready")
                    detail = f"SWV MUX run completed. Scans: {n_scans}"
                    if last_csv_path:
                        detail += f"\nLast data saved to: {last_csv_path}"
                    self.log_message("SWV MUX run completed successfully.")
                    messagebox.showinfo("Run Complete", detail)
                else:
                    self.update_status("Ready (last run failed)")
                    self.log_message("SWV MUX run failed.")
                    messagebox.showerror("Run Failed", "SWV MUX run failed. Check the log for details.")

            self.root.after(0, finalize)

        threading.Thread(target=worker, daemon=True).start()
    
    def refresh_queue_display(self):
        for item in self.queue_tree.get_children(): self.queue_tree.delete(item)
        for i, item in enumerate(self.measurement_queue):
            self.queue_tree.insert('', 'end', iid=str(i), text=str(i+1), values=(item['type'], item['status'].upper(), item.get('details', '')))

    def _serialize_queue_item(self, item):
        data = {
            'type': item.get('type'),
            'status': item.get('status', 'pending'),
            'details': item.get('details'),
        }
        item_type = data['type']
        if item_type == 'PAUSE':
            data['pause_seconds'] = item.get('pause_seconds', 0.0)
        elif item_type == 'ALERT':
            data['alert_message'] = item.get('alert_message', '')
        elif item_type and item_type.startswith('PUMP_'):
            action = item.get('pump_action') or {}
            data['pump_action'] = {
                'name': action.get('name'),
                'params': dict(action.get('params') or {}),
            }
        else:
            if 'script_path' in item:
                data['script_path'] = item['script_path']
        return data

    def save_queue(self):
        if not self.measurement_queue:
            messagebox.showwarning("Empty Queue", "No items to save.")
            return
        if self.is_running:
            messagebox.showwarning("Queue Running", "Stop the queue before saving.")
            return

        default_name = f"queue_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        file_path = filedialog.asksaveasfilename(
            title="Save Queue",
            defaultextension=".json",
            filetypes=(("Queue Files", "*.json"), ("All Files", "*.*")),
            initialdir=str(self.base_path),
            initialfile=default_name,
        )
        if not file_path:
            return

        payload = {
            'metadata': {
                'saved_at': datetime.now().isoformat(timespec='seconds'),
                'version': 1,
            },
            'items': [self._serialize_queue_item(item) for item in self.measurement_queue],
        }

        try:
            with open(file_path, 'w', encoding='utf-8') as handle:
                json.dump(payload, handle, indent=2)
        except OSError as exc:
            messagebox.showerror("Save Failed", f"Could not save queue:\n{exc}")
            return

        messagebox.showinfo("Queue Saved", f"Queue saved to:\n{file_path}")

    def load_queue(self):
        if self.is_running:
            messagebox.showwarning("Queue Running", "Stop the queue before loading.")
            return

        file_path = filedialog.askopenfilename(
            title="Load Queue",
            defaultextension=".json",
            filetypes=(("Queue Files", "*.json"), ("All Files", "*.*")),
            initialdir=str(self.base_path),
        )
        if not file_path:
            return

        try:
            with open(file_path, 'r', encoding='utf-8') as handle:
                payload = json.load(handle)
            items = payload.get('items')
            if not isinstance(items, list):
                raise ValueError("Queue file missing 'items' list")
        except Exception as exc:
            messagebox.showerror("Load Failed", f"Could not load queue:\n{exc}")
            return

        new_queue = []
        skipped = 0
        for raw_item in items:
            if not isinstance(raw_item, dict):
                skipped += 1
                continue
            item_type = raw_item.get('type')
            if not item_type:
                skipped += 1
                continue

            queue_item = {'type': item_type, 'status': 'pending'}
            details = raw_item.get('details')

            if item_type == 'PAUSE':
                try:
                    seconds = float(raw_item.get('pause_seconds', 0.0))
                except (TypeError, ValueError):
                    skipped += 1
                    continue
                queue_item['pause_seconds'] = seconds
                queue_item['details'] = details or f'Pause for {seconds:.1f} sec'
            elif item_type == 'ALERT':
                message = raw_item.get('alert_message')
                if not isinstance(message, str) or not message.strip():
                    skipped += 1
                    continue
                queue_item['alert_message'] = message.strip()
                queue_item['details'] = details or "Alert pause"
            elif item_type.startswith('PUMP_'):
                action = raw_item.get('pump_action') or {}
                action_name = action.get('name')
                if not action_name:
                    skipped += 1
                    continue
                params = action.get('params') or {}
                queue_item['pump_action'] = {
                    'name': action_name,
                    'params': dict(params),
                }
                queue_item['details'] = details or f'Pump action {action_name}'
            else:
                script_path = raw_item.get('script_path')
                if not script_path:
                    skipped += 1
                    continue
                queue_item['script_path'] = script_path
                queue_item['details'] = details or Path(script_path).name
                if not Path(script_path).exists():
                    self.log_message(f"Warning: queue file references missing script -> {script_path}")

            new_queue.append(queue_item)

        if not new_queue:
            messagebox.showwarning("Load Queue", "No valid queue items found in the selected file.")
            return

        self.measurement_queue = new_queue
        self.refresh_queue_display()
        self.update_status(f"Queue loaded ({len(new_queue)} items)")
        if skipped:
            self.log_message(f"Queue load skipped {skipped} invalid item(s) from {file_path}.")
        messagebox.showinfo("Queue Loaded", f"Loaded {len(new_queue)} queue item(s).")

    def run_queue(self):
        self.reset_queue_reorder()
        if not self.measurement_queue: messagebox.showwarning("Empty Queue", "No items in queue"); return
        if self.is_running: messagebox.showwarning("Already Running", "Queue is already running"); return
        self.is_running = True
        self.clear_log()
        self.queue_thread = threading.Thread(target=self.execute_queue, args=(0,), daemon=True)
        self.queue_thread.start()

    def run_queue_from_selected(self):
        self.reset_queue_reorder()
        if not self.measurement_queue:
            messagebox.showwarning("Empty Queue", "No items in queue")
            return
        if self.is_running:
            messagebox.showwarning("Already Running", "Queue is already running")
            return
        selected = self.queue_tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Select a queue item to start from.")
            return
        try:
            index = self.queue_tree.index(selected[0])
        except Exception:
            messagebox.showerror("Selection Error", "Unable to determine selected queue item.")
            return
        if index < 0 or index >= len(self.measurement_queue):
            messagebox.showerror("Selection Error", "Selected queue item is out of range.")
            return
        self.is_running = True
        self.clear_log()
        self.queue_thread = threading.Thread(target=self.execute_queue, args=(index,), daemon=True)
        self.queue_thread.start()

    def execute_queue(self, start_index=0):
        for i, item in enumerate(list(self.measurement_queue)[start_index:], start=start_index):
            if not self.is_running: self.log_message("Queue execution stopped by user."); break
            self.measurement_queue[i]['status'] = 'running'
            self.root.after(0, self.refresh_queue_display)
            self.root.after(0, self.update_status, f"Running: {item['type']} - {item.get('details', '')}")
            csv_path = None
            success = False
            try:
                if item['type'] == 'PAUSE':
                    seconds = float(item.get('pause_seconds', 0))
                    self.log_message(f"Queue pause start: {seconds:.1f} sec")
                    pause_completed = self.execute_pause(seconds)
                    self.measurement_queue[i]['status'] = 'completed' if pause_completed else 'stopped'
                    success = pause_completed
                    if pause_completed:
                        self.log_message(f"Queue pause complete: {seconds:.1f} sec")
                    else:
                        self.log_message("Queue pause cancelled before completion.")
                elif item['type'] == 'ALERT':
                    message = item.get('alert_message', 'Paused - click OK to continue.')
                    self.log_message("Queue alert pause started")
                    alert_completed = self.execute_alert_pause(message)
                    self.measurement_queue[i]['status'] = 'completed' if alert_completed else 'stopped'
                    success = alert_completed
                    if alert_completed:
                        self.log_message("Queue alert pause complete")
                    else:
                        self.log_message("Queue alert pause cancelled before completion.")
                elif item['type'].startswith('PUMP_'):
                    success = self.execute_pump_action(item)
                    self.measurement_queue[i]['status'] = 'completed' if success else 'failed'
                else:
                    self.root.after(0, self.start_live_plot, f"{item['type']} (live)", None, item['type'])
                    try:
                        self.current_runner = SerialMeasurementRunner(
                            Path(item['script_path']),
                            log_callback=self.log_message,
                            data_callback=self.queue_live_point,
                        )
                        success, csv_path = self.current_runner.execute()
                        self.measurement_queue[i]['status'] = 'completed' if success else 'failed'
                    finally:
                        self.current_runner = None
                        self.root.after(0, self.stop_live_plot)
            except Exception as e:
                self.measurement_queue[i]['status'] = 'failed'
                self.log_message(f"CRITICAL ERROR in queue execution: {e}")
            
            # If the measurement was successful and created a file, plot it
            if csv_path:
                self.root.after(0, self.plot_data, csv_path, self.last_live_plot_color)

            self.root.after(0, self.refresh_queue_display)
            time.sleep(1)

        self.is_running = False
        self.root.after(0, self.update_status, "Queue Complete")
    
    def stop_queue(self):
        if not self.is_running: return
        self.is_running = False
        if self.current_runner: self.current_runner.stop()
        self.update_status("Queue Stopped")
    
    def clear_queue(self):
        self.reset_queue_reorder()
        if self.is_running: messagebox.showwarning("Queue Running", "Cannot clear queue while running"); return
        self.measurement_queue = []
        self.refresh_queue_display()
        self.update_status("Queue Cleared")

    def delete_selected_queue_item(self):
        self.reset_queue_reorder()
        if self.is_running:
            messagebox.showwarning("Queue Running", "Cannot delete items while running")
            return
        selected = self.queue_tree.selection()
        if not selected:
            messagebox.showwarning("No Selection", "Select a queue item to delete.")
            return
        try:
            index = self.queue_tree.index(selected[0])
        except Exception:
            messagebox.showerror("Selection Error", "Unable to determine selected queue item.")
            return
        if index < 0 or index >= len(self.measurement_queue):
            messagebox.showerror("Selection Error", "Selected queue item is out of range.")
            return
        removed = self.measurement_queue.pop(index)
        self.refresh_queue_display()
        self.update_status("Queue item deleted")
        self.log_message(f"Queue item deleted: {removed.get('details', removed.get('type', 'Unknown'))}")

    def _on_queue_drag_start(self, event):
        if self.is_running:
            return
        item = self.queue_tree.identify_row(event.y)
        if not item:
            return
        self.queue_drag_item = item
        if not self.queue_reorder_pending:
            self.queue_reorder_snapshot = list(self.measurement_queue)

    def _on_queue_drag_motion(self, event):
        if self.is_running or not self.queue_drag_item:
            return
        target = self.queue_tree.identify_row(event.y)
        if not target or target == self.queue_drag_item:
            return
        target_index = self.queue_tree.index(target)
        self.queue_tree.move(self.queue_drag_item, '', target_index)
        self.queue_reorder_pending = True

    def _on_queue_drag_release(self, event):
        if self.queue_reorder_pending:
            self.update_status("Queue reorder pending - click Confirm Move")
        self.queue_drag_item = None

    def confirm_queue_reorder(self):
        if not self.queue_reorder_pending or not self.queue_reorder_snapshot:
            messagebox.showinfo("No Changes", "No pending queue reorder to confirm.")
            return
        try:
            order = [int(iid) for iid in self.queue_tree.get_children()]
        except Exception:
            messagebox.showerror("Reorder Error", "Failed to read queue order.")
            return
        if any(i < 0 or i >= len(self.queue_reorder_snapshot) for i in order):
            messagebox.showerror("Reorder Error", "Queue order is out of range.")
            return
        self.measurement_queue = [self.queue_reorder_snapshot[i] for i in order]
        self.queue_reorder_snapshot = None
        self.queue_reorder_pending = False
        self.refresh_queue_display()
        self.update_status("Queue reordered")

    def reset_queue_reorder(self):
        if self.queue_reorder_pending:
            self.queue_reorder_pending = False
            self.queue_reorder_snapshot = None
            self.refresh_queue_display()

    def update_status(self, message):
        self.status_label.config(text=f"Status: {message}")

    def execute_pause(self, seconds: float) -> bool:
        total = max(0.0, float(seconds))
        if total <= 0:
            self.root.after(0, self.update_status, "Pause complete")
            return True

        start_time = time.time()
        while self.is_running:
            elapsed = time.time() - start_time
            remaining = total - elapsed
            if remaining <= 0:
                break
            remaining = max(0.0, remaining)

            def update(rem=remaining):
                if self.is_running:
                    self.update_status(f"Pausing: {rem:.1f} sec remaining")

            self.root.after(0, update)
            time.sleep(min(0.5, remaining))

        if not self.is_running:
            return False

        self.root.after(0, self.update_status, "Pause complete")
        return True

    def execute_alert_pause(self, message: str) -> bool:
        if not self.is_running:
            return False
        done = threading.Event()

        def show_alert():
            messagebox.showinfo("Paused", message)
            done.set()

        self.root.after(0, show_alert)
        while self.is_running and not done.is_set():
            done.wait(timeout=0.2)
        return done.is_set()

    def execute_pump_action(self, item):
        if not PUMP_AVAILABLE or self.pump_ctrl is None:
            self.log_message("Pump backend unavailable; cannot execute queued action.")
            return False

        action_info = item.get('pump_action') or {}
        action_name = action_info.get('name')
        params = action_info.get('params') or {}
        details = item.get('details', f"Pump action {action_name}")

        if not action_name:
            self.log_message("Invalid pump queue item: missing action name.")
            return False

        if not self._pump_require_connection():
            return False

        def log_both(message):
            self.log_message(message)
            self.pump_log(message)

        log_both(f"Queue start -> {details}")

        self.set_pump_busy(True)
        try:
            if action_name == 'INIT':
                self.pump_ctrl.initialize()
                log_both("Queue init complete")
                return True
            if action_name == 'SET_SPEED':
                speed = int(params.get('speed'))
                self.pump_ctrl.set_speed(speed)
                log_both(f"Queue set speed done (S{speed}R)")
                return True
            if action_name == 'VALVE':
                port = int(params.get('port'))
                self.pump_ctrl.valve_to(port)
                log_both(f"Queue valve move complete (I{port}R)")
                return True
            if action_name == 'ASPIRATE':
                volume = float(params.get('volume'))
                speed = int(params.get('speed'))
                self.pump_ctrl.set_speed(speed)
                self.pump_ctrl.aspirate_ul(volume)
                log_both(f"Queue aspirate complete ({volume:.2f} µL @ S{speed}R)")
                return True
            if action_name == 'DISPENSE':
                volume = float(params.get('volume'))
                speed = int(params.get('speed'))
                self.pump_ctrl.set_speed(speed)
                self.pump_ctrl.dispense_ul(volume)
                log_both(f"Queue dispense complete ({volume:.2f} µL @ S{speed}R)")
                return True

            log_both(f"Unsupported pump action: {action_name}")
            return False
        except Exception as exc:
            log_both(f"Pump action failed: {exc}")
            return False
        finally:
            self.set_pump_busy(False)


def main():
    root = tk.Tk()
    app = ElectrochemGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
