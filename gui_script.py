#!/usr/bin/env python3
# electrochemistry_automation_gui.py

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
import warnings
from typing import Dict, List, Optional

# --- Matplotlib and Pandas imports for plotting ---
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


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

PREFERRED_SYRINGE_UL = 1000.0
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
        return [MScriptVar(var) for var in line[1:-1].split(';')]
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
    def __init__(self, script_path, log_callback=print):
        self.script_path = Path(script_path)
        self.data_points = []
        self.connection = None
        self.log = log_callback # Callback to log messages to the GUI
        self.is_running = True

        self.data_base_path = Path("measurement_data")
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
            
            while self.is_running:
                try:
                    line = self.connection.readline()
                    if not line: continue
                    text = line.decode('utf-8', errors='ignore').strip()
                    if not text: continue
                    self.log(text)
                    if text.startswith('P'): self.parse_data_line(text)
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
            self.log(f"Error running script: {e}")
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
        except Exception as e:
            self.log(f"Error parsing data package: {line} -> {e}")

    def save_data_to_csv(self):
        if not self.data_points:
            self.log("No data to save")
            return None
        base_name = self.script_path.stem
        timestamp = datetime.now().strftime('%H%M%S')
        csv_filename = self.data_folder / f"{base_name}_{timestamp}.csv"
        with open(csv_filename, 'w', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=['potential', 'current'])
            writer.writerow({'potential': 'Potential (V)', 'current': 'Current (µA)'})
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
        n_scans = int(self.swv_params['n_scans'].get())
        if n_scans < 1:
            raise ValueError("Number of scans must be at least 1")
        
        min_pot = min(begin_v, end_v) - amp_v
        max_pot = max(begin_v, end_v) + amp_v
        min_pot_mv, max_pot_mv = int(min_pot * 1000), int(max_pot * 1000)

        script_parts = [
            "e", "var c", "var p", "var f", "var r", "set_pgstat_mode 2", "set_max_bandwidth 1600",
            f"set_range_minmax da {min_pot_mv}m {max_pot_mv}m",
            "set_range ba 5m", "set_autoranging ba 100n 5m", "cell_on"
        ]
        
        if float(cond_time) > 0:
            script_parts.extend([
                f"# Equilibrate at {cond_pot} for {cond_time}s",
                f"set_e {cond_pot}", f"wait {cond_time}"
            ])
            
        swv_command = f"meas_loop_swv p c f r {begin} {end} {step} {amplitude} {frequency}"
        if n_scans > 1:
            swv_command += f" nscans({n_scans})"

        script_parts.extend([
            "# SWV measurement loop",
            swv_command,
            "\tpck_start", "\tpck_add p", "\tpck_add c", "\tpck_add f", "\tpck_add r",
            "\tpck_end", "endloop", "on_finished:", "cell_off"
        ])

        return "\n".join(script_parts)
    
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
        params = [("Begin Potential (V):", "begin_potential", "0"), ("Vertex 1 (V):", "vertex1", "-0.5"), ("Vertex 2 (V):", "vertex2", "0.5"), ("Step Potential (V):", "step_potential", "0.002"), ("Scan Rate (V/s):", "scan_rate", "0.1"), ("Number of Scans:", "n_scans", "1"), ("Conditioning Potential (V):", "cond_potential", "0"), ("Conditioning Time (s):", "cond_time", "0")]
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
        params = [("Begin Potential (V):", "begin_potential", "-0.5"), ("End Potential (V):", "end_potential", "0.5"), ("Step Potential (V):", "step_potential", "0.002"), ("Amplitude (V):", "amplitude", "0.02"), ("Frequency (Hz):", "frequency", "15"), ("Number of Scans:", "n_scans", "1"), ("Conditioning Potential (V):", "cond_potential", "0"), ("Conditioning Time (s):", "cond_time", "0")]
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

        button_frame = ttk.Frame(self.params_frame)
        button_frame.grid(row=1, column=0, columnspan=2, pady=20)
        ttk.Button(button_frame, text="Add Pause to Queue", command=self.add_pause_to_queue).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Run Pause Now", command=self.run_pause_immediately).pack(side='left', padx=5)

    def setup_pump_tab(self):
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
        ttk.Button(control_frame, text="Stop", command=self.stop_queue).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Save Queue", command=self.save_queue).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Load Queue", command=self.load_queue).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Clear Queue", command=self.clear_queue).pack(side='left', padx=5)
        self.queue_tree = ttk.Treeview(top_frame, columns=('Type', 'Status', 'Details'), show='tree headings', height=8)
        self.queue_tree.heading('#0', text='#'); self.queue_tree.heading('Type', text='Type'); self.queue_tree.heading('Status', text='Status'); self.queue_tree.heading('Details', text='Details')
        self.queue_tree.column('#0', width=50); self.queue_tree.column('Type', width=150); self.queue_tree.column('Status', width=100); self.queue_tree.column('Details', width=400)
        self.queue_tree.pack(fill='both', expand=True, padx=10, pady=5)
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

        # Create a Matplotlib figure and axis
        self.fig = Figure(figsize=(8, 6), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.ax.set_title('Voltammogram')
        self.ax.set_xlabel('Potential (V)')
        self.ax.set_ylabel('Current (µA)')
        self.ax.grid(True)

        # Create a canvas to embed the plot in Tkinter
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.plotter_frame)
        self.canvas.draw()
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

    def plot_data(self, csv_path):
        """Reads a CSV file and plots the voltammogram."""
        try:
            df = self._read_csv_with_fallback(csv_path)
        except Exception as exc:
            self.log_message(f"Plot error: failed to read {csv_path}: {exc}")
            messagebox.showerror("Plot Error", f"Failed to read data: {exc}")
            self.update_status("Plot failed: see log for details")
            return

        potential_col = self._find_column(df, ("Potential (V)",))
        current_col = self._find_column(df, ("Current (µA)", "Current (uA)", "Current (�A)"))

        if not potential_col or not current_col:
            message = "Plot error: CSV file must contain 'Potential (V)' and 'Current (µA)' columns."
            self.log_message(message)
            messagebox.showerror("Plot Error", message)
            self.update_status("Plot failed: missing required columns")
            return

        try:
            self.ax.clear()
            self.ax.plot(df[potential_col], df[current_col])
            self.ax.set_title('Voltammogram')
            self.ax.set_xlabel('Potential (V)')
            self.ax.set_ylabel('Current (µA)')
            self.ax.grid(visible=True, which='major', linestyle='-')
            self.ax.grid(visible=True, which='minor', linestyle='--', alpha=0.2)
            self.ax.minorticks_on()
            self.canvas.draw()
            self.notebook.select(self.plotter_frame)
        except Exception as exc:
            self.log_message(f"Plot error: failed to render {csv_path}: {exc}")
            messagebox.showerror("Plot Error", f"Failed to render plot: {exc}")
            self.update_status("Plot failed: see log for details")
    def clear_params_frame(self):
        for widget in self.params_frame.winfo_children(): widget.destroy()

    def generate_cv_script(self):
        try:
            script = self.create_cv_methodscript()
            self.current_script = script
            self.update_script_preview(script)
            return script
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate script: {str(e)}")
            return None

    def generate_swv_script(self):
        try:
            script = self.create_swv_methodscript()
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
        script = self.generate_cv_script()
        if script: self.add_to_queue("CV", script)
    
    def add_swv_to_queue(self):
        script = self.generate_swv_script()
        if script: self.add_to_queue("SWV", script)

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

    def run_cv_immediately(self):
        script = self.generate_cv_script()
        if script: self.run_script_immediately("CV", script)

    def run_swv_immediately(self):
        script = self.generate_swv_script()
        if script: self.run_script_immediately("SWV", script)

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

    def run_script_immediately(self, technique, script):
        if self.is_running:
            messagebox.showwarning("Busy", "Another measurement is currently running. Stop it or wait for it to finish before starting a new run.")
            return

        try:
            filepath, filename = self.save_script_file(technique, script)
        except Exception as e:
            messagebox.showerror("File Error", f"Failed to save {technique} script: {e}")
            return

        self.clear_log()
        self.is_running = True
        self.update_status(f"Running: {technique} - {filename}")
        self.log_message(f"Starting immediate {technique} run ({filename})")

        def worker():
            success = False
            csv_path = None
            stopped_by_user = False
            runner = None
            try:
                runner = SerialMeasurementRunner(Path(filepath), log_callback=self.log_message)
                self.current_runner = runner
                success, csv_path = runner.execute()
                stopped_by_user = not runner.is_running
            except Exception as exc:
                self.log_message(f"CRITICAL ERROR executing {technique}: {exc}")
            finally:
                self.current_runner = None

                def finalize():
                    self.is_running = False
                    if stopped_by_user:
                        self.update_status("Ready (stopped)")
                        if csv_path:
                            self.plot_data(csv_path)
                        detail = f"{technique} run was stopped. Script: {filename}"
                        if csv_path:
                            detail += f"\nData saved to: {csv_path}"
                        self.log_message(f"{technique} run stopped by user.")
                        messagebox.showinfo("Run Stopped", detail)
                    elif success:
                        self.update_status("Ready")
                        if csv_path:
                            self.plot_data(csv_path)
                        detail = f"{technique} run completed. Script: {filename}"
                        if csv_path:
                            detail += f"\nData saved to: {csv_path}"
                        self.log_message(f"{technique} run completed successfully.")
                        messagebox.showinfo("Run Complete", detail)
                    else:
                        self.update_status("Ready (last run failed)")
                        if csv_path:
                            self.plot_data(csv_path)
                        self.log_message(f"{technique} run failed.")
                        messagebox.showerror("Run Failed", f"{technique} run failed. Check the log for details.")

                self.root.after(0, finalize)

        threading.Thread(target=worker, daemon=True).start()

    def save_script_file(self, technique, script):
        date_folder = self.base_path / datetime.now().strftime('%Y-%m-%d')
        date_folder.mkdir(exist_ok=True)
        slug = technique.lower().replace(' ', '_')
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
    
    def refresh_queue_display(self):
        for item in self.queue_tree.get_children(): self.queue_tree.delete(item)
        for i, item in enumerate(self.measurement_queue):
            self.queue_tree.insert('', 'end', text=str(i+1), values=(item['type'], item['status'].upper(), item.get('details', '')))

    def _serialize_queue_item(self, item):
        data = {
            'type': item.get('type'),
            'status': item.get('status', 'pending'),
            'details': item.get('details'),
        }
        item_type = data['type']
        if item_type == 'PAUSE':
            data['pause_seconds'] = item.get('pause_seconds', 0.0)
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
        if not self.measurement_queue: messagebox.showwarning("Empty Queue", "No items in queue"); return
        if self.is_running: messagebox.showwarning("Already Running", "Queue is already running"); return
        self.is_running = True
        self.clear_log()
        self.queue_thread = threading.Thread(target=self.execute_queue, daemon=True)
        self.queue_thread.start()

    def execute_queue(self):
        for i, item in enumerate(list(self.measurement_queue)):
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
                elif item['type'].startswith('PUMP_'):
                    success = self.execute_pump_action(item)
                    self.measurement_queue[i]['status'] = 'completed' if success else 'failed'
                else:
                    self.current_runner = SerialMeasurementRunner(Path(item['script_path']), log_callback=self.log_message)
                    success, csv_path = self.current_runner.execute()
                    self.measurement_queue[i]['status'] = 'completed' if success else 'failed'
                    self.current_runner = None
            except Exception as e:
                self.measurement_queue[i]['status'] = 'failed'
                self.log_message(f"CRITICAL ERROR in queue execution: {e}")
            
            # If the measurement was successful and created a file, plot it
            if csv_path:
                self.root.after(0, self.plot_data, csv_path)

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
        if self.is_running: messagebox.showwarning("Queue Running", "Cannot clear queue while running"); return
        self.measurement_queue = []
        self.refresh_queue_display()
        self.update_status("Queue Cleared")

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

