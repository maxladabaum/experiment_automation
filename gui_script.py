#!/usr/bin/env python3
# electrochemistry_automation_gui.py

import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import json
import os
from datetime import datetime
from pathlib import Path
import threading
import time
import sys
import serial
import serial.tools.list_ports
import csv
import math
import collections
import warnings

# --- Matplotlib and Pandas imports for plotting ---
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


# Import the local tecancavro library
try:
    from tecancavro import XCaliburD, TecanAPISerial
    PUMP_AVAILABLE = True
except ImportError:
    PUMP_AVAILABLE = False
    print("Warning: tecancavro library not found. Pump features disabled.")

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
    def parse_metadata(tokens: list[str]) -> dict[str, int]:
        """Parse the (optional) metadata."""
        metadata = {}
        for token in tokens:
            if (len(token) == 2) and (token[0] == '1'):
                metadata['status'] = int(token[1], 16)
            if (len(token) == 3) and (token[0] == '2'):
                metadata['cr'] = int(token[1:], 16)
        return metadata


def parse_mscript_data_package(line: str) -> list[MScriptVar]:
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
        if unit in ['V', 'V/s']: # V/s for scan rate is treated like V
            # For values between -1 and 1 (exclusive of 0), use milli suffix
            if abs(val) < 1.0 and val != 0:
                return f"{int(val * 1000)}m"
            else:
                return f"{int(val)}" # For whole numbers like 0, 1, -2 etc.
        elif unit == 'Hz':
            return f"{int(val)}"
        return value_str # Fallback
    except (ValueError, TypeError):
        return value_str # Return original if conversion fails

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
        elif len(candidates) > 1:
            self.log(f"Multiple devices found: {candidates}")
            self.log(f"Using first device: {candidates[0]}")
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
        
        self.pump = None
        self.pump_com = None
        
        self.setup_gui()
        
    def setup_gui(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        self.method_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.method_frame, text="Method Creation")
        self.setup_method_tab()
        
        self.queue_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.queue_frame, text="Queue & Execution")
        self.setup_queue_tab()
        
        if PUMP_AVAILABLE:
            self.pump_frame = ttk.Frame(self.notebook)
            self.notebook.add(self.pump_frame, text="Pump Control")
            self.setup_pump_tab()
        
        self.script_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.script_frame, text="Script Preview")
        self.setup_script_tab()
        
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
            "e", "var c", "var p", "var f", "var r", "set_pgstat_mode 2", "set_max_bandwidth 1600",
            f"set_range_minmax da {min_pot_mv}m {max_pot_mv}m",
            "set_range ba 5m", "set_autoranging ba 100n 5m", "cell_on"
        ]
        
        if float(cond_time) > 0:
            script_parts.extend([
                f"# Equilibrate at {cond_pot} for {cond_time}s",
                f"set_e {cond_pot}", f"wait {cond_time}"
            ])
            
        script_parts.extend([
            "# SWV measurement loop",
            f"meas_loop_swv p c f r {begin} {end} {step} {amplitude} {frequency}",
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
        ttk.Button(button_frame, text="Add to Queue", command=self.add_cv_to_queue).pack(side='left', padx=5)
        
    def show_swv_params(self):
        self.clear_params_frame()
        self.current_technique = "SWV"
        self.swv_params = {}
        params = [("Begin Potential (V):", "begin_potential", "-0.5"), ("End Potential (V):", "end_potential", "0.5"), ("Step Potential (V):", "step_potential", "0.002"), ("Amplitude (V):", "amplitude", "0.02"), ("Frequency (Hz):", "frequency", "15"), ("Conditioning Potential (V):", "cond_potential", "0"), ("Conditioning Time (s):", "cond_time", "0")]
        for i, (label, key, default) in enumerate(params):
            ttk.Label(self.params_frame, text=label).grid(row=i, column=0, sticky='w', pady=2)
            entry = ttk.Entry(self.params_frame, width=15); entry.insert(0, default); entry.grid(row=i, column=1, pady=2)
            self.swv_params[key] = entry
        button_frame = ttk.Frame(self.params_frame); button_frame.grid(row=len(params), column=0, columnspan=2, pady=20)
        ttk.Button(button_frame, text="Generate Script", command=self.generate_swv_script).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Add to Queue", command=self.add_swv_to_queue).pack(side='left', padx=5)
        
    def setup_pump_tab(self):
        pass

    def setup_queue_tab(self):
        main_pane = ttk.PanedWindow(self.queue_frame, orient=tk.VERTICAL); main_pane.pack(fill='both', expand=True)
        top_frame = ttk.Frame(main_pane); main_pane.add(top_frame, weight=1)
        bottom_frame = ttk.Frame(main_pane); main_pane.add(bottom_frame, weight=1)
        control_frame = ttk.Frame(top_frame); control_frame.pack(pady=10, fill='x', padx=10)
        ttk.Button(control_frame, text="Run Queue", command=self.run_queue).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Stop", command=self.stop_queue).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Clear Queue", command=self.clear_queue).pack(side='left', padx=5)
        self.queue_tree = ttk.Treeview(top_frame, columns=('Type', 'Status', 'Details'), show='tree headings', height=8)
        self.queue_tree.heading('#0', text='#'); self.queue_tree.heading('Type', text='Type'); self.queue_tree.heading('Status', text='Status'); self.queue_tree.heading('Details', text='Details')
        self.queue_tree.column('#0', width=50); self.queue_tree.column('Type', width=150); self.queue_tree.column('Status', width=100); self.queue_tree.column('Details', width=400)
        self.queue_tree.pack(fill='both', expand=True, padx=10, pady=5)
        log_frame = ttk.LabelFrame(bottom_frame, text="Live Output Log"); log_frame.pack(fill='both', expand=True, padx=10, pady=5)
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10); self.log_text.pack(fill='both', expand=True)
        self.log_text.config(state='disabled')
        self.status_label = ttk.Label(self.queue_frame, text="Status: Ready", relief='sunken'); self.status_label.pack(side='bottom', fill='x', padx=10, pady=5)

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

    def plot_data(self, csv_path):
        """Reads a CSV file and plots the voltammogram."""
        try:
            df = pd.read_csv(csv_path)
            potential_col = 'Potential (V)'
            current_col = 'Current (µA)'

            if potential_col in df.columns and current_col in df.columns:
                self.ax.clear()  # Clear the previous plot
                self.ax.plot(df[potential_col], df[current_col])
                
                # Apply styling from the user's example
                self.ax.set_title('Voltammogram')
                self.ax.set_xlabel('Potential (V)')
                self.ax.set_ylabel('Current (µA)')
                self.ax.grid(visible=True, which='major', linestyle='-')
                self.ax.grid(visible=True, which='minor', linestyle='--', alpha=0.2)
                self.ax.minorticks_on()
                
                self.canvas.draw() # Redraw the canvas
                self.notebook.select(self.plotter_frame) # Switch to the plot tab
            else:
                messagebox.showerror("Plot Error", "CSV file must contain 'Potential (V)' and 'Current (µA)' columns.")
        except Exception as e:
            messagebox.showerror("Plot Error", f"Failed to plot data: {e}")

    def clear_params_frame(self):
        for widget in self.params_frame.winfo_children(): widget.destroy()

    def generate_cv_script(self):
        try:
            script = self.create_cv_methodscript()
            self.current_script = script
            self.update_script_preview(script)
            self.notebook.select(self.notebook.index(self.script_frame))
            return script
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate script: {str(e)}")
            return None

    def generate_swv_script(self):
        try:
            script = self.create_swv_methodscript()
            self.current_script = script
            self.update_script_preview(script)
            self.notebook.select(self.notebook.index(self.script_frame))
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

    def add_to_queue(self, technique, script):
        date_folder = self.base_path / datetime.now().strftime('%Y-%m-%d'); date_folder.mkdir(exist_ok=True)
        filename = f"{len(list(date_folder.glob('*.ms'))) + 1:03d}_{technique.lower()}.ms"
        filepath = date_folder / filename
        with open(filepath, 'w') as f: f.write(script)
        queue_item = {'type': technique, 'script_path': str(filepath), 'status': 'pending', 'details': filename}
        self.measurement_queue.append(queue_item)
        self.refresh_queue_display()
        messagebox.showinfo("Success", f"{technique} added to queue\nSaved as: {filename}")
    
    def refresh_queue_display(self):
        for item in self.queue_tree.get_children(): self.queue_tree.delete(item)
        for i, item in enumerate(self.measurement_queue):
            self.queue_tree.insert('', 'end', text=str(i+1), values=(item['type'], item['status'].upper(), item.get('details', '')))

    def run_queue(self):
        if not self.measurement_queue: messagebox.showwarning("Empty Queue", "No items in queue"); return
        if self.is_running: messagebox.showwarning("Already Running", "Queue is already running"); return
        self.is_running = True
        self.log_text.config(state='normal'); self.log_text.delete(1.0, tk.END); self.log_text.config(state='disabled')
        self.queue_thread = threading.Thread(target=self.execute_queue, daemon=True)
        self.queue_thread.start()

    def execute_queue(self):
        for i, item in enumerate(list(self.measurement_queue)):
            if not self.is_running: self.log_message("Queue execution stopped by user."); break
            self.measurement_queue[i]['status'] = 'running'
            self.root.after(0, self.refresh_queue_display)
            self.root.after(0, self.update_status, f"Running: {item['type']} - {item.get('details', '')}")
            csv_path = None
            try:
                if item['type'].startswith('PUMP_'): self.execute_pump_action(item)
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

    def execute_pump_action(self, item):
        pass


def main():
    root = tk.Tk()
    app = ElectrochemGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
