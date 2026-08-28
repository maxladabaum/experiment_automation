#!/usr/bin/env python3
# mux_testing_fixed.py
#
# Fixes applied:
# 1) Ensures there is exactly ONE ElectrochemGUI class definition containing save_queue/load_queue/etc.
# 2) Integrates the EmStat MUX live-plot GUI as a RIGHT-SIDE panel inside the "Method Creation" tab
#    (to the right of your technique/parameter UI).
# 3) Live-plotter runs responsively: serial RX happens in a background thread; UI stays responsive.
# 4) Live plot shows ALL channels (curves) on ONE axis, each with its own color (matplotlib default cycle).

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
from typing import Dict, List, Optional, Tuple

import pandas as pd
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


# =========================
# Pump integration (unchanged from your script; guarded by try/except)
# =========================
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


# =========================
# PalmSens MethodSCRIPT parser integration (your mscript-adapted code)
# =========================
VarType = collections.namedtuple('VarType', ['id', 'name', 'unit'])

SI_PREFIX_FACTOR = {
    'a': 1e-18, 'f': 1e-15, 'p': 1e-12, 'n': 1e-9, 'u': 1e-6,
    'm': 1e-3, ' ': 1e0, 'k': 1e3, 'M': 1e6, 'G': 1e9,
    'T': 1e12, 'P': 1e15, 'E': 1e18, 'i': 1e0,
}

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
    if var_id in MSCRIPT_VAR_TYPES_DICT:
        return MSCRIPT_VAR_TYPES_DICT[var_id]
    warnings.warn(f'Unsupported VarType id "{var_id}"!')
    return VarType(var_id, 'unknown', '')


class MScriptVar:
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
        assert len(var) == 7
        return int(var, 16) - (2 ** 27)

    @staticmethod
    def parse_metadata(tokens: List[str]) -> Dict[str, int]:
        metadata = {}
        for token in tokens:
            if (len(token) == 2) and (token[0] == '1'):
                metadata['status'] = int(token[1], 16)
            if (len(token) == 3) and (token[0] == '2'):
                metadata['cr'] = int(token[1:], 16)
        return metadata


def parse_mscript_data_package(line: str) -> Optional[List[MScriptVar]]:
    # expects full line ending with '\n'
    if line.startswith('P') and line.endswith('\n'):
        return [MScriptVar(var) for var in line[1:-1].split(';') if var.strip()]
    return None


# =========================
# Helper function to convert float to SI string
# =========================
def to_si_string(value_str, unit='V'):
    try:
        val = float(value_str)
    except (ValueError, TypeError):
        return value_str

    if unit in ['V', 'V/s']:
        if val == 0:
            return "0"
        milli_value = val * 1000.0
        formatted = f"{milli_value:.12f}".rstrip('0').rstrip('.')
        if formatted in ('', '-0', '+0'):
            formatted = '0'
        return f"{formatted}m"
    if unit == 'Hz':
        if float(val).is_integer():
            return f"{int(val)}"
        return f"{val:g}"
    return value_str


# =========================
# SerialMeasurementRunner (your existing class; unchanged except minor robustness)
# =========================
class SerialMeasurementRunner:
    def __init__(self, script_path, log_callback=print):
        self.script_path = Path(script_path)
        self.data_points = []
        self.connection = None
        self.log = log_callback
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
            if port is None:
                return False
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
                    if not line:
                        continue
                    text = line.decode('utf-8', errors='ignore').strip()
                    if not text:
                        continue
                    self.log(text)
                    if text.startswith('P'):
                        self.parse_data_line(text)
                    if text in ['*', 'Measurement completed', 'Script completed']:
                        self.log("\nMeasurement completed")
                        break
                    if text.startswith('!'):
                        self.log(f"Device error: {text}")
                        if "abort" in text.lower():
                            break
                except serial.SerialException as e:
                    self.log(f"Serial Error: {e}")
                    break

            if not self.is_running:
                self.log("Measurement stopped by user.")
            return True
        except Exception as e:
            self.log(f"Error running script: {e}")
            return False

    def parse_data_line(self, line):
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
            except Exception as e:
                self.log(f"Error on disconnect: {e}")

    def execute(self):
        self.log("=" * 60)
        self.log(f"Starting measurement for: {self.script_path.name}")
        self.log("=" * 60)
        csv_path = None
        try:
            with open(self.script_path, 'r') as f:
                script = f.read()
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


# =========================
# NEW: EmStat MUX Live Plotter (embedded panel)
# =========================
class EmstatSerialClient:
    """
    Background serial reader + simple send_script framing for MethodSCRIPT.
    This keeps the Tk GUI responsive while measurement runs.
    """
    def __init__(self, log_cb=None):
        self.ser: Optional[serial.Serial] = None
        self.stop_event = threading.Event()
        self.rx_thread: Optional[threading.Thread] = None
        self.rx_lines: "collections.deque[str]" = collections.deque(maxlen=20000)
        self._lock = threading.Lock()
        self.log_cb = log_cb or (lambda msg: None)

    def open(self, port: str, baudrate: int, timeout: float = 0.1):
        self.ser = serial.Serial(
            port=port,
            baudrate=baudrate,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            timeout=timeout,
            write_timeout=1.0,
        )
        self.stop_event.clear()
        self.rx_thread = threading.Thread(target=self._reader_loop, daemon=True)
        self.rx_thread.start()

    def close(self):
        self.stop_event.set()
        if self.rx_thread and self.rx_thread.is_alive():
            self.rx_thread.join(timeout=1.0)
        try:
            if self.ser and self.ser.is_open:
                self.ser.close()
        except Exception:
            pass
        self.ser = None

    def is_open(self) -> bool:
        return bool(self.ser and self.ser.is_open)

    def _reader_loop(self):
        assert self.ser is not None
        buf = b""
        while not self.stop_event.is_set():
            try:
                chunk = self.ser.read(4096)
                if not chunk:
                    continue
                buf += chunk
                while b"\n" in buf:
                    line, buf = buf.split(b"\n", 1)
                    txt = line.decode("utf-8", errors="replace")
                    with self._lock:
                        self.rx_lines.append(txt)
            except Exception as e:
                self.log_cb(f"[MUX RX ERROR] {e}")
                break

    def pop_all_lines(self) -> List[str]:
        out = []
        with self._lock:
            while self.rx_lines:
                out.append(self.rx_lines.popleft())
        return out

    def send_script(self, script_text: str, mode: str = "execute"):
        if not self.is_open():
            raise RuntimeError("Serial port is not open")
        assert self.ser is not None

        script_text = script_text.replace("\r\n", "\n").replace("\r", "\n")
        raw_lines = [ln.rstrip() for ln in script_text.split("\n")]
        lines = [ln for ln in raw_lines if ln.strip() != ""]

        cmd = "e" if mode == "execute" else "l"
        if lines:
            first = lines[0].strip()
            if first in ("e", "l"):
                cmd = first
                lines = lines[1:]

        self.ser.write((cmd + "\n").encode("utf-8"))
        for ln in lines:
            self.ser.write((ln + "\n").encode("utf-8"))
        self.ser.write(b"\n")  # end of script
        self.ser.flush()


class MuxLivePlotterPanel(ttk.Frame):
    """
    Embedded panel:
    - Connect/disconnect to EmStat
    - Paste MethodSCRIPT and execute
    - Live plot ALL channels on one axis with different colors
      (channels are inferred by "split when X decreases" => next channel)
    """
    def __init__(self, parent, *, log_cb=None):
        super().__init__(parent)
        self.log_cb = log_cb or (lambda msg: None)

        self.client = EmstatSerialClient(log_cb=self.log_cb)

        # curve storage: channel -> {"x":[], "y":[]}
        self.curves: Dict[int, Dict[str, List[float]]] = {}
        self._current_curve = 0
        self._last_x: Optional[float] = None

        # settings
        self.x_var_index = tk.IntVar(value=0)
        self.y_var_index = tk.IntVar(value=1)
        self.enable_split = tk.BooleanVar(value=True)
        self.split_threshold = tk.DoubleVar(value=0.05)
        self.max_curves = tk.IntVar(value=16)

        self._build_ui()
        self.after(50, self._poll_serial)

    @staticmethod
    def _list_ports() -> List[str]:
        return [p.device for p in serial.tools.list_ports.comports()]

    def _build_ui(self):
        # --- connection row
        conn = ttk.LabelFrame(self, text="MUX Live Plotter (EmStat)")
        conn.pack(fill="x", padx=8, pady=6)

        ttk.Label(conn, text="Port:").grid(row=0, column=0, padx=4, pady=4, sticky="w")
        self.port_combo = ttk.Combobox(conn, width=16, values=self._list_ports())
        self.port_combo.grid(row=0, column=1, padx=4, pady=4, sticky="w")
        if self.port_combo["values"]:
            self.port_combo.current(0)

        ttk.Button(conn, text="Refresh", command=self._refresh_ports).grid(row=0, column=2, padx=4, pady=4)

        ttk.Label(conn, text="Baud:").grid(row=0, column=3, padx=4, pady=4, sticky="e")
        self.baud_entry = ttk.Entry(conn, width=10)
        self.baud_entry.insert(0, "230400")
        self.baud_entry.grid(row=0, column=4, padx=4, pady=4, sticky="w")

        self.btn_connect = ttk.Button(conn, text="Connect", command=self._connect)
        self.btn_connect.grid(row=0, column=5, padx=4, pady=4)

        self.btn_disconnect = ttk.Button(conn, text="Disconnect", command=self._disconnect, state="disabled")
        self.btn_disconnect.grid(row=0, column=6, padx=4, pady=4)

        self.status_lbl = ttk.Label(conn, text="Disconnected")
        self.status_lbl.grid(row=0, column=7, padx=8, pady=4, sticky="w")

        # --- script box + controls
        script_box = ttk.LabelFrame(self, text="MethodSCRIPT (paste & run)")
        script_box.pack(fill="both", expand=True, padx=8, pady=6)

        self.script_txt = tk.Text(script_box, height=10, wrap="none")
        self.script_txt.pack(fill="both", expand=True, padx=6, pady=6)

        btnrow = ttk.Frame(script_box)
        btnrow.pack(fill="x", padx=6, pady=(0, 6))
        ttk.Button(btnrow, text="Send & Execute", command=self._send_execute).pack(side="left")
        ttk.Button(btnrow, text="Clear Live Data", command=self._clear_data).pack(side="left", padx=6)

        # --- parse settings
        settings = ttk.LabelFrame(self, text="Packet mapping (pck_add order)")
        settings.pack(fill="x", padx=8, pady=6)
        r1 = ttk.Frame(settings)
        r1.pack(fill="x", padx=6, pady=4)
        ttk.Label(r1, text="X var index:").pack(side="left")
        ttk.Spinbox(r1, from_=0, to=64, textvariable=self.x_var_index, width=5).pack(side="left", padx=6)
        ttk.Label(r1, text="Y var index:").pack(side="left")
        ttk.Spinbox(r1, from_=0, to=64, textvariable=self.y_var_index, width=5).pack(side="left", padx=6)
        ttk.Label(r1, text="Max curves:").pack(side="left")
        ttk.Spinbox(r1, from_=1, to=256, textvariable=self.max_curves, width=6).pack(side="left", padx=6)

        r2 = ttk.Frame(settings)
        r2.pack(fill="x", padx=6, pady=4)
        ttk.Checkbutton(r2, text="Auto-split on X reset", variable=self.enable_split).pack(side="left")
        ttk.Label(r2, text="Threshold:").pack(side="left", padx=(10, 0))
        ttk.Entry(r2, textvariable=self.split_threshold, width=8).pack(side="left", padx=6)
        ttk.Label(r2, text="(e.g., 0.05 V)").pack(side="left")

        # --- plot
        plot_frame = ttk.LabelFrame(self, text="Live Plot (all channels)")
        plot_frame.pack(fill="both", expand=True, padx=8, pady=6)

        self.fig = Figure(figsize=(5, 4), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.ax.set_xlabel("X")
        self.ax.set_ylabel("Y")
        self.ax.grid(True)

        self.canvas = FigureCanvasTkAgg(self.fig, master=plot_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        # keep a line per curve
        self.lines: Dict[int, any] = {}

        # Prefill demo script (your MUX16 LSV loop example)
        demo = """e
set_gpio_cfg 0x3FFi 1
set_gpio 0x11i
var i
var c
var p
store_var i 0i aa
set_pgstat_chan 0
set_pgstat_mode 2
set_max_bandwidth 400
set_pot_range -1 1
set_cr 1m
set_autoranging 10u 1m
cell_on
loop i <= 0xFFi
 set_gpio i
 set_e -1000m
 wait 100m
 meas_loop_lsv p c -1 1 10m 1
 pck_start
 pck_add p
 pck_add c
 pck_end
 endloop
 add_var i 0x11i
endloop
on_finished:
cell_off
"""
        self.script_txt.insert("1.0", demo)

    def _refresh_ports(self):
        ports = self._list_ports()
        self.port_combo["values"] = ports
        if ports:
            self.port_combo.current(0)

    def _connect(self):
        port = self.port_combo.get().strip()
        if not port:
            messagebox.showerror("MUX Connect", "Select a serial port.")
            return
        try:
            baud = int(self.baud_entry.get().strip())
        except Exception:
            messagebox.showerror("MUX Connect", "Invalid baud rate.")
            return
        try:
            self.client.open(port=port, baudrate=baud)
        except Exception as e:
            messagebox.showerror("MUX Connect", f"Failed to open {port}: {e}")
            return

        self.status_lbl.config(text=f"Connected: {port} @ {baud}")
        self.btn_connect.config(state="disabled")
        self.btn_disconnect.config(state="normal")
        self.log_cb(f"[MUX] Connected to {port} @ {baud}")

    def _disconnect(self):
        self.client.close()
        self.status_lbl.config(text="Disconnected")
        self.btn_connect.config(state="normal")
        self.btn_disconnect.config(state="disabled")
        self.log_cb("[MUX] Disconnected")

    def _send_execute(self):
        if not self.client.is_open():
            messagebox.showerror("MUX Send", "Not connected.")
            return
        script = self.script_txt.get("1.0", "end").strip("\n")
        if not script.strip():
            messagebox.showerror("MUX Send", "Paste a MethodSCRIPT first.")
            return

        # reset split state for each run
        self._current_curve = 0
        self._last_x = None
        self.curves.clear()
        self._redraw_all_lines()

        try:
            self.client.send_script(script, mode="execute")
        except Exception as e:
            messagebox.showerror("MUX Send", f"Failed to send script: {e}")
            return

        self.log_cb("[MUX] Script sent. Live plotting...")

    def _clear_data(self):
        self.curves.clear()
        self._current_curve = 0
        self._last_x = None
        self._redraw_all_lines()
        self.log_cb("[MUX] Cleared live data")

    def _poll_serial(self):
        # Drain serial lines and parse packets
        for line in self.client.pop_all_lines():
            stripped = line.rstrip("\r")
            if not stripped:
                continue
            if stripped.startswith("P"):
                self._handle_packet(stripped)
            # optional: log T lines / others, but avoid spamming
            # elif stripped.startswith("T"):
            #     self.log_cb("[MUX] " + stripped[1:])
        self.after(50, self._poll_serial)

    def _handle_packet(self, packet_line: str):
        # We can reuse the same parser (expects newline)
        vars_list = parse_mscript_data_package(packet_line + "\n")
        if not vars_list:
            return

        xi = int(self.x_var_index.get())
        yi = int(self.y_var_index.get())
        if xi < 0 or yi < 0 or xi >= len(vars_list) or yi >= len(vars_list):
            return

        x = vars_list[xi].value
        y = vars_list[yi].value

        if self.enable_split.get() and self._last_x is not None:
            thr = float(self.split_threshold.get())
            # split when X decreases enough (typical LSV/CV reset)
            if x < (self._last_x - thr):
                self._current_curve += 1
                self._last_x = None

        maxc = max(1, int(self.max_curves.get()))
        if self._current_curve >= maxc:
            self._current_curve = maxc - 1

        c = self.curves.setdefault(self._current_curve, {"x": [], "y": []})
        c["x"].append(x)
        c["y"].append(y)
        self._last_x = x

        self._redraw_all_lines()

    def _redraw_all_lines(self):
        # Ensure we have a line object for each curve index present
        self.ax.clear()
        self.ax.grid(True)

        # label axes based on current var ids if possible
        # (best-effort; we don’t have the packet here, so keep generic)
        self.ax.set_xlabel("X")
        self.ax.set_ylabel("Y")

        # Plot all curves (each gets next color in matplotlib cycle)
        for curve_idx in sorted(self.curves.keys()):
            c = self.curves[curve_idx]
            label = f"Ch {curve_idx + 1}"
            self.ax.plot(c["x"], c["y"], linewidth=1, label=label)

        if self.curves:
            self.ax.legend(loc="best", fontsize=8)

        self.canvas.draw()


# =========================
# MAIN GUI
# =========================
class ElectrochemGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Electrochemistry Automation System")
        self.root.geometry("1600x900")

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

    # -------------------------
    # GUI layout
    # -------------------------
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

        self.plotter_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.plotter_frame, text="Plotter")
        self.setup_plotter_tab()

    # -------------------------
    # Method creation scripts
    # -------------------------
    def create_cv_methodscript(self):
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

        min_pot = min(begin_v, end_v) - amp_v
        max_pot = max(begin_v, end_v) + amp_v
        min_pot_mv, max_pot_mv = int(min_pot * 1000), int(max_pot * 1000)

        script_parts = [
            "e", "var c", "var p", "var f", "var r",
            "set_pgstat_mode 2", "set_max_bandwidth 1600",
            f"set_range_minmax da {min_pot_mv}m {max_pot_mv}m",
            "set_range ba 5m", "set_autoranging ba 100n 5m", "cell_on"
        ]

        if float(cond_time) > 0:
            script_parts.extend([
                f"# Equilibrate at {cond_pot} for {cond_time}s",
                f"set_e {cond_pot}", f"wait {cond_time}"
            ])

        if n_scans > 1:
            script_parts.extend([
                f"# SWV measurement loop for {n_scans} scans",
                "var scan_num", "store_var scan_num 0i i", f"loop scan_num < {n_scans}i",
                "\tadd_var scan_num 1i",
                '\tsend_string "C"', '\tsend_var scan_num',
                f"\tmeas_loop_swv p c f r {begin} {end} {step} {amplitude} {frequency}",
                "\t\tpck_start",
                "\t\t\tpck_add p",
                "\t\t\tpck_add c",
                "\t\t\tpck_add f",
                "\t\t\tpck_add r",
                "\t\tpck_end",
                "\tendloop",
                '\tsend_string "-"',
                "endloop"
            ])
        else:
            script_parts.extend([
                f"meas_loop_swv p c f r {begin} {end} {step} {amplitude} {frequency}",
                "\tpck_start",
                "\t\tpck_add p",
                "\t\tpck_add c",
                "\t\tpck_add f",
                "\t\tpck_add r",
                "\tpck_end",
                "endloop"
            ])

        script_parts.extend(["on_finished:", "cell_off"])
        return "\n".join(script_parts)

    # -------------------------
    # Method tab layout (UPDATED: adds live plotter to the RIGHT)
    # -------------------------
    def setup_method_tab(self):
        # 3-column grid:
        # [Technique buttons] [Parameters] [MUX Live Plotter panel]
        self.method_frame.columnconfigure(0, weight=0)
        self.method_frame.columnconfigure(1, weight=1)
        self.method_frame.columnconfigure(2, weight=1)
        self.method_frame.rowconfigure(0, weight=1)

        # LEFT: technique picker
        left_frame = ttk.Frame(self.method_frame)
        left_frame.grid(row=0, column=0, sticky="nsw", padx=6, pady=6)

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

        # MIDDLE: parameters (your existing params_frame)
        self.params_frame = ttk.LabelFrame(self.method_frame, text="Parameters", padding=10)
        self.params_frame.grid(row=0, column=1, sticky="nsew", padx=6, pady=6)

        # RIGHT: embedded MUX live plotter panel
        mux_frame = ttk.LabelFrame(self.method_frame, text="MUX Live Plotter", padding=0)
        mux_frame.grid(row=0, column=2, sticky="nsew", padx=6, pady=6)

        self.mux_panel = MuxLivePlotterPanel(
            mux_frame,
            log_cb=self.log_message  # reuse the main log function
        )
        self.mux_panel.pack(fill="both", expand=True)

        # Default technique
        self.show_cv_params()

    def check_device(self):
        ports = list(serial.tools.list_ports.comports())
        if ports:
            self.device_status.config(text="Devices found (check console)", foreground="green")
            print("Available serial devices:\n" + "\n".join([f"{p.device}: {p.description}" for p in ports]))
        else:
            self.device_status.config(text="No devices found", foreground="red")

    def clear_params_frame(self):
        for widget in self.params_frame.winfo_children():
            widget.destroy()

    def show_cv_params(self):
        self.clear_params_frame()
        self.current_technique = "CV"
        self.cv_params = {}
        params = [
            ("Begin Potential (V):", "begin_potential", "0"),
            ("Vertex 1 (V):", "vertex1", "-0.5"),
            ("Vertex 2 (V):", "vertex2", "0.5"),
            ("Step Potential (V):", "step_potential", "0.002"),
            ("Scan Rate (V/s):", "scan_rate", "0.1"),
            ("Number of Scans:", "n_scans", "1"),
            ("Conditioning Potential (V):", "cond_potential", "0"),
            ("Conditioning Time (s):", "cond_time", "0"),
        ]
        for i, (label, key, default) in enumerate(params):
            ttk.Label(self.params_frame, text=label).grid(row=i, column=0, sticky='w', pady=2)
            entry = ttk.Entry(self.params_frame, width=15)
            entry.insert(0, default)
            entry.grid(row=i, column=1, pady=2)
            self.cv_params[key] = entry

        button_frame = ttk.Frame(self.params_frame)
        button_frame.grid(row=len(params), column=0, columnspan=2, pady=20)
        ttk.Button(button_frame, text="Generate Script", command=self.generate_cv_script).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Run Now", command=self.run_cv_immediately).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Add to Queue", command=self.add_cv_to_queue).pack(side='left', padx=5)

    def show_swv_params(self):
        self.clear_params_frame()
        self.current_technique = "SWV"
        self.swv_params = {}
        params = [
            ("Begin Potential (V):", "begin_potential", "-0.5"),
            ("End Potential (V):", "end_potential", "0.5"),
            ("Step Potential (V):", "step_potential", "0.002"),
            ("Amplitude (V):", "amplitude", "0.02"),
            ("Frequency (Hz):", "frequency", "15"),
            ("Number of Scans:", "n_scans", "1"),
            ("Conditioning Potential (V):", "cond_potential", "0"),
            ("Conditioning Time (s):", "cond_time", "0"),
        ]
        for i, (label, key, default) in enumerate(params):
            ttk.Label(self.params_frame, text=label).grid(row=i, column=0, sticky='w', pady=2)
            entry = ttk.Entry(self.params_frame, width=15)
            entry.insert(0, default)
            entry.grid(row=i, column=1, pady=2)
            self.swv_params[key] = entry

        button_frame = ttk.Frame(self.params_frame)
        button_frame.grid(row=len(params), column=0, columnspan=2, pady=20)
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

    # -------------------------
    # Script preview tab
    # -------------------------
    def setup_script_tab(self):
        text_frame = ttk.Frame(self.script_frame)
        text_frame.pack(fill='both', expand=True, padx=10, pady=5)
        self.script_text = tk.Text(text_frame, wrap='none', font=('Courier', 11))
        self.script_text.pack(fill='both', expand=True)

    def update_script_preview(self, script):
        self.script_text.delete(1.0, tk.END)
        self.script_text.insert(1.0, script)

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

    # -------------------------
    # Plotter tab (CSV)
    # -------------------------
    def setup_plotter_tab(self):
        plot_controls = ttk.Frame(self.plotter_frame)
        plot_controls.pack(side='top', fill='x', pady=5, padx=5)
        ttk.Button(plot_controls, text="Load and Plot CSV", command=self.load_and_plot_csv).pack(side='left')

        self.fig = Figure(figsize=(8, 6), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.ax.set_title('Voltammogram')
        self.ax.set_xlabel('Potential (V)')
        self.ax.set_ylabel('Current (µA)')
        self.ax.grid(True)

        self.canvas = FigureCanvasTkAgg(self.fig, master=self.plotter_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)

    def load_and_plot_csv(self):
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

    # -------------------------
    # Queue & Execution tab (IMPORTANT: contains save_queue/load_queue that your error complained about)
    # -------------------------
    def setup_queue_tab(self):
        main_pane = ttk.PanedWindow(self.queue_frame, orient=tk.VERTICAL)
        main_pane.pack(fill='both', expand=True)

        top_frame = ttk.Frame(main_pane)
        main_pane.add(top_frame, weight=1)

        bottom_frame = ttk.Frame(main_pane)
        main_pane.add(bottom_frame, weight=1)

        control_frame = ttk.Frame(top_frame)
        control_frame.pack(pady=10, fill='x', padx=10)

        ttk.Button(control_frame, text="Run Queue", command=self.run_queue).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Stop", command=self.stop_queue).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Save Queue", command=self.save_queue).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Load Queue", command=self.load_queue).pack(side='left', padx=5)
        ttk.Button(control_frame, text="Clear Queue", command=self.clear_queue).pack(side='left', padx=5)

        self.queue_tree = ttk.Treeview(top_frame, columns=('Type', 'Status', 'Details'), show='tree headings', height=8)
        self.queue_tree.heading('#0', text='#')
        self.queue_tree.heading('Type', text='Type')
        self.queue_tree.heading('Status', text='Status')
        self.queue_tree.heading('Details', text='Details')
        self.queue_tree.column('#0', width=50)
        self.queue_tree.column('Type', width=150)
        self.queue_tree.column('Status', width=100)
        self.queue_tree.column('Details', width=400)
        self.queue_tree.pack(fill='both', expand=True, padx=10, pady=5)

        log_frame = ttk.LabelFrame(bottom_frame, text="Live Output Log")
        log_frame.pack(fill='both', expand=True, padx=10, pady=5)
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10)
        self.log_text.pack(fill='both', expand=True)
        self.log_text.config(state='disabled')

        self.status_label = ttk.Label(self.queue_frame, text="Status: Ready", relief='sunken')
        self.status_label.pack(side='bottom', fill='x', padx=10, pady=5)

    def clear_log(self):
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')

    def log_message(self, message):
        def append_message():
            try:
                self.log_text.config(state='normal')
                self.log_text.insert(tk.END, message + '\n')
                self.log_text.see(tk.END)
                self.log_text.config(state='disabled')
            except Exception:
                pass
        self.root.after(0, append_message)
        print(message)

    def update_status(self, message):
        self.status_label.config(text=f"Status: {message}")

    # ---- Queue persistence (THIS FIXES YOUR AttributeError) ----
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

    # ---- Queue execution ----
    def refresh_queue_display(self):
        for item in self.queue_tree.get_children():
            self.queue_tree.delete(item)
        for i, item in enumerate(self.measurement_queue):
            self.queue_tree.insert('', 'end', text=str(i+1),
                                   values=(item.get('type', ''),
                                           str(item.get('status', 'pending')).upper(),
                                           item.get('details', '')))

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

    def add_cv_to_queue(self):
        script = self.generate_cv_script()
        if script:
            self.add_to_queue("CV", script)

    def add_swv_to_queue(self):
        script = self.generate_swv_script()
        if script:
            self.add_to_queue("SWV", script)

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
        if script:
            self.run_script_immediately("CV", script)

    def run_swv_immediately(self):
        script = self.generate_swv_script()
        if script:
            self.run_script_immediately("SWV", script)

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
            messagebox.showwarning("Busy", "Another measurement is currently running.")
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

    def run_queue(self):
        if not self.measurement_queue:
            messagebox.showwarning("Empty Queue", "No items in queue")
            return
        if self.is_running:
            messagebox.showwarning("Already Running", "Queue is already running")
            return
        self.is_running = True
        self.clear_log()
        self.queue_thread = threading.Thread(target=self.execute_queue, daemon=True)
        self.queue_thread.start()

    def execute_queue(self):
        for i, item in enumerate(list(self.measurement_queue)):
            if not self.is_running:
                self.log_message("Queue execution stopped by user.")
                break

            self.measurement_queue[i]['status'] = 'running'
            self.root.after(0, self.refresh_queue_display)
            self.root.after(0, self.update_status, f"Running: {item.get('type','')} - {item.get('details', '')}")

            csv_path = None
            success = False
            try:
                if item.get('type') == 'PAUSE':
                    seconds = float(item.get('pause_seconds', 0))
                    self.log_message(f"Queue pause start: {seconds:.1f} sec")
                    pause_completed = self.execute_pause(seconds)
                    self.measurement_queue[i]['status'] = 'completed' if pause_completed else 'stopped'
                    success = pause_completed
                else:
                    self.current_runner = SerialMeasurementRunner(Path(item['script_path']), log_callback=self.log_message)
                    success, csv_path = self.current_runner.execute()
                    self.measurement_queue[i]['status'] = 'completed' if success else 'failed'
                    self.current_runner = None
            except Exception as e:
                self.measurement_queue[i]['status'] = 'failed'
                self.log_message(f"CRITICAL ERROR in queue execution: {e}")

            if csv_path:
                self.root.after(0, self.plot_data, csv_path)

            self.root.after(0, self.refresh_queue_display)
            time.sleep(1)

        self.is_running = False
        self.root.after(0, self.update_status, "Queue Complete")

    def stop_queue(self):
        if not self.is_running:
            return
        self.is_running = False
        if self.current_runner:
            self.current_runner.stop()
        self.update_status("Queue Stopped")

    def clear_queue(self):
        if self.is_running:
            messagebox.showwarning("Queue Running", "Cannot clear queue while running")
            return
        self.measurement_queue = []
        self.refresh_queue_display()
        self.update_status("Queue Cleared")

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

    # -------------------------
    # Pump tab (kept minimal here; your original is long—leave as-is in your repo if you want)
    # If you need the FULL pump UI exactly as in your script, keep your existing setup_pump_tab()
    # and pump methods; this stub prevents crashes if PumpCtrl exists.
    # -------------------------
    def setup_pump_tab(self):
        if not PUMP_AVAILABLE or PumpCtrl is None:
            ttk.Label(self.pump_frame, text="Pump controls unavailable.").pack(pady=20)
            return
        ttk.Label(self.pump_frame, text="Pump tab is enabled, but this file uses a minimal stub.").pack(pady=20)
        ttk.Label(self.pump_frame, text="If you want your full pump UI, paste your original setup_pump_tab() here.").pack(pady=6)


def main():
    root = tk.Tk()
    app = ElectrochemGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
