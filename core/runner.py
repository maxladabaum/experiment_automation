"""
core/runner.py — SerialMeasurementRunner.

Handles all serial communication with the PalmSens device:
  - port auto-detection
  - connecting / disconnecting
  - sending a MethodSCRIPT and streaming back data lines
  - parsing data packets via core.mscript_parser
  - saving results to CSV

Zero GUI imports.  All user-facing output goes through the log_callback
and data_callback callables so the GUI can wire them to whatever it likes.
"""

import csv
import time
import traceback
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional, Tuple

import serial
import serial.tools.list_ports

from .mscript_parser import parse_mscript_data_package
from config import DATA_DIR, DEVICE_KEYWORDS, DEVICE_BAUDRATE


class SerialMeasurementRunner:
    """Run a single MethodSCRIPT measurement over a serial port.

    Parameters
    ----------
    script_path:
        Path to the ``.ms`` file to run.
    log_callback:
        Callable that accepts a ``str`` — used for all log output.
        Defaults to ``print``.
    data_callback:
        Optional callable that receives a ``{'potential': float,
        'current': float}`` dict for each parsed data point (used for
        live plotting).
    pump_com_port:
        The COM port used by the pump (as a string like ``"COM8"`` or
        an int) so it can be deprioritised when auto-detecting the
        measurement device.
    """

    def __init__(
        self,
        script_path,
        log_callback: Callable[[str], None] = print,
        data_callback: Optional[Callable[[dict], None]] = None,
        pump_com_port=None,
    ):
        self.script_path    = Path(script_path)
        self.data_points    = []
        self.connection     = None
        self.log            = log_callback
        self.data_callback  = data_callback
        self.is_running     = True
        self.partial_packet = ""
        self._pump_com_port = pump_com_port

        # Prepare per-day data folder
        self._data_base = Path(DATA_DIR)
        self._data_base.mkdir(exist_ok=True)
        date_folder = self._data_base / datetime.now().strftime("%Y-%m-%d")
        date_folder.mkdir(exist_ok=True)
        self.data_folder = date_folder

    # ── Port discovery ────────────────────────────────────────────────────────

    def find_device_port(self) -> Optional[str]:
        self.log("Scanning for devices...")
        ports = serial.tools.list_ports.comports(include_links=False)
        candidates = []
        for port in ports:
            self.log(f"  Found port: {port.description} ({port.device})")
            if any(kw in port.description for kw in DEVICE_KEYWORDS):
                candidates.append(port.device)

        if not candidates:
            self.log("ERROR: No measurement device found")
            return None

        # Deprioritise the pump port so we never accidentally send a script there
        pump_upper = None
        if self._pump_com_port is not None:
            try:
                pump_upper = f"COM{int(self._pump_com_port)}".upper()
            except (TypeError, ValueError):
                pump_upper = str(self._pump_com_port).upper()

        candidates.sort(key=lambda dev: (pump_upper is not None and dev.upper() == pump_upper, dev))

        if len(candidates) > 1:
            self.log(f"Multiple devices found: {candidates}")
        self.log(f"Using device: {candidates[0]}")
        return candidates[0]

    # ── Connection ────────────────────────────────────────────────────────────

    def connect(self, port: Optional[str] = None) -> bool:
        if port is None:
            port = self.find_device_port()
        if port is None:
            return False
        try:
            self.log(f"Connecting to {port}...")
            self.connection = serial.Serial(
                port=port, baudrate=DEVICE_BAUDRATE, timeout=1, write_timeout=1
            )
            time.sleep(2)
            self.connection.reset_input_buffer()
            self.connection.reset_output_buffer()
            self.connection.write(b"t\n")
            response = self.connection.readline()
            if response:
                self.log(f"Device responded: {response.decode('utf-8', errors='ignore').strip()}")
                return True
            self.log("No response from device")
            return False
        except Exception as exc:
            self.log(f"Connection failed: {exc}")
            return False

    def disconnect(self):
        if self.connection and self.connection.is_open:
            try:
                self.connection.close()
                self.log("Disconnected from device")
            except Exception as exc:
                self.log(f"Error on disconnect: {exc}")

    def stop(self):
        """Signal the runner to stop after the current data read."""
        self.is_running = False

    # ── Script execution ──────────────────────────────────────────────────────

    def run_script(self, script: str) -> bool:
        if not self.connection:
            self.log("ERROR: Not connected to device")
            return False
        try:
            self.log("Sending script to device...")
            for line in script.strip().split("\n"):
                self.connection.write((line + "\n").encode("utf-8"))
                time.sleep(0.01)
            self.connection.write(b"\n")
            self.log("Script sent. Collecting data...")
            self.log("-" * 40)

            empty_reads = 0
            while self.is_running:
                try:
                    raw = self.connection.readline()
                    if not raw:
                        if self.partial_packet:
                            empty_reads += 1
                            if empty_reads >= 3:
                                self.log("Warning: incomplete data packet timed out")
                                break
                        continue
                    empty_reads = 0
                    text = raw.decode("utf-8", errors="ignore").rstrip("\r\n")
                    if not text:
                        continue

                    # Discard stale partial packet on new 'P' line
                    if self.partial_packet:
                        self.log("Warning: dropped incomplete data packet")
                        self.partial_packet = ""

                    self.log(text)

                    if text.startswith("P"):
                        if not self._is_complete_packet(text):
                            self.partial_packet = text
                            continue
                        self._parse_data_line(text)

                    if text in ("*", "Measurement completed", "Script completed"):
                        self.log("\nMeasurement completed")
                        break

                    if text.startswith("!"):
                        self.log(f"Device error: {text}")
                        if "abort" in text.lower():
                            break

                except serial.SerialException as exc:
                    self.log(f"Serial Error: {exc}")
                    break

            if not self.is_running:
                self.log("Measurement stopped by user.")
            return True

        except Exception as exc:
            self.log(f"Error running script: {type(exc).__name__}: {exc}")
            self.log(traceback.format_exc())
            return False

    # ── Data parsing ──────────────────────────────────────────────────────────

    def _parse_data_line(self, line: str):
        package = parse_mscript_data_package(line + "\n")
        if not package:
            return
        try:
            point = {}
            for var in package:
                if var.id in ("ab", "da"):
                    point["potential"] = var.value
                elif var.id == "ba":
                    point["current"] = var.value * 1e6   # A → µA
            if "potential" in point and "current" in point:
                self.data_points.append(point)
                if self.data_callback:
                    try:
                        self.data_callback(point)
                    except Exception as exc:
                        self.log(f"Live plot callback error: {exc}")
        except Exception as exc:
            self.log(f"Error parsing data package: {line!r} → {exc}")

    @staticmethod
    def _is_complete_packet(line: str) -> bool:
        if not line.startswith("P"):
            return False
        parts = line[1:].split(";")
        return bool(parts) and all(len(p) >= 10 for p in parts)

    # ── CSV output ────────────────────────────────────────────────────────────

    def save_data_to_csv(self, meas_tag: Optional[str] = None) -> Optional[Path]:
        """Write collected data points to a CSV file.

        Parameters
        ----------
        meas_tag:
            Sequential tag supplied by :class:`~core.session.SessionState`
            (e.g. ``"meas_007"``).  If omitted a timestamp fallback is used.

        Returns
        -------
        Path to the written file, or ``None`` if there was no data.
        """
        if not self.data_points:
            self.log("No data to save")
            return None

        base     = self.script_path.stem
        tag      = meas_tag or datetime.now().strftime("meas_%H%M%S")
        csv_path = self.data_folder / f"{base}_{tag}.csv"

        with open(csv_path, "w", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=["potential", "current"])
            writer.writerow({"potential": "Potential (V)", "current": "Current (uA)"})
            writer.writerows(self.data_points)

        self.log(f"\nData saved to: {csv_path}")
        return csv_path

    # ── High-level entry point ────────────────────────────────────────────────

    def execute(self, meas_tag: Optional[str] = None) -> Tuple[bool, Optional[Path]]:
        """Connect, send the script, collect data, save CSV, disconnect.

        Returns
        -------
        (success: bool, csv_path: Path | None)
        """
        self.log("=" * 60)
        self.log(f"Starting measurement for: {self.script_path.name}")
        self.log("=" * 60)

        try:
            script = self.script_path.read_text()
        except Exception as exc:
            self.log(f"ERROR: Failed to read script: {exc}")
            return False, None

        if not self.connect():
            self.log("ERROR: Failed to connect to device")
            return False, None

        csv_path = None
        success  = False
        try:
            if self.run_script(script):
                if self.data_points:
                    csv_path = self.save_data_to_csv(meas_tag=meas_tag)
                self.log(f"Total data points: {len(self.data_points)}")
                success = True
        finally:
            self.disconnect()

        return success, csv_path
