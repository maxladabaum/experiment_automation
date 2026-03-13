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
import math
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
    data_folder:
        Optional override for where CSV and raw packet logs are written.
    save_raw_packets:
        If True, write raw device output lines to a ``*_raw.txt`` file.
    invert_current:
        If True, multiply measured currents by -1 before processing.
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
        data_folder: Optional[Path] = None,
        save_raw_packets: bool = False,
        simulate_measurements: bool = False,
        invert_current: bool = False,
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
        self.save_raw_packets = bool(save_raw_packets)
        self.simulate_measurements = bool(simulate_measurements)
        self.invert_current = bool(invert_current)
        self._raw_fh = None
        self._fallback_tag_counter = 0

        # Prepare per-day data folder
        if data_folder is not None:
            self.data_folder = Path(data_folder)
            self.data_folder.mkdir(parents=True, exist_ok=True)
        else:
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
        if pump_upper is not None:
            self.log(f"Using first device: {candidates[0]} (pump port {pump_upper} deprioritized)")
        else:
            self.log(f"Using first device: {candidates[0]}")
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

            buffer = b""
            packet_start_time = None
            packet_timeout_sec = 1.0
            idle_timeout_sec = 5.0
            last_data_time = time.time()
            measurement_completed = False

            while self.is_running:
                try:
                    waiting = self.connection.in_waiting
                    chunk = self.connection.read(waiting or 1)

                    if not chunk:
                        if measurement_completed:
                            break

                        if (time.time() - last_data_time) >= idle_timeout_sec:
                            self.log("Warning: data stream idle timed out")
                            break

                        if buffer:
                            if packet_start_time is None:
                                packet_start_time = time.time()
                            elif (time.time() - packet_start_time) >= packet_timeout_sec:
                                self.log("Warning: incomplete data packet timed out (dropping buffer)")
                                buffer = b""
                                packet_start_time = None
                        continue

                    last_data_time = time.time()
                    buffer += chunk
                    packet_start_time = None

                    while b"\n" in buffer:
                        line_bytes, _, buffer = buffer.partition(b"\n")
                        text_line = line_bytes.decode("utf-8", errors="ignore").rstrip("\r")
                        if not text_line:
                            continue

                        if self.partial_packet:
                            if text_line.startswith("P"):
                                self.log("Warning: dropped incomplete data packet")
                                self.partial_packet = ""
                            else:
                                combined = self.partial_packet + text_line
                                if self._is_complete_packet(combined):
                                    self._parse_data_line(combined)
                                    self.partial_packet = ""
                                else:
                                    self.partial_packet = combined
                                continue

                        if not text_line.startswith("P"):
                            self.log(text_line)

                        if self._raw_fh is not None:
                            try:
                                self._raw_fh.write(text_line + "\n")
                            except OSError:
                                self.log("Warning: failed to write raw packet log.")
                                self._raw_fh = None

                        if text_line.startswith("P"):
                            if not self._is_complete_packet(text_line):
                                self.partial_packet = text_line
                                continue
                            self._parse_data_line(text_line)

                        if text_line in ("*", "Measurement completed", "Script completed"):
                            self.log("\nMeasurement completed")
                            measurement_completed = True
                            self.partial_packet = ""
                            buffer = b""
                            break

                        if text_line.startswith("!"):
                            self.log(f"Device error: {text_line}")
                            if "abort" in text_line.lower():
                                break

                    if measurement_completed:
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
            currents = []
            for var in package:
                if var.id in ("ab", "da"):
                    point["potential"] = var.value
                elif var.id == "ba":
                    current_amp = var.value
                    if self.invert_current:
                        current_amp = -current_amp
                    currents.append(current_amp * 1e6)   # A -> uA

            if currents:
                if len(currents) >= 3:
                    # PalmSens SWV typically emits three WE current values.
                    # Empirically: [diff (negated), reverse, forward].
                    point["current_diff"] = -currents[0]
                    point["current_reverse"] = currents[1]
                    point["current_forward"] = currents[2]
                    point["current"] = point["current_diff"]
                elif len(currents) == 2:
                    point["current_forward"] = currents[0]
                    point["current_reverse"] = currents[1]
                    point["current"] = currents[1]
                else:
                    point["current"] = currents[0]

            if "potential" in point and "current" in point:
                self.data_points.append(point)
                if self.data_callback:
                    try:
                        self.data_callback(point)
                    except Exception as exc:
                        self.log(f"Live plot callback error: {exc}")
        except Exception as exc:
            self.log(f"Error parsing data package: {line!r} -> {exc}")

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
        tag      = meas_tag or self._fallback_meas_tag()
        csv_path = self.data_folder / f"{base}_{tag}.csv"

        with open(csv_path, "w", newline="") as fh:
            fieldnames = ["potential", "current"]
            label_map = {
                "potential": "Potential (V)",
                "current": "Current (uA)",
            }
            if any("current_forward" in dp for dp in self.data_points):
                fieldnames.append("current_forward")
                label_map["current_forward"] = "Current Forward (uA)"
            if any("current_reverse" in dp for dp in self.data_points):
                fieldnames.append("current_reverse")
                label_map["current_reverse"] = "Current Reverse (uA)"
            if any("current_diff" in dp for dp in self.data_points):
                fieldnames.append("current_diff")
                label_map["current_diff"] = "Current Diff (uA)"

            writer = csv.DictWriter(fh, fieldnames=fieldnames)
            writer.writerow({key: label_map.get(key, key) for key in fieldnames})
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

        tag = meas_tag or self._fallback_meas_tag()
        self._save_used_method_copy(script, tag)
        if self.simulate_measurements:
            self.log("Simulation mode: skipping serial connection and simulating data.")
            return self._execute_simulated_run(script, tag)

        if not self.connect():
            self.log("ERROR: Failed to connect to device")
            return False, None

        if self.save_raw_packets:
            raw_path = self.data_folder / f"{self.script_path.stem}_{tag}_raw.txt"
            try:
                self._raw_fh = open(raw_path, "w", encoding="utf-8", newline="\n")
                self.log(f"Raw packet log: {raw_path}")
            except OSError as exc:
                self.log(f"Warning: could not open raw packet log: {exc}")
                self._raw_fh = None

        csv_path = None
        success  = False
        try:
            if self.run_script(script):
                if self.data_points:
                    csv_path = self.save_data_to_csv(meas_tag=tag)
                self.log(f"Total data points: {len(self.data_points)}")
                success = True
        finally:
            self.disconnect()
            if self._raw_fh is not None:
                try:
                    self._raw_fh.close()
                except OSError:
                    pass
                self._raw_fh = None

        return success, csv_path

    def _fallback_meas_tag(self) -> str:
        """Build a collision-resistant fallback tag when SessionState is not used."""
        self._fallback_tag_counter += 1
        now = datetime.now()
        return f"meas_{now:%Y%m%d_%H%M%S}_{self._fallback_tag_counter:04d}"

    def _save_used_method_copy(self, script: str, tag: str) -> Optional[Path]:
        """Save the exact method text used for this run under data_folder/methods_used."""
        try:
            methods_used_dir = self.data_folder / "methods_used"
            methods_used_dir.mkdir(parents=True, exist_ok=True)
            base_name = f"{self.script_path.stem}_{tag}.ms"
            path = methods_used_dir / base_name
            if path.exists():
                idx = 2
                while True:
                    candidate = methods_used_dir / f"{self.script_path.stem}_{tag}_{idx:02d}.ms"
                    if not candidate.exists():
                        path = candidate
                        break
                    idx += 1
            path.write_text(script, encoding="utf-8")
            self.log(f"Method snapshot: {path}")
            return path
        except Exception as exc:
            # Snapshot logging should never block measurement execution.
            self.log(f"Warning: could not save method snapshot: {exc}")
            return None

    def _execute_simulated_run(self, script: str, tag: str) -> Tuple[bool, Optional[Path]]:
        points = self._generate_simulated_points(script)
        if not points:
            self.log("Simulation generated no points.")
            return True, None

        for point in points:
            if not self.is_running:
                self.log("Simulated measurement stopped by user.")
                break
            self.data_points.append(point)
            if self.data_callback:
                try:
                    self.data_callback(point)
                except Exception as exc:
                    self.log(f"Live plot callback error: {exc}")
            time.sleep(0.01)

        csv_path = self.save_data_to_csv(meas_tag=tag) if self.data_points else None
        self.log(f"Total data points: {len(self.data_points)}")
        return True, csv_path

    def _generate_simulated_points(self, script: str) -> list:
        script_lower = script.lower()
        if "meas_loop_swv" in script_lower:
            return self._sim_swv_points()
        return self._sim_cv_points()

    @staticmethod
    def _sim_cv_points() -> list:
        points = []
        n_half = 80
        start, vertex, end = -0.5, 0.5, -0.5
        for i in range(n_half):
            frac = i / max(1, n_half - 1)
            p = start + (vertex - start) * frac
            c = 25.0 * math.sin((p + 0.5) * math.pi) + 3.0 * math.sin(8.0 * p)
            points.append({"potential": p, "current": c})
        for i in range(n_half):
            frac = i / max(1, n_half - 1)
            p = vertex + (end - vertex) * frac
            c = -20.0 * math.sin((p + 0.5) * math.pi) + 2.5 * math.sin(7.5 * p)
            points.append({"potential": p, "current": c})
        return points

    @staticmethod
    def _sim_swv_points() -> list:
        points = []
        n = 120
        start, end = -0.5, 0.5
        for i in range(n):
            frac = i / max(1, n - 1)
            p = start + (end - start) * frac
            peak = 40.0 * math.exp(-((p - 0.1) ** 2) / 0.02)
            baseline = 4.0 * math.sin(10.0 * p)
            forward = peak + baseline
            reverse = -0.5 * peak + 0.6 * baseline
            diff = forward - reverse
            points.append(
                {
                    "potential": p,
                    "current": diff,
                    "current_forward": forward,
                    "current_reverse": reverse,
                    "current_diff": diff,
                }
            )
        return points
