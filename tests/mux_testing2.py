import threading
import queue
import time
import re
import math
from dataclasses import dataclass, field
from typing import List, Dict, Optional

import tkinter as tk
from tkinter import ttk, messagebox

import serial
import serial.tools.list_ports

import matplotlib
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure


# -------------------------
# Parsing helpers (per MethodSCRIPT manual)
# -------------------------

SI_FACTORS = {
    'a': 1e-18,
    'f': 1e-15,
    'p': 1e-12,
    'n': 1e-9,
    'u': 1e-6,
    'm': 1e-3,
    ' ': 1.0,
    'k': 1e3,
    'M': 1e6,
    'G': 1e9,
    'T': 1e12,
    'P': 1e15,
    'E': 1e18,
    'i': 1.0,  # integer marker (handled separately)
}

OFFSET_2P27 = 2 ** 27  # 0x8000000


def decode_ms_value(hhhhhhhp: str) -> float:
    """
    Decode a MethodSCRIPT measurement data value of the form HHHHHHHp where:
      - HHHHHHH is (typically) 7 hex chars representing a 28-bit unsigned value with offset 2^27
      - p is SI prefix character or 'i' for integer
    """
    s = hhhhhhhp.strip()
    if s.endswith("nan"):
        return float("nan")
    if len(s) < 2:
        return float("nan")

    prefix = s[-1]
    hex_part = s[:-1].strip()
    if not re.fullmatch(r"[0-9A-Fa-f]+", hex_part):
        return float("nan")

    raw = int(hex_part, 16)
    signed = raw - OFFSET_2P27

    if prefix == 'i':
        return float(signed)

    factor = SI_FACTORS.get(prefix, 1.0)
    return signed * factor


@dataclass
class PacketVar:
    vartype: str
    value_str: str
    value: float
    meta: str = ""


@dataclass
class MeasurementPacket:
    vars: List[PacketVar] = field(default_factory=list)
    raw_line: str = ""


def parse_measurement_packet_line(line: str) -> Optional[MeasurementPacket]:
    """
    Parse a measurement data package line that starts with 'P':
      P <Var1> ; <Var2> ; ...
    Each variable is:
      ttHHHHHHHp[,metadata...]
    """
    if not line.startswith("P"):
        return None

    body = line[1:].strip()
    if not body:
        return MeasurementPacket(vars=[], raw_line=line)

    parts = body.split(";")
    out = MeasurementPacket(vars=[], raw_line=line)

    for part in parts:
        part = part.strip()
        if not part:
            continue

        if "," in part:
            main, meta = part.split(",", 1)
            meta = meta.strip()
        else:
            main, meta = part, ""

        main = main.strip()
        if len(main) < 4:
            continue

        vartype = main[:2]
        value_str = main[2:]

        try:
            value = decode_ms_value(value_str)
        except Exception:
            value = float("nan")

        out.vars.append(PacketVar(vartype=vartype, value_str=value_str, value=value, meta=meta))

    return out


# -------------------------
# EmStat Pico connection + IO (serial protocol)
# -------------------------

class EmstatSerialClient:
    """
    Implements the communication pattern:
      - Send 'e\\n' or 'l\\n'
      - Send script lines each terminated by '\\n'
      - Send an empty line ('\\n') to end script
    """

    def __init__(self):
        self.ser: Optional[serial.Serial] = None
        self.rx_thread: Optional[threading.Thread] = None
        self.stop_event = threading.Event()
        self.rx_queue: "queue.Queue[str]" = queue.Queue()

    def open(self, port: str, baudrate: int, timeout: float = 0.1) -> None:
        self.ser = serial.Serial(
            port=port,
            baudrate=baudrate,
            bytesize=serial.EIGHTBITS,
            parity=serial.PARITY_NONE,
            stopbits=serial.STOPBITS_ONE,
            timeout=timeout,
        )
        self.stop_event.clear()
        self.rx_thread = threading.Thread(target=self._reader_loop, daemon=True)
        self.rx_thread.start()

    def close(self) -> None:
        self.stop_event.set()
        if self.rx_thread and self.rx_thread.is_alive():
            self.rx_thread.join(timeout=1.0)
        if self.ser and self.ser.is_open:
            try:
                self.ser.close()
            except Exception:
                pass
        self.ser = None

    def is_open(self) -> bool:
        return bool(self.ser and self.ser.is_open)

    def _reader_loop(self) -> None:
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
                    try:
                        txt = line.decode("utf-8", errors="replace")
                    except Exception:
                        txt = repr(line)
                    self.rx_queue.put(txt)
            except Exception as e:
                self.rx_queue.put(f"[RX ERROR] {e}")
                break

    def send_script(self, script_text: str, mode: str = "execute") -> None:
        """
        mode: 'execute' or 'load'
        Accepts pasted scripts that may or may not include a leading 'e'/'l' line.
        Ensures protocol framing: command line + script lines + empty line terminator.
        """
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
        self.ser.write(("\n").encode("utf-8"))
        self.ser.flush()


# -------------------------
# GUI + plotting (ALL CHANNELS ON ONE PLOT)
# -------------------------

class EmstatGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("EmStat Pico MUX16 — MethodSCRIPT Sender + Live Plotter (All Channels)")
        self.geometry("1200x800")

        self.client = EmstatSerialClient()

        # curves[ch] = {"x": [...], "y": [...]}
        self.curves: Dict[int, Dict[str, List[float]]] = {}

        # Based on pck_add order
        # Demo packet order: i, p, c -> X=1, Y=2, Channel=0
        self.x_var_index = tk.IntVar(value=1)
        self.y_var_index = tk.IntVar(value=2)
        self.channel_var_index = tk.IntVar(value=0)
        self.max_curves = tk.IntVar(value=16)

        # UI responsiveness during send
        self._tx_busy = False

        # Plot throttle
        self.plot_fps = 12.0
        self._last_draw_time = 0.0
        self._dirty_channels: set[int] = set()
        self._last_autoscale_time = 0.0

        self._build_ui()
        self.after(30, self._poll_rx_queue)
        self.after(50, self._plot_pump)

    def _build_ui(self):
        # Top: connection row
        conn = ttk.Frame(self)
        conn.pack(fill="x", padx=10, pady=8)

        ttk.Label(conn, text="Port:").pack(side="left")
        self.port_combo = ttk.Combobox(conn, width=25, values=self._list_ports())
        self.port_combo.pack(side="left", padx=6)
        if self.port_combo["values"]:
            self.port_combo.current(0)

        ttk.Button(conn, text="Refresh", command=self._refresh_ports).pack(side="left", padx=6)

        ttk.Label(conn, text="Baud:").pack(side="left", padx=(18, 0))
        self.baud_entry = ttk.Entry(conn, width=10)
        self.baud_entry.insert(0, "230400")
        self.baud_entry.pack(side="left", padx=6)

        self.connect_btn = ttk.Button(conn, text="Connect", command=self._connect)
        self.connect_btn.pack(side="left", padx=6)

        self.disconnect_btn = ttk.Button(conn, text="Disconnect", command=self._disconnect, state="disabled")
        self.disconnect_btn.pack(side="left", padx=6)

        self.status_lbl = ttk.Label(conn, text="Disconnected")
        self.status_lbl.pack(side="left", padx=18)

        # Middle: script editor + controls + log/plot
        mid = ttk.Panedwindow(self, orient="horizontal")
        mid.pack(fill="both", expand=True, padx=10, pady=8)

        left = ttk.Frame(mid)
        right = ttk.Frame(mid)
        mid.add(left, weight=3)
        mid.add(right, weight=2)

        ttk.Label(left, text="MethodSCRIPT (paste here):").pack(anchor="w")
        self.script_txt = tk.Text(left, height=18, wrap="none")
        self.script_txt.pack(fill="both", expand=True)

        btnrow = ttk.Frame(left)
        btnrow.pack(fill="x", pady=6)

        ttk.Button(btnrow, text="Send & Execute", command=self._send_execute).pack(side="left")
        ttk.Button(btnrow, text="Send & Load", command=lambda: self._send_execute(load_only=True)).pack(side="left", padx=6)
        ttk.Button(btnrow, text="Clear Data/Plots", command=self._clear_data).pack(side="left", padx=6)

        settings = ttk.LabelFrame(left, text="Plot parsing settings (based on pck_add order)")
        settings.pack(fill="x", pady=6)

        r1 = ttk.Frame(settings)
        r1.pack(fill="x", padx=8, pady=4)
        ttk.Label(r1, text="X var index:").pack(side="left")
        ttk.Spinbox(r1, from_=0, to=64, textvariable=self.x_var_index, width=5).pack(side="left", padx=6)
        ttk.Label(r1, text="Y var index:").pack(side="left", padx=(12, 0))
        ttk.Spinbox(r1, from_=0, to=64, textvariable=self.y_var_index, width=5).pack(side="left", padx=6)

        r2 = ttk.Frame(settings)
        r2.pack(fill="x", padx=8, pady=4)
        ttk.Label(r2, text="Channel var index:").pack(side="left")
        ttk.Spinbox(r2, from_=0, to=64, textvariable=self.channel_var_index, width=5).pack(side="left", padx=6)
        ttk.Label(r2, text="Max channels:").pack(side="left", padx=(12, 0))
        ttk.Spinbox(r2, from_=1, to=256, textvariable=self.max_curves, width=6).pack(side="left", padx=6)

        # Right: log + single plot
        ttk.Label(right, text="Device log / output:").pack(anchor="w")
        self.log_txt = tk.Text(right, height=12, wrap="word")
        self.log_txt.pack(fill="x", expand=False)

        plot_frame = ttk.LabelFrame(right, text="All channels (live)")
        plot_frame.pack(fill="both", expand=True, pady=(8, 0))

        self.fig = Figure(figsize=(5, 4), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.ax.grid(True)

        (self._dummy_line,) = self.ax.plot([], [])  # kept for safety; not used

        self.canvas = FigureCanvasTkAgg(self.fig, master=plot_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(fill="both", expand=True)

        # One line per channel, created lazily up to max_curves
        self.lines: Dict[int, any] = {}

        self._ensure_channel_lines(self.max_curves.get())

    def _list_ports(self) -> List[str]:
        return [p.device for p in serial.tools.list_ports.comports()]

    def _refresh_ports(self):
        ports = self._list_ports()
        self.port_combo["values"] = ports
        if ports:
            self.port_combo.current(0)

    def _connect(self):
        port = self.port_combo.get().strip()
        if not port:
            messagebox.showerror("Connect", "Select a serial port.")
            return
        try:
            baud = int(self.baud_entry.get().strip())
        except Exception:
            messagebox.showerror("Connect", "Invalid baud rate.")
            return

        try:
            self.client.open(port=port, baudrate=baud)
        except Exception as e:
            messagebox.showerror("Connect", f"Failed to open {port}: {e}")
            return

        self.status_lbl.config(text=f"Connected: {port} @ {baud}")
        self.connect_btn.config(state="disabled")
        self.disconnect_btn.config(state="normal")
        self._log(f"[INFO] Connected to {port} @ {baud}\n")

    def _disconnect(self):
        self.client.close()
        self.status_lbl.config(text="Disconnected")
        self.connect_btn.config(state="normal")
        self.disconnect_btn.config(state="disabled")
        self._log("[INFO] Disconnected\n")

    # -------------------------
    # Threaded send so UI stays responsive
    # -------------------------
    def _send_execute(self, load_only: bool = False):
        if not self.client.is_open():
            messagebox.showerror("Send", "Not connected.")
            return

        if self._tx_busy:
            messagebox.showwarning("Busy", "Already sending/running a script.")
            return

        script = self.script_txt.get("1.0", "end").strip("\n")
        if not script.strip():
            messagebox.showerror("Send", "Paste a MethodSCRIPT first.")
            return

        mode = "load" if load_only else "execute"

        self._tx_busy = True
        self._log(f"[TX] Sending script ({mode})...\n")

        def worker():
            try:
                self.client.send_script(script, mode=mode)
                self.after(0, lambda: self._log(f"[TX] Sent script ({mode}). Waiting for data...\n"))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Send", f"Failed to send script: {e}"))
            finally:
                self._tx_busy = False

        threading.Thread(target=worker, daemon=True).start()

    def _clear_data(self):
        self.curves.clear()
        self._dirty_channels.clear()
        for ch, ln in self.lines.items():
            ln.set_data([], [])
        self.ax.relim()
        self.ax.autoscale_view()
        self.canvas.draw()
        self._log("[INFO] Cleared data/plots\n")

    def _ensure_channel_lines(self, n: int):
        """
        Ensure we have line objects for channels 1..n (stored as 0-based keys 0..n-1).
        Uses matplotlib's default color cycle automatically (different colors).
        """
        n = max(1, int(n))
        for k in range(n):
            if k in self.lines:
                continue
            # Label for legend: Channel 1..n
            (ln,) = self.ax.plot([], [], linewidth=1, label=f"Ch {k+1}")
            self.lines[k] = ln

        # Rebuild legend (cheap enough at creation time)
        self.ax.legend(loc="best", fontsize=8)

    def _poll_rx_queue(self):
        while True:
            try:
                line = self.client.rx_queue.get_nowait()
            except queue.Empty:
                break

            stripped = line.rstrip("\r")

            if stripped == "":
                self._log("[RX] <end>\n")
                continue

            if stripped.startswith("[RX ERROR]"):
                self._log(stripped + "\n")
                continue

            if stripped.startswith("P"):
                pkt = parse_measurement_packet_line(stripped)
                if pkt:
                    self._handle_packet(pkt)
            elif stripped.startswith("T"):
                self._log("[RX] " + stripped[1:] + "\n")
            else:
                self._log("[RX] " + stripped + "\n")

        self.after(30, self._poll_rx_queue)

    def _handle_packet(self, pkt: MeasurementPacket):
        xi = self.x_var_index.get()
        yi = self.y_var_index.get()
        ci = self.channel_var_index.get()

        if xi < 0 or yi < 0 or ci < 0:
            return
        if xi >= len(pkt.vars) or yi >= len(pkt.vars) or ci >= len(pkt.vars):
            return

        x = pkt.vars[xi].value
        y = pkt.vars[yi].value
        ch_val = pkt.vars[ci].value

        if not math.isfinite(ch_val):
            return

        maxc = max(1, int(self.max_curves.get()))
        self._ensure_channel_lines(maxc)

        # Channel mapping:
        # - If value is 1..maxc -> use ch-1
        # - Otherwise fold into range
        ch_int = int(round(ch_val))
        if 1 <= ch_int <= maxc:
            key = ch_int - 1
        else:
            key = ch_int % maxc

        c = self.curves.setdefault(key, {"x": [], "y": []})
        c["x"].append(x)
        c["y"].append(y)

        self._dirty_channels.add(key)

        # Light axis labeling
        xvt = pkt.vars[xi].vartype
        yvt = pkt.vars[yi].vartype
        self.ax.set_xlabel(f"X (idx {xi}, vt {xvt})")
        self.ax.set_ylabel(f"Y (idx {yi}, vt {yvt})")

    def _plot_pump(self):
        now = time.time()
        min_dt = 1.0 / max(1e-6, float(self.plot_fps))

        if (now - self._last_draw_time) >= min_dt and self._dirty_channels:
            do_autoscale = (now - self._last_autoscale_time) > 0.5  # autoscale at most 2 Hz

            # Update only lines that received new data
            for k in list(self._dirty_channels):
                if k not in self.curves:
                    continue
                ln = self.lines.get(k)
                if ln is None:
                    continue
                c = self.curves[k]
                ln.set_data(c["x"], c["y"])

            if do_autoscale:
                self.ax.relim()
                self.ax.autoscale_view()
                self._last_autoscale_time = now

            self.canvas.draw()

            self._dirty_channels.clear()
            self._last_draw_time = now

        self.after(50, self._plot_pump)

    def _log(self, msg: str):
        self.log_txt.insert("end", msg)
        self.log_txt.see("end")


if __name__ == "__main__":
    app = EmstatGUI()

    # Demo script includes i in packet so plots can route per-channel live:
    # Packet order is: i, p, c -> X index 1, Y index 2, Channel index 0
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
 pck_add i
 pck_add p
 pck_add c
 pck_end
 endloop
 add_var i 0x11i
endloop
on_finished:
cell_off
"""
    app.script_txt.insert("1.0", demo)
    app.mainloop()

