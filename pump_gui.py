# pump_gui.py — Cavro Centris Controller (Real/Sim) — fixed early logging
# - Real mode (Windows): uses PumpCommServer.PumpComm (pure, dev=1)
# - Sim mode (any OS): built-in simulator (no hardware)
#
# Usage:
#   python pump_gui.py --sim     # simulation anywhere (Mac/Win/Linux)
#   python pump_gui.py           # real hardware on Windows (DLLs installed)

import sys, time, threading, argparse
import tkinter as tk
from tkinter import ttk, messagebox

# Optional COM imports (Windows real mode). In sim mode we don't need them.
HAS_COM = False
try:
    import pythoncom
    from win32com.client import gencache
    HAS_COM = True
except Exception:
    HAS_COM = False

PROGID = "PumpCommServer.PumpComm"

DEFAULT_COM_PORT = 8
DEFAULT_BAUD     = 9600
DEFAULT_DEV      = 1
DEFAULT_STEPS    = 100000   # "100K"
DEFAULT_SYRINGE  = 1250.0   # µL
SPEED_MIN, SPEED_MAX = 1, 40

# ============================= Simulator backend =============================
class SimPumpComm:
    def __init__(self, steps_per_stroke=DEFAULT_STEPS, syringe_ul=DEFAULT_SYRINGE):
        self.connected = False
        self.dev = DEFAULT_DEV
        self.com_port = None
        self.speed = 20
        self.valve_port = 1
        self.steps_per_stroke = int(steps_per_stroke)
        self.syringe_ul = float(syringe_ul)
        self.plunger_steps = 0
        self._last_answer = ""
        self._lock = threading.Lock()

    def PumpInitComm(self, com_port):
        with self._lock:
            self.com_port = int(com_port); self.connected = True

    def PumpExitComm(self):
        with self._lock:
            self.connected = False; self.com_port = None

    def PumpSendCommand(self, cmd, dev, _ans=""):
        return self._process(cmd)

    def PumpSendNoWait(self, cmd, dev):
        return self._process(cmd, extra_wait=False)

    def PumpGetLastAnswer(self, dev=None):
        return self._last_answer

    def _set_answer(self, s=""):
        self._last_answer = s
        return s

    def _duration_plunger(self, steps):
        steps_per_sec_at_20 = 1500.0
        rate = steps_per_sec_at_20 * (self.speed / 20.0)
        if rate <= 1: rate = 1.0
        return max(0.05, float(steps) / rate)

    def _duration_valve(self, from_port, to_port):
        if from_port == to_port: return 0.1
        base = 0.5
        return max(0.2, base * (0.8 if self.speed >= 25 else 1.0))

    def _process(self, cmd, extra_wait=True):
        with self._lock:
            if not self.connected: return self._set_answer("")
            c = cmd.strip().upper()

            if c == "ZR":
                time.sleep(0.8); self.plunger_steps = 0; return self._set_answer("")

            if c.startswith("S") and c.endswith("R") and c[1:-1].isdigit():
                val = int(c[1:-1]); self.speed = max(SPEED_MIN, min(SPEED_MAX, val))
                time.sleep(0.05); return self._set_answer("")

            if c.startswith("I") and c.endswith("R") and c[1:-1].isdigit():
                port = int(c[1:-1])
                if 1 <= port <= 9:
                    time.sleep(self._duration_valve(self.valve_port, port))
                    self.valve_port = port
                else:
                    time.sleep(0.05)
                return self._set_answer("")

            if c.startswith("A") and c.endswith("R") and c[1:-1].isdigit():
                target = max(0, int(c[1:-1]))
                target = min(self.steps_per_stroke, target)
                move = max(0, target - self.plunger_steps)
                time.sleep(self._duration_plunger(move))
                self.plunger_steps = target
                return self._set_answer("")

            if c.startswith("D") and c.endswith("R") and c[1:-1].isdigit():
                steps = max(0, int(c[1:-1]))
                move = min(self.plunger_steps, steps)
                time.sleep(self._duration_plunger(move))
                self.plunger_steps = max(0, self.plunger_steps - move)
                return self._set_answer("")

            time.sleep(0.02); return self._set_answer("")

# ========================== Controller (real or sim) =========================
class PumpCtrl:
    def __init__(self, use_sim=False, log_cb=None):
        self.use_sim = use_sim
        self.log_cb = log_cb
        self.connected = False
        self.dev = DEFAULT_DEV
        self.com_port = DEFAULT_COM_PORT
        self.baud = DEFAULT_BAUD
        self.steps_per_stroke = DEFAULT_STEPS
        self.syringe_ul = DEFAULT_SYRINGE
        self.current_speed = None
        self._backend = None
        self._plunger_steps = 0

    def _log(self, s):
        if self.log_cb: self.log_cb(s)

    def _sync_backend_plunger(self):
        backend = self._backend
        if backend is None:
            return
        if hasattr(backend, 'plunger_steps'):
            try:
                backend.plunger_steps = int(self._plunger_steps)
            except Exception:
                pass

    def _set_plunger_steps(self, steps: int):
        clamped = max(0, min(int(steps), self.steps_per_stroke))
        self._plunger_steps = clamped
        self._sync_backend_plunger()


    def connect(self, com_port:int, baud:int, dev:int):
        if self.connected: return
        self.com_port, self.baud, self.dev = int(com_port), int(baud), int(dev)

        if self.use_sim or not HAS_COM:
            self._backend = SimPumpComm(self.steps_per_stroke, self.syringe_ul)
            self._backend.PumpInitComm(self.com_port)
            self.connected = True
            self._sync_backend_plunger()
            self._log(f"[SIM] Connected (COM{self.com_port})")
            return

        self._log(f"Connecting (real) -> COM{self.com_port} @ {self.baud}, dev={self.dev}")
        self._backend = gencache.EnsureDispatch(PROGID)
        try:
            try:
                self._backend.EnableLog = True
                self._backend.LogComPort = True
                self._backend.CommandAckTimeout = 18
                self._backend.CommandRetryCount = 3
                try: self._backend.BaudRate = self.baud
                except Exception: pass
            except Exception: pass

            self._backend.PumpInitComm(self.com_port)
            self.connected = True
            self._sync_backend_plunger()
            self._log("Connected.")
        except Exception as e:
            self._backend = None
            raise RuntimeError(f"Connect failed: {e}")

    def disconnect(self):
        if not self.connected: return
        try: self._backend.PumpExitComm()
        except Exception: pass
        self.connected = False; self._backend = None
        self._log("Disconnected.")

    def _send(self, cmd, wait_s=1.0):
        if not self.connected: raise RuntimeError("Not connected.")
        try:
            self._backend.PumpSendCommand(cmd, self.dev, "")
        except Exception:
            try: self._backend.PumpSendNoWait(cmd, self.dev)
            except Exception: pass
        time.sleep(wait_s)
        try: return (self._backend.PumpGetLastAnswer(self.dev) or "").strip()
        except Exception: return ""

    def initialize(self):
        ans = self._send("ZR", 1.2)
        self._set_plunger_steps(0)
        return ans

    def valve_to(self, port):
        return self._send(f"I{int(port)}R", 0.8)

    def configure_calibration(self, steps_per_stroke: int, syringe_ul: float):
        steps = max(1, int(steps_per_stroke))
        volume = max(1e-6, float(syringe_ul))
        current_volume = self.current_volume_ul
        self.steps_per_stroke = steps
        self.syringe_ul = volume
        self._set_plunger_steps(min(self._ul_to_steps(current_volume), self.steps_per_stroke))
        if self._backend is not None:
            if hasattr(self._backend, "steps_per_stroke"):
                self._backend.steps_per_stroke = self.steps_per_stroke
            if hasattr(self._backend, "syringe_ul"):
                self._backend.syringe_ul = self.syringe_ul

    def set_speed(self, s:int, settle=0.15):
        s = max(SPEED_MIN, min(SPEED_MAX, int(s)))
        self.current_speed = s
        return self._send(f"S{s}R", settle)

    def _ul_to_steps(self, ul:float) -> int:
        return max(0, int(round(self.steps_per_stroke * (float(ul) / float(self.syringe_ul)))))

    def _steps_to_ul(self, steps: int) -> float:
        if self.steps_per_stroke <= 0:
            return 0.0
        return (float(steps) / float(self.steps_per_stroke)) * float(self.syringe_ul)

    @property
    def plunger_steps(self) -> int:
        return self._plunger_steps

    def steps_for_volume(self, ul: float) -> int:
        return self._ul_to_steps(ul)

    def volume_for_steps(self, steps: int) -> float:
        return self._steps_to_ul(steps)

    @property
    def current_volume_ul(self) -> float:
        return self._steps_to_ul(self._plunger_steps)

    @property
    def remaining_capacity_ul(self) -> float:
        remaining_steps = max(0, self.steps_per_stroke - self._plunger_steps)
        return self._steps_to_ul(remaining_steps)

    def aspirate_ul(self, ul:float):
        volume = float(ul)
        if volume < 0:
            raise ValueError("Cannot aspirate a negative volume.")
        delta_steps = self._ul_to_steps(volume)
        if delta_steps <= 0:
            return ""
        target_steps = self._plunger_steps + delta_steps
        if target_steps > self.steps_per_stroke:
            remaining = self.remaining_capacity_ul
            raise ValueError(
                f"Requested {volume:.2f} uL exceeds remaining syringe capacity of {remaining:.2f} uL (syringe volume {self.syringe_ul:.2f} uL)."
            )
        if self.current_speed is not None:
            self._send(f"S{self.current_speed}R", 0.05)
        ans = self._send(f"A{target_steps}R", 1.0)
        self._set_plunger_steps(target_steps)
        return ans

    def dispense_ul(self, ul:float):
        volume = float(ul)
        if volume < 0:
            raise ValueError("Cannot dispense a negative volume.")
        delta_steps = self._ul_to_steps(volume)
        if delta_steps <= 0:
            return ""
        if delta_steps > self._plunger_steps:
            available = self.current_volume_ul
            raise ValueError(
                f"Requested dispense of {volume:.2f} uL exceeds loaded volume of {available:.2f} uL."
            )
        if self.current_speed is not None:
            self._send(f"S{self.current_speed}R", 0.05)
        ans = self._send(f"D{delta_steps}R", 1.0)
        new_steps = self._plunger_steps - delta_steps
        self._set_plunger_steps(new_steps)
        return ans

# ================================== GUI =====================================
class PumpGUI(tk.Tk):
    def __init__(self, force_sim=False):
        super().__init__()
        self.title("Cavro Centris Controller (Real/Sim)")
        self.geometry("740x600")
        self.resizable(False, False)

        self.force_sim = force_sim or (not HAS_COM)
        self.ctrl = PumpCtrl(use_sim=self.force_sim, log_cb=self.log)
        self._busy = False
        self._early_logs = []   # buffer logs before widget exists

        self._build_ui()

    # ---- logging: tolerate early calls before log_text exists ----
    def log(self, msg:str):
        if hasattr(self, "log_text"):
            self.log_text.configure(state="normal")
            self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        else:
            self._early_logs.append(msg)

    def _flush_early_logs(self):
        if not hasattr(self, "log_text"): return
        if not self._early_logs: return
        self.log_text.configure(state="normal")
        for m in self._early_logs:
            self.log_text.insert("end", m + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        self._early_logs.clear()

    def set_busy(self, busy:bool):
        self._busy = busy
        state = "disabled" if busy else "normal"
        for w in self.disable_group:
            try: w.configure(state=state)
            except Exception: pass
        self.update_idletasks()

    def threaded(self, fn, *args, **kwargs):
        if self._busy: return
        def run():
            if HAS_COM and not self.var_sim.get():
                pythoncom.CoInitialize()
            try:
                self.set_busy(True); fn(*args, **kwargs)
            except Exception as e:
                self.log(f"ERROR: {e}"); messagebox.showerror("Error", str(e))
            finally:
                self.set_busy(False)
                if HAS_COM and not self.var_sim.get():
                    pythoncom.CoUninitialize()
        threading.Thread(target=run, daemon=True).start()

    def _build_ui(self):
        pad = {"padx":6, "pady":4}

        # --- Build LOG FIRST so self.log() is safe immediately ---
        f_log = ttk.LabelFrame(self, text="Log")
        f_log.place(x=10, y=430, width=720, height=160)
        self.log_text = tk.Text(f_log, height=8, state="disabled")
        self.log_text.pack(fill="both", expand=True, padx=6, pady=6)

        # Connection
        f_conn = ttk.LabelFrame(self, text="Connection")
        f_conn.place(x=10, y=10, width=720, height=120)

        self.var_sim = tk.BooleanVar(value=self.force_sim)
        chk = ttk.Checkbutton(f_conn, text="Simulate (no hardware)", variable=self.var_sim)
        chk.grid(row=0, column=0, columnspan=2, **pad, sticky="w")
        if not HAS_COM:
            chk.configure(state="disabled")
            self.log("COM not available; defaulting to SIM mode.")

        ttk.Label(f_conn, text="COM port:").grid(row=1, column=0, **pad, sticky="e")
        self.var_com = tk.IntVar(value=DEFAULT_COM_PORT)
        self.ent_com = ttk.Spinbox(f_conn, from_=1, to=60, width=6, textvariable=self.var_com)
        self.ent_com.grid(row=1, column=1, **pad)

        ttk.Label(f_conn, text="Baud:").grid(row=1, column=2, **pad, sticky="e")
        self.var_baud = tk.IntVar(value=DEFAULT_BAUD)
        self.ent_baud = ttk.Combobox(f_conn, values=[9600, 38400], width=8, textvariable=self.var_baud)
        self.ent_baud.grid(row=1, column=3, **pad); self.ent_baud.set(str(DEFAULT_BAUD))

        ttk.Label(f_conn, text="Device #:").grid(row=1, column=4, **pad, sticky="e")
        self.var_dev = tk.IntVar(value=DEFAULT_DEV)
        self.ent_dev = ttk.Spinbox(f_conn, from_=0, to=30, width=6, textvariable=self.var_dev)
        self.ent_dev.grid(row=1, column=5, **pad)

        self.btn_connect = ttk.Button(f_conn, text="Connect", command=lambda: self.threaded(self.on_connect))
        self.btn_disconnect = ttk.Button(f_conn, text="Disconnect", command=lambda: self.threaded(self.on_disconnect))
        self.btn_connect.grid(row=2, column=0, columnspan=2, **pad)
        self.btn_disconnect.grid(row=2, column=2, columnspan=2, **pad)

        # Calibration
        f_cal = ttk.LabelFrame(self, text="Calibration (µL ↔ steps)")
        f_cal.place(x=10, y=135, width=720, height=80)
        ttk.Label(f_cal, text="Steps/stroke:").grid(row=0, column=0, **pad, sticky="e")
        self.var_steps = tk.IntVar(value=int(DEFAULT_STEPS))
        self.ent_steps = ttk.Entry(f_cal, width=10, textvariable=self.var_steps)
        self.ent_steps.grid(row=0, column=1, **pad)
        ttk.Label(f_cal, text="Syringe (µL):").grid(row=0, column=2, **pad, sticky="e")
        self.var_syr = tk.DoubleVar(value=float(DEFAULT_SYRINGE))
        self.ent_syr = ttk.Entry(f_cal, width=10, textvariable=self.var_syr)
        self.ent_syr.grid(row=0, column=3, **pad)
        self.btn_apply = ttk.Button(f_cal, text="Apply", command=self.on_apply_cal)
        self.btn_apply.grid(row=0, column=4, **pad)

        # Actions
        f_act = ttk.LabelFrame(self, text="Actions")
        f_act.place(x=10, y=220, width=720, height=200)

        self.btn_init = ttk.Button(f_act, text="Initialize (ZR)", command=lambda: self.threaded(self.do_init))
        self.btn_init.grid(row=0, column=0, **pad)

        ttk.Label(f_act, text="Volume (µL):").grid(row=0, column=1, **pad, sticky="e")
        self.var_vol = tk.DoubleVar(value=50.0)
        self.ent_vol = ttk.Entry(f_act, width=10, textvariable=self.var_vol)
        self.ent_vol.grid(row=0, column=2, **pad)

        ttk.Label(f_act, text="Plunger speed (SnnR):").grid(row=0, column=3, **pad, sticky="e")
        self.var_speed = tk.IntVar(value=20)
        self.ent_speed = ttk.Spinbox(f_act, from_=SPEED_MIN, to=SPEED_MAX, width=6, textvariable=self.var_speed)
        self.ent_speed.grid(row=0, column=4, **pad)
        self.btn_setspeed = ttk.Button(f_act, text="Set Speed", command=lambda: self.threaded(self.do_set_speed))
        self.btn_setspeed.grid(row=0, column=5, **pad)

        ttk.Label(f_act, text="Valve port:").grid(row=1, column=0, **pad, sticky="e")
        self.var_port = tk.IntVar(value=1)
        self.ent_port = ttk.Spinbox(f_act, from_=1, to=9, width=6, textvariable=self.var_port)
        self.ent_port.grid(row=1, column=1, **pad)
        self.btn_valve = ttk.Button(f_act, text="Move Valve (I#R)", command=lambda: self.threaded(self.do_valve))
        self.btn_valve.grid(row=1, column=2, **pad)

        grid = ttk.LabelFrame(f_act, text="Valve quick")
        grid.grid(row=2, column=0, columnspan=6, padx=6, pady=(6, 2))
        self.valve_buttons = []
        for i in range(1, 10):
            b = ttk.Button(grid, text=str(i), width=3, command=lambda p=i: self.threaded(self.do_valve_num, p))
            b.grid(row=(i-1)//5, column=(i-1)%5, padx=3, pady=3)
            self.valve_buttons.append(b)

        self.btn_asp  = ttk.Button(f_act, text="Aspirate", command=lambda: self.threaded(self.do_asp))
        self.btn_disp = ttk.Button(f_act, text="Dispense", command=lambda: self.threaded(self.do_disp))
        self.btn_asp.grid(row=3, column=3, **pad)
        self.btn_disp.grid(row=3, column=4, **pad)

        # controls to disable during actions
        self.disable_group = [
            self.btn_connect, self.btn_disconnect, self.btn_init, self.btn_asp, self.btn_disp,
            self.btn_valve, self.btn_apply, self.ent_com, self.ent_baud, self.ent_dev,
            self.ent_steps, self.ent_syr, self.ent_vol, self.ent_port, self.btn_setspeed,
            self.ent_speed, *self.valve_buttons
        ]

        # flush any early logs now that log_text exists
        self._flush_early_logs()

    # ---------------- handlers (run in worker thread) ----------------
    def on_connect(self):
        try:
            self.ctrl.use_sim = bool(self.var_sim.get()) or (not HAS_COM)
            mode = "[SIM]" if self.ctrl.use_sim else "[REAL]"
            self.log(f"{mode} Connecting…")
            self.ctrl.connect(self.var_com.get(), int(self.var_baud.get()), self.var_dev.get())
            if self.ctrl.use_sim:
                self.log("Sim mode ready: speed/valve/A-D behave with realistic timing.")
            else:
                self.log("Real mode ready. Tip: set plunger speed (SnnR) before A/D.")
        except Exception as e:
            messagebox.showerror("Connect failed", str(e)); self.log(f"Connect failed: {e}")

    def on_disconnect(self): self.ctrl.disconnect()

    def on_apply_cal(self):
        try:
            self.ctrl.steps_per_stroke = int(self.var_steps.get())
            self.ctrl.syringe_ul = float(self.var_syr.get())
            self.log(f"Applied: steps/stroke={self.ctrl.steps_per_stroke}, syringe={self.ctrl.syringe_ul:.0f} µL")
        except Exception as e:
            messagebox.showerror("Invalid calibration", str(e))

    def do_init(self):
        if not self.ctrl.connected: return self.log("Not connected.")
        self.log("Initialize (ZR)…"); _ = self.ctrl.initialize(); self.log("Init done.")

    def do_set_speed(self):
        if not self.ctrl.connected: return self.log("Not connected.")
        s = int(self.var_speed.get()); self.log(f"Set plunger speed: S{s}R"); _ = self.ctrl.set_speed(s)

    def do_valve(self):
        if not self.ctrl.connected: return self.log("Not connected.")
        p = int(self.var_port.get()); self.log(f"Valve -> {p} (I{p}R)"); _ = self.ctrl.valve_to(p); self.log("Valve move done.")

    def do_valve_num(self, port:int):
        if not self.ctrl.connected: return self.log("Not connected.")
        self.log(f"Valve -> {port} (I{port}R)"); _ = self.ctrl.valve_to(port); self.log("Valve move done.")

    def do_asp(self):
        if not self.ctrl.connected: return self.log("Not connected.")
        v = float(self.var_vol.get()); s = int(self.var_speed.get())
        self.log(f"Aspirate {v:.2f} µL @ S{s}R"); _ = self.ctrl.set_speed(s); _ = self.ctrl.aspirate_ul(v); self.log("Aspirate done.")

    def do_disp(self):
        if not self.ctrl.connected: return self.log("Not connected.")
        v = float(self.var_vol.get()); s = int(self.var_speed.get())
        self.log(f"Dispense  {v:.2f} µL @ S{s}R"); _ = self.ctrl.set_speed(s); _ = self.ctrl.dispense_ul(v); self.log("Dispense done.")

# ================================== main ====================================
def main():
    ap = argparse.ArgumentParser(description="Cavro Centris GUI (Real/Sim)")
    ap.add_argument("--sim", action="store_true", help="Run with built-in simulator (no hardware)")
    args = ap.parse_args()

    app = PumpGUI(force_sim=args.sim)
    app.mainloop()

if __name__ == "__main__":
    main()
