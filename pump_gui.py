# pump_gui.py
# Minimal Cavro Centris GUI (Windows, 32-bit Python) using PumpCommServer.PumpComm
# - Pure command style to dev=1 (works on your setup)
# - Valve moves use I#R
# - µL<->steps defaults: 100K steps/stroke, 1250 µL syringe (editable)
#
# Requirements: pywin32 (already installed in your 32-bit venv)

import tkinter as tk
from tkinter import ttk, messagebox
from win32com.client import gencache
import pythoncom
import threading
import time

DEFAULT_COM_PORT = 8
DEFAULT_BAUD     = 9600
DEFAULT_DEV      = 1
DEFAULT_STEPS    = 100000   # "100K"
DEFAULT_SYRINGE  = 1250     # µL

PROGID = "PumpCommServer.PumpComm"

# ---------- Low-level controller (same pattern that worked for you) ----------
class PumpCtrl:
    def __init__(self):
        self.pump = None
        self.connected = False
        self.com_port = DEFAULT_COM_PORT
        self.baud = DEFAULT_BAUD
        self.dev = DEFAULT_DEV
        self.steps_per_stroke = DEFAULT_STEPS
        self.syringe_ul = DEFAULT_SYRINGE

    def connect(self, com_port:int, baud:int, dev:int, log_cb=None):
        if self.connected:
            return
        self.com_port = int(com_port)
        self.baud = int(baud)
        self.dev = int(dev)

        if log_cb: log_cb(f"Connecting -> COM{self.com_port} @ {self.baud}, dev={self.dev}")
        self.pump = gencache.EnsureDispatch(PROGID)
        try:
            # "nice-to-have" properties (ignore failures)
            self.pump.EnableLog = True
            self.pump.LogComPort = True
            self.pump.CommandAckTimeout = 18
            self.pump.CommandRetryCount = 3
            try:
                self.pump.BaudRate = self.baud
            except Exception:
                pass
        except Exception:
            pass
        self.pump.PumpInitComm(self.com_port)
        self.connected = True
        if log_cb: log_cb("Connected.")

    def disconnect(self, log_cb=None):
        if not self.connected:
            return
        try:
            self.pump.PumpExitComm()
        except Exception:
            pass
        self.pump = None
        self.connected = False
        if log_cb: log_cb("Disconnected.")

    def _send_pure(self, cmd:str, wait_s:float=1.0):
        """Pure style with 3-arg PumpSendCommand; fallback to NoWait + short sleep."""
        if not self.connected:
            raise RuntimeError("Not connected.")
        try:
            ans = self.pump.PumpSendCommand(cmd, self.dev, "")
            time.sleep(wait_s)
            return ans or ""
        except pythoncom.com_error:
            self.pump.PumpSendNoWait(cmd, self.dev)
            time.sleep(wait_s)
            try:
                return self.pump.PumpGetLastAnswer(self.dev) or ""
            except pythoncom.com_error:
                return ""

    # --- Public actions ---
    def initialize(self):
        return self._send_pure("ZR", 1.5)

    def valve_to(self, port:int):
        return self._send_pure(f"I{int(port)}R", 0.9)

    def _ul_to_steps(self, ul:float)->int:
        return max(0, int(round(self.steps_per_stroke * (float(ul)/float(self.syringe_ul)))))

    def aspirate_ul(self, ul:float):
        steps = self._ul_to_steps(ul)
        return self._send_pure(f"A{steps}R", 1.0)

    def dispense_ul(self, ul:float):
        steps = self._ul_to_steps(ul)
        return self._send_pure(f"D{steps}R", 1.0)

# ---------- GUI ----------
class PumpGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Cavro Centris Controller")
        self.geometry("700x520")
        self.resizable(False, False)

        self.ctrl = PumpCtrl()
        self._busy = False
        self._build_ui()

    # --- UI helpers ---
    def log(self, msg:str):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def set_busy(self, busy:bool):
        self._busy = busy
        state = "disabled" if busy else "normal"
        for w in self.disable_group:
            try:
                w.configure(state=state)
            except Exception:
                pass
        self.update_idletasks()

# at the top you already have: import pythoncom
# ...

    def threaded(self, fn, *args, **kwargs):
        """Run pump actions off the UI thread; make sure COM is initialized."""
        if self._busy:
            return
        def run():
            pythoncom.CoInitialize()          # <-- ADD
            try:
                self.set_busy(True)
                fn(*args, **kwargs)
            except Exception as e:
                self.log(f"ERROR: {e}")
                messagebox.showerror("Error", str(e))
            finally:
                self.set_busy(False)
                pythoncom.CoUninitialize()     # <-- ADD
        threading.Thread(target=run, daemon=True).start()


    # --- UI building ---
    def _build_ui(self):
        pad = {"padx": 6, "pady": 4}

        # Connection frame
        f_conn = ttk.LabelFrame(self, text="Connection")
        f_conn.place(x=10, y=10, width=680, height=90)

        ttk.Label(f_conn, text="COM port:").grid(row=0, column=0, **pad, sticky="e")
        self.var_com = tk.IntVar(value=DEFAULT_COM_PORT)
        self.ent_com = ttk.Spinbox(f_conn, from_=1, to=60, width=6, textvariable=self.var_com)
        self.ent_com.grid(row=0, column=1, **pad)

        ttk.Label(f_conn, text="Baud:").grid(row=0, column=2, **pad, sticky="e")
        self.var_baud = tk.IntVar(value=DEFAULT_BAUD)
        self.ent_baud = ttk.Combobox(f_conn, values=[9600, 38400], width=8, textvariable=self.var_baud)
        self.ent_baud.grid(row=0, column=3, **pad)
        self.ent_baud.set(str(DEFAULT_BAUD))

        ttk.Label(f_conn, text="Device #:").grid(row=0, column=4, **pad, sticky="e")
        self.var_dev = tk.IntVar(value=DEFAULT_DEV)
        self.ent_dev = ttk.Spinbox(f_conn, from_=0, to=30, width=6, textvariable=self.var_dev)
        self.ent_dev.grid(row=0, column=5, **pad)

        self.btn_connect = ttk.Button(f_conn, text="Connect", command=lambda: self.threaded(self.on_connect))
        self.btn_disconnect = ttk.Button(f_conn, text="Disconnect", command=lambda: self.threaded(self.on_disconnect))
        self.btn_connect.grid(row=1, column=0, columnspan=2, **pad)
        self.btn_disconnect.grid(row=1, column=2, columnspan=2, **pad)

        # Calibration frame
        f_cal = ttk.LabelFrame(self, text="Calibration (µL ↔ steps)")
        f_cal.place(x=10, y=105, width=680, height=75)

        ttk.Label(f_cal, text="Steps/stroke:").grid(row=0, column=0, **pad, sticky="e")
        self.var_steps = tk.IntVar(value=DEFAULT_STEPS)
        self.ent_steps = ttk.Entry(f_cal, width=10, textvariable=self.var_steps)
        self.ent_steps.grid(row=0, column=1, **pad)

        ttk.Label(f_cal, text="Syringe (µL):").grid(row=0, column=2, **pad, sticky="e")
        self.var_syr = tk.IntVar(value=DEFAULT_SYRINGE)
        self.ent_syr = ttk.Entry(f_cal, width=10, textvariable=self.var_syr)
        self.ent_syr.grid(row=0, column=3, **pad)

        self.btn_apply = ttk.Button(f_cal, text="Apply", command=self.on_apply_cal)
        self.btn_apply.grid(row=0, column=4, **pad)

        # Actions frame
        f_act = ttk.LabelFrame(self, text="Actions")
        f_act.place(x=10, y=185, width=680, height=150)

        self.btn_init = ttk.Button(f_act, text="Initialize (ZR)", command=lambda: self.threaded(self.do_init))
        self.btn_init.grid(row=0, column=0, **pad)

        ttk.Label(f_act, text="Volume (µL):").grid(row=0, column=1, **pad, sticky="e")
        self.var_vol = tk.DoubleVar(value=50.0)
        self.ent_vol = ttk.Entry(f_act, width=10, textvariable=self.var_vol)
        self.ent_vol.grid(row=0, column=2, **pad)

        self.btn_asp = ttk.Button(f_act, text="Aspirate", command=lambda: self.threaded(self.do_asp))
        self.btn_disp = ttk.Button(f_act, text="Dispense", command=lambda: self.threaded(self.do_disp))
        self.btn_asp.grid(row=0, column=3, **pad)
        self.btn_disp.grid(row=0, column=4, **pad)

        ttk.Label(f_act, text="Valve port:").grid(row=1, column=0, **pad, sticky="e")
        self.var_port = tk.IntVar(value=1)
        self.ent_port = ttk.Spinbox(f_act, from_=1, to=9, width=6, textvariable=self.var_port)
        self.ent_port.grid(row=1, column=1, **pad)

        self.btn_valve = ttk.Button(f_act, text="Move Valve (I#R)", command=lambda: self.threaded(self.do_valve))
        self.btn_valve.grid(row=1, column=2, **pad)

        # Quick valve grid (1..9)
        grid = ttk.LabelFrame(f_act, text="Valve quick")
        grid.grid(row=2, column=0, columnspan=5, padx=6, pady=(6, 2))
        self.valve_buttons = []
        for i in range(1, 10):
            b = ttk.Button(grid, text=str(i), width=3, command=lambda p=i: self.threaded(self.do_valve_num, p))
            b.grid(row=(i-1)//5, column=(i-1)%5, padx=3, pady=3)
            self.valve_buttons.append(b)

        # Log frame
        f_log = ttk.LabelFrame(self, text="Log")
        f_log.place(x=10, y=340, width=680, height=170)

        self.log_text = tk.Text(f_log, height=8, state="disabled")
        self.log_text.pack(fill="both", expand=True, padx=6, pady=6)

        # Controls we disable during actions
        self.disable_group = [
            self.btn_connect, self.btn_disconnect, self.btn_init, self.btn_asp, self.btn_disp,
            self.btn_valve, self.btn_apply, self.ent_com, self.ent_baud, self.ent_dev,
            self.ent_steps, self.ent_syr, self.ent_vol, self.ent_port, *self.valve_buttons
        ]

    # --- Button handlers (run in thread via self.threaded) ---
    def on_connect(self):
        try:
            self.ctrl.connect(self.var_com.get(), int(self.var_baud.get()), self.var_dev.get(), self.log)
        except Exception as e:
            messagebox.showerror("Connect failed", str(e))
            self.log(f"Connect failed: {e}")

    def on_disconnect(self):
        self.ctrl.disconnect(self.log)

    def on_apply_cal(self):
        try:
            self.ctrl.steps_per_stroke = int(self.var_steps.get())
            self.ctrl.syringe_ul = float(self.var_syr.get())
            self.log(f"Applied calibration: steps/stroke={self.ctrl.steps_per_stroke}, syringe={self.ctrl.syringe_ul} µL")
        except Exception as e:
            messagebox.showerror("Invalid calibration", str(e))

    def do_init(self):
        if not self.ctrl.connected:
            return self.log("Not connected.")
        self.log("Init (ZR)...")
        _ = self.ctrl.initialize()
        self.log("Init done.")

    def do_valve(self):
        if not self.ctrl.connected:
            return self.log("Not connected.")
        port = int(self.var_port.get())
        self.log(f"Valve -> {port} (I{port}R)")
        _ = self.ctrl.valve_to(port)
        self.log("Valve move done.")

    def do_valve_num(self, port:int):
        if not self.ctrl.connected:
            return self.log("Not connected.")
        self.log(f"Valve -> {port} (I{port}R)")
        _ = self.ctrl.valve_to(port)
        self.log("Valve move done.")

    def do_asp(self):
        if not self.ctrl.connected:
            return self.log("Not connected.")
        vol = float(self.var_vol.get())
        self.log(f"Aspirate {vol:.2f} µL")
        _ = self.ctrl.aspirate_ul(vol)
        self.log("Aspirate done.")

    def do_disp(self):
        if not self.ctrl.connected:
            return self.log("Not connected.")
        vol = float(self.var_vol.get())
        self.log(f"Dispense {vol:.2f} µL")
        _ = self.ctrl.dispense_ul(vol)
        self.log("Dispense done.")

if __name__ == "__main__":
    app = PumpGUI()
    app.mainloop()
