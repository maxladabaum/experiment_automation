"""
gui/tab_pump.py — Pump Control tab.

Wraps the pump hardware controls (connect, calibrate, init, valve,
aspirate, dispense) and exposes "Queue …" buttons that delegate to
the QueueTab via the on_add_to_queue callback.

Works on 64-bit Python with no hardware: shows a disabled stub when
PumpCtrl is unavailable.
"""

import threading
from tkinter import messagebox
import tkinter as tk
from tkinter import ttk

try:
    import pythoncom   # type: ignore
    HAS_PYTHONCOM = True
except ImportError:
    HAS_PYTHONCOM = False

from config import (
    PUMP_DEFAULT_COM_PORT, PUMP_DEFAULT_BAUD, PUMP_DEFAULT_DEV,
    PUMP_SPEED_MIN, PUMP_SPEED_MAX,
    PREFERRED_STEPS_PER_STROKE, PREFERRED_SYRINGE_UL,
)


class PumpTab:
    """Manages the 'Pump Control' notebook tab.

    Parameters
    ----------
    parent_frame:
        The ``ttk.Frame`` added to the notebook for this tab.
    pump_ctrl:
        A :class:`~pump.pump_ctrl.PumpCtrl` instance, or ``None`` if
        the pump backend is unavailable.
    on_add_to_queue:
        Callable ``(item: dict) → None`` wired to QueueTab.add_item.
    root:
        Root ``tk.Tk`` window for ``after()`` calls.
    """

    def __init__(self, parent_frame, pump_ctrl, on_add_to_queue, root):
        self._frame         = parent_frame
        self._ctrl          = pump_ctrl
        self._add_to_queue  = on_add_to_queue
        self._root          = root
        self._busy          = False
        self._early_logs:list = []
        self._log_text      = None
        self._disable_group: list = []

        if pump_ctrl is None:
            ttk.Label(self._frame,
                      text="Pump controls unavailable (64-bit / no hardware).",
                      foreground="gray").pack(pady=40)
            return

        self._build()

    # ── Build ─────────────────────────────────────────────────────────────────

    def _build(self):
        pad = {"padx": 6, "pady": 4}
        container = ttk.Frame(self._frame)
        container.pack(fill="both", expand=True, padx=10, pady=10)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(3, weight=1)

        # ── Connection ────────────────────────────────────────────────────────
        conn = ttk.LabelFrame(container, text="Connection")
        conn.grid(row=0, column=0, sticky="ew")

        self._var_sim  = tk.BooleanVar(value=self._ctrl.use_sim)
        chk = ttk.Checkbutton(conn, text="Simulate (no hardware)",
                               variable=self._var_sim)
        chk.grid(row=0, column=0, columnspan=2, **pad, sticky="w")

        ttk.Label(conn, text="COM port:").grid(row=1, column=0, **pad, sticky="e")
        self._var_com  = tk.IntVar(value=int(PUMP_DEFAULT_COM_PORT))
        ttk.Spinbox(conn, from_=1, to=60, width=6,
                    textvariable=self._var_com).grid(row=1, column=1, **pad)

        ttk.Label(conn, text="Baud:").grid(row=1, column=2, **pad, sticky="e")
        self._var_baud = tk.StringVar(value=str(PUMP_DEFAULT_BAUD))
        combo = ttk.Combobox(conn, values=["9600", "38400"], width=8,
                             textvariable=self._var_baud)
        combo.grid(row=1, column=3, **pad)
        combo.set(str(PUMP_DEFAULT_BAUD))

        ttk.Label(conn, text="Device #:").grid(row=1, column=4, **pad, sticky="e")
        self._var_dev  = tk.IntVar(value=int(PUMP_DEFAULT_DEV))
        ttk.Spinbox(conn, from_=0, to=30, width=6,
                    textvariable=self._var_dev).grid(row=1, column=5, **pad)

        btn_connect = ttk.Button(conn, text="Connect",
            command=lambda: self._launch(self._do_connect,
                lambda: bool(self._var_sim.get()),
                lambda: int(self._var_com.get()),
                lambda: int(self._var_baud.get()),
                lambda: int(self._var_dev.get())))
        btn_connect.grid(row=2, column=0, columnspan=2, **pad)

        btn_disconnect = ttk.Button(conn, text="Disconnect",
            command=lambda: self._threaded(self._do_disconnect))
        btn_disconnect.grid(row=2, column=2, columnspan=2, **pad)

        # ── Calibration ───────────────────────────────────────────────────────
        cal = ttk.LabelFrame(container, text="Calibration (µL ↔ steps)")
        cal.grid(row=1, column=0, sticky="ew", pady=(10, 0))

        ttk.Label(cal, text="Steps/stroke:").grid(row=0, column=0, **pad, sticky="e")
        self._var_steps  = tk.IntVar(value=int(PREFERRED_STEPS_PER_STROKE))
        ttk.Entry(cal, width=10,
                  textvariable=self._var_steps).grid(row=0, column=1, **pad)

        ttk.Label(cal, text="Syringe (µL):").grid(row=0, column=2, **pad, sticky="e")
        self._var_syr    = tk.DoubleVar(value=float(PREFERRED_SYRINGE_UL))
        ttk.Entry(cal, width=10,
                  textvariable=self._var_syr).grid(row=0, column=3, **pad)

        ttk.Button(cal, text="Apply",
            command=lambda: self._launch(self._do_apply_cal,
                lambda: int(self._var_steps.get()),
                lambda: float(self._var_syr.get()))
        ).grid(row=0, column=4, **pad)

        # ── Actions ───────────────────────────────────────────────────────────
        act = ttk.LabelFrame(container, text="Actions")
        act.grid(row=2, column=0, sticky="ew", pady=(10, 0))

        ttk.Button(act, text="Initialize (ZR)",
            command=lambda: self._threaded(self._do_init)
        ).grid(row=0, column=0, **pad)

        ttk.Button(act, text="Queue Init",
            command=self._queue_init).grid(row=0, column=1, **pad)

        ttk.Label(act, text="Volume (µL):").grid(row=0, column=2, **pad, sticky="e")
        self._var_vol   = tk.DoubleVar(value=50.0)
        ttk.Entry(act, width=10,
                  textvariable=self._var_vol).grid(row=0, column=3, **pad)

        ttk.Label(act, text="Speed (SnnR):").grid(row=0, column=4, **pad, sticky="e")
        self._var_speed = tk.IntVar(value=20)
        ttk.Spinbox(act, from_=PUMP_SPEED_MIN, to=PUMP_SPEED_MAX, width=6,
                    textvariable=self._var_speed).grid(row=0, column=5, **pad)

        ttk.Button(act, text="Set Speed",
            command=lambda: self._launch(self._do_set_speed,
                lambda: int(self._var_speed.get()))
        ).grid(row=0, column=6, **pad)

        ttk.Button(act, text="Queue Speed",
            command=self._queue_set_speed).grid(row=0, column=7, **pad)

        ttk.Label(act, text="Valve port:").grid(row=1, column=0, **pad, sticky="e")
        self._var_valve = tk.IntVar(value=1)
        ttk.Spinbox(act, from_=1, to=9, width=6,
                    textvariable=self._var_valve).grid(row=1, column=1, **pad)

        ttk.Button(act, text="Move Valve",
            command=lambda: self._launch(self._do_valve,
                lambda: int(self._var_valve.get()))
        ).grid(row=1, column=2, **pad)

        ttk.Button(act, text="Queue Valve",
            command=self._queue_valve).grid(row=1, column=3, **pad)

        # Valve quick-select
        vq = ttk.LabelFrame(act, text="Valve quick")
        vq.grid(row=2, column=0, columnspan=8, padx=6, pady=(6, 2))
        for i in range(1, 10):
            ttk.Button(vq, text=str(i), width=3,
                command=lambda p=i: self._threaded(self._do_valve, p)
            ).grid(row=(i - 1) // 5, column=(i - 1) % 5, padx=3, pady=3)

        ttk.Button(act, text="Aspirate",
            command=lambda: self._launch(self._do_aspirate,
                lambda: float(self._var_vol.get()),
                lambda: int(self._var_speed.get()))
        ).grid(row=3, column=2, **pad)
        ttk.Button(act, text="Queue Aspirate",
            command=self._queue_aspirate).grid(row=3, column=3, **pad)

        ttk.Button(act, text="Dispense",
            command=lambda: self._launch(self._do_dispense,
                lambda: float(self._var_vol.get()),
                lambda: int(self._var_speed.get()))
        ).grid(row=3, column=4, **pad)
        ttk.Button(act, text="Queue Dispense",
            command=self._queue_dispense).grid(row=3, column=5, **pad)

        # ── Log ───────────────────────────────────────────────────────────────
        log_frame = ttk.LabelFrame(container, text="Log")
        log_frame.grid(row=3, column=0, sticky="nsew", pady=(10, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self._log_text = tk.Text(log_frame, height=10, state="disabled")
        self._log_text.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)

        self._flush_early_logs()

    # ── Logging ───────────────────────────────────────────────────────────────

    def log(self, msg: str):
        if self._log_text is None:
            self._early_logs.append(msg)
            return
        def _append():
            self._log_text.configure(state="normal")
            self._log_text.insert("end", msg + "\n")
            self._log_text.see("end")
            self._log_text.configure(state="disabled")
        self._root.after(0, _append)

    def _flush_early_logs(self):
        if self._log_text is None or not self._early_logs:
            return
        msgs = self._early_logs[:]
        self._early_logs.clear()
        def _flush():
            self._log_text.configure(state="normal")
            for m in msgs:
                self._log_text.insert("end", m + "\n")
            self._log_text.see("end")
            self._log_text.configure(state="disabled")
        self._root.after(0, _flush)

    # ── Threading ─────────────────────────────────────────────────────────────

    def _threaded(self, fn, *args):
        if self._ctrl is None:
            messagebox.showerror("Pump Error", "Pump backend unavailable."); return
        if self._busy:
            return
        sim_mode = bool(self._var_sim.get()) if hasattr(self, "_var_sim") else True

        def run():
            use_com = HAS_PYTHONCOM and not sim_mode
            if use_com:
                try: pythoncom.CoInitialize()
                except Exception: use_com = False
            try:
                self._set_busy(True)
                fn(*args)
            except Exception as exc:
                self.log(f"ERROR: {exc}")
                self._root.after(0, lambda: messagebox.showerror("Pump Error", str(exc)))
            finally:
                self._set_busy(False)
                if use_com:
                    try: pythoncom.CoUninitialize()
                    except Exception: pass

        threading.Thread(target=run, daemon=True).start()

    def _launch(self, target, *factories):
        if self._ctrl is None:
            messagebox.showerror("Pump Error", "Pump backend unavailable."); return
        try:
            vals = [f() for f in factories]
        except (ValueError, tk.TclError) as exc:
            messagebox.showerror("Invalid value", str(exc)); return
        self._threaded(target, *vals)

    def _set_busy(self, busy: bool):
        self._busy = busy
    
    def autoconnect(self):
        if self._ctrl is None:
            return
        if getattr(self._ctrl, "connected", False):
            return
        if self._busy:
            return
        if not hasattr(self, "_var_sim"):
            return
        self._launch(
            self._do_connect,
            lambda: bool(self._var_sim.get()),
            lambda: int(self._var_com.get()),
            lambda: int(self._var_baud.get()),
            lambda: int(self._var_dev.get()),
        )

    # ── Pump action implementations ───────────────────────────────────────────

    def _require_conn(self) -> bool:
        if not (self._ctrl and self._ctrl.connected):
            self.log("Not connected."); return False
        return True

    def _do_connect(self, sim, com, baud, dev):
        try:
            self._ctrl.use_sim = bool(sim)
            mode = "[SIM]" if self._ctrl.use_sim else "[REAL]"
            self.log(f"{mode} Connecting…")
            self._ctrl.connect(com, baud, dev)
            self.log("Connected.")
        except Exception as exc:
            self._root.after(0, lambda: messagebox.showerror("Connect failed", str(exc)))
            self.log(f"Connect failed: {exc}")

    def _do_disconnect(self):
        if self._ctrl:
            self._ctrl.disconnect()

    def _do_apply_cal(self, steps, syr):
        if self._ctrl:
            self._ctrl.configure_calibration(int(steps), float(syr))
            self.log(f"Calibration applied: {steps} steps, {syr:.0f} µL")

    def _do_init(self):
        if not self._require_conn(): return
        self.log("Initialize (ZR)…"); self._ctrl.initialize(); self.log("Init done.")

    def _do_set_speed(self, s):
        if not self._require_conn(): return
        self.log(f"Set speed: S{s}R"); self._ctrl.set_speed(s)

    def _do_valve(self, port):
        if not self._require_conn(): return
        self.log(f"Valve → {port}"); self._ctrl.valve_to(port); self.log("Valve done.")

    def _do_aspirate(self, vol, spd):
        if not self._require_conn(): return
        self.log(f"Aspirate {vol:.2f} µL @ S{spd}R")
        self._ctrl.set_speed(spd); self._ctrl.aspirate_ul(vol); self.log("Aspirate done.")

    def _do_dispense(self, vol, spd):
        if not self._require_conn(): return
        self.log(f"Dispense {vol:.2f} µL @ S{spd}R")
        self._ctrl.set_speed(spd); self._ctrl.dispense_ul(vol); self.log("Dispense done.")

    # ── Queue pump actions ────────────────────────────────────────────────────

    def _pump_queue_item(self, action_name: str, params: dict, details: str):
        self._add_to_queue({
            "type":        f"PUMP_{action_name}",
            "status":      "pending",
            "details":     details,
            "pump_action": {"name": action_name, "params": params},
        })

    def _queue_init(self):
        self._pump_queue_item("INIT", {}, "Pump: Initialize (ZR)")

    def _queue_set_speed(self):
        try:
            s = int(self._var_speed.get())
        except (ValueError, tk.TclError) as exc:
            messagebox.showerror("Invalid speed", str(exc)); return
        self._pump_queue_item("SET_SPEED", {"speed": s}, f"Pump: Set Speed S{s}R")

    def _queue_valve(self):
        try:
            p = int(self._var_valve.get())
        except (ValueError, tk.TclError) as exc:
            messagebox.showerror("Invalid port", str(exc)); return
        self._pump_queue_item("VALVE", {"port": p}, f"Pump: Valve → {p}")

    def _queue_aspirate(self):
        try:
            vol = float(self._var_vol.get())
            spd = int(self._var_speed.get())
        except (ValueError, tk.TclError) as exc:
            messagebox.showerror("Invalid parameters", str(exc)); return
        self._pump_queue_item(
            "ASPIRATE", {"volume": vol, "speed": spd},
            f"Pump: Aspirate {vol:.2f} µL @ S{spd}R"
        )

    def _queue_dispense(self):
        try:
            vol = float(self._var_vol.get())
            spd = int(self._var_speed.get())
        except (ValueError, tk.TclError) as exc:
            messagebox.showerror("Invalid parameters", str(exc)); return
        self._pump_queue_item(
            "DISPENSE", {"volume": vol, "speed": spd},
            f"Pump: Dispense {vol:.2f} µL @ S{spd}R"
        )
