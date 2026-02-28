"""
gui/tab_queue.py — Queue & Execution tab.

Responsible for:
  - Displaying the measurement queue in a Treeview
  - Copy / paste / duplicate / delete / reorder queue items
  - Save / load queue to JSON
  - Running / stopping the queue
  - Executing each queue item type (measurement, pause, alert, pump)
  - Session info bar (measurement counter, script registry size)
"""

import copy
import json
import threading
import time
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk, scrolledtext

from core.runner import SerialMeasurementRunner
from core.session import SessionState


class QueueTab:
    """Manages the 'Queue & Execution' notebook tab.

    Parameters
    ----------
    parent_frame:
        The ``ttk.Frame`` added to the notebook for this tab.
    session:
        Shared :class:`~core.session.SessionState`.
    plotter:
        Reference to :class:`~gui.tab_plotter.PlotterTab` for live plotting.
    pump_ctrl:
        Optional pump controller (may be ``None`` on 64-bit / no hardware).
    root:
        The root ``tk.Tk`` window — needed for ``root.after()``.
    """

    def __init__(self, parent_frame, session: SessionState, plotter, pump_ctrl, root):
        self._frame      = parent_frame
        self._session    = session
        self._plotter    = plotter
        self._pump_ctrl  = pump_ctrl
        self._root       = root

        self._queue_thread = None
        self._reorder_pending  = False
        self._reorder_snapshot = None
        self._drag_item        = None
        self._clipboard:list   = []

        self._build()

    # ── Build ─────────────────────────────────────────────────────────────────

    def _build(self):
        pane = ttk.PanedWindow(self._frame, orient=tk.VERTICAL)
        pane.pack(fill="both", expand=True)

        top    = ttk.Frame(pane); pane.add(top, weight=1)
        bottom = ttk.Frame(pane); pane.add(bottom, weight=1)

        # ── Control bar ───────────────────────────────────────────────────────
        ctrl = ttk.Frame(top)
        ctrl.pack(pady=8, fill="x", padx=10)

        ttk.Button(ctrl, text="▶ Run Queue",       command=self.run_queue).pack(side="left", padx=4)
        ttk.Button(ctrl, text="▶ From Selected",   command=self.run_from_selected).pack(side="left", padx=4)
        ttk.Button(ctrl, text="⏹ Stop",            command=self.stop_queue).pack(side="left", padx=4)
        ttk.Separator(ctrl, orient="vertical").pack(side="left", fill="y", padx=6)
        ttk.Button(ctrl, text="💾 Save",            command=self.save_queue).pack(side="left", padx=4)
        ttk.Button(ctrl, text="📂 Load",            command=self.load_queue).pack(side="left", padx=4)
        ttk.Separator(ctrl, orient="vertical").pack(side="left", fill="y", padx=6)
        ttk.Button(ctrl, text="📋 Copy",            command=self.copy_selected).pack(side="left", padx=2)
        ttk.Button(ctrl, text="📌 Paste",           command=self.paste_after_selected).pack(side="left", padx=2)
        ttk.Button(ctrl, text="⧉ Duplicate",       command=self.duplicate_selected).pack(side="left", padx=2)
        ttk.Separator(ctrl, orient="vertical").pack(side="left", fill="y", padx=6)
        ttk.Button(ctrl, text="🗑 Delete",          command=self.delete_selected).pack(side="left", padx=2)
        ttk.Button(ctrl, text="✓ Confirm Move",    command=self.confirm_reorder).pack(side="left", padx=4)
        ttk.Button(ctrl, text="🗑 Clear All",       command=self.clear_queue).pack(side="left", padx=4)

        # ── Treeview ──────────────────────────────────────────────────────────
        cols = ("Type", "Status", "Details")
        self._tree = ttk.Treeview(top, columns=cols, show="tree headings", height=10)
        self._tree.heading("#0",      text="#")
        self._tree.heading("Type",    text="Type")
        self._tree.heading("Status",  text="Status")
        self._tree.heading("Details", text="Details")
        self._tree.column("#0",      width=50)
        self._tree.column("Type",    width=150)
        self._tree.column("Status",  width=100)
        self._tree.column("Details", width=400)
        self._tree.pack(fill="both", expand=True, padx=10, pady=5)

        # Drag reorder
        self._tree.bind("<ButtonPress-1>",   self._drag_start)
        self._tree.bind("<B1-Motion>",       self._drag_motion)
        self._tree.bind("<ButtonRelease-1>", self._drag_release)

        # Right-click context menu
        self._ctx = tk.Menu(self._tree, tearoff=0)
        self._ctx.add_command(label="📋 Copy",        command=self.copy_selected)
        self._ctx.add_command(label="📌 Paste After", command=self.paste_after_selected)
        self._ctx.add_command(label="⧉ Duplicate",   command=self.duplicate_selected)
        self._ctx.add_separator()
        self._ctx.add_command(label="🗑 Delete",      command=self.delete_selected)
        self._tree.bind("<Button-3>", self._show_ctx)
        self._tree.bind("<Control-c>", lambda e: self.copy_selected())
        self._tree.bind("<Control-v>", lambda e: self.paste_after_selected())
        self._tree.bind("<Control-d>", lambda e: self.duplicate_selected())

        # ── Log panel ─────────────────────────────────────────────────────────
        log_frame = ttk.LabelFrame(bottom, text="Live Output Log")
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self._log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=10)
        self._log_text.pack(fill="both", expand=True)
        self._log_text.config(state="disabled")

        # ── Session info bar ──────────────────────────────────────────────────
        info_bar = ttk.Frame(self._frame)
        info_bar.pack(side="bottom", fill="x", padx=10, pady=(0, 2))
        self._lbl_counter  = ttk.Label(info_bar, text="Measurements this session: 0",
                                       foreground="#555")
        self._lbl_counter.pack(side="left", padx=8)
        self._lbl_registry = ttk.Label(info_bar, text="Script registry: 0 unique",
                                       foreground="#555")
        self._lbl_registry.pack(side="left", padx=8)
        ttk.Button(info_bar, text="Reset Counter",
                   command=self._reset_counter).pack(side="right", padx=4)
        ttk.Button(info_bar, text="Clear Registry",
                   command=self._clear_registry).pack(side="right", padx=4)

        # ── Status bar ────────────────────────────────────────────────────────
        self._status = ttk.Label(self._frame, text="Status: Ready", relief="sunken")
        self._status.pack(side="bottom", fill="x", padx=10, pady=5)

    # ── Public API (used by app.py and MethodTab) ─────────────────────────────

    def add_item(self, item: dict):
        """Append a queue item dict and refresh the display."""
        self._session.measurement_queue.append(item)
        self.refresh()

    def refresh(self):
        """Rebuild the Treeview from session.measurement_queue."""
        for row in self._tree.get_children():
            self._tree.delete(row)
        for i, item in enumerate(self._session.measurement_queue):
            self._tree.insert(
                "", "end", iid=str(i), text=str(i + 1),
                values=(item["type"], item["status"].upper(), item.get("details", "")),
            )

    def set_status(self, msg: str):
        self._status.config(text=f"Status: {msg}")

    def log(self, msg: str):
        def _append():
            self._log_text.config(state="normal")
            self._log_text.insert(tk.END, msg + "\n")
            self._log_text.see(tk.END)
            self._log_text.config(state="disabled")
        self._root.after(0, _append)

    def clear_log(self):
        self._log_text.config(state="normal")
        self._log_text.delete("1.0", tk.END)
        self._log_text.config(state="disabled")

    def refresh_labels(self):
        """Update session info bar labels."""
        self._lbl_counter.config(
            text=f"Measurements this session: {self._session.measurement_counter}")
        self._lbl_registry.config(
            text=f"Script registry: {self._session.registry.size} unique")

    # ── Session info bar buttons ──────────────────────────────────────────────

    def _reset_counter(self):
        self._session.reset_counter()
        self.refresh_labels()

    def _clear_registry(self):
        self._session.registry.clear()
        self.refresh_labels()

    # ── Copy / paste / duplicate ──────────────────────────────────────────────

    def _selected_indices(self) -> list:
        return sorted(
            self._tree.index(iid) for iid in self._tree.selection()
            if iid
        )

    def _show_ctx(self, event):
        row = self._tree.identify_row(event.y)
        if row:
            self._tree.selection_set(row)
        try:
            self._ctx.tk_popup(event.x_root, event.y_root)
        finally:
            self._ctx.grab_release()

    def copy_selected(self):
        idxs = self._selected_indices()
        if not idxs:
            messagebox.showwarning("No Selection", "Select item(s) to copy.")
            return
        self._clipboard = [copy.deepcopy(self._session.measurement_queue[i]) for i in idxs]
        self.set_status(f"Copied {len(self._clipboard)} item(s)")

    def paste_after_selected(self):
        if self._session.is_running:
            messagebox.showwarning("Queue Running", "Stop before editing.")
            return
        if not self._clipboard:
            messagebox.showwarning("Empty Clipboard", "Copy items first.")
            return
        idxs = self._selected_indices()
        pos  = (idxs[-1] + 1) if idxs else len(self._session.measurement_queue)
        new  = [copy.deepcopy(i) for i in self._clipboard]
        for item in new:
            item["status"] = "pending"
        self._session.measurement_queue[pos:pos] = new
        self.refresh()
        self.set_status(f"Pasted {len(new)} item(s) at position {pos + 1}")

    def duplicate_selected(self):
        if self._session.is_running:
            messagebox.showwarning("Queue Running", "Stop before editing.")
            return
        idxs = self._selected_indices()
        if not idxs:
            messagebox.showwarning("No Selection", "Select item(s) to duplicate.")
            return
        for idx in reversed(idxs):
            dupe = copy.deepcopy(self._session.measurement_queue[idx])
            dupe["status"] = "pending"
            self._session.measurement_queue.insert(idx + 1, dupe)
        self.refresh()
        self.set_status(f"Duplicated {len(idxs)} item(s)")

    def delete_selected(self):
        if self._session.is_running:
            messagebox.showwarning("Queue Running", "Stop before editing.")
            return
        idxs = self._selected_indices()
        if not idxs:
            messagebox.showwarning("No Selection", "Select item to delete.")
            return
        for idx in reversed(idxs):
            removed = self._session.measurement_queue.pop(idx)
            self.log(f"Deleted: {removed.get('details', removed.get('type'))}")
        self.refresh()
        self.set_status(f"Deleted {len(idxs)} item(s)")

    def clear_queue(self):
        self._reset_reorder()
        if self._session.is_running:
            messagebox.showwarning("Queue Running", "Stop before clearing.")
            return
        self._session.measurement_queue.clear()
        self.refresh()
        self.set_status("Queue cleared")

    # ── Drag reorder ──────────────────────────────────────────────────────────

    def _drag_start(self, event):
        if self._session.is_running:
            return
        item = self._tree.identify_row(event.y)
        if item:
            self._drag_item = item
            if not self._reorder_pending:
                self._reorder_snapshot = list(self._session.measurement_queue)

    def _drag_motion(self, event):
        if self._session.is_running or not self._drag_item:
            return
        target = self._tree.identify_row(event.y)
        if target and target != self._drag_item:
            self._tree.move(self._drag_item, "", self._tree.index(target))
            self._reorder_pending = True

    def _drag_release(self, event):
        if self._reorder_pending:
            self.set_status("Queue reorder pending — click ✓ Confirm Move")
        self._drag_item = None

    def confirm_reorder(self):
        if not self._reorder_pending or not self._reorder_snapshot:
            messagebox.showinfo("No Changes", "No pending reorder.")
            return
        try:
            order = [int(iid) for iid in self._tree.get_children()]
        except Exception:
            messagebox.showerror("Reorder Error", "Failed to read queue order.")
            return
        if any(i < 0 or i >= len(self._reorder_snapshot) for i in order):
            messagebox.showerror("Reorder Error", "Queue order out of range.")
            return
        self._session.measurement_queue = [self._reorder_snapshot[i] for i in order]
        self._reorder_snapshot = None
        self._reorder_pending  = False
        self.refresh()
        self.set_status("Queue reordered")

    def _reset_reorder(self):
        if self._reorder_pending:
            self._reorder_pending  = False
            self._reorder_snapshot = None
            self.refresh()

    # ── Save / Load ───────────────────────────────────────────────────────────

    def save_queue(self):
        if not self._session.measurement_queue:
            messagebox.showwarning("Empty Queue", "Nothing to save."); return
        if self._session.is_running:
            messagebox.showwarning("Running", "Stop the queue first."); return
        path = filedialog.asksaveasfilename(
            title="Save Queue",
            defaultextension=".json",
            filetypes=(("Queue Files", "*.json"), ("All", "*.*")),
            initialfile=f"queue_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
        )
        if not path:
            return
        payload = {
            "metadata": {"saved_at": datetime.now().isoformat(timespec="seconds"),
                         "version": 1},
            "items": [self._serialize(i) for i in self._session.measurement_queue],
        }
        try:
            with open(path, "w", encoding="utf-8") as fh:
                json.dump(payload, fh, indent=2)
            messagebox.showinfo("Saved", f"Queue saved to:\n{path}")
        except OSError as exc:
            messagebox.showerror("Save Failed", str(exc))

    def load_queue(self):
        if self._session.is_running:
            messagebox.showwarning("Running", "Stop the queue first."); return
        path = filedialog.askopenfilename(
            title="Load Queue",
            defaultextension=".json",
            filetypes=(("Queue Files", "*.json"), ("All", "*.*")),
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as fh:
                payload = json.load(fh)
            items = payload.get("items")
            if not isinstance(items, list):
                raise ValueError("Queue file missing 'items' list")
        except Exception as exc:
            messagebox.showerror("Load Failed", str(exc)); return

        new_queue, skipped = [], 0
        for raw in items:
            item = self._deserialize(raw)
            if item is None:
                skipped += 1
            else:
                new_queue.append(item)

        if not new_queue:
            messagebox.showwarning("Load Queue", "No valid items found."); return

        self._session.measurement_queue = new_queue
        self.refresh()
        self.set_status(f"Queue loaded ({len(new_queue)} items)")
        if skipped:
            self.log(f"Queue load skipped {skipped} invalid item(s).")
        messagebox.showinfo("Queue Loaded", f"Loaded {len(new_queue)} item(s).")

    @staticmethod
    def _serialize(item: dict) -> dict:
        data = {k: item.get(k) for k in ("type", "status", "details")}
        t = data["type"]
        if t == "PAUSE":
            data["pause_seconds"] = item.get("pause_seconds", 0.0)
        elif t == "ALERT":
            data["alert_message"] = item.get("alert_message", "")
        elif t and t.startswith("PUMP_"):
            action = item.get("pump_action") or {}
            data["pump_action"] = {"name": action.get("name"),
                                   "params": dict(action.get("params") or {})}
        else:
            if "script_path" in item:
                data["script_path"] = item["script_path"]
        return data

    @staticmethod
    def _deserialize(raw: dict):
        if not isinstance(raw, dict):
            return None
        t = raw.get("type")
        if not t:
            return None
        item = {"type": t, "status": "pending"}
        details = raw.get("details")
        if t == "PAUSE":
            try:
                item["pause_seconds"] = float(raw.get("pause_seconds", 0.0))
            except (TypeError, ValueError):
                return None
            item["details"] = details or f"Pause for {item['pause_seconds']:.1f} sec"
        elif t == "ALERT":
            msg = raw.get("alert_message")
            if not isinstance(msg, str) or not msg.strip():
                return None
            item["alert_message"] = msg.strip()
            item["details"]       = details or "Alert pause"
        elif t.startswith("PUMP_"):
            action = raw.get("pump_action") or {}
            if not action.get("name"):
                return None
            item["pump_action"] = {"name": action["name"],
                                   "params": dict(action.get("params") or {})}
            item["details"] = details or f"Pump action {action['name']}"
        else:
            sp = raw.get("script_path")
            if not sp:
                return None
            item["script_path"] = sp
            item["details"]     = details or Path(sp).name
        return item

    # ── Run queue ─────────────────────────────────────────────────────────────

    def run_queue(self):
        self._reset_reorder()
        if not self._session.measurement_queue:
            messagebox.showwarning("Empty Queue", "No items in queue."); return
        if self._session.is_running:
            messagebox.showwarning("Already Running", "Queue already running."); return
        self._session.is_running = True
        self.clear_log()
        self._queue_thread = threading.Thread(
            target=self._execute_queue, args=(0,), daemon=True
        )
        self._queue_thread.start()

    def run_from_selected(self):
        self._reset_reorder()
        if not self._session.measurement_queue:
            messagebox.showwarning("Empty Queue", "No items in queue."); return
        if self._session.is_running:
            messagebox.showwarning("Already Running", "Queue already running."); return
        sel = self._tree.selection()
        if not sel:
            messagebox.showwarning("No Selection", "Select a queue item to start from.")
            return
        try:
            idx = self._tree.index(sel[0])
        except Exception:
            messagebox.showerror("Selection Error", "Could not determine selected item.")
            return
        self._session.is_running = True
        self.clear_log()
        self._queue_thread = threading.Thread(
            target=self._execute_queue, args=(idx,), daemon=True
        )
        self._queue_thread.start()

    def stop_queue(self):
        if not self._session.is_running:
            return
        self._session.is_running = False
        self._session.stop_current_runner()
        self.set_status("Queue Stopped")

    def _execute_queue(self, start_index: int = 0):
        queue = list(self._session.measurement_queue)
        for i, item in enumerate(queue[start_index:], start=start_index):
            if not self._session.is_running:
                self.log("Queue execution stopped by user."); break

            self._session.measurement_queue[i]["status"] = "running"
            self._root.after(0, self.refresh)
            self._root.after(0, self.set_status,
                             f"Running: {item['type']} — {item.get('details', '')}")

            csv_path = None
            success  = False
            try:
                t = item["type"]
                if t == "PAUSE":
                    ok = self._exec_pause(float(item.get("pause_seconds", 0)))
                    self._session.measurement_queue[i]["status"] = "completed" if ok else "stopped"
                    success = ok

                elif t == "ALERT":
                    ok = self._exec_alert(item.get("alert_message", "Paused — click OK."))
                    self._session.measurement_queue[i]["status"] = "completed" if ok else "stopped"
                    success = ok

                elif t.startswith("PUMP_"):
                    ok = self._exec_pump(item)
                    self._session.measurement_queue[i]["status"] = "completed" if ok else "failed"
                    success = ok

                else:
                    self._root.after(0, self._plotter.start_live,
                                     f"{item['type']} (live)", None, item["type"])
                    try:
                        meas_tag = self._session.next_meas_tag()
                        self.log(f"[Tag] {meas_tag}")
                        self._root.after(0, self.refresh_labels)
                        runner = SerialMeasurementRunner(
                            Path(item["script_path"]),
                            log_callback=self.log,
                            data_callback=self._plotter.push_live_point,
                        )
                        self._session.current_runner = runner
                        success, csv_path = runner.execute(meas_tag=meas_tag)
                        self._session.measurement_queue[i]["status"] = (
                            "completed" if success else "failed"
                        )
                    finally:
                        self._session.current_runner = None
                        self._root.after(0, self._plotter.stop_live)

            except Exception as exc:
                self._session.measurement_queue[i]["status"] = "failed"
                self.log(f"CRITICAL ERROR in queue: {exc}")

            if csv_path:
                self._root.after(0, self._plotter.plot_data, csv_path,
                                 self._session.last_live_plot_color)
            self._root.after(0, self.refresh)
            time.sleep(1)

        self._session.is_running = False
        self._root.after(0, self.set_status, "Queue Complete")

    # ── Pause / alert helpers ─────────────────────────────────────────────────

    def _exec_pause(self, seconds: float) -> bool:
        total = max(0.0, seconds)
        start = time.time()
        while self._session.is_running:
            elapsed   = time.time() - start
            remaining = total - elapsed
            if remaining <= 0:
                break
            rem = max(0.0, remaining)
            self._root.after(0, self.set_status, f"Pausing: {rem:.1f} sec remaining")
            time.sleep(min(0.5, rem))
        if not self._session.is_running:
            return False
        self._root.after(0, self.set_status, "Pause complete")
        return True

    def _exec_alert(self, message: str) -> bool:
        if not self._session.is_running:
            return False
        done = threading.Event()
        self._root.after(0, lambda: (messagebox.showinfo("Paused", message), done.set()))
        while self._session.is_running and not done.is_set():
            done.wait(timeout=0.2)
        return done.is_set()

    # ── Pump execution ────────────────────────────────────────────────────────

    def _exec_pump(self, item: dict) -> bool:
        if self._pump_ctrl is None:
            self.log("Pump backend unavailable — skipping pump action.")
            return False
        action_info = item.get("pump_action") or {}
        name        = action_info.get("name")
        params      = action_info.get("params") or {}
        details     = item.get("details", f"Pump {name}")

        if not name:
            self.log("Invalid pump item: missing action name."); return False
        if not self._pump_ctrl.connected:
            self.log("Pump not connected."); return False

        self.log(f"Queue pump → {details}")
        try:
            if name == "INIT":
                self._pump_ctrl.initialize(); return True
            if name == "SET_SPEED":
                self._pump_ctrl.set_speed(int(params["speed"])); return True
            if name == "VALVE":
                self._pump_ctrl.valve_to(int(params["port"])); return True
            if name == "ASPIRATE":
                self._pump_ctrl.set_speed(int(params["speed"]))
                self._pump_ctrl.aspirate_ul(float(params["volume"])); return True
            if name == "DISPENSE":
                self._pump_ctrl.set_speed(int(params["speed"]))
                self._pump_ctrl.dispense_ul(float(params["volume"])); return True
            self.log(f"Unsupported pump action: {name}"); return False
        except Exception as exc:
            self.log(f"Pump action failed: {exc}"); return False
