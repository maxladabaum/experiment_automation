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
import re
import threading
import time
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk, scrolledtext, simpledialog
from typing import Optional
import math

from core.queue_eta import (
    estimate_item_seconds,
    estimate_queue_eta,
    estimate_running_queue_eta,
    eta_finish_time,
    format_duration,
)
from core.runner import SerialMeasurementRunner
from core.bo_session import BOIntegrationSession, load_bo_config, normalize_bo_config, parse_channels, validate_bo_config
from methods import library_map
from core.session import SessionState
from gui.widgets import FlowFrame


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
        self._last_selected    = None
        self._last_queue_path  = None
        self._active_alert = None
        self._completion_callbacks = []

        self._build()

    # ── Build ─────────────────────────────────────────────────────────────────

    def _build(self):
        pane = ttk.PanedWindow(self._frame, orient=tk.VERTICAL)
        pane.pack(fill="both", expand=True)

        top    = ttk.Frame(pane); pane.add(top, weight=1)
        bottom = ttk.Frame(pane); pane.add(bottom, weight=1)

        # ── Control bar ───────────────────────────────────────────────────────
        ctrl = FlowFrame(top)
        ctrl.pack(pady=8, fill="x", padx=10)

        ctrl.add(ttk.Button(ctrl, text="Run Queue", command=self.run_queue))
        ctrl.add(ttk.Button(ctrl, text="From Selected", command=self.run_from_selected))
        ctrl.add(ttk.Button(ctrl, text="Stop", command=self.stop_queue))
        ctrl.separator()
        ctrl.add(ttk.Button(ctrl, text="Save", command=self.save_queue))
        ctrl.add(ttk.Button(ctrl, text="Load", command=self.load_queue))
        ctrl.separator()
        ctrl.add(ttk.Button(ctrl, text="Copy", command=self.copy_selected))
        ctrl.add(ttk.Button(ctrl, text="Paste", command=self.paste_after_selected))
        ctrl.add(ttk.Button(ctrl, text="Duplicate", command=self.duplicate_selected))
        ctrl.separator()
        ctrl.add(ttk.Button(ctrl, text="Delete", command=self.delete_selected))
        ctrl.add(ttk.Button(ctrl, text="Confirm Move", command=self.confirm_reorder))
        ctrl.add(ttk.Button(ctrl, text="Clear All", command=self.clear_queue))

        # ── Treeview ──────────────────────────────────────────────────────────
        cols = ("Type", "Status", "Details")
        tree_frame = ttk.Frame(top)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self._tree = ttk.Treeview(
            tree_frame, columns=cols, show="tree headings", height=10, selectmode="extended"
        )
        self._tree.heading("#0",      text="#")
        self._tree.heading("Type",    text="Type")
        self._tree.heading("Status",  text="Status")
        self._tree.heading("Details", text="Details")
        self._tree.column("#0",      width=50)
        self._tree.column("Type",    width=150)
        self._tree.column("Status",  width=100)
        self._tree.column("Details", width=400)
        self._tree.pack(side="left", fill="both", expand=True)
        tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self._tree.yview)
        tree_scroll.pack(side="right", fill="y")
        self._tree.configure(yscrollcommand=tree_scroll.set)
        self._tree.tag_configure("bo", background="#e8ddff")
        self._tree.tag_configure("alert", background="#f8d7da")
        self._tree.tag_configure("pump", background="#eef3f8")
        self._tree.tag_configure("default", background="white")

        # Drag reorder
        self._tree.bind("<ButtonPress-1>",   self._drag_start)
        self._tree.bind("<B1-Motion>",       self._drag_motion)
        self._tree.bind("<ButtonRelease-1>", self._drag_release)
        self._tree.bind("<Shift-Button-1>",  self._select_range)

        # Right-click context menu
        self._ctx = tk.Menu(self._tree, tearoff=0)
        self._ctx.add_command(label="📋 Copy",        command=self.copy_selected)
        self._ctx.add_command(label="📌 Paste After", command=self.paste_after_selected)
        self._ctx.add_command(label="⧉ Duplicate",   command=self.duplicate_selected)
        self._ctx.add_command(label="Select Range…",  command=self._select_range_prompt)
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
        ttk.Button(info_bar, text="Queue ETA",
                   command=self.show_queue_eta).pack(side="right", padx=4)
        ttk.Button(info_bar, text="Clear Registry",
                   command=self._clear_registry).pack(side="right", padx=4)

        # ── Status bar ────────────────────────────────────────────────────────
        self._status = ttk.Label(self._frame, text="Status: Ready", relief="sunken")
        self._status.pack(side="bottom", fill="x", padx=10, pady=5)

    # ── Public API (used by app.py and MethodTab) ─────────────────────────────

    def add_item(self, item: dict):
        """Append a queue item dict and refresh the display."""
        prepared = item
        if isinstance(item, dict) and "script_path" not in item and "method_ref" in item:
            resolved = self._deserialize(item)
            if resolved is not None:
                prepared = resolved
        self._session.measurement_queue.append(prepared)
        self.refresh()

    def refresh(self):
        """Rebuild the Treeview from session.measurement_queue."""
        for row in self._tree.get_children():
            self._tree.delete(row)
        for i, item in enumerate(self._session.measurement_queue):
            tag = self._row_tag_for_item(item)
            self._tree.insert(
                "", "end", iid=str(i), text=str(i + 1),
                values=(item["type"], item["status"].upper(), item.get("details", "")),
                tags=(tag,),
                open=(str(item.get("type") or "").upper() == "BO_AUTO_LOOP"),
            )
            if str(item.get("type") or "").upper() == "BO_AUTO_LOOP":
                for j, progress in enumerate(item.get("bo_progress") or []):
                    self._tree.insert(
                        str(i),
                        "end",
                        iid=f"{i}:bo:{j}",
                        text=f"{i + 1}.{j + 1}",
                        values=(
                            progress.get("type", "BO_STEP"),
                            str(progress.get("status", "")).upper(),
                            progress.get("details", ""),
                        ),
                        tags=("bo",),
                    )

    @staticmethod
    def _row_tag_for_item(item: dict) -> str:
        item_type = str(item.get("type") or "").upper()
        if item_type == "BO_AUTO_LOOP":
            return "bo"
        if item_type in ("PAUSE", "ALERT"):
            return "alert"
        if item_type.startswith("PUMP_"):
            return "pump"
        return "default"

    def set_status(self, msg: str):
        self._status.config(text=f"Status: {msg}")

    def _set_bo_live_details(self, item: dict, details: str):
        item["details"] = details
        self._root.after(0, self.refresh)
        self._root.after(0, self.set_status, f"Running: {details}")

    def _append_bo_progress(self, item: dict | None, step_type: str, status: str, details: str) -> Optional[dict]:
        if item is None:
            return None
        progress = item.setdefault("bo_progress", [])
        record = {
            "type": step_type,
            "status": status,
            "details": details,
            "updated_at": datetime.now().isoformat(timespec="seconds"),
        }
        progress.append(record)
        if len(progress) > 300:
            del progress[:-300]
        self._root.after(0, self.refresh)
        return record

    def _update_bo_progress(self, record: Optional[dict], status: str, details: Optional[str] = None):
        if record is None:
            return
        record["status"] = status
        if details is not None:
            record["details"] = details
        record["updated_at"] = datetime.now().isoformat(timespec="seconds")
        self._root.after(0, self.refresh)

    def log(self, msg: str):
        session_mgr = getattr(self._session, "session_manager", None)
        if session_mgr is not None:
            session_mgr.log(msg)
            return
        self._append_log_gui(msg)

    def _append_log_gui(self, msg: str):
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

    def add_completion_callback(self, callback):
        """Register a callback called after a queue run reaches its final state."""
        if callable(callback) and callback not in self._completion_callbacks:
            self._completion_callbacks.append(callback)


    # ── Session info bar buttons ──────────────────────────────────────────────

    def _reset_counter(self):
        self._session.reset_counter()
        self.refresh_labels()

    def _clear_registry(self):
        self._session.registry.clear()
        self.refresh_labels()

    # ── Copy / paste / duplicate ──────────────────────────────────────────────

    def show_queue_eta(self):
        """Show an on-demand ETA estimate for the queue or current selection."""
        queue = self._session.measurement_queue
        if not queue:
            messagebox.showinfo("Queue ETA", "Queue is empty.")
            return

        idxs = self._selected_indices()
        lines = []
        if self._session.is_running:
            lines.extend(self._build_live_eta_lines())
            if idxs:
                start_index = idxs[0]
                scope = f"selected item #{start_index + 1} onward"
                lines.append("")
                lines.extend(self._build_static_eta_lines(start_index, scope))
        else:
            start_index = idxs[0] if idxs else 0
            scope = f"selected item #{start_index + 1} onward" if idxs else "entire queue"
            lines.extend(self._build_static_eta_lines(start_index, scope))

        messagebox.showinfo("Queue ETA", "\n".join(lines))

    def get_slack_eta_text(self) -> str:
        queue = self._session.measurement_queue
        if not queue:
            return self._with_slackbot_loves_baris("Queue is empty.")
        if self._session.is_running:
            return self._with_slackbot_loves_baris(self._build_slack_remaining_text())
        return self._with_slackbot_loves_baris(
            f"Remaining measurements: {self._count_measurement_items(queue, start_index=0)}"
        )

    @staticmethod
    def _with_slackbot_loves_baris(text: str) -> str:
        return f"{text}\nslackbot loves baris"

    def _build_slack_remaining_text(self) -> str:
        queue = self._session.measurement_queue
        status = self._session.get_queue_status()
        if str(status.get("bo_mode") or "").strip().lower() == "paired_response":
            completed = self._coerce_int(status.get("bo_completed_measurements")) or 0
            total_measurements = self._coerce_int(status.get("bo_total_measurements"))
            cycle = self._coerce_int(status.get("bo_cycle_current"))
            total_cycles = self._coerce_int(status.get("bo_cycle_total"))
            observed = self._coerce_int(status.get("bo_observed_sets")) or 0
            total_sets = self._coerce_int(status.get("bo_total_sets"))
            remaining = (
                max(0, total_measurements - completed)
                if total_measurements is not None
                else self._count_measurement_items(queue, start_index=0)
            )
            phase = str(status.get("bo_phase") or status.get("active_step_details") or "").strip()
            cycle_text = f"cycle {cycle}/{total_cycles}" if cycle is not None and total_cycles is not None else "cycle ?"
            set_text = f" | observed sets: {observed}/{total_sets}" if total_sets is not None else ""
            phase_text = f" | phase: {phase}" if phase else ""
            return f"Paired BO status: {cycle_text}{phase_text} | remaining measurements: {remaining}{set_text}"
        active_index = self._coerce_int(status.get("active_queue_index"))
        current_index = self._coerce_int(status.get("current_index"))
        total = self._coerce_int(status.get("total"))
        remaining = self._count_measurement_items(queue, start_index=active_index if active_index is not None else 0)
        if current_index is not None and total is not None:
            return f"Queue status: step {current_index}/{total} | remaining measurements: {remaining}"
        return f"Remaining measurements: {remaining}"

    @staticmethod
    def _count_measurement_items(queue, start_index: int = 0) -> int:
        count = 0
        for item in list(queue)[max(0, int(start_index or 0)):]:
            item_type = str((item or {}).get("type") or "").strip().upper()
            if item_type in {"CV", "SWV", "DPV", "LSV", "EIS", "CUSTOM", "CUSTOM_MUX"}:
                count += 1
            elif item_type == "BO_AUTO_LOOP":
                try:
                    block = ((item or {}).get("bo_block") or {})
                    target = int(block.get("target_iterations", 0) or 0)
                    if str(block.get("objective") or "").strip().lower() == "paired_response":
                        target *= max(1, int(block.get("batch_size", 1) or 1)) * 2
                except Exception:
                    target = 0
                count += max(0, target)
        return count

    def _build_static_eta_lines(self, start_index: int, scope: str) -> list:
        eta = estimate_queue_eta(
            self._session.measurement_queue,
            start_index=start_index,
            step_delay_seconds=getattr(self._session, "step_delay", 0.0) or 0.0,
        )
        finish_at = eta_finish_time(eta.total_seconds)
        lines = [
            f"Estimate scope: {scope}",
            f"Predicted duration: {format_duration(eta.total_seconds)}",
            f"Estimated finish: {finish_at.strftime('%Y-%m-%d %I:%M:%S %p')}",
        ]
        return lines + self._eta_caveat_lines(eta.unknown_item_count, eta.excluded_alert_count)

    def _build_live_eta_lines(self) -> list:
        queue = self._session.measurement_queue
        status = self._session.get_queue_status()
        active_index = self._coerce_int(status.get("active_queue_index"))
        next_index = self._coerce_int(status.get("next_queue_index"))
        total = self._coerce_int(status.get("total"))
        current_index = self._coerce_int(status.get("current_index"))
        if active_index is None or active_index < 0 or active_index >= len(queue):
            return self._build_static_eta_lines(0, "entire queue")
        if next_index is None:
            next_index = min(active_index + 1, len(queue))

        state = str(status.get("state") or "running").lower()
        details = status.get("active_step_details") or status.get("current_label") or "(unknown)"
        elapsed_seconds = self._elapsed_since(status.get("active_step_started_at"))
        estimated_seconds = self._coerce_float(status.get("active_step_estimated_seconds"))
        include_next_step_delay = str(status.get("active_step_type") or "").upper() != "STEP_DELAY"

        eta = estimate_running_queue_eta(
            queue,
            next_index=next_index,
            current_step_elapsed_seconds=elapsed_seconds,
            current_step_estimated_seconds=estimated_seconds,
            step_delay_seconds=getattr(self._session, "step_delay", 0.0) or 0.0,
            include_next_step_delay=include_next_step_delay,
        )

        step_label = f"{current_index}/{total}" if current_index is not None and total is not None else "?"
        lines = [
            "Estimate scope: active run",
            f"Current step: {step_label} | {details}",
        ]
        if eta.current_step_predictable and eta.current_step_remaining_seconds is not None:
            lines.append(
                f"Current step remaining: {format_duration(eta.current_step_remaining_seconds)}"
            )
        elif state == "waiting_alert":
            lines.append("Current step remaining: waiting for alert acknowledgment")
        else:
            lines.append("Current step remaining: unknown")

        lines.append(
            f"Remaining after this step: {format_duration(eta.remaining_after_current_seconds)}"
        )
        if eta.total_remaining_seconds is not None:
            finish_at = eta_finish_time(eta.total_remaining_seconds)
            lines.append(f"Total remaining: {format_duration(eta.total_remaining_seconds)}")
            lines.append(f"Estimated finish: {finish_at.strftime('%Y-%m-%d %I:%M:%S %p')}")
        else:
            lines.append("Total remaining: unknown until the current step finishes")

        return lines + self._eta_caveat_lines(eta.unknown_item_count, eta.excluded_alert_count)

    @staticmethod
    def _eta_caveat_lines(unknown_item_count: int, excluded_alert_count: int) -> list:
        lines = []
        if unknown_item_count:
            lines.append(f"Unknown items not counted: {unknown_item_count}")
        if excluded_alert_count:
            lines.append(f"Alert pauses treated as 0 sec: {excluded_alert_count}")
        if not unknown_item_count and not excluded_alert_count:
            lines.append("All queued items were estimated.")
        return lines

    @staticmethod
    def _coerce_int(value) -> Optional[int]:
        try:
            return int(value)
        except (TypeError, ValueError):
            return None

    @staticmethod
    def _coerce_float(value) -> Optional[float]:
        try:
            return float(value)
        except (TypeError, ValueError):
            return None

    @staticmethod
    def _elapsed_since(timestamp_text) -> float:
        if not timestamp_text:
            return 0.0
        try:
            started_at = datetime.fromisoformat(str(timestamp_text))
        except (TypeError, ValueError):
            return 0.0
        return max(0.0, (datetime.now() - started_at).total_seconds())

    def _selected_indices(self) -> list:
        return sorted(
            self._tree.index(iid) for iid in self._tree.selection()
            if iid
        )

    def _select_range(self, event):
        row = self._tree.identify_row(event.y)
        if not row:
            return
        if self._last_selected is None:
            self._tree.selection_set(row)
            self._last_selected = row
            return
        try:
            start = self._tree.index(self._last_selected)
            end = self._tree.index(row)
        except Exception:
            self._tree.selection_set(row)
            self._last_selected = row
            return
        if start > end:
            start, end = end, start
        self._tree.selection_set(self._tree.get_children()[start:end + 1])
        self._last_selected = row

    def _select_range_prompt(self):
        total = len(self._tree.get_children())
        if total == 0:
            return
        start = simpledialog.askinteger(
            "Select Range",
            f"Start row (1-{total}):",
            minvalue=1, maxvalue=total
        )
        if start is None:
            return
        end = simpledialog.askinteger(
            "Select Range",
            f"End row (1-{total}):",
            minvalue=1, maxvalue=total
        )
        if end is None:
            return
        if start > end:
            start, end = end, start
        children = self._tree.get_children()
        self._tree.selection_set(children[start - 1:end])
        self._last_selected = children[end - 1]

    def _show_ctx(self, event):
        row = self._tree.identify_row(event.y)
        if row:
            self._tree.selection_set(row)
            self._last_selected = row
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
            self.log(f"Queue item deleted: {removed.get('details', removed.get('type'))}")
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
        self.log("Queue cleared.")

    # ── Drag reorder ──────────────────────────────────────────────────────────

    def _drag_start(self, event):
        if self._session.is_running:
            return
        item = self._tree.identify_row(event.y)
        if item:
            self._last_selected = item
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
            self.log(f"Queue saved: {path}")
            self._last_queue_path = path
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
            with open(path, "r", encoding="utf-8-sig") as fh:
                payload = json.load(fh)
            if isinstance(payload, list):
                items = payload
            else:
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
        self.log(f"Queue loaded: {path} ({len(new_queue)} items)")
        self._last_queue_path = path
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
        elif t == "BO_AUTO_LOOP":
            data["bo_block"] = dict(item.get("bo_block") or {})
        elif t and t.startswith("PUMP_"):
            action = item.get("pump_action") or {}
            data["pump_action"] = {"name": action.get("name"),
                                   "params": dict(action.get("params") or {})}
        else:
            if "script_path" in item:
                data["script_path"] = item["script_path"]
            if "method_ref" in item:
                data["method_ref"] = dict(item.get("method_ref") or {})
            if "bo_ref" in item:
                data["bo_ref"] = dict(item.get("bo_ref") or {})
            for key in ("meas_tag", "csv_path", "completed_at", "failed_at"):
                if key in item:
                    data[key] = item.get(key)
        return data

    def _deserialize(self, raw: dict):
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
        elif t == "BO_AUTO_LOOP":
            block = raw.get("bo_block")
            if not isinstance(block, dict):
                return None
            item["bo_block"] = dict(block)
            item["details"] = details or self._format_bo_block_details(block)
        elif t.startswith("PUMP_"):
            action = raw.get("pump_action") or {}
            if not action.get("name"):
                return None
            item["pump_action"] = {"name": action["name"],
                                   "params": dict(action.get("params") or {})}
            item["details"] = details or f"Pump action {action['name']}"
        else:
            sp = raw.get("script_path")
            method_ref = raw.get("method_ref") or {}

            if sp:
                # Prefer exact library_map entry when provided.
                hash_key = method_ref.get("hash_key")
                if hash_key:
                    resolved = library_map.lookup(hash_key)
                    if resolved is not None:
                        sp = str(resolved)

                # Prefer MUX-specific library file if method_ref requests a channel.
                mux = method_ref.get("mux_channel")
                mux_ch = None
                if mux not in (None, "", 0, "0"):
                    try:
                        mux_ch = int(mux)
                    except (TypeError, ValueError):
                        mux_ch = None

                if mux_ch is not None and 1 <= mux_ch <= 16:
                    technique = method_ref.get("technique") or t
                    params = method_ref.get("params")
                    resolved = None
                    if isinstance(params, dict):
                        try:
                            mux_key = library_map.compute_hash(technique, params, mux_ch)
                            resolved = library_map.lookup(mux_key)
                        except Exception:
                            resolved = None
                    if resolved is not None:
                        sp = str(resolved)
                        item["details"] = details or f"{Path(sp).name} (MUX ch {mux_ch})"

            if not sp:
                hash_key = method_ref.get("hash_key")
                if hash_key:
                    path = library_map.lookup(hash_key)
                    if path is None:
                        return None
                    mux = method_ref.get("mux_channel")
                    if mux not in (None, "", 0, "0"):
                        try:
                            mux_ch = int(mux)
                        except (TypeError, ValueError):
                            mux_ch = None

                        if mux_ch is not None and 1 <= mux_ch <= 16:
                            technique = method_ref.get("technique") or t
                            params = method_ref.get("params")
                            resolved = None

                            if isinstance(params, dict):
                                try:
                                    mux_key = library_map.compute_hash(technique, params, mux_ch)
                                    resolved = library_map.lookup(mux_key)
                                except Exception:
                                    resolved = None

                            if resolved is None:
                                # Fallback: wrap the referenced base script with the requested channel.
                                try:
                                    base_script = path.read_text(encoding="utf-8")
                                    wrapped = self._wrap_mux(
                                        self._strip_first_mux_header(base_script),
                                        mux_ch,
                                    )
                                    mux_note = self._compose_mux_note(
                                        method_ref=method_ref,
                                        mux_channel=mux_ch,
                                        fallback=f"MUX ch {mux_ch}",
                                    )
                                    saved_path, _ = self._session.registry.save_script(
                                        technique=technique,
                                        script=wrapped,
                                        params=params if isinstance(params, dict) else None,
                                        mux_channel=mux_ch,
                                        note=mux_note,
                                    )
                                    resolved = saved_path
                                except Exception as exc:
                                    self.log(f"Failed to generate MUX ch {mux_ch} script from method_ref: {exc}")
                                    return None

                            sp = str(resolved)
                            item["details"] = details or f"{Path(sp).name} (MUX ch {mux_ch})"
                        else:
                            sp = str(path)
                            item["details"] = details or path.name
                    else:
                        sp = str(path)
                        item["details"] = details or path.name
                else:
                    return None

            item["script_path"] = sp
            if "method_ref" in raw and isinstance(raw.get("method_ref"), dict):
                item["method_ref"] = dict(raw.get("method_ref") or {})
            if "bo_ref" in raw and isinstance(raw.get("bo_ref"), dict):
                item["bo_ref"] = dict(raw.get("bo_ref") or {})
            for key in ("meas_tag", "csv_path", "completed_at", "failed_at"):
                if key in raw:
                    item[key] = raw.get(key)
            item["details"]     = item.get("details") or details or Path(sp).name
        return item

    @staticmethod
    def _mux_channel_address(channel: int) -> int:
        idx = channel - 1
        return (idx << 4) | idx

    @classmethod
    def _wrap_mux(cls, base_script: str, channel: int) -> str:
        lines = base_script.splitlines()
        header = lines[0].strip() if lines and lines[0].strip() in ("e", "l") else "e"
        rest = lines[1:] if lines and lines[0].strip() in ("e", "l") else lines
        addr = cls._mux_channel_address(channel)
        prefix = [
            header,
            "# MUX16 channel select",
            "set_gpio_cfg 0x3FFi 1",
            f"set_gpio {addr}i",
        ]
        return "\n".join(prefix + rest)

    @staticmethod
    def _strip_first_mux_header(script: str) -> str:
        lines = script.splitlines()
        cfg_idx = None
        gpio_idx = None
        for i, line in enumerate(lines):
            s = line.strip()
            if cfg_idx is None and s == "set_gpio_cfg 0x3FFi 1":
                cfg_idx = i
                continue
            if cfg_idx is not None and gpio_idx is None and s.startswith("set_gpio ") and not s.startswith("set_gpio_cfg"):
                gpio_idx = i
                break
        if cfg_idx is not None and gpio_idx is not None:
            del lines[gpio_idx]
            del lines[cfg_idx]
        return "\n".join(lines)

    @staticmethod
    def _extract_mux_from_script(script: str) -> Optional[int]:
        """Read first set_gpio value and decode nibble-pair channel (0x11 -> ch2)."""
        for line in script.splitlines():
            s = line.strip()
            if not s.startswith("set_gpio ") or s.startswith("set_gpio_cfg"):
                continue
            token = s[len("set_gpio "):].strip()
            if token.endswith("i"):
                token = token[:-1]
            try:
                value = int(token, 16) if token.lower().startswith("0x") else int(token)
            except ValueError:
                continue
            lo = value & 0x0F
            hi = (value >> 4) & 0x0F
            if lo == hi and 0 <= lo <= 15:
                return lo + 1
            return None
        return None

    def _compose_mux_note(self, method_ref: dict, mux_channel: int, fallback: str) -> str:
        """Build note using original method note (if any) + current channel tag."""
        base_note = ""
        if isinstance(method_ref, dict):
            hash_key = method_ref.get("hash_key")
            if hash_key:
                try:
                    entry = library_map.all_entries().get(hash_key) or {}
                    base_note = (entry.get("note") or "").strip()
                except Exception:
                    base_note = ""

        tag = f"MUX ch {mux_channel}"
        if base_note:
            if re.search(r"\bMUX\s*ch\s*\d+\b", base_note, flags=re.IGNORECASE):
                return re.sub(
                    r"\bMUX\s*ch\s*\d+\b",
                    tag,
                    base_note,
                    flags=re.IGNORECASE,
                )
            return f"{base_note} | {tag}"
        return fallback

    # ── Run queue ─────────────────────────────────────────────────────────────

    def run_queue(self):
        self._reset_reorder()
        if not self._session.measurement_queue:
            messagebox.showwarning("Empty Queue", "No items in queue."); return
        if self._session.is_running:
            messagebox.showwarning("Already Running", "Queue already running."); return
        self._session.is_running = True
        self._session.update_queue_status(
            state="running",
            current_index=0,
            total=len(self._session.measurement_queue),
            current_label="(starting)",
            started_at=datetime.now().isoformat(timespec="seconds"),
            queue_start_index=0,
            active_queue_index=None,
            next_queue_index=0,
            active_step_started_at=None,
            active_step_estimated_seconds=None,
            active_step_type=None,
            active_step_details=None,
            bo_mode=None,
            bo_cycle_current=None,
            bo_cycle_total=None,
            bo_phase=None,
            bo_completed_measurements=None,
            bo_total_measurements=None,
            bo_observed_sets=None,
            bo_total_sets=None,
        )
        self.clear_log()
        self.log("Queue start requested.")
        self.log(f"Measurement simulation: {'ON' if self._session.simulate_measurements else 'OFF'}")
        self._announce_queue_start(start_index=0)
        self._copy_queue_file("run_queue")
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
        self.run_from_index(idx)

    def run_from_index(self, idx: int):
        """Start queue execution at an explicit index without replaying prior items."""
        self._reset_reorder()
        if not self._session.measurement_queue:
            messagebox.showwarning("Empty Queue", "No items in queue."); return
        if self._session.is_running:
            messagebox.showwarning("Already Running", "Queue already running."); return
        try:
            idx = int(idx)
        except (TypeError, ValueError):
            messagebox.showerror("Queue Error", "Invalid queue start index.")
            return
        if idx < 0 or idx >= len(self._session.measurement_queue):
            messagebox.showerror("Queue Error", "Queue start index is out of range.")
            return
        self._session.is_running = True
        self._session.update_queue_status(
            state="running",
            current_index=0,
            total=len(self._session.measurement_queue) - idx,
            current_label="(starting)",
            started_at=datetime.now().isoformat(timespec="seconds"),
            queue_start_index=idx,
            active_queue_index=None,
            next_queue_index=idx,
            active_step_started_at=None,
            active_step_estimated_seconds=None,
            active_step_type=None,
            active_step_details=None,
            bo_mode=None,
            bo_cycle_current=None,
            bo_cycle_total=None,
            bo_phase=None,
            bo_completed_measurements=None,
            bo_total_measurements=None,
            bo_observed_sets=None,
            bo_total_sets=None,
        )
        self.clear_log()
        self.log(f"Queue start requested from item {idx + 1}.")
        self.log(f"Measurement simulation: {'ON' if self._session.simulate_measurements else 'OFF'}")
        self._announce_queue_start(start_index=idx)
        self._copy_queue_file("run_queue_from_selected")
        self._queue_thread = threading.Thread(
            target=self._execute_queue, args=(idx,), daemon=True
        )
        self._queue_thread.start()

    def stop_queue(self):
        if not self._session.is_running:
            return
        self.log("Queue stop requested.")
        self._session.is_running = False
        self._session.stop_current_runner()
        self._session.update_queue_status(state="stopping")
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
            item_eta = estimate_item_seconds(item)
            if str(item.get("type", "")).strip().upper() == "ALERT":
                item_eta = None
            self._session.update_queue_status(
                state="running",
                current_index=(i - start_index + 1),
                total=len(queue) - start_index,
                current_label=(item.get("details") or item.get("type") or ""),
                queue_start_index=start_index,
                active_queue_index=i,
                next_queue_index=min(i + 1, len(queue)),
                active_step_started_at=datetime.now().isoformat(timespec="seconds"),
                active_step_estimated_seconds=item_eta,
                active_step_type=item.get("type"),
                active_step_details=(item.get("details") or item.get("type") or ""),
            )
            self.log(f"Queue start -> {item.get('details', item.get('type'))}")

            csv_path = None
            success  = False
            try:
                t = item["type"]
                if t == "PAUSE":
                    ok = self._exec_pause(float(item.get("pause_seconds", 0)))
                    self._session.measurement_queue[i]["status"] = "completed" if ok else "stopped"
                    success = ok

                elif t == "ALERT":
                    alert_msg = item.get("alert_message", "Paused — click OK.")
                    session_mgr = getattr(self._session, "session_manager", None)
                    if session_mgr is not None:
                        session_mgr.notify_slack(f"Queue alert: {alert_msg}")
                    ok = self._exec_alert(alert_msg)
                    self._session.measurement_queue[i]["status"] = "completed" if ok else "stopped"
                    success = ok

                elif t.startswith("PUMP_"):
                    ok = self._exec_pump(item)
                    self._session.measurement_queue[i]["status"] = "completed" if ok else "failed"
                    if not ok:
                        self.log(f"Queue item FAILED: {t} | {item.get('details', '')}")
                    success = ok

                else:
                    self._ensure_mux_script_for_item(item)
                    self._root.after(0, self._plotter.start_live,
                                     f"{item['type']} (live)", None, item["type"])
                    try:
                        mux_channel = self._extract_mux_channel(item)
                        meas_tag = self._session.next_meas_tag_with_mux(mux_channel)
                        self._session.measurement_queue[i]["meas_tag"] = meas_tag
                        self.log(f"[Tag] {meas_tag}")
                        self._root.after(0, self.refresh_labels)
                        data_folder = None
                        if self._session.session_manager is not None:
                            data_folder = self._session.session_manager.require_experiment()
                            if data_folder is None:
                                self._session.measurement_queue[i]["status"] = "failed"
                                self._root.after(0, self.refresh)
                                break
                        runner = SerialMeasurementRunner(
                            Path(item["script_path"]),
                            log_callback=self.log,
                            data_callback=self._plotter.push_live_point,
                            data_folder=data_folder,
                            save_raw_packets=self._session.save_raw_packets,
                            simulate_measurements=self._session.simulate_measurements,
                            invert_current=(item.get("type") == "SWV"),
                            device_port=self._session.device_port,
                        )
                        self._session.current_runner = runner
                        success, csv_path = runner.execute(meas_tag=meas_tag)
                        if csv_path:
                            self._session.measurement_queue[i]["csv_path"] = str(csv_path)
                        if success:
                            self._session.measurement_queue[i]["status"] = "completed"
                            self._session.measurement_queue[i]["completed_at"] = datetime.now().isoformat(timespec="seconds")
                        else:
                            self._session.measurement_queue[i]["status"] = "failed"
                            self._session.measurement_queue[i]["failed_at"] = datetime.now().isoformat(timespec="seconds")
                            self.log(f"Queue item FAILED: {item['type']} | {item.get('details', meas_tag)}")
                    finally:
                        self._session.current_runner = None
                        self._root.after(0, self._plotter.stop_live)

            except Exception as exc:
                self._session.measurement_queue[i]["status"] = "failed"
                self.log(f"CRITICAL ERROR in queue: {exc}")

            if csv_path:
                self._root.after(0, self._plotter.plot_data, csv_path,
                                 self._session.last_live_plot_color, None, True, False)
            self._root.after(0, self.refresh)
            step_delay = getattr(self._session, "step_delay", 0.0) or 0.0
            if step_delay > 0 and i < len(queue) - 1:
                next_step_number = i - start_index + 2
                delay_label = f"Inter-step delay before step {next_step_number}"
                self._session.update_queue_status(
                    state="step_delay",
                    current_index=(i - start_index + 1),
                    total=len(queue) - start_index,
                    current_label=delay_label,
                    queue_start_index=start_index,
                    active_queue_index=i,
                    next_queue_index=i + 1,
                    active_step_started_at=datetime.now().isoformat(timespec="seconds"),
                    active_step_estimated_seconds=step_delay,
                    active_step_type="STEP_DELAY",
                    active_step_details=delay_label,
                )
                if not self._exec_pause(step_delay):
                    break

        self._session.is_running = False
        self._session.update_queue_status(
            state="idle",
            current_label="Queue Complete",
            active_queue_index=None,
            next_queue_index=None,
            active_step_started_at=None,
            active_step_estimated_seconds=None,
            active_step_type=None,
            active_step_details=None,
            bo_mode=None,
            bo_cycle_current=None,
            bo_cycle_total=None,
            bo_phase=None,
            bo_completed_measurements=None,
            bo_total_measurements=None,
            bo_observed_sets=None,
            bo_total_sets=None,
        )
        self.log("Queue completed.")
        self._root.after(0, self.set_status, "Queue Complete")
        self._announce_queue_end(start_index=start_index)
        self._notify_completion_callbacks(start_index=start_index)

    def _notify_completion_callbacks(self, start_index: int):
        ran = self._session.measurement_queue[start_index:]
        summary = {
            "start_index": start_index,
            "total": len(ran),
            "completed": sum(1 for item in ran if item.get("status") == "completed"),
            "failed": sum(1 for item in ran if item.get("status") == "failed"),
            "stopped": sum(1 for item in ran if item.get("status") == "stopped"),
            "items": [dict(item) for item in ran],
        }
        for callback in list(self._completion_callbacks):
            try:
                self._root.after(0, lambda cb=callback, data=dict(summary): cb(data))
            except Exception as exc:
                self.log(f"Queue completion callback failed: {exc}")

    def _announce_queue_start(self, start_index: int):
        session_mgr = getattr(self._session, "session_manager", None)
        if session_mgr is None:
            return
        ran = self._session.measurement_queue[start_index:]
        if self._is_bo_only_run(ran):
            return
        total = max(0, len(self._session.measurement_queue) - start_index)
        session_name = (
            session_mgr.current_session_path.name
            if session_mgr.current_session_path is not None
            else "(none)"
        )
        experiment_name = (
            session_mgr.current_experiment_path.name
            if session_mgr.current_experiment_path is not None
            else "(none)"
        )
        msg = (
            f"Queue started: {total} item(s). "
            f"Session={session_name}; Experiment={experiment_name}."
        )
        session_mgr.notify_slack(msg)

    def _announce_queue_end(self, start_index: int):
        session_mgr = getattr(self._session, "session_manager", None)
        if session_mgr is None:
            return

        ran = self._session.measurement_queue[start_index:]
        if not ran:
            return
        if self._is_bo_only_run(ran):
            return

        total = len(ran)
        completed = sum(1 for item in ran if item.get("status") == "completed")
        failed = sum(1 for item in ran if item.get("status") == "failed")
        stopped = sum(1 for item in ran if item.get("status") == "stopped")

        if stopped > 0:
            state = "STOPPED"
        elif failed > 0:
            state = "FAILED"
        else:
            state = "COMPLETED"

        self._session.update_queue_status(
            state=state.lower(),
            current_index=total,
            total=total,
            current_label="(finished)",
            queue_start_index=None,
            active_queue_index=None,
            next_queue_index=None,
            active_step_started_at=None,
            active_step_estimated_seconds=None,
            active_step_type=None,
            active_step_details=None,
        )

        session_name = (
            session_mgr.current_session_path.name
            if session_mgr.current_session_path is not None
            else "(none)"
        )
        experiment_name = (
            session_mgr.current_experiment_path.name
            if session_mgr.current_experiment_path is not None
            else "(none)"
        )
        msg = (
            f"Queue {state}: completed={completed}/{total}, "
            f"failed={failed}, stopped={stopped}. "
            f"Session={session_name}; Experiment={experiment_name}."
        )
        session_mgr.notify_slack(msg)

    @staticmethod
    def _is_bo_only_run(items) -> bool:
        if not items:
            return False
        for item in items:
            item_type = str((item or {}).get("type") or "").strip().upper()
            if item_type == "BO_AUTO_LOOP":
                continue
            bo_ref = (item or {}).get("bo_ref")
            if isinstance(bo_ref, dict) and str(bo_ref.get("session_id") or "").strip():
                continue
            return False
        return True

    def _ensure_mux_script_for_item(self, item: dict):
        """Auto-correct script_path to requested MUX channel before execution."""
        mux_channel = self._extract_mux_channel(item)
        if mux_channel is None:
            return
        script_path = item.get("script_path")
        if not script_path:
            return

        src = Path(script_path)
        try:
            base_script = src.read_text(encoding="utf-8")
        except Exception as exc:
            self.log(f"Warning: could not read script for MUX verification: {exc}")
            return

        current_mux = self._extract_mux_from_script(base_script)
        if current_mux == mux_channel:
            return

        wrapped = self._wrap_mux(self._strip_first_mux_header(base_script), mux_channel)
        method_ref = item.get("method_ref") or {}
        params = method_ref.get("params")
        try:
            mux_note = self._compose_mux_note(
                method_ref=method_ref,
                mux_channel=mux_channel,
                fallback=f"MUX ch {mux_channel}",
            )
            saved_path, saved_name = self._session.registry.save_script(
                technique=item.get("type", ""),
                script=wrapped,
                params=params if isinstance(params, dict) else None,
                mux_channel=mux_channel,
                note=mux_note,
            )
            item["script_path"] = str(saved_path)
            self.log(
                f"Adjusted script for MUX ch {mux_channel}: {src.name} -> {saved_name}"
            )
        except Exception as exc:
            self.log(f"Warning: failed to adjust script for MUX ch {mux_channel}: {exc}")

    def _queue_payload(self) -> dict:
        return {
            "metadata": {"saved_at": datetime.now().isoformat(timespec="seconds"),
                         "version": 1},
            "items": [self._serialize(i) for i in self._session.measurement_queue],
        }

    def _copy_queue_file(self, prefix: str):
        session_mgr = getattr(self._session, "session_manager", None)
        exp_path = getattr(session_mgr, "current_experiment_path", None) if session_mgr else None
        if exp_path is None:
            return
        try:
            queue_dir = Path(exp_path) / "queue_files"
            queue_dir.mkdir(parents=True, exist_ok=True)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            suffix = ""
            if self._last_queue_path:
                try:
                    suffix = f"_{Path(self._last_queue_path).name}"
                except Exception:
                    suffix = ""
            filename = f"{prefix}_{ts}{suffix}"
            dst = queue_dir / filename
            with open(dst, "w", encoding="utf-8") as fh:
                json.dump(self._queue_payload(), fh, indent=2)
            self.log(f"Queue file copied to: {dst}")
        except Exception as exc:
            self.log(f"Queue file copy failed: {exc}")

    @staticmethod
    def _extract_mux_channel(item: dict) -> Optional[int]:
        method_ref = item.get("method_ref") or {}
        mux = method_ref.get("mux_channel")
        if mux is not None:
            try:
                return int(mux)
            except (TypeError, ValueError):
                pass
        details = str(item.get("details") or "")
        m = re.search(r"\bMUX\s*ch\s*(\d+)\b", details, flags=re.IGNORECASE)
        if m:
            try:
                return int(m.group(1))
            except ValueError:
                return None
        return None

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
        self._active_alert = {"event": done, "window": None, "message": message}
        self._session.update_queue_status(
            state="waiting_alert",
            current_label=message,
            active_step_details=message,
        )
        self._root.after(0, lambda: self._show_alert_window(message, done))
        try:
            while self._session.is_running and not done.is_set():
                done.wait(timeout=0.2)
            return done.is_set()
        finally:
            self._root.after(0, self._close_active_alert_window)
            self._active_alert = None

    def _show_alert_window(self, message: str, done: threading.Event):
        if done.is_set():
            return
        win = tk.Toplevel(self._root)
        win.title("Queue Alert")
        win.transient(self._root)
        win.resizable(False, False)
        win.grab_set()
        container = ttk.Frame(win, padding=(18, 16, 18, 14))
        container.pack(fill="both", expand=True)
        ttk.Label(container, text=message, justify="left", wraplength=460).pack(
            padx=18, pady=(16, 10), fill="x"
        )
        ttk.Button(container, text="OK", command=self._acknowledge_active_alert).pack(pady=(0, 4))
        win.protocol("WM_DELETE_WINDOW", self._acknowledge_active_alert)
        win.update_idletasks()
        win.minsize(max(360, win.winfo_reqwidth()), max(140, win.winfo_reqheight()))
        self._center_alert_window(win)
        win.lift()
        win.focus_force()
        if isinstance(self._active_alert, dict):
            self._active_alert["window"] = win

    def _center_alert_window(self, win):
        try:
            self._root.update_idletasks()
            root_x = self._root.winfo_rootx()
            root_y = self._root.winfo_rooty()
            root_w = self._root.winfo_width()
            root_h = self._root.winfo_height()
            win_w = win.winfo_width()
            win_h = win.winfo_height()
            x = root_x + max(0, (root_w - win_w) // 2)
            y = root_y + max(0, (root_h - win_h) // 2)
            win.geometry(f"+{x}+{y}")
        except Exception:
            pass

    def _acknowledge_active_alert(self):
        alert = self._active_alert if isinstance(self._active_alert, dict) else None
        if not alert:
            return
        event = alert.get("event")
        if event is not None and not event.is_set():
            event.set()
        self._close_active_alert_window()

    def _close_active_alert_window(self):
        alert = self._active_alert if isinstance(self._active_alert, dict) else None
        win = alert.get("window") if alert else None
        if win is not None:
            try:
                win.grab_release()
            except Exception:
                pass
            try:
                win.destroy()
            except Exception:
                pass
            alert["window"] = None

    def resume_active_alert(self, command: str = "") -> str:
        alert = self._active_alert if isinstance(self._active_alert, dict) else None
        if not alert:
            return "No alert step is waiting for confirmation."
        normalized = (command or "").strip().lower()
        if normalized not in ("continue", "resume", "continue queue", "resume queue", "proceed", "ok"):
            return (
                "Alert step is waiting. Mention me with `continue` or `resume` to continue the queue."
            )
        event = alert.get("event")
        if event is None or event.is_set():
            return "No alert step is waiting for confirmation."
        self.log(f"Queue alert resumed remotely: {alert.get('message', '')}")
        self._root.after(0, self._acknowledge_active_alert)
        return "Queue continued past the alert step."

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

    @staticmethod
    def _format_bo_block_details(block: dict) -> str:
        target = int(block.get("target_iterations", 1) or 1)
        channels = (block.get("channels_override") or "").strip() or "config channels"
        config_name = Path(str(block.get("bo_config_path") or "BO config")).name
        if str(block.get("objective") or "").lower() == "paired_response":
            batch = max(1, int(block.get("batch_size", 1) or 1))
            target_eq = max(0.0, float(block.get("target_equilibration_seconds", 0.0) or 0.0))
            buffer_eq = max(0.0, float(block.get("buffer_equilibration_seconds", 0.0) or 0.0))
            eq_text = f" | eq target {target_eq:g}s, buffer {buffer_eq:g}s" if (target_eq or buffer_eq) else ""
            return f"{config_name} | paired {target} cycles x {batch} methods{eq_text} | {channels}"
        return f"{config_name} | {target} iter | {channels}"

    def _execute_measurement_item(self, item: dict):
        self._ensure_mux_script_for_item(item)
        self._root.after(0, self._plotter.start_live, f"{item['type']} (live)", None, item["type"])
        csv_path = None
        success = False
        try:
            mux_channel = self._extract_mux_channel(item)
            meas_tag = self._session.next_meas_tag_with_mux(mux_channel)
            item["meas_tag"] = meas_tag
            self.log(f"[Tag] {meas_tag}")
            self._root.after(0, self.refresh_labels)
            data_folder = None
            if self._session.session_manager is not None:
                data_folder = self._session.session_manager.require_experiment()
                if data_folder is None:
                    return False, None
            runner = SerialMeasurementRunner(
                Path(item["script_path"]),
                log_callback=self.log,
                data_callback=self._plotter.push_live_point,
                data_folder=data_folder,
                save_raw_packets=self._session.save_raw_packets,
                simulate_measurements=self._session.simulate_measurements,
                invert_current=(item.get("type") == "SWV"),
                device_port=self._session.device_port,
            )
            self._session.current_runner = runner
            success, csv_path = runner.execute(meas_tag=meas_tag)
            return success, csv_path
        finally:
            self._session.current_runner = None
            self._root.after(0, self._plotter.stop_live)

    def _run_bo_analysis(self, bo_session: BOIntegrationSession, block: dict, suggestion=None, phase: str | None = None) -> Path:
        session_mgr = getattr(self._session, "session_manager", None)
        exp_path = session_mgr.require_experiment() if session_mgr is not None else None
        if exp_path is None:
            raise RuntimeError("An active experiment folder is required for BO analysis")
        output_dir = Path(str(block.get("analysis_output_dir") or "")) if str(block.get("analysis_output_dir") or "").strip() else Path(exp_path) / "bo_analysis"
        return bo_session.run_pending_analysis(
            folders=[exp_path],
            output_dir=output_dir,
            analysis=dict(block.get("analysis") or {}),
            suggestion=suggestion,
            phase=phase,
        )

    def _run_bo_render_break(self, bo_session: BOIntegrationSession, iteration: int | None) -> None:
        """Give the GUI an acquisition-free window to load BO results."""
        callback = getattr(self._session, "_bo_live_refresh_callback", None)
        if not callable(callback):
            return
        self._session.update_queue_status(
            state="render_break",
            current_label=f"BO iteration {iteration} render break",
            active_step_type="RENDER_BREAK",
            active_step_details="Loading Bayesian optimization results and records",
            active_step_started_at=datetime.now().isoformat(timespec="seconds"),
            active_step_estimated_seconds=None,
        )
        rendered = threading.Event()

        def refresh():
            try:
                callback({
                    "record_dir": str(bo_session.record_dir),
                    "session_id": bo_session.session_id,
                    "iteration": iteration,
                    "event": "observation_imported",
                })
            finally:
                rendered.set()

        self._root.after(0, refresh)
        # Rendering runs on Tk's thread. Waiting here prevents the next device
        # acquisition from overlapping that work; stop requests remain responsive.
        while self._session.is_running and not rendered.wait(timeout=0.05):
            pass

    def _run_bo_queue_items(self, queue_items: list, label: str, progress: dict | None = None):
        completed = 0
        failed = 0
        stopped = 0
        recorded_items = []
        progress = dict(progress or {})
        bo_parent_item = progress.get("bo_parent_item")
        phase_label = str(progress.get("bo_phase") or label)
        cycle_current = progress.get("bo_cycle_current")
        cycle_total = progress.get("bo_cycle_total")
        cycle_text = f"Cycle {cycle_current}/{cycle_total}" if cycle_current and cycle_total else "BO"
        total_items = len(queue_items)
        for idx, sub_item in enumerate(queue_items):
            if not self._session.is_running:
                stopped += 1
                sub_item["status"] = "stopped"
                recorded_items.append(dict(sub_item))
                break
            if bo_parent_item is not None:
                self._set_bo_live_details(
                    bo_parent_item,
                    f"{cycle_text} | {phase_label} | {idx + 1}/{total_items}: {sub_item.get('details', sub_item.get('type'))}",
                )
            status_update = {}
            if progress:
                completed_before = int(progress.get("completed_before", 0) or 0)
                status_update = {
                    "bo_mode": progress.get("bo_mode"),
                    "bo_cycle_current": progress.get("bo_cycle_current"),
                    "bo_cycle_total": progress.get("bo_cycle_total"),
                    "bo_phase": progress.get("bo_phase"),
                    "bo_completed_measurements": completed_before + idx,
                    "bo_total_measurements": progress.get("bo_total_measurements"),
                    "bo_observed_sets": progress.get("bo_observed_sets"),
                    "bo_total_sets": progress.get("bo_total_sets"),
                }
            self._session.update_queue_status(
                state="running",
                current_label=f"{label}: {sub_item.get('details', sub_item.get('type'))}",
                active_step_type=sub_item.get("type"),
                active_step_details=sub_item.get("details"),
                active_step_started_at=datetime.now().isoformat(timespec="seconds"),
                active_step_estimated_seconds=estimate_item_seconds(sub_item),
                **status_update,
            )
            ok, csv_path = self._execute_measurement_item(sub_item)
            if ok:
                completed += 1
                sub_item["status"] = "completed"
                sub_item["completed_at"] = datetime.now().isoformat(timespec="seconds")
                if csv_path:
                    sub_item["csv_path"] = str(csv_path)
                    self._root.after(0, self._plotter.plot_data, csv_path, self._session.last_live_plot_color, None, True, False)
            else:
                if not self._session.is_running:
                    stopped += 1
                    sub_item["status"] = "stopped"
                else:
                    failed += 1
                    sub_item["status"] = "failed"
                    sub_item["failed_at"] = datetime.now().isoformat(timespec="seconds")
                recorded_items.append(dict(sub_item))
                break
            recorded_items.append(dict(sub_item))
        if bo_parent_item is not None:
            summary_status = "completed" if failed == 0 and stopped == 0 else ("stopped" if stopped else "failed")
            self._append_bo_progress(
                bo_parent_item,
                "BO_MEASURE",
                summary_status,
                f"{cycle_text} | {phase_label}: completed {completed}/{total_items}",
            )
        return completed, failed, stopped, recorded_items

    def _execute_bo_operational_items(self, items: list, label: str, bo_parent_item: Optional[dict] = None) -> bool:
        total_items = len(items)
        for idx, sub_item in enumerate(items):
            if not self._session.is_running:
                return False
            t = str(sub_item.get("type") or "").upper()
            details = str(sub_item.get("details") or t)
            if bo_parent_item is not None:
                self._set_bo_live_details(bo_parent_item, f"{label}: step {idx + 1}/{total_items} | {details}")
            progress_record = self._append_bo_progress(
                bo_parent_item,
                t,
                "running",
                f"{label}: step {idx + 1}/{total_items} | {details}",
            )
            self._session.update_queue_status(
                state="running",
                current_label=f"{label}: {details}",
                active_step_type=t,
                active_step_details=details,
                active_step_started_at=datetime.now().isoformat(timespec="seconds"),
                active_step_estimated_seconds=estimate_item_seconds(sub_item),
            )
            if t == "PAUSE":
                ok = self._exec_pause(float(sub_item.get("pause_seconds", 0.0) or 0.0))
            elif t == "ALERT":
                ok = self._exec_alert(str(sub_item.get("alert_message") or "Fluid exchange checkpoint"))
            elif t.startswith("PUMP_"):
                ok = self._exec_pump(sub_item)
            else:
                raise RuntimeError(f"Fluid exchange block contains unsupported item type: {t}")
            if not ok:
                self._update_bo_progress(progress_record, "failed")
                return False
            self._update_bo_progress(progress_record, "completed")
        return True

    def _execute_bo_equilibration_pause(self, seconds: float, label: str, bo_parent_item: Optional[dict] = None) -> bool:
        seconds = max(0.0, float(seconds or 0.0))
        if seconds <= 0.0:
            return True
        self.log(f"{label}: equilibrating for {seconds:g} sec before measurement")
        if bo_parent_item is not None:
            self._set_bo_live_details(bo_parent_item, f"{label}: pause {seconds:g} sec")
        progress_record = self._append_bo_progress(bo_parent_item, "PAUSE", "running", f"{label}: pause {seconds:g} sec")
        self._session.update_queue_status(
            state="running",
            current_label=label,
            active_step_type="PAUSE",
            active_step_details=f"Equilibration pause for {seconds:g} sec",
            active_step_started_at=datetime.now().isoformat(timespec="seconds"),
            active_step_estimated_seconds=seconds,
        )
        ok = self._exec_pause(seconds)
        self._update_bo_progress(progress_record, "completed" if ok else "stopped")
        return ok

    @staticmethod
    def _load_bo_exchange_items(block: dict, key: str, label: str) -> list:
        path_text = str(block.get(key) or "").strip()
        if not path_text:
            return []
        path = Path(path_text)
        if not path.exists():
            raise FileNotFoundError(f"{label} block not found: {path}")
        with open(path, "r", encoding="utf-8-sig") as fh:
            payload = json.load(fh)
        items = payload.get("items") if isinstance(payload, dict) else payload
        if not isinstance(items, list):
            raise ValueError(f"{label} block must contain an items list")
        allowed = []
        for raw in items:
            item = dict(raw)
            t = str(item.get("type") or "").upper()
            if t.startswith("PUMP_") or t in ("PAUSE", "ALERT"):
                item.setdefault("status", "pending")
                allowed.append(item)
            else:
                raise ValueError(f"{label} block item is not a pump/pause/alert step: {t}")
        return allowed

    @staticmethod
    def _paired_bo_batch_span(
        completed_observations: int,
        target_observations: int,
        batch_size: int,
        warmup_observations: int,
    ) -> tuple[int, int]:
        """Return the next suggestion count and number of logical cycles it covers."""
        remaining = max(0, int(target_observations) - int(completed_observations))
        if remaining <= 0:
            return 0, 0
        batch_size = max(1, int(batch_size))
        warmup_remaining = max(0, int(warmup_observations) - int(completed_observations))
        if warmup_remaining > 0:
            count = min(remaining, warmup_remaining)
        else:
            count = min(remaining, batch_size)
        cycle_span = max(1, int(math.ceil(count / float(batch_size))))
        return count, cycle_span

    @staticmethod
    def _paired_bo_warmup_parameter_sets(config: dict) -> int:
        """Return the common per-group warmup prefix that can run without BO feedback."""
        groups = config.get("channel_groups") or []
        if groups:
            warmups = [
                max(0, int(group.get("n_initial_points", config.get("n_initial_points", 0)) or 0))
                for group in groups
            ]
            return min(warmups) if warmups else 0
        return max(0, int(config.get("n_initial_points", 0) or 0))

    @staticmethod
    def _paired_bo_execution_order(suggestions: list) -> list:
        return sorted(
            suggestions,
            key=lambda suggestion: (int(suggestion.iteration), int(suggestion.group_id)),
        )

    def _exec_bo_auto_loop(self, item: dict) -> bool:
        block = dict(item.get("bo_block") or {})
        config_path = str(block.get("bo_config_path") or "").strip()
        if not config_path:
            raise RuntimeError("BO block is missing its BO config path.")
        config = load_bo_config(config_path)
        if str(block.get("channels_override") or "").strip():
            config["channels"] = parse_channels(block.get("channels_override"))
        if str(block.get("objective") or "").strip():
            config["objective"] = str(block.get("objective")).strip()
        if isinstance(block.get("config_overrides"), dict):
            for key, value in dict(block.get("config_overrides") or {}).items():
                config[key] = value
        paired_mode = str(block.get("objective") or "").strip().lower() == "paired_response"
        if paired_mode:
            batch_size_for_warmup = max(1, int(block.get("batch_size", 1) or 1))
            if block.get("paired_warmup_cycles") is not None:
                warmup_cycles = max(0, int(block.get("paired_warmup_cycles") or 0))
                config["paired_warmup_cycles"] = warmup_cycles
                config["paired_batch_size"] = batch_size_for_warmup
                config["n_initial_points"] = warmup_cycles * batch_size_for_warmup
        analysis_cfg = dict((config.get("analysis") or {}))
        analysis_cfg.update(dict(block.get("analysis") or {}))
        if block.get("analysis_file_glob"):
            analysis_cfg["file_glob"] = str(block.get("analysis_file_glob"))
        config["analysis"] = analysis_cfg
        scoring_override = block.get("scoring")
        if isinstance(scoring_override, dict):
            scoring_cfg = dict(config.get("scoring") or {})
            for key, value in scoring_override.items():
                if isinstance(value, dict) and isinstance(scoring_cfg.get(key), dict):
                    merged = dict(scoring_cfg.get(key) or {})
                    merged.update(value)
                    scoring_cfg[key] = merged
                else:
                    scoring_cfg[key] = value
            config["scoring"] = scoring_cfg
        config = normalize_bo_config(config)
        errors = validate_bo_config(config)
        if errors:
            raise RuntimeError("; ".join(errors))

        session_mgr = getattr(self._session, "session_manager", None)
        exp_path = session_mgr.require_experiment() if session_mgr is not None else None
        if exp_path is None:
            return False

        target_iterations = int(block.get("target_iterations", 1) or 1)
        if target_iterations < 1:
            raise RuntimeError("BO block target iterations must be at least 1.")

        analysis_output_dir = str(block.get("analysis_output_dir") or (Path(exp_path) / "bo_analysis"))
        bo_session = BOIntegrationSession(config, exp_path, config_path=config_path, analysis_output_dir=analysis_output_dir)
        item["bo_session_id"] = bo_session.session_id
        item["bo_record_dir"] = str(bo_session.record_dir)
        item["bo_progress"] = []
        self.log(f"BO block started: {item.get('details', '')}")
        self._append_bo_progress(item, "BO_START", "running", f"BO session started: {bo_session.session_id}")
        halfway_iteration = max(1, int(math.ceil(target_iterations / 2.0)))
        halfway_notified = False

        if paired_mode:
            group_count = len(config.get("channel_groups") or [{"channels": config.get("channels", [])}])
            batch_size = max(1, int(block.get("batch_size", 1) or 1))
            target_cycles = target_iterations
            target_parameter_sets = target_cycles * batch_size
            target_observations = target_parameter_sets * group_count
            warmup_observations = min(
                target_parameter_sets,
                self._paired_bo_warmup_parameter_sets(config),
            )
            halfway_cycle = max(1, int(math.ceil(target_cycles / 2.0)))
            completed_cycles = 0
            target_equilibration_seconds = max(0.0, float(block.get("target_equilibration_seconds", 0.0) or 0.0))
            buffer_equilibration_seconds = max(0.0, float(block.get("buffer_equilibration_seconds", 0.0) or 0.0))
            target_exchange_items = self._load_bo_exchange_items(
                block,
                "target_exchange_block_path",
                "Target exchange",
            )
            buffer_exchange_items = self._load_bo_exchange_items(
                block,
                "buffer_exchange_block_path",
                "Return-to-buffer exchange",
            )
            while self._session.is_running and completed_cycles < target_cycles:
                suggestion_count, cycle_span = self._paired_bo_batch_span(
                    len(bo_session.observations) // group_count,
                    target_parameter_sets,
                    batch_size,
                    warmup_observations,
                )
                if suggestion_count <= 0:
                    break
                cycle_index = completed_cycles + 1
                cycle_end = min(target_cycles, completed_cycles + cycle_span)
                is_consolidated_warmup = (
                    len(bo_session.observations) // group_count < warmup_observations
                    and suggestion_count > batch_size
                )
                cycle_label = (
                    f"Warmup cycles {cycle_index}-{cycle_end}"
                    if is_consolidated_warmup
                    else f"Cycle {cycle_index}"
                )
                suggestions = self._paired_bo_execution_order(bo_session.ask_batch(suggestion_count))
                self._set_bo_live_details(
                    item,
                    f"{cycle_label}/{target_cycles} | preparing suggestions | iterations {suggestions[0].iteration}-{suggestions[-1].iteration}",
                )
                cycle_progress_record = self._append_bo_progress(
                    item,
                    "BO_WARMUP" if is_consolidated_warmup else "BO_CYCLE",
                    "running",
                    f"{cycle_label}/{target_cycles}: iterations {suggestions[0].iteration}-{suggestions[-1].iteration}",
                )
                self.log(
                    f"BO paired {cycle_label.lower()}/{target_cycles}: {len(suggestions)} suggestion(s), "
                    f"iterations {suggestions[0].iteration}-{suggestions[-1].iteration}"
                )

                buffer_items = []
                target_items = []
                suggestion_metadata = {}
                for suggestion in suggestions:
                    logical_cycle = ((int(suggestion.iteration) - 1) // batch_size) + 1
                    batch_index = ((int(suggestion.iteration) - 1) % batch_size) + 1
                    buffer = bo_session.build_queue_items(self._session.registry, suggestion, phase="buffer")
                    target = bo_session.build_queue_items(self._session.registry, suggestion, phase="target")
                    bo_session.record_queued(suggestion, buffer + target)
                    suggestion_metadata[int(suggestion.iteration)] = {
                        "paired_cycle": logical_cycle,
                        "paired_batch_index": batch_index,
                        "buffer_trace_number": len(buffer_items) + 1,
                        "target_trace_number": len(target_items) + 1,
                    }
                    buffer_items.extend(buffer)
                    target_items.extend(target)
                measurements_per_set = (
                    (len(buffer_items) + len(target_items)) // max(1, len(suggestions))
                )
                total_measurements = target_observations * measurements_per_set
                completed_measurements = len(bo_session.observations) * measurements_per_set
                live_refresh = getattr(self._session, "_bo_live_refresh_callback", None)
                if callable(live_refresh):
                    self._root.after(
                        0,
                        lambda cb=live_refresh, data={
                            "record_dir": str(bo_session.record_dir),
                            "session_id": bo_session.session_id,
                            "iteration": suggestions[0].iteration if suggestions else None,
                            "event": "batch_queued",
                        }: cb(dict(data)),
                    )

                b_completed, b_failed, b_stopped, b_recorded = self._run_bo_queue_items(
                    buffer_items,
                    "BO buffer batch",
                    progress={
                        "bo_mode": "paired_response",
                        "bo_parent_item": item,
                        "bo_cycle_current": cycle_index,
                        "bo_cycle_total": target_cycles,
                        "bo_phase": "buffer measurements",
                        "completed_before": completed_measurements,
                        "bo_total_measurements": total_measurements,
                        "bo_observed_sets": len(bo_session.observations),
                        "bo_total_sets": target_observations,
                    },
                )
                bo_session.record_queue_completion({
                    "start_index": None,
                    "total": len(b_recorded),
                    "completed": b_completed,
                    "failed": b_failed,
                    "stopped": b_stopped,
                    "items": b_recorded,
                })
                if b_failed or b_stopped or not self._session.is_running:
                    return False
                completed_measurements += len(buffer_items)

                if target_exchange_items:
                    self._session.update_queue_status(
                        bo_mode="paired_response",
                        bo_cycle_current=cycle_index,
                        bo_cycle_total=target_cycles,
                        bo_phase="buffer-to-target exchange",
                        bo_completed_measurements=completed_measurements,
                        bo_total_measurements=total_measurements,
                        bo_observed_sets=len(bo_session.observations),
                        bo_total_sets=target_observations,
                    )
                    self.log(f"BO paired batch: running buffer-to-target exchange block ({len(target_exchange_items)} step(s))")
                    if not self._execute_bo_operational_items(
                        [copy.deepcopy(x) for x in target_exchange_items],
                        f"{cycle_label}/{target_cycles} buffer-to-target exchange",
                        bo_parent_item=item,
                    ):
                        return False
                else:
                    self.log("BO paired batch: no buffer-to-target exchange block configured; continuing to target measurements")

                self._session.update_queue_status(
                    bo_mode="paired_response",
                    bo_cycle_current=cycle_index,
                    bo_cycle_total=target_cycles,
                    bo_phase="target equilibration",
                    bo_completed_measurements=completed_measurements,
                    bo_total_measurements=total_measurements,
                    bo_observed_sets=len(bo_session.observations),
                    bo_total_sets=target_observations,
                )
                if not self._execute_bo_equilibration_pause(
                    target_equilibration_seconds,
                    f"BO paired {cycle_label.lower()}/{target_cycles}: target equilibration",
                    bo_parent_item=item,
                ):
                    return False

                t_completed, t_failed, t_stopped, t_recorded = self._run_bo_queue_items(
                    target_items,
                    "BO target batch",
                    progress={
                        "bo_mode": "paired_response",
                        "bo_parent_item": item,
                        "bo_cycle_current": cycle_index,
                        "bo_cycle_total": target_cycles,
                        "bo_phase": "target measurements",
                        "completed_before": completed_measurements,
                        "bo_total_measurements": total_measurements,
                        "bo_observed_sets": len(bo_session.observations),
                        "bo_total_sets": target_observations,
                    },
                )
                bo_session.record_queue_completion({
                    "start_index": None,
                    "total": len(t_recorded),
                    "completed": t_completed,
                    "failed": t_failed,
                    "stopped": t_stopped,
                    "items": t_recorded,
                })
                if t_failed or t_stopped or not self._session.is_running:
                    return False
                completed_measurements += len(target_items)

                if buffer_exchange_items:
                    self._session.update_queue_status(
                        bo_mode="paired_response",
                        bo_cycle_current=cycle_index,
                        bo_cycle_total=target_cycles,
                        bo_phase="target-to-buffer exchange",
                        bo_completed_measurements=completed_measurements,
                        bo_total_measurements=total_measurements,
                        bo_observed_sets=len(bo_session.observations),
                        bo_total_sets=target_observations,
                    )
                    self.log(f"BO paired batch: running target-to-buffer exchange block ({len(buffer_exchange_items)} step(s))")
                    if not self._execute_bo_operational_items(
                        [copy.deepcopy(x) for x in buffer_exchange_items],
                        f"{cycle_label}/{target_cycles} target-to-buffer exchange",
                        bo_parent_item=item,
                    ):
                        return False
                else:
                    self.log("BO paired batch: no target-to-buffer exchange block configured; next cycle will start in current fluid")

                if cycle_end < target_cycles:
                    self._session.update_queue_status(
                        bo_mode="paired_response",
                        bo_cycle_current=cycle_index,
                        bo_cycle_total=target_cycles,
                        bo_phase="buffer equilibration",
                        bo_completed_measurements=completed_measurements,
                        bo_total_measurements=total_measurements,
                        bo_observed_sets=len(bo_session.observations),
                        bo_total_sets=target_observations,
                    )
                    if not self._execute_bo_equilibration_pause(
                        buffer_equilibration_seconds,
                        f"BO paired {cycle_label.lower()}/{target_cycles}: buffer equilibration",
                        bo_parent_item=item,
                    ):
                        return False

                self._set_bo_live_details(
                    item,
                    f"Cycle {cycle_index}/{target_cycles} | analysis | importing iterations {suggestions[0].iteration}-{suggestions[-1].iteration}",
                )
                self._session.update_queue_status(
                    bo_mode="paired_response",
                    bo_cycle_current=cycle_index,
                    bo_cycle_total=target_cycles,
                    bo_phase="analysis",
                    bo_completed_measurements=completed_measurements,
                    bo_total_measurements=total_measurements,
                    bo_observed_sets=len(bo_session.observations),
                    bo_total_sets=target_observations,
                )
                for suggestion in suggestions:
                    self._set_bo_live_details(
                        item,
                        f"Cycle {cycle_index}/{target_cycles} | analysis | importing iteration {suggestion.iteration}",
                    )
                    metadata = dict(suggestion_metadata.get(int(suggestion.iteration), {}))
                    metadata["target_trace_number"] = len(buffer_items) + int(metadata.get("target_trace_number", 0) or 0)
                    buffer_summary = self._run_bo_analysis(bo_session, block, suggestion=suggestion, phase="buffer")
                    target_summary = self._run_bo_analysis(bo_session, block, suggestion=suggestion, phase="target")
                    obs = bo_session.import_paired_analysis(
                        suggestion,
                        buffer_summary,
                        target_summary,
                        notes="Imported from recipe paired-response BO block",
                        **metadata,
                    )
                    self.log(
                        f"BO iteration {obs['iteration']} paired complete: "
                        f"Q_run={obs['Q_run']:.3f}, "
                        f"mean delta={float(obs['quality'].get('mean_delta_peak_height_uA', 0.0)):.4g} uA"
                    )
                    if callable(live_refresh):
                        self._run_bo_render_break(bo_session, obs.get("iteration"))
                completed_cycles = cycle_end
                self._update_bo_progress(
                    cycle_progress_record,
                    "completed",
                    f"{cycle_label}/{target_cycles}: iterations {suggestions[0].iteration}-{suggestions[-1].iteration} complete",
                )
                self._set_bo_live_details(
                    item,
                    f"Cycle {cycle_index}/{target_cycles} complete | {len(bo_session.observations)}/{target_observations} parameter sets observed",
                )
                self._session.update_queue_status(
                    bo_mode="paired_response",
                    bo_cycle_current=cycle_index,
                    bo_cycle_total=target_cycles,
                    bo_phase="cycle complete",
                    bo_completed_measurements=completed_measurements,
                    bo_total_measurements=total_measurements,
                    bo_observed_sets=len(bo_session.observations),
                    bo_total_sets=target_observations,
                )
                if session_mgr is not None and not halfway_notified and completed_cycles >= halfway_cycle:
                    halfway_notified = True
                    session_mgr.notify_slack(
                        f"BO paired-response progress: cycle {completed_cycles}/{target_cycles} complete "
                        f"(halfway). {len(bo_session.observations)}/{target_observations} hyperparameter sets observed. "
                        f"Session={bo_session.session_id}; Experiment={Path(exp_path).name}."
                    )

            completed_iterations = len(bo_session.observations)
            item["details"] = (
                f"{self._format_bo_block_details(block)} | done "
                f"{completed_cycles}/{target_cycles} cycles, {completed_iterations}/{target_observations} methods"
            )
            self._append_bo_progress(
                item,
                "BO_DONE",
                "completed" if completed_cycles >= target_cycles else "stopped",
                f"Paired BO done: {completed_cycles}/{target_cycles} cycles, {completed_iterations}/{target_observations} parameter sets",
            )
            self._root.after(0, self.refresh)
            best = bo_session.best_observation()
            if best is not None:
                item["bo_best_q"] = float(best.get("Q_run", 0.0))
                self.log(f"BO block best Q_run={item['bo_best_q']:.3f}")
            if session_mgr is not None and completed_cycles >= target_cycles:
                best_text = ""
                if best is not None:
                    best_text = f" Best Q_run={float(best.get('Q_run', 0.0)):.3f} at iter {int(best.get('iteration', 0) or 0)}."
                session_mgr.notify_slack(
                    f"BO paired-response session completed: {completed_cycles}/{target_cycles} cycles, "
                    f"{completed_iterations}/{target_observations} hyperparameter sets. "
                    f"Session={bo_session.session_id}; Experiment={Path(exp_path).name}.{best_text}"
                )
            return self._session.is_running and completed_cycles >= target_cycles

        classic_groups = config.get("channel_groups") or [
            {"id": 1, "name": "Group 1", "channels": config.get("channels", [])}
        ]
        def group_completed(group):
            group_id = int(group.get("id", 1))
            return sum(
                1 for obs in bo_session.observations
                if int(obs.get("group_id", 1)) == group_id
            )

        while self._session.is_running and any(
            group_completed(group) < target_iterations for group in classic_groups
        ):
            group = min(classic_groups, key=group_completed)
            suggestion = bo_session.ask_next_for_group(int(group.get("id", 1)))
            self.log(
                f"BO {suggestion.group_name} iteration {suggestion.iteration}: generating queue items"
            )
            queue_items = bo_session.build_queue_items(self._session.registry, suggestion)
            bo_session.record_queued(suggestion, queue_items)

            completed = 0
            failed = 0
            stopped = 0
            recorded_items = []
            for sub_item in queue_items:
                if not self._session.is_running:
                    stopped += 1
                    sub_item["status"] = "stopped"
                    recorded_items.append(dict(sub_item))
                    break
                self._session.update_queue_status(
                    state="running",
                    current_label=f"BO iter {suggestion.iteration}: {sub_item.get('details', sub_item.get('type'))}",
                    active_step_type=sub_item.get("type"),
                    active_step_details=sub_item.get("details"),
                    active_step_started_at=datetime.now().isoformat(timespec="seconds"),
                    active_step_estimated_seconds=estimate_item_seconds(sub_item),
                )
                ok, csv_path = self._execute_measurement_item(sub_item)
                if ok:
                    completed += 1
                    sub_item["status"] = "completed"
                    sub_item["completed_at"] = datetime.now().isoformat(timespec="seconds")
                    if csv_path:
                        sub_item["csv_path"] = str(csv_path)
                        self._root.after(0, self._plotter.plot_data, csv_path, self._session.last_live_plot_color, None, True, False)
                else:
                    if not self._session.is_running:
                        stopped += 1
                        sub_item["status"] = "stopped"
                    else:
                        failed += 1
                        sub_item["status"] = "failed"
                        sub_item["failed_at"] = datetime.now().isoformat(timespec="seconds")
                    recorded_items.append(dict(sub_item))
                    break
                recorded_items.append(dict(sub_item))

            queue_summary = {
                "start_index": None,
                "total": len(recorded_items),
                "completed": completed,
                "failed": failed,
                "stopped": stopped,
                "items": recorded_items,
            }
            bo_session.record_queue_completion(queue_summary)
            if failed or stopped or not self._session.is_running:
                if session_mgr is not None:
                    state = "stopped" if stopped or not self._session.is_running else "failed"
                    session_mgr.notify_slack(
                        f"BO session {state}: iter {suggestion.iteration}/{target_iterations}. "
                        f"Session={bo_session.session_id}; Experiment={Path(exp_path).name}."
                    )
                return False

            summary_path = self._run_bo_analysis(bo_session, block, suggestion=suggestion)
            obs = bo_session.import_analysis(
                summary_path,
                notes="Imported from recipe BO block",
                suggestion=suggestion,
            )
            self.log(
                f"BO {obs['group_name']} iteration {obs['iteration']} complete: "
                f"Q_run={obs['Q_run']:.3f}"
            )
            self._run_bo_render_break(bo_session, obs.get("iteration"))
            if session_mgr is not None and not halfway_notified and int(obs["iteration"]) >= halfway_iteration:
                halfway_notified = True
                session_mgr.notify_slack(
                    f"BO progress: iter {obs['iteration']}/{target_iterations} complete "
                    f"(halfway). Q_run={float(obs['Q_run']):.3f}. "
                    f"Session={bo_session.session_id}; Experiment={Path(exp_path).name}."
                )

        completed_iterations = len(bo_session.observations)
        expected_observations = target_iterations * len(classic_groups)
        item["details"] = (
            f"{self._format_bo_block_details(block)} | done "
            f"{completed_iterations}/{expected_observations} group iterations"
        )
        best = bo_session.best_observation()
        if best is not None:
            item["bo_best_q"] = float(best.get("Q_run", 0.0))
            self.log(f"BO block best Q_run={item['bo_best_q']:.3f}")
        if session_mgr is not None and completed_iterations >= expected_observations:
            best_text = ""
            if best is not None:
                best_text = f" Best Q_run={float(best.get('Q_run', 0.0)):.3f} at iter {int(best.get('iteration', 0) or 0)}."
            session_mgr.notify_slack(
                f"BO session completed: {completed_iterations}/{expected_observations} group iterations. "
                f"Session={bo_session.session_id}; Experiment={Path(exp_path).name}.{best_text}"
            )
        return self._session.is_running and completed_iterations >= expected_observations

    def _execute_queue(self, start_index: int = 0):
        queue = list(self._session.measurement_queue)
        for i, item in enumerate(queue[start_index:], start=start_index):
            if not self._session.is_running:
                self.log("Queue execution stopped by user.")
                break

            self._session.measurement_queue[i]["status"] = "running"
            self._root.after(0, self.refresh)
            self._root.after(0, self.set_status, f"Running: {item['type']} - {item.get('details', '')}")
            item_eta = estimate_item_seconds(item)
            if str(item.get("type", "")).strip().upper() == "ALERT":
                item_eta = None
            self._session.update_queue_status(
                state="running",
                current_index=(i - start_index + 1),
                total=len(queue) - start_index,
                current_label=(item.get("details") or item.get("type") or ""),
                queue_start_index=start_index,
                active_queue_index=i,
                next_queue_index=min(i + 1, len(queue)),
                active_step_started_at=datetime.now().isoformat(timespec="seconds"),
                active_step_estimated_seconds=item_eta,
                active_step_type=item.get("type"),
                active_step_details=(item.get("details") or item.get("type") or ""),
            )
            self.log(f"Queue start -> {item.get('details', item.get('type'))}")

            csv_path = None
            try:
                t = str(item.get("type") or "")
                if t == "PAUSE":
                    ok = self._exec_pause(float(item.get("pause_seconds", 0)))
                    self._session.measurement_queue[i]["status"] = "completed" if ok else "stopped"
                elif t == "ALERT":
                    alert_msg = item.get("alert_message", "Paused - click OK.")
                    session_mgr = getattr(self._session, "session_manager", None)
                    if session_mgr is not None:
                        session_mgr.notify_slack(f"Queue alert: {alert_msg}")
                    ok = self._exec_alert(alert_msg)
                    self._session.measurement_queue[i]["status"] = "completed" if ok else "stopped"
                elif t == "BO_AUTO_LOOP":
                    ok = self._exec_bo_auto_loop(item)
                    self._session.measurement_queue[i]["status"] = "completed" if ok else ("stopped" if not self._session.is_running else "failed")
                    stamp_key = "completed_at" if ok else "failed_at"
                    self._session.measurement_queue[i][stamp_key] = datetime.now().isoformat(timespec="seconds")
                elif t.startswith("PUMP_"):
                    ok = self._exec_pump(item)
                    self._session.measurement_queue[i]["status"] = "completed" if ok else "failed"
                    if not ok:
                        self.log(f"Queue item FAILED: {t} | {item.get('details', '')}")
                else:
                    ok, csv_path = self._execute_measurement_item(item)
                    if item.get("meas_tag"):
                        self._session.measurement_queue[i]["meas_tag"] = item.get("meas_tag")
                    if csv_path:
                        self._session.measurement_queue[i]["csv_path"] = str(csv_path)
                    if ok:
                        self._session.measurement_queue[i]["status"] = "completed"
                        self._session.measurement_queue[i]["completed_at"] = datetime.now().isoformat(timespec="seconds")
                    else:
                        self._session.measurement_queue[i]["status"] = "failed"
                        self._session.measurement_queue[i]["failed_at"] = datetime.now().isoformat(timespec="seconds")
                        self.log(f"Queue item FAILED: {item['type']} | {item.get('details', item.get('meas_tag', ''))}")
            except Exception as exc:
                self._session.measurement_queue[i]["status"] = "failed"
                self.log(f"CRITICAL ERROR in queue: {exc}")

            if csv_path:
                self._root.after(0, self._plotter.plot_data, csv_path, self._session.last_live_plot_color, None, True, False)
            self._root.after(0, self.refresh)
            step_delay = getattr(self._session, "step_delay", 0.0) or 0.0
            if step_delay > 0 and i < len(queue) - 1:
                next_step_number = i - start_index + 2
                delay_label = f"Inter-step delay before step {next_step_number}"
                self._session.update_queue_status(
                    state="step_delay",
                    current_index=(i - start_index + 1),
                    total=len(queue) - start_index,
                    current_label=delay_label,
                    queue_start_index=start_index,
                    active_queue_index=i,
                    next_queue_index=i + 1,
                    active_step_started_at=datetime.now().isoformat(timespec="seconds"),
                    active_step_estimated_seconds=step_delay,
                    active_step_type="STEP_DELAY",
                    active_step_details=delay_label,
                )
                if not self._exec_pause(step_delay):
                    break

        self._session.is_running = False
        self._session.update_queue_status(
            state="idle",
            current_label="Queue Complete",
            active_queue_index=None,
            next_queue_index=None,
            active_step_started_at=None,
            active_step_estimated_seconds=None,
            active_step_type=None,
            active_step_details=None,
            bo_mode=None,
            bo_cycle_current=None,
            bo_cycle_total=None,
            bo_phase=None,
            bo_completed_measurements=None,
            bo_total_measurements=None,
            bo_observed_sets=None,
            bo_total_sets=None,
        )
        self.log("Queue completed.")
        self._root.after(0, self.set_status, "Queue Complete")
        self._announce_queue_end(start_index=start_index)
        self._notify_completion_callbacks(start_index=start_index)
