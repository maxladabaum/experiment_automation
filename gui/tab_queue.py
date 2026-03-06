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

from core.runner import SerialMeasurementRunner
from methods import library_map
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
        self._last_selected    = None
        self._last_queue_path  = None

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
        self.log(f"Queue add: {item.get('details', item.get('type'))}")

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
        elif t and t.startswith("PUMP_"):
            action = item.get("pump_action") or {}
            data["pump_action"] = {"name": action.get("name"),
                                   "params": dict(action.get("params") or {})}
        else:
            if "script_path" in item:
                data["script_path"] = item["script_path"]
            if "method_ref" in item:
                data["method_ref"] = dict(item.get("method_ref") or {})
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
                method_ref = raw.get("method_ref") or {}
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
        self.clear_log()
        self.log("Queue start requested.")
        self.log(f"Measurement simulation: {'ON' if self._session.simulate_measurements else 'OFF'}")
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
        self._session.is_running = True
        self.clear_log()
        self.log("Queue start from selected requested.")
        self.log(f"Measurement simulation: {'ON' if self._session.simulate_measurements else 'OFF'}")
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
                    ok = self._exec_alert(item.get("alert_message", "Paused — click OK."))
                    self._session.measurement_queue[i]["status"] = "completed" if ok else "stopped"
                    success = ok

                elif t.startswith("PUMP_"):
                    ok = self._exec_pump(item)
                    self._session.measurement_queue[i]["status"] = "completed" if ok else "failed"
                    success = ok

                else:
                    self._ensure_mux_script_for_item(item)
                    self._root.after(0, self._plotter.start_live,
                                     f"{item['type']} (live)", None, item["type"])
                    try:
                        mux_channel = self._extract_mux_channel(item)
                        meas_tag = self._session.next_meas_tag_with_mux(mux_channel)
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
            step_delay = getattr(self._session, "step_delay", 0.0) or 0.0
            if step_delay > 0 and i < len(queue) - 1:
                if not self._exec_pause(step_delay):
                    break

        self._session.is_running = False
        self.log("Queue completed.")
        self._root.after(0, self.set_status, "Queue Complete")

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
