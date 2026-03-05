"""
gui/tab_recipe_maker.py — Recipe Maker tab.

Provides a lightweight recipe builder that mirrors the Queue tab layout:
  - Recipe list (Treeview) with add/remove/reorder controls
  - Pump step editor (speed/volume/port/pause)
  - Method library browser (from methods/library_map.json) with search/filter

This tab does not execute; it only composes recipe items for later use.
"""

import copy
import json
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk

from methods import library_map
from config import BLOCKS_DIR


class RecipeMakerTab:
    """Manages the 'Recipe Maker' notebook tab."""

    def __init__(self, parent_frame):
        self._frame = parent_frame
        self._recipe: list = []
        self._clipboard: list = []
        self._method_entries: dict = {}
        self._method_iid_to_key: dict = {}
        self._style = ttk.Style(self._frame)
        self._repo_root = Path(__file__).resolve().parents[1]
        self._recipe_root = self._repo_root / "recipe_maker"
        self._blocks_dir = (self._repo_root / BLOCKS_DIR).resolve()
        self._recipe_root.mkdir(parents=True, exist_ok=True)
        self._blocks_dir.mkdir(parents=True, exist_ok=True)
        self._build()

    # ── Build ──────────────────────────────────────────────────────────────

    def _build(self):
        pane = ttk.PanedWindow(self._frame, orient=tk.VERTICAL)
        pane.pack(fill="both", expand=True)

        top = ttk.Frame(pane); pane.add(top, weight=2)
        bottom = ttk.Frame(pane); pane.add(bottom, weight=1)

        # ── Control bar
        ctrl = ttk.Frame(top)
        ctrl.pack(pady=8, fill="x", padx=10)

        ttk.Button(ctrl, text="Add Pump Step",
                   command=self._add_pump_step).pack(side="left", padx=4)
        ttk.Button(ctrl, text="Add Method Step",
                   command=self._add_method_step).pack(side="left", padx=4)
        ttk.Separator(ctrl, orient="vertical").pack(side="left", fill="y", padx=6)
        ttk.Button(ctrl, text="Move Up",
                   command=lambda: self._move_selected(-1)).pack(side="left", padx=4)
        ttk.Button(ctrl, text="Move Down",
                   command=lambda: self._move_selected(1)).pack(side="left", padx=4)
        ttk.Separator(ctrl, orient="vertical").pack(side="left", fill="y", padx=6)
        ttk.Button(ctrl, text="Copy",
                   command=self._copy_selected).pack(side="left", padx=4)
        ttk.Button(ctrl, text="Paste",
                   command=self._paste_after_selected).pack(side="left", padx=4)
        ttk.Button(ctrl, text="Duplicate",
                   command=self._duplicate_selected).pack(side="left", padx=4)
        ttk.Button(ctrl, text="Delete",
                   command=self._delete_selected).pack(side="left", padx=4)
        ttk.Separator(ctrl, orient="vertical").pack(side="left", fill="y", padx=6)
        ttk.Button(ctrl, text="Save",
                   command=self._save_recipe).pack(side="left", padx=4)
        ttk.Button(ctrl, text="Load",
                   command=self._load_recipe).pack(side="left", padx=4)
        ttk.Button(ctrl, text="Clear",
                   command=self._clear_recipe).pack(side="left", padx=4)


        # ── Recipe Treeview
        cols = ("Type", "Block", "Details")
        self._style.configure("Recipe.Treeview", background="white", fieldbackground="white")
        self._style.map("Recipe.Treeview", background=[("selected", "#cce4ff")])
        self._tree = ttk.Treeview(
            top, columns=cols, show="tree headings", height=10, style="Recipe.Treeview"
        )
        self._tree.heading("#0", text="#")
        self._tree.heading("Type", text="Type")
        self._tree.heading("Block", text="Block")
        self._tree.heading("Details", text="Details")
        self._tree.column("#0", width=50)
        self._tree.column("Type", width=160)
        self._tree.column("Block", width=180)
        self._tree.column("Details", width=420)
        self._tree.pack(fill="both", expand=True, padx=10, pady=5)
        self._tree.tag_configure("volt", background="#dff5d8")
        self._tree.tag_configure("block", background="#fff3cd")
        self._tree.tag_configure("alert", background="#f8d7da")
        self._tree.tag_configure("default", background="#f2f2f2")

        legend = ttk.Frame(top)
        legend.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Label(legend, text="Legend:").pack(side="left")
        self._legend_chip(legend, "#dff5d8", "Voltammetry (CV/SWV)")
        self._legend_chip(legend, "#fff3cd", "Block step")
        self._legend_chip(legend, "#f8d7da", "Alert/Pause")
        self._legend_chip(legend, "#f2f2f2", "Other")

        self._ctx = tk.Menu(self._tree, tearoff=0)
        self._ctx.add_command(label="Copy", command=self._copy_selected)
        self._ctx.add_command(label="Paste After", command=self._paste_after_selected)
        self._ctx.add_command(label="Duplicate", command=self._duplicate_selected)
        self._ctx.add_separator()
        self._ctx.add_command(label="Delete", command=self._delete_selected)
        self._tree.bind("<Button-3>", self._show_ctx)
        self._tree.bind("<Control-c>", lambda e: self._copy_selected())
        self._tree.bind("<Control-v>", lambda e: self._paste_after_selected())
        self._tree.bind("<Control-d>", lambda e: self._duplicate_selected())

        # ── Bottom pane: editors / library
        bottom_nb = ttk.Notebook(bottom)
        bottom_nb.pack(fill="both", expand=True, padx=10, pady=8)

        pump_tab = ttk.Frame(bottom_nb)
        method_tab = ttk.Frame(bottom_nb)
        block_tab = ttk.Frame(bottom_nb)
        bottom_nb.add(pump_tab, text="Pump Steps")
        bottom_nb.add(method_tab, text="Method Library")
        bottom_nb.add(block_tab, text="Blocks")

        self._build_pump_editor(pump_tab)
        self._build_method_library(method_tab)
        self._build_blocks_library(block_tab)

    def _legend_chip(self, parent, color: str, text: str):
        swatch = tk.Canvas(parent, width=12, height=12, highlightthickness=0)
        swatch.create_rectangle(0, 0, 12, 12, fill=color, outline="#777")
        swatch.pack(side="left", padx=(8, 2))
        ttk.Label(parent, text=text).pack(side="left", padx=(0, 6))

    # ── Pump editor ────────────────────────────────────────────────────────

    def _build_pump_editor(self, parent):
        pad = {"padx": 6, "pady": 4}

        ttk.Label(parent, text="Pump action:").grid(row=0, column=0, **pad, sticky="e")
        self._pump_action = tk.StringVar(value="INIT")
        ttk.Combobox(
            parent,
            textvariable=self._pump_action,
            values=["INIT", "SET_SPEED", "VALVE", "ASPIRATE", "DISPENSE", "PAUSE", "ALERT"],
            width=16,
            state="readonly",
        ).grid(row=0, column=1, **pad, sticky="w")

        ttk.Label(parent, text="Speed:").grid(row=0, column=2, **pad, sticky="e")
        self._pump_speed = tk.IntVar(value=20)
        ttk.Entry(parent, width=8, textvariable=self._pump_speed).grid(row=0, column=3, **pad, sticky="w")

        ttk.Label(parent, text="Volume (uL):").grid(row=0, column=4, **pad, sticky="e")
        self._pump_volume = tk.DoubleVar(value=100.0)
        ttk.Entry(parent, width=10, textvariable=self._pump_volume).grid(row=0, column=5, **pad, sticky="w")

        ttk.Label(parent, text="Valve port:").grid(row=0, column=6, **pad, sticky="e")
        self._pump_port = tk.IntVar(value=1)
        ttk.Entry(parent, width=8, textvariable=self._pump_port).grid(row=0, column=7, **pad, sticky="w")

        ttk.Label(parent, text="Pause (sec):").grid(row=0, column=8, **pad, sticky="e")
        self._pause_seconds = tk.DoubleVar(value=10.0)
        ttk.Entry(parent, width=10, textvariable=self._pause_seconds).grid(row=0, column=9, **pad, sticky="w")

        ttk.Label(parent, text="Alert message:").grid(row=1, column=0, **pad, sticky="e")
        self._alert_message = tk.StringVar(value="Check setup")
        ttk.Entry(parent, width=50, textvariable=self._alert_message).grid(
            row=1, column=1, columnspan=6, **pad, sticky="w"
        )

        ttk.Label(
            parent,
            text="Tip: Only relevant fields are used based on action type.",
            foreground="#666",
        ).grid(row=2, column=0, columnspan=10, padx=6, pady=(0, 6), sticky="w")

    def _add_pump_step(self):
        action = self._pump_action.get().strip().upper()
        if not action:
            return

        if action == "PAUSE":
            seconds = float(self._pause_seconds.get())
            item = {
                "type": "PAUSE",
                "status": "pending",
                "details": f"Pause for {seconds:.1f} sec",
                "pause_seconds": seconds,
            }
            self._recipe.append(item)
            self._refresh()
            return
        if action == "ALERT":
            msg = (self._alert_message.get() or "").strip()
            if not msg:
                messagebox.showerror("Invalid alert", "Alert message cannot be empty.")
                return
            item = {
                "type": "ALERT",
                "status": "pending",
                "details": f"Alert: {msg}",
                "alert_message": msg,
            }
            self._recipe.append(item)
            self._refresh()
            return

        details = ""
        pump_action = {"name": action, "params": {}}
        if action == "INIT":
            details = "Pump: Initialize (ZR)"
        elif action == "SET_SPEED":
            speed = int(self._pump_speed.get())
            details = f"Pump: Set speed S{speed}R"
            pump_action["params"]["speed"] = speed
        elif action == "VALVE":
            port = int(self._pump_port.get())
            details = f"Pump: Valve -> {port}"
            pump_action["params"]["port"] = port
        elif action == "ASPIRATE":
            speed = int(self._pump_speed.get())
            volume = float(self._pump_volume.get())
            details = f"Pump: Aspirate {volume:.2f} uL @ S{speed}R"
            pump_action["params"].update({"speed": speed, "volume": volume})
        elif action == "DISPENSE":
            speed = int(self._pump_speed.get())
            volume = float(self._pump_volume.get())
            details = f"Pump: Dispense {volume:.2f} uL @ S{speed}R"
            pump_action["params"].update({"speed": speed, "volume": volume})
        else:
            messagebox.showerror("Invalid action", f"Unsupported pump action: {action}")
            return

        item = {
            "type": f"PUMP_{action}",
            "status": "pending",
            "details": details,
            "pump_action": pump_action,
        }
        self._recipe.append(item)
        self._refresh()

    # ── Method library ─────────────────────────────────────────────────────

    def _build_method_library(self, parent):
        pad = {"padx": 6, "pady": 4}

        top = ttk.Frame(parent)
        top.pack(fill="x", padx=6, pady=6)

        ttk.Label(top, text="Search:").pack(side="left")
        self._method_search = tk.StringVar()
        self._method_search.trace_add("write", lambda *_: self._refresh_methods())
        ttk.Entry(top, textvariable=self._method_search, width=30).pack(side="left", padx=6)

        ttk.Label(top, text="Technique:").pack(side="left", padx=(10, 0))
        self._tech_filter = tk.StringVar(value="ALL")
        ttk.Combobox(
            top,
            textvariable=self._tech_filter,
            values=["ALL", "CV", "SWV"],
            state="readonly",
            width=8,
        ).pack(side="left", padx=6)
        self._tech_filter.trace_add("write", lambda *_: self._refresh_methods())

        ttk.Button(top, text="Refresh",
                   command=self._load_method_map).pack(side="left", padx=6)

        sweep = ttk.Frame(parent)
        sweep.pack(fill="x", padx=6, pady=(0, 4))

        ttk.Label(sweep, text="Sweep Start:").grid(row=0, column=0, **pad, sticky="e")
        self._sweep_start = tk.IntVar(value=1)
        ttk.Entry(sweep, width=6, textvariable=self._sweep_start).grid(row=0, column=1, **pad, sticky="w")

        ttk.Label(sweep, text="End:").grid(row=0, column=2, **pad, sticky="e")
        self._sweep_end = tk.IntVar(value=16)
        ttk.Entry(sweep, width=6, textvariable=self._sweep_end).grid(row=0, column=3, **pad, sticky="w")

        ttk.Label(sweep, text="Step:").grid(row=0, column=4, **pad, sticky="e")
        self._sweep_step = tk.IntVar(value=1)
        ttk.Entry(sweep, width=6, textvariable=self._sweep_step).grid(row=0, column=5, **pad, sticky="w")

        self._sweep_reverse = tk.BooleanVar(value=False)
        ttk.Checkbutton(sweep, text="Reverse", variable=self._sweep_reverse).grid(
            row=0, column=6, **pad, sticky="w"
        )

        ttk.Label(sweep, text="Custom order:").grid(row=1, column=0, **pad, sticky="e")
        self._sweep_custom = tk.StringVar(value="")
        ttk.Entry(sweep, width=44, textvariable=self._sweep_custom).grid(
            row=1, column=1, columnspan=5, **pad, sticky="we"
        )
        ttk.Label(sweep, text="e.g. 1,3,5,2,4").grid(row=1, column=6, **pad, sticky="w")

        ttk.Button(
            sweep,
            text="Add Channel Sweep Block",
            command=self._add_method_sweep_block,
        ).grid(row=0, column=7, rowspan=2, padx=(12, 6), pady=4, sticky="ns")

        cols = ("Hash", "Technique", "Params")
        self._method_tree = ttk.Treeview(parent, columns=cols, show="headings", height=8)
        self._method_tree.heading("Hash", text="Hash")
        self._method_tree.heading("Technique", text="Technique")
        self._method_tree.heading("Params", text="Params")
        self._method_tree.column("Hash", width=140)
        self._method_tree.column("Technique", width=100)
        self._method_tree.column("Params", width=520)
        self._method_tree.pack(fill="both", expand=True, padx=6, pady=6)

        self._load_method_map()

        hint = ttk.Label(
            parent,
            text=(
                "Select a method and use Add Method Step, or configure channels and use "
                "'Add Channel Sweep Block'."
            ),
            foreground="#666",
        )
        hint.pack(side="bottom", anchor="w", padx=8, pady=(0, 6))

    def _load_method_map(self):
        self._method_entries = library_map.all_entries()
        self._refresh_methods()

    def _refresh_methods(self):
        for row in self._method_tree.get_children():
            self._method_tree.delete(row)
        self._method_iid_to_key.clear()

        search = (self._method_search.get() or "").strip().lower()
        tech = (self._tech_filter.get() or "ALL").upper()

        for key, entry in sorted(self._method_entries.items()):
            technique = entry.get("technique", "")
            if tech != "ALL" and technique.upper() != tech:
                continue

            params = entry.get("params", {})
            params_str = ", ".join(f"{k}={v}" for k, v in params.items())
            hay = f"{key} {technique} {params_str}".lower()
            if search and search not in hay:
                continue

            iid = self._method_tree.insert(
                "", "end",
                values=(key, technique, params_str),
            )
            self._method_iid_to_key[iid] = key

    def _selected_method_entry(self):
        sel = self._method_tree.selection()
        if not sel:
            return None
        key = self._method_iid_to_key.get(sel[0])
        if not key:
            return None
        entry = self._method_entries.get(key)
        if not entry:
            return None
        return key, entry

    def _add_method_step(self):
        selected = self._selected_method_entry()
        if not selected:
            messagebox.showwarning("No selection", "Select a method from the library list.")
            return
        key, entry = selected
        technique = entry.get("technique", "")
        params = entry.get("params", {})
        mux_channel = entry.get("mux_channel")

        details = f"{key}.ms"
        item = {
            "type": technique,
            "status": "pending",
            "details": details,
            "method_ref": {
                "hash_key": key,
                "technique": technique,
                "params": params,
                "mux_channel": mux_channel,
            },
        }
        self._recipe.append(item)
        self._refresh()

    def _parse_sweep_channels(self):
        custom = (self._sweep_custom.get() or "").strip()
        if custom:
            tokens = custom.replace(";", ",").split(",")
            channels = []
            for tok in tokens:
                t = tok.strip()
                if not t:
                    continue
                try:
                    ch = int(t)
                except ValueError:
                    raise ValueError(f"Invalid channel in custom order: '{t}'")
                if ch < 1 or ch > 16:
                    raise ValueError("Channel numbers must be between 1 and 16.")
                channels.append(ch)
            if not channels:
                raise ValueError("Custom order is empty.")
            return channels

        start = int(self._sweep_start.get())
        end = int(self._sweep_end.get())
        step = abs(int(self._sweep_step.get()))
        if step == 0:
            raise ValueError("Step must be >= 1.")
        if start < 1 or start > 16 or end < 1 or end > 16:
            raise ValueError("Sweep start/end must be between 1 and 16.")

        if start <= end:
            channels = list(range(start, end + 1, step))
        else:
            channels = list(range(start, end - 1, -step))
        if self._sweep_reverse.get():
            channels.reverse()
        if not channels:
            raise ValueError("Sweep channel list is empty.")
        return channels

    def _add_method_sweep_block(self):
        selected = self._selected_method_entry()
        if not selected:
            messagebox.showwarning("No selection", "Select a method from the library list.")
            return
        try:
            channels = self._parse_sweep_channels()
        except Exception as exc:
            messagebox.showerror("Invalid sweep channels", str(exc))
            return

        key, entry = selected
        technique = entry.get("technique", "")
        params = copy.deepcopy(entry.get("params", {}))
        block_name = f"Sweep {technique} ({len(channels)} ch)"
        for ch in channels:
            item = {
                "type": technique,
                "status": "pending",
                "details": f"{key}.ms | MUX ch {ch}",
                "block_name": block_name,
                "method_ref": {
                    "hash_key": key,
                    "technique": technique,
                    "params": copy.deepcopy(params),
                    "mux_channel": ch,
                },
            }
            self._recipe.append(item)
        self._refresh()

    # ── Recipe list ops ────────────────────────────────────────────────────

    def _row_tag_for_item(self, item: dict) -> str:
        item_type = (item.get("type") or "").upper()
        if item_type in ("CV", "SWV"):
            return "volt"
        if item_type in ("PAUSE", "ALERT"):
            return "alert"
        if item.get("block_name") or item.get("block_ref"):
            return "block"
        return "default"

    def _refresh(self):
        for row in self._tree.get_children():
            self._tree.delete(row)
        for i, item in enumerate(self._recipe):
            tag = self._row_tag_for_item(item)
            self._tree.insert(
                "", "end", iid=str(i), text=str(i + 1),
                values=(
                    item.get("type", ""),
                    item.get("block_name", ""),
                    item.get("details", ""),
                ),
                tags=(tag,),
            )

    def _selected_indices(self):
        return sorted(
            self._tree.index(iid) for iid in self._tree.selection() if iid
        )

    def _copy_selected(self):
        indices = self._selected_indices()
        if not indices:
            return
        self._clipboard = [copy.deepcopy(self._recipe[i]) for i in indices]

    def _paste_after_selected(self):
        if not self._clipboard:
            return
        indices = self._selected_indices()
        insert_at = (max(indices) + 1) if indices else len(self._recipe)
        for item in self._clipboard:
            self._recipe.insert(insert_at, copy.deepcopy(item))
            insert_at += 1
        self._refresh()

    def _duplicate_selected(self):
        indices = self._selected_indices()
        if not indices:
            return
        insert_at = max(indices) + 1
        for i in indices:
            self._recipe.insert(insert_at, copy.deepcopy(self._recipe[i]))
            insert_at += 1
        self._refresh()

    def _show_ctx(self, event):
        self._ctx.tk_popup(event.x_root, event.y_root)

    def _delete_selected(self):
        indices = self._selected_indices()
        if not indices:
            return
        for i in reversed(indices):
            self._recipe.pop(i)
        self._refresh()

    def _move_selected(self, delta: int):
        indices = self._selected_indices()
        if not indices:
            return
        if delta < 0:
            for i in indices:
                if i <= 0:
                    continue
                self._recipe[i - 1], self._recipe[i] = self._recipe[i], self._recipe[i - 1]
        else:
            for i in reversed(indices):
                if i >= len(self._recipe) - 1:
                    continue
                self._recipe[i + 1], self._recipe[i] = self._recipe[i], self._recipe[i + 1]
        self._refresh()

    def _clear_recipe(self):
        if not self._recipe:
            return
        if not messagebox.askyesno("Clear recipe", "Remove all recipe items?"):
            return
        self._recipe.clear()
        self._refresh()

    # ── Save / load ────────────────────────────────────────────────────────

    def _save_recipe(self):
        if not self._recipe:
            messagebox.showwarning("Empty", "No recipe items to save.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            initialdir=str(self._recipe_root),
        )
        if not path:
            return
        payload = {"items": self._recipe}
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)

    def _load_recipe(self):
        path = filedialog.askopenfilename(
            filetypes=[("JSON", "*.json")],
            initialdir=str(self._recipe_root),
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                payload = json.load(f)
            items = payload.get("items", [])
            if not isinstance(items, list):
                raise ValueError("Invalid recipe format: items is not a list.")
            self._recipe = items
            self._refresh()
        except Exception as exc:
            messagebox.showerror("Load failed", str(exc))

    # â”€â”€ Blocks â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    def _build_blocks_library(self, parent):
        top = ttk.Frame(parent)
        top.pack(fill="x", padx=6, pady=6)
        ttk.Button(top, text="Refresh Blocks",
                   command=self._load_blocks).pack(side="left", padx=4)
        ttk.Button(top, text="Add Block",
                   command=self._add_selected_block).pack(side="left", padx=4)

        cols = ("Block", "Items")
        self._block_tree = ttk.Treeview(parent, columns=cols, show="headings", height=8)
        self._block_tree.heading("Block", text="Block")
        self._block_tree.heading("Items", text="Items")
        self._block_tree.column("Block", width=200)
        self._block_tree.column("Items", width=560)
        self._block_tree.pack(fill="both", expand=True, padx=6, pady=6)

        self._blocks: dict = {}
        self._block_iid_to_name: dict = {}
        self._load_blocks()

        hint = ttk.Label(
            parent,
            text="Blocks are predefined sequences stored in recipe_maker/default_blocks/.",
            foreground="#666",
        )
        hint.pack(side="bottom", anchor="w", padx=8, pady=(0, 6))

    def _load_blocks(self):
        self._blocks.clear()
        self._block_iid_to_name.clear()
        for row in self._block_tree.get_children():
            self._block_tree.delete(row)

        blocks_dir = self._blocks_dir
        if not blocks_dir.exists():
            return

        files = list(blocks_dir.glob("*.json")) + list(blocks_dir.glob("*.JSON"))
        seen = set()
        for path in sorted(files):
            norm = path.resolve().as_posix().lower()
            if norm in seen:
                continue
            seen.add(norm)
            try:
                payload = json.loads(path.read_text(encoding="utf-8-sig"))
            except Exception as exc:
                self._block_tree.insert(
                    "", "end",
                    values=(f"Invalid JSON: {path.name}", str(exc)),
                )
                continue

            items = payload.get("items", [])
            if not isinstance(items, list):
                self._block_tree.insert(
                    "", "end",
                    values=(f"Invalid block: {path.name}", "Missing items[]"),
                )
                continue

            name = payload.get("name") or path.stem
            self._blocks[name] = items
            summary = ", ".join(item.get("type", "") for item in items[:5])
            if len(items) > 5:
                summary += f" (+{len(items) - 5})"
            iid = self._block_tree.insert("", "end", values=(name, summary))
            self._block_iid_to_name[iid] = name

        if not self._blocks:
            self._block_tree.insert(
                "", "end",
                values=("No blocks found", str(blocks_dir)),
            )

    def _add_selected_block(self):
        sel = self._block_tree.selection()
        if not sel:
            messagebox.showwarning("No selection", "Select a block to add.")
            return
        name = self._block_iid_to_name.get(sel[0])
        if not name:
            return
        items = self._blocks.get(name, [])
        if not items:
            return
        for item in items:
            cloned = dict(item)
            cloned.setdefault("status", "pending")
            cloned["block_name"] = name
            self._recipe.append(cloned)
        self._refresh()
