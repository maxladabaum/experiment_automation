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
from tkinter import ttk, simpledialog

from methods import library_map
from config import (
    BLOCKS_DIR,
    BO_ANALYSIS_FILE_GLOB,
    BO_ANALYSIS_OUTPUT_DIR,
    BO_DEFAULT_CONFIG_PATH,
)
from core.bo_session import load_bo_config, normalize_bo_config


class RecipeMakerTab:
    """Manages the 'Recipe Maker' notebook tab."""

    def __init__(self, parent_frame, on_send_to_queue=None):
        self._frame = parent_frame
        self._on_send_to_queue = on_send_to_queue
        self._recipe: list = []
        self._clipboard: list = []
        self._method_entries: dict = {}
        self._method_iid_to_key: dict = {}
        self._last_selected = None
        self._style = ttk.Style(self._frame)
        self._repo_root = Path(__file__).resolve().parents[1]
        self._recipe_root = self._repo_root / "recipe_maker"
        self._default_blocks_dir = (self._repo_root / BLOCKS_DIR).resolve()
        self._custom_blocks_dir = (self._repo_root / "recipe_maker" / "custom_blocks").resolve()
        self._saved_blocks_dir = (self._repo_root / "recipe_maker" / "saved_recipes").resolve()
        self._recipe_root.mkdir(parents=True, exist_ok=True)
        self._default_blocks_dir.mkdir(parents=True, exist_ok=True)
        self._custom_blocks_dir.mkdir(parents=True, exist_ok=True)
        self._saved_blocks_dir.mkdir(parents=True, exist_ok=True)
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
        ttk.Separator(ctrl, orient="vertical").pack(side="left", fill="y", padx=6)
        ttk.Button(ctrl, text="Send to Queue",
                   command=self._send_to_queue).pack(side="left", padx=4)


        # ── Recipe Treeview
        cols = ("Type", "Block", "Details")
        self._style.configure("Recipe.Treeview", background="white", fieldbackground="white")
        self._style.map("Recipe.Treeview", background=[("selected", "#cce4ff")])
        tree_frame = ttk.Frame(top)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self._tree = ttk.Treeview(
            tree_frame,
            columns=cols,
            show="tree headings",
            height=10,
            style="Recipe.Treeview",
            selectmode="extended",
        )
        self._tree.heading("#0", text="#")
        self._tree.heading("Type", text="Type")
        self._tree.heading("Block", text="Block")
        self._tree.heading("Details", text="Details")
        self._tree.column("#0", width=50)
        self._tree.column("Type", width=160)
        self._tree.column("Block", width=180)
        self._tree.column("Details", width=420)
        self._tree.pack(side="left", fill="both", expand=True)
        tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self._tree.yview)
        tree_scroll.pack(side="right", fill="y")
        self._tree.configure(yscrollcommand=tree_scroll.set)
        self._tree.tag_configure("volt", background="#dff5d8")
        self._tree.tag_configure("block", background="#fff3cd")
        self._tree.tag_configure("alert", background="#f8d7da")
        self._tree.tag_configure("bo", background="#e8ddff")
        self._tree.tag_configure("default", background="#f2f2f2")

        legend = ttk.Frame(top)
        legend.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Label(legend, text="Legend:").pack(side="left")
        self._legend_chip(legend, "#dff5d8", "Voltammetry (CV/SWV)")
        self._legend_chip(legend, "#fff3cd", "Block step")
        self._legend_chip(legend, "#f8d7da", "Alert/Pause")
        self._legend_chip(legend, "#e8ddff", "BO loop")
        self._legend_chip(legend, "#f2f2f2", "Other")

        self._ctx = tk.Menu(self._tree, tearoff=0)
        self._ctx.add_command(label="Copy", command=self._copy_selected)
        self._ctx.add_command(label="Paste After", command=self._paste_after_selected)
        self._ctx.add_command(label="Duplicate", command=self._duplicate_selected)
        self._ctx.add_command(label="Select Range…", command=self._select_range_prompt)
        self._ctx.add_separator()
        self._ctx.add_command(label="Delete", command=self._delete_selected)
        self._tree.bind("<Button-3>", self._show_ctx)
        self._tree.bind("<Shift-Button-1>", self._select_range)
        self._tree.bind("<Double-1>", self._on_tree_double_click)
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
        try:
            item = self._build_pump_item(
                action=action,
                speed=int(self._pump_speed.get()),
                volume=float(self._pump_volume.get()),
                port=int(self._pump_port.get()),
                pause=float(self._pause_seconds.get()),
                alert=(self._alert_message.get() or "").strip(),
            )
        except Exception as exc:
            messagebox.showerror("Invalid pump step", str(exc))
            return
        self._recipe.append(item)
        self._refresh()

    def _default_bo_analysis_config(self) -> dict:
        try:
            config = normalize_bo_config(load_bo_config(BO_DEFAULT_CONFIG_PATH))
            return dict(config.get("analysis") or {})
        except Exception:
            return {}

    def _default_bo_block(self) -> dict:
        analysis_cfg = self._default_bo_analysis_config()
        block = {
            "bo_config_path": str(BO_DEFAULT_CONFIG_PATH),
            "analysis_output_dir": str(BO_ANALYSIS_OUTPUT_DIR),
            "analysis_file_glob": str(BO_ANALYSIS_FILE_GLOB),
            "target_iterations": 3,
            "channels_override": "",
            "analysis": analysis_cfg,
        }
        return block

    def _bo_details(self, block: dict) -> str:
        target = int(block.get("target_iterations", 1) or 1)
        channels = (block.get("channels_override") or "").strip() or "config channels"
        config_name = Path(str(block.get("bo_config_path") or "BO config")).name
        return f"{config_name} | {target} iter | {channels}"

    def _add_bo_loop_step(self):
        block = self._default_bo_block()
        item = {
            "type": "BO_AUTO_LOOP",
            "status": "pending",
            "details": self._bo_details(block),
            "bo_block": block,
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

        ttk.Label(top, text="View:").pack(side="left", padx=(10, 0))
        self._mux_filter = tk.StringVar(value="ALL")
        ttk.Combobox(
            top,
            textvariable=self._mux_filter,
            values=["ALL", "BASE", "MUX"],
            state="readonly",
            width=8,
        ).pack(side="left", padx=6)
        self._mux_filter.trace_add("write", lambda *_: self._refresh_methods())

        ttk.Button(top, text="Refresh",
                   command=self._load_method_map).pack(side="left", padx=6)
        ttk.Button(top, text="Delete Method",
                   command=self._delete_method_family).pack(side="left", padx=6)
        ttk.Button(top, text="Clear MUX Methods",
                   command=self._clear_mux_methods).pack(side="left", padx=6)

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

        ttk.Label(sweep, text="Repeats/ch:").grid(row=0, column=7, **pad, sticky="e")
        self._sweep_repeats = tk.IntVar(value=1)
        ttk.Entry(sweep, width=6, textvariable=self._sweep_repeats).grid(
            row=0, column=8, **pad, sticky="w"
        )

        ttk.Label(sweep, text="Custom order:").grid(row=1, column=0, **pad, sticky="e")
        self._sweep_custom = tk.StringVar(value="")
        ttk.Entry(sweep, width=44, textvariable=self._sweep_custom).grid(
            row=1, column=1, columnspan=5, **pad, sticky="we"
        )
        ttk.Label(sweep, text="e.g. 1,3,5,2,4").grid(row=1, column=6, columnspan=3, **pad, sticky="w")

        ttk.Button(
            sweep,
            text="Add Channel Sweep Block",
            command=self._add_method_sweep_block,
        ).grid(row=0, column=9, rowspan=2, padx=(12, 6), pady=4, sticky="ns")

        cols = ("Hash", "Note", "Technique", "Params")
        self._method_tree = ttk.Treeview(parent, columns=cols, show="headings", height=8)
        self._method_tree.heading("Hash", text="Hash")
        self._method_tree.heading("Note", text="Note")
        self._method_tree.heading("Technique", text="Technique")
        self._method_tree.heading("Params", text="Params")
        self._method_tree.column("Hash", width=140)
        self._method_tree.column("Note", width=220)
        self._method_tree.column("Technique", width=100)
        self._method_tree.column("Params", width=320)
        self._method_tree.pack(fill="both", expand=True, padx=6, pady=6)
        self._method_tree.bind("<Double-1>", self._on_method_tree_double_click)

        self._load_method_map()

        hint = ttk.Label(
            parent,
            text=(
                "Select a method and use Add Method Step, double-click to edit its note, "
                "or configure channels and use 'Add Channel Sweep Block'."
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
        view = (getattr(self, "_mux_filter", tk.StringVar(value="ALL")).get() or "ALL").upper()

        for key, entry in sorted(self._method_entries.items()):
            technique = entry.get("technique", "")
            note = entry.get("note", "")
            mux_raw = entry.get("mux_channel")
            is_mux = mux_raw not in (None, "", 0, "0")
            if tech != "ALL" and technique.upper() != tech:
                continue
            if view == "BASE" and is_mux:
                continue
            if view == "MUX" and not is_mux:
                continue

            params = entry.get("params", {})
            params_str = ", ".join(f"{k}={v}" for k, v in params.items())
            hay = f"{key} {note} {technique} {params_str}".lower()
            if search and search not in hay:
                continue

            iid = self._method_tree.insert(
                "", "end",
                values=(key, note, technique, params_str),
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

    def _edit_method_note(self):
        selected = self._selected_method_entry()
        if not selected:
            messagebox.showwarning("No selection", "Select a method from the library list.")
            return
        key, entry = selected
        current_note = entry.get("note", "")
        new_note = simpledialog.askstring(
            "Edit Library Note",
            "Update note for this method family:",
            initialvalue=current_note,
            parent=self._frame,
        )
        if new_note is None:
            return
        changed = library_map.update_family_note(key, new_note)
        self._load_method_map()
        if changed:
            messagebox.showinfo("Updated", f"Updated note for {changed} method entr{'y' if changed == 1 else 'ies'}.")

    def _on_method_tree_double_click(self, event):
        row = self._method_tree.identify_row(event.y)
        if not row:
            return
        self._method_tree.selection_set(row)
        self._edit_method_note()

    def _delete_method_family(self):
        selected = self._selected_method_entry()
        if not selected:
            messagebox.showwarning("No selection", "Select a method from the library list.")
            return
        key, entry = selected
        family = library_map.family_keys(key)
        technique = entry.get("technique", "")
        prompt = (
            f"Delete method '{key}' from the library"
            f"{' and its mux variants' if len(family) > 1 else ''}?\n\n"
            f"Technique: {technique}\n"
            f"Entries to delete: {len(family)}"
        )
        if not messagebox.askyesno("Delete Method", prompt):
            return
        deleted = library_map.delete_family(key)
        self._load_method_map()
        if deleted:
            messagebox.showinfo("Deleted", f"Deleted {deleted} method entr{'y' if deleted == 1 else 'ies'} from the library.")

    def _clear_mux_methods(self):
        keys = library_map.mux_method_keys()
        if not keys:
            messagebox.showinfo("Clear MUX Methods", "There are no MUX methods to clear.")
            return
        prompt = (
            f"Delete all {len(keys)} MUX-specific method entr"
            f"{'y' if len(keys) == 1 else 'ies'} from the library?\n\n"
            "Base methods will be kept. Saved recipes that point directly to a deleted "
            "MUX method may need to regenerate it from the base method."
        )
        if not messagebox.askyesno("Clear MUX Methods", prompt):
            return
        deleted = library_map.delete_mux_methods()
        self._load_method_map()
        messagebox.showinfo(
            "Clear MUX Methods",
            f"Deleted {deleted} MUX method entr{'y' if deleted == 1 else 'ies'}.",
        )

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

    def _parse_sweep_repeats(self):
        repeats = int(self._sweep_repeats.get())
        if repeats < 1:
            raise ValueError("Repeats/ch must be >= 1.")
        if repeats > 1000:
            raise ValueError("Repeats/ch is too large (max 1000).")
        return repeats

    def _add_method_sweep_block(self):
        selected = self._selected_method_entry()
        if not selected:
            messagebox.showwarning("No selection", "Select a method from the library list.")
            return
        try:
            channels = self._parse_sweep_channels()
            repeats = self._parse_sweep_repeats()
        except Exception as exc:
            messagebox.showerror("Invalid sweep settings", str(exc))
            return

        key, entry = selected
        technique = entry.get("technique", "")
        params = copy.deepcopy(entry.get("params", {}))
        block_name = f"Sweep {technique} ({len(channels)} ch x {repeats})"
        for ch in channels:
            for rep in range(1, repeats + 1):
                rep_suffix = f" | rep {rep}/{repeats}" if repeats > 1 else ""
                item = {
                    "type": technique,
                    "status": "pending",
                    "details": f"{key}.ms | MUX ch {ch}{rep_suffix}",
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
        if item_type == "BO_AUTO_LOOP":
            return "bo"
        if item_type in ("PAUSE", "ALERT"):
            return "alert"
        if item.get("block_name") or item.get("block_ref"):
            return "block"
        return "default"

    def _refresh(self):
        self._apply_pump_speeds()
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
        row = self._tree.identify_row(event.y)
        if row:
            self._tree.selection_set(row)
            self._last_selected = row
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

    def _send_to_queue(self):
        if not self._recipe:
            messagebox.showwarning("Empty", "No recipe items to send.")
            return
        if not callable(self._on_send_to_queue):
            messagebox.showwarning("Unavailable", "Queue is not available.")
            return
        for item in self._recipe:
            cloned = copy.deepcopy(item)
            cloned["status"] = "pending"
            self._on_send_to_queue(cloned)

    def _build_pump_item(
        self,
        *,
        action: str,
        speed: int,
        volume: float,
        port: int,
        pause: float,
        alert: str,
    ) -> dict:
        action = (action or "").strip().upper()
        if not action:
            raise ValueError("Pump action is required.")

        if action == "PAUSE":
            seconds = float(pause)
            return {
                "type": "PAUSE",
                "status": "pending",
                "details": f"Pause for {seconds:.1f} sec",
                "pause_seconds": seconds,
            }
        if action == "ALERT":
            msg = (alert or "").strip()
            if not msg:
                raise ValueError("Alert message cannot be empty.")
            return {
                "type": "ALERT",
                "status": "pending",
                "details": f"Alert: {msg}",
                "alert_message": msg,
            }

        details = ""
        pump_action = {"name": action, "params": {}}
        if action == "INIT":
            details = "Pump: Initialize (ZR)"
        elif action == "SET_SPEED":
            details = f"Pump: Set speed S{int(speed)}R"
            pump_action["params"]["speed"] = int(speed)
        elif action == "VALVE":
            details = f"Pump: Valve -> {int(port)}"
            pump_action["params"]["port"] = int(port)
        elif action == "ASPIRATE":
            details = f"Pump: Aspirate {float(volume):.2f} uL @ S{int(speed)}R"
            pump_action["params"].update({"speed": int(speed), "volume": float(volume)})
        elif action == "DISPENSE":
            details = f"Pump: Dispense {float(volume):.2f} uL @ S{int(speed)}R"
            pump_action["params"].update({"speed": int(speed), "volume": float(volume)})
        else:
            raise ValueError(f"Unsupported pump action: {action}")

        return {
            "type": f"PUMP_{action}",
            "status": "pending",
            "details": details,
            "pump_action": pump_action,
        }

    def _apply_pump_speeds(self):
        current_speed = None
        for item in self._recipe:
            item_type = (item.get("type") or "").upper()
            if item_type == "PUMP_SET_SPEED":
                params = (item.get("pump_action") or {}).get("params") or {}
                try:
                    current_speed = int(params.get("speed"))
                except (TypeError, ValueError):
                    current_speed = None
                if current_speed is not None:
                    item["details"] = f"Pump: Set speed S{current_speed}R"
                continue

            if item_type in ("PUMP_ASPIRATE", "PUMP_DISPENSE") and current_speed is not None:
                action_info = item.get("pump_action")
                if not isinstance(action_info, dict):
                    action_info = {"name": item_type.replace("PUMP_", ""), "params": {}}
                    item["pump_action"] = action_info
                params = action_info.get("params")
                if not isinstance(params, dict):
                    params = {}
                    action_info["params"] = params

                params["speed"] = current_speed
                try:
                    volume = float(params.get("volume"))
                except (TypeError, ValueError):
                    volume = None

                label = "Aspirate" if item_type == "PUMP_ASPIRATE" else "Dispense"
                if volume is None:
                    item["details"] = f"Pump: {label} @ S{current_speed}R"
                else:
                    item["details"] = f"Pump: {label} {volume:.2f} uL @ S{current_speed}R"

    def _on_tree_double_click(self, event):
        row = self._tree.identify_row(event.y)
        if not row:
            return
        try:
            idx = self._tree.index(row)
        except Exception:
            return
        if idx < 0 or idx >= len(self._recipe):
            return
        item = self._recipe[idx]
        if (item.get("type") or "").upper() == "BO_AUTO_LOOP":
            self._edit_bo_step(idx)
            return
        if not self._is_pump_editable(item):
            return
        self._edit_pump_step(idx)

    def _is_pump_editable(self, item: dict) -> bool:
        item_type = (item.get("type") or "").upper()
        return item_type.startswith("PUMP_") or item_type in ("PAUSE", "ALERT")

    def _extract_pump_fields(self, item: dict) -> dict:
        item_type = (item.get("type") or "").upper()
        if item_type == "PAUSE":
            return {
                "action": "PAUSE",
                "speed": 20,
                "volume": 100.0,
                "port": 1,
                "pause": float(item.get("pause_seconds", 10.0)),
                "alert": "Check setup",
            }
        if item_type == "ALERT":
            return {
                "action": "ALERT",
                "speed": 20,
                "volume": 100.0,
                "port": 1,
                "pause": 10.0,
                "alert": str(item.get("alert_message") or ""),
            }

        action_info = item.get("pump_action") or {}
        action = (action_info.get("name") or item_type.replace("PUMP_", "")).upper()
        params = action_info.get("params") or {}
        return {
            "action": action,
            "speed": int(params.get("speed", 20)),
            "volume": float(params.get("volume", 100.0)),
            "port": int(params.get("port", 1)),
            "pause": 10.0,
            "alert": "Check setup",
        }

    def _edit_pump_step(self, index: int):
        item = self._recipe[index]
        fields = self._extract_pump_fields(item)

        win = tk.Toplevel(self._frame)
        win.title("Edit Pump Step")
        win.transient(self._frame.winfo_toplevel())
        win.grab_set()

        pad = {"padx": 6, "pady": 4}
        ttk.Label(win, text="Pump action:").grid(row=0, column=0, **pad, sticky="e")
        action_var = tk.StringVar(value=fields["action"])
        ttk.Combobox(
            win,
            textvariable=action_var,
            values=["INIT", "SET_SPEED", "VALVE", "ASPIRATE", "DISPENSE", "PAUSE", "ALERT"],
            width=16,
            state="readonly",
        ).grid(row=0, column=1, **pad, sticky="w")

        ttk.Label(win, text="Speed:").grid(row=0, column=2, **pad, sticky="e")
        speed_var = tk.IntVar(value=fields["speed"])
        ttk.Entry(win, width=8, textvariable=speed_var).grid(row=0, column=3, **pad, sticky="w")

        ttk.Label(win, text="Volume (uL):").grid(row=0, column=4, **pad, sticky="e")
        volume_var = tk.DoubleVar(value=fields["volume"])
        ttk.Entry(win, width=10, textvariable=volume_var).grid(row=0, column=5, **pad, sticky="w")

        ttk.Label(win, text="Valve port:").grid(row=0, column=6, **pad, sticky="e")
        port_var = tk.IntVar(value=fields["port"])
        ttk.Entry(win, width=8, textvariable=port_var).grid(row=0, column=7, **pad, sticky="w")

        ttk.Label(win, text="Pause (sec):").grid(row=0, column=8, **pad, sticky="e")
        pause_var = tk.DoubleVar(value=fields["pause"])
        ttk.Entry(win, width=10, textvariable=pause_var).grid(row=0, column=9, **pad, sticky="w")

        ttk.Label(win, text="Alert message:").grid(row=1, column=0, **pad, sticky="e")
        alert_var = tk.StringVar(value=fields["alert"])
        ttk.Entry(win, width=50, textvariable=alert_var).grid(
            row=1, column=1, columnspan=6, **pad, sticky="w"
        )

        btns = ttk.Frame(win)
        btns.grid(row=2, column=0, columnspan=10, pady=(6, 8))

        def _apply():
            try:
                new_item = self._build_pump_item(
                    action=action_var.get(),
                    speed=int(speed_var.get()),
                    volume=float(volume_var.get()),
                    port=int(port_var.get()),
                    pause=float(pause_var.get()),
                    alert=(alert_var.get() or "").strip(),
                )
            except Exception as exc:
                messagebox.showerror("Invalid pump step", str(exc))
                return

            for key in ("block_name", "block_ref"):
                if key in item and key not in new_item:
                    new_item[key] = item[key]
            if "status" in item:
                new_item["status"] = item.get("status")

            self._recipe[index] = new_item
            self._refresh()
            win.destroy()

        ttk.Button(btns, text="Update", command=_apply).pack(side="left", padx=6)
        ttk.Button(btns, text="Cancel", command=win.destroy).pack(side="left", padx=6)
        win.bind("<Return>", lambda _e: _apply())
        win.bind("<Escape>", lambda _e: win.destroy())

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

    # -----Blocks ---------------------------------------------

    def _build_blocks_library(self, parent):
        top = ttk.Frame(parent)
        top.pack(fill="x", padx=6, pady=6)
        ttk.Button(top, text="Refresh Blocks",
                   command=self._load_blocks).pack(side="left", padx=4)
        ttk.Button(top, text="Add Block",
                   command=self._add_selected_block).pack(side="left", padx=4)
        ttk.Label(top, text="View:").pack(side="left", padx=(12, 2))
        self._block_filter = tk.StringVar(value="All")
        ttk.Combobox(
            top,
            textvariable=self._block_filter,
            values=["All", "Default", "Custom", "Saved"],
            state="readonly",
            width=10,
        ).pack(side="left", padx=4)
        self._block_filter.trace_add("write", lambda *_: self._load_blocks())

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
            text=(
                "Blocks are predefined sequences stored in recipe_maker/default_blocks/, "
                "recipe_maker/custom_blocks/, and recipe_maker/saved_recipes/."
            ),
            foreground="#666",
        )
        hint.pack(side="bottom", anchor="w", padx=8, pady=(0, 6))

    def _load_blocks(self):
        self._blocks.clear()
        self._block_iid_to_name.clear()
        for row in self._block_tree.get_children():
            self._block_tree.delete(row)

        view = (getattr(self, "_block_filter", tk.StringVar(value="All")).get() or "All").lower()
        if view == "default":
            blocks_dirs = [self._default_blocks_dir]
        elif view == "custom":
            blocks_dirs = [self._custom_blocks_dir]
        elif view == "saved":
            blocks_dirs = [self._saved_blocks_dir]
        else:
            blocks_dirs = [self._default_blocks_dir, self._custom_blocks_dir, self._saved_blocks_dir]
        files = []
        for blocks_dir in blocks_dirs:
            if not blocks_dir.exists():
                continue
            files.extend(list(blocks_dir.glob("*.json")) + list(blocks_dir.glob("*.JSON")))
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
            if name in self._blocks:
                name = f"{name} ({path.parent.name})"
            self._blocks[name] = items
            summary = ", ".join(item.get("type", "") for item in items[:5])
            if len(items) > 5:
                summary += f" (+{len(items) - 5})"
            iid = self._block_tree.insert("", "end", values=(name, summary))
            self._block_iid_to_name[iid] = name

        if not self._blocks:
            self._block_tree.insert(
                "", "end",
                values=("No blocks found", "recipe_maker/default_blocks, custom_blocks, or saved_recipes"),
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

    def _edit_bo_step(self, index: int):
        item = self._recipe[index]
        block = copy.deepcopy(item.get("bo_block") or self._default_bo_block())
        default_analysis_output = str(BO_ANALYSIS_OUTPUT_DIR)
        if not str(block.get("analysis_output_dir") or "").strip():
            block["analysis_output_dir"] = default_analysis_output

        win = tk.Toplevel(self._frame)
        win.title("Edit BO Loop")
        win.transient(self._frame.winfo_toplevel())
        win.grab_set()

        pad = {"padx": 6, "pady": 4}
        ttk.Label(win, text="BO config:").grid(row=0, column=0, **pad, sticky="e")
        cfg_var = tk.StringVar(value=str(block.get("bo_config_path") or ""))
        ttk.Entry(win, width=56, textvariable=cfg_var).grid(row=0, column=1, columnspan=3, **pad, sticky="we")
        ttk.Button(
            win,
            text="Browse",
            command=lambda: self._set_string_from_dialog(
                cfg_var,
                filedialog.askopenfilename(
                    title="Choose BO config",
                    filetypes=[("JSON", "*.json"), ("All files", "*.*")],
                    initialdir=str(Path(cfg_var.get()).parent if cfg_var.get() else self._repo_root),
                ),
            ),
        ).grid(row=0, column=4, **pad, sticky="w")

        ttk.Label(win, text="Analysis output:").grid(row=1, column=0, **pad, sticky="e")
        out_var = tk.StringVar(value=str(block.get("analysis_output_dir") or ""))
        ttk.Entry(win, width=56, textvariable=out_var).grid(row=1, column=1, columnspan=3, **pad, sticky="we")
        ttk.Button(
            win,
            text="Browse",
            command=lambda: self._set_string_from_dialog(
                out_var,
                filedialog.askdirectory(title="Choose BO analysis output folder"),
            ),
        ).grid(row=1, column=4, **pad, sticky="w")

        ttk.Label(win, text="Target iterations:").grid(row=2, column=0, **pad, sticky="e")
        target_var = tk.IntVar(value=int(block.get("target_iterations", 3) or 3))
        ttk.Entry(win, width=8, textvariable=target_var).grid(row=2, column=1, **pad, sticky="w")
        ttk.Label(win, text="Channels override:").grid(row=3, column=0, **pad, sticky="e")
        channels_var = tk.StringVar(value=str(block.get("channels_override") or ""))
        ttk.Entry(win, width=24, textvariable=channels_var).grid(row=3, column=1, **pad, sticky="w")
        ttk.Label(win, text="Glob:").grid(row=3, column=2, **pad, sticky="e")
        glob_var = tk.StringVar(value=str(block.get("analysis_file_glob") or BO_ANALYSIS_FILE_GLOB))
        ttk.Entry(win, width=18, textvariable=glob_var).grid(row=3, column=3, **pad, sticky="w")

        analysis = block.get("analysis") or {}
        ttk.Label(win, text="Crop min/max (V):").grid(row=4, column=0, **pad, sticky="e")
        crop_min_var = tk.StringVar(value=str(analysis.get("crop_min_v", -0.6)))
        crop_max_var = tk.StringVar(value=str(analysis.get("crop_max_v", -0.1)))
        ttk.Entry(win, width=8, textvariable=crop_min_var).grid(row=4, column=1, **pad, sticky="w")
        ttk.Entry(win, width=8, textvariable=crop_max_var).grid(row=4, column=1, padx=(76, 6), pady=4, sticky="w")
        ttk.Label(win, text="Smooth win/poly:").grid(row=4, column=2, **pad, sticky="e")
        smooth_win_var = tk.StringVar(value=str(analysis.get("smooth_window", 15)))
        smooth_poly_var = tk.StringVar(value=str(analysis.get("smooth_polyorder", 2)))
        ttk.Entry(win, width=8, textvariable=smooth_win_var).grid(row=4, column=3, **pad, sticky="w")
        ttk.Entry(win, width=8, textvariable=smooth_poly_var).grid(row=4, column=3, padx=(76, 6), pady=4, sticky="w")

        ttk.Label(win, text="Minima window (V):").grid(row=5, column=0, **pad, sticky="e")
        minima_var = tk.StringVar(value=str(analysis.get("minima_search_window_v", 0.30)))
        ttk.Entry(win, width=10, textvariable=minima_var).grid(row=5, column=1, **pad, sticky="w")
        ttk.Label(win, text="Min peak height (uA):").grid(row=5, column=2, **pad, sticky="e")
        min_peak_var = tk.StringVar(value="" if analysis.get("min_peak_height_ua") in (None, "") else str(analysis.get("min_peak_height_ua")))
        ttk.Entry(win, width=10, textvariable=min_peak_var).grid(row=5, column=3, **pad, sticky="w")

        ttk.Label(win, text="Min start V:").grid(row=6, column=0, **pad, sticky="e")
        min_start_var = tk.StringVar(value=str(analysis.get("min_start_voltage_v", -0.6)))
        ttk.Entry(win, width=10, textvariable=min_start_var).grid(row=6, column=1, **pad, sticky="w")
        ttk.Label(win, text="Scan windows:").grid(row=6, column=2, **pad, sticky="e")
        scan_windows_var = tk.StringVar(value=str(analysis.get("scan_windows", "")))
        ttk.Entry(win, width=24, textvariable=scan_windows_var).grid(row=6, column=3, **pad, sticky="w")

        prominent_var = tk.BooleanVar(value=bool(analysis.get("use_prominent_minima", False)))
        double_corr_var = tk.BooleanVar(value=bool(analysis.get("use_double_correction", True)))
        skew_var = tk.BooleanVar(value=bool(analysis.get("compute_skew", False)))
        wavelet_energy_var = tk.BooleanVar(value=bool(analysis.get("compute_wavelet_energy", False)))
        wavelet_trace_var = tk.BooleanVar(value=bool(analysis.get("compute_wavelet_denoised_trace", False)))
        wavelet_corr_var = tk.BooleanVar(value=bool(analysis.get("use_wavelet_for_correction", False)))
        ttk.Checkbutton(win, text="Prominent minima", variable=prominent_var).grid(row=7, column=0, columnspan=2, **pad, sticky="w")
        ttk.Checkbutton(win, text="Double correction", variable=double_corr_var).grid(row=7, column=2, columnspan=2, **pad, sticky="w")
        ttk.Checkbutton(win, text="Compute skew", variable=skew_var).grid(row=8, column=0, columnspan=2, **pad, sticky="w")
        ttk.Checkbutton(win, text="Wavelet energy", variable=wavelet_energy_var).grid(row=8, column=2, columnspan=2, **pad, sticky="w")
        ttk.Checkbutton(win, text="Wavelet trace", variable=wavelet_trace_var).grid(row=9, column=0, columnspan=2, **pad, sticky="w")
        ttk.Checkbutton(win, text="Wavelet correction", variable=wavelet_corr_var).grid(row=9, column=2, columnspan=2, **pad, sticky="w")

        btns = ttk.Frame(win)
        btns.grid(row=10, column=0, columnspan=5, pady=(8, 10))

        def _apply():
            try:
                new_analysis = {
                    "crop_min_v": float(crop_min_var.get()),
                    "crop_max_v": float(crop_max_var.get()),
                    "smooth_window": int(smooth_win_var.get()),
                    "smooth_polyorder": int(smooth_poly_var.get()),
                    "minima_search_window_v": float(minima_var.get()),
                    "min_peak_height_ua": None if not str(min_peak_var.get()).strip() else float(min_peak_var.get()),
                    "min_start_voltage_v": float(min_start_var.get()),
                    "scan_windows": scan_windows_var.get().strip(),
                    "use_prominent_minima": bool(prominent_var.get()),
                    "use_double_correction": bool(double_corr_var.get()),
                    "compute_skew": bool(skew_var.get()),
                    "compute_wavelet_energy": bool(wavelet_energy_var.get()),
                    "compute_wavelet_denoised_trace": bool(wavelet_trace_var.get()),
                    "use_wavelet_for_correction": bool(wavelet_corr_var.get()),
                }
                new_block = {
                    "bo_config_path": cfg_var.get().strip(),
                    "analysis_output_dir": out_var.get().strip(),
                    "analysis_file_glob": glob_var.get().strip() or "*.json",
                    "target_iterations": int(target_var.get()),
                    "channels_override": channels_var.get().strip(),
                    "analysis": new_analysis,
                }
                if not new_block["bo_config_path"]:
                    raise ValueError("BO config path is required.")
                if new_block["target_iterations"] < 1:
                    raise ValueError("Target iterations must be at least 1.")
            except Exception as exc:
                messagebox.showerror("Invalid BO step", str(exc))
                return

            item["bo_block"] = new_block
            item["details"] = self._bo_details(new_block)
            item["status"] = item.get("status", "pending")
            self._recipe[index] = item
            self._refresh()
            win.destroy()

        ttk.Button(btns, text="Update", command=_apply).pack(side="left", padx=6)
        ttk.Button(btns, text="Cancel", command=win.destroy).pack(side="left", padx=6)
        win.bind("<Return>", lambda _e: _apply())
        win.bind("<Escape>", lambda _e: win.destroy())

    @staticmethod
    def _set_string_from_dialog(var: tk.StringVar, value):
        if value:
            var.set(value)
