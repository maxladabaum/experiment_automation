"""
gui/tab_automated_titration.py — BO-driven automated titration recipe builder.
"""

from __future__ import annotations

import copy
import math
import tkinter as tk
from tkinter import messagebox, ttk

from config import PREFERRED_SYRINGE_UL
from core.bo_session import PARAMETER_ORDER, build_swv_script, wrap_mux
from core.titration import (
    calculate_titration_plan,
    parse_concentrations,
    split_transfer,
)
from gui.widgets import ScrollableFrame


class AutomatedTitrationTab:
    """Build pump/SWV recipes using the best parameters from each BO group."""

    ACCENT_DARK = "#3f3a63"
    ACCENT_LIGHT = "#e7e3ff"

    def __init__(
        self,
        parent_frame,
        *,
        session=None,
        on_get_best_parameters=None,
        on_send_to_queue=None,
        on_run_queue=None,
        on_lock_for_bo=None,
    ):
        self._frame = parent_frame
        self._session = session
        self._get_best_parameters = on_get_best_parameters
        self._send_queue_item = on_send_to_queue
        self._run_queue = on_run_queue
        self._on_lock_for_bo = on_lock_for_bo
        self._parameter_groups = []
        self._manual_channel_params = {}
        self._bo_setup_groups = []
        self._bo_locked_settings = None
        self._bo_locked_plan = None
        self._recipe = []
        self._plan = []

        self._status_var = tk.StringVar(
            value="Waiting for optimized parameter groups from Bayesian Optimization."
        )

        self._stock_port_var = tk.StringVar(value="5")
        self._buffer_port_var = tk.StringVar(value="6")
        self._mix_port_var = tk.StringVar(value="4")
        self._flow_cell_port_var = tk.StringVar(value="1")
        self._waste_port_var = tk.StringVar(value="2")
        self._air_port_var = tk.StringVar(value="9")
        self._mix_line_volume_var = tk.StringVar(value="110")
        self._stock_air_spacer_var = tk.StringVar(value="100")
        self._mix_line_air_push_var = tk.StringVar(value="250")
        self._pump_speed_var = tk.StringVar(value="20")
        self._initial_buffer_speed_var = tk.StringVar(value="20")
        self._final_cleanup_speed_var = tk.StringVar(value="20")
        self._syringe_capacity_var = tk.StringVar(value=f"{PREFERRED_SYRINGE_UL:g}")
        self._mix_volume_var = tk.StringVar(value="200")
        self._mix_cycles_var = tk.StringVar(value="3")
        self._equilibration_var = tk.StringVar(value="0")

        self._stock_concentration_var = tk.StringVar(value="10000")
        self._initial_buffer_volume_var = tk.StringVar(value="10000")
        self._aliquot_volume_var = tk.StringVar(value="500")
        self._concentrations_var = tk.StringVar(value="10, 25, 50, 100")
        self._replicates_var = tk.StringVar(value="1")
        self._skip_initial_buffer_var = tk.BooleanVar(value=False)

        self._build()

    def _build(self):
        root = ttk.Frame(self._frame)
        root.pack(fill="both", expand=True)

        banner = tk.Frame(root, bg=self.ACCENT_DARK, height=58)
        banner.pack(side="top", fill="x")
        banner.pack_propagate(False)
        tk.Label(
            banner,
            text="Automated Titration",
            bg=self.ACCENT_DARK,
            fg="white",
            font=("Arial", 16, "bold"),
        ).pack(side="left", padx=(16, 10), pady=12)
        tk.Label(
            banner,
            text="Build pump and SWV sequences from optimized channel-group parameters",
            bg=self.ACCENT_DARK,
            fg=self.ACCENT_LIGHT,
            font=("Arial", 10),
        ).pack(side="left", padx=8, pady=15)

        scroller = ScrollableFrame(root, min_width=1040)
        scroller.pack(fill="both", expand=True)
        content = scroller.content

        ttk.Label(
            content,
            text=(
                "The calculator tracks the liquid and kanamycin remaining in the mixing "
                "tube after every flow-cell aliquot, then calculates the exact 10 mM "
                "stock addition needed for each requested concentration."
            ),
            wraplength=980,
            justify="left",
        ).pack(fill="x", padx=14, pady=(14, 8))

        self._build_parameter_groups(content)
        self._build_manual_parameters(content)
        self._build_setup(content)
        self._build_calculation_preview(content)
        self._build_recipe_preview(content)

        ttk.Label(root, textvariable=self._status_var, relief="sunken").pack(
            side="bottom", fill="x", padx=8, pady=(0, 8)
        )

    def _build_parameter_groups(self, parent):
        source = ttk.LabelFrame(parent, text="Optimized Parameter Groups", padding=10)
        source.pack(fill="x", padx=14, pady=6)

        source_actions = ttk.Frame(source)
        source_actions.pack(fill="x", pady=(0, 8))
        ttk.Button(
            source_actions,
            text="Receive from Bayesian Optimization",
            command=self._receive_best_parameters,
        ).pack(side="left")
        ttk.Label(
            source_actions,
            text="Uses the best recorded Q value for each configured channel group.",
            foreground="#666666",
        ).pack(side="left", padx=10)

        columns = (
            "Group", "Channels", "Score", "Begin (V)", "End (V)", "Step (V)",
            "Amplitude (V)", "Frequency (Hz)", "Conditioning",
        )
        host = ttk.Frame(source)
        host.pack(fill="x")
        self._parameter_tree = ttk.Treeview(
            host, columns=columns, show="headings", height=5
        )
        for column in columns:
            self._parameter_tree.heading(column, text=column)
            width = 170 if column == "Conditioning" else 105
            self._parameter_tree.column(column, width=width, minwidth=80, anchor="center")
        self._parameter_tree.pack(side="top", fill="x", expand=True)
        horizontal = ttk.Scrollbar(
            host, orient="horizontal", command=self._parameter_tree.xview
        )
        horizontal.pack(side="bottom", fill="x")
        self._parameter_tree.configure(xscrollcommand=horizontal.set)

    def _build_manual_parameters(self, parent):
        manual = ttk.LabelFrame(parent, text="Manual SWV Settings by Channel", padding=10)
        manual.pack(fill="x", padx=14, pady=6)

        actions = ttk.Frame(manual)
        actions.pack(fill="x", pady=(0, 8))
        ttk.Button(
            actions,
            text="Edit Selected Channel",
            command=self._edit_manual_channel,
        ).pack(side="left")
        ttk.Button(
            actions,
            text="Set All Channels",
            command=self._edit_all_manual_channels,
        ).pack(side="left", padx=(6, 0))
        ttk.Label(
            actions,
            text=(
                "After receiving optimized groups, select or double-click a channel "
                "to define its manual SWV values."
            ),
            foreground="#666666",
        ).pack(side="left", padx=10)

        columns = (
            "Channel", "Begin (V)", "End (V)", "Step (V)", "Amplitude (V)",
            "Frequency (Hz)", "Conditioning (V)", "Conditioning (s)",
        )
        self._manual_tree = ttk.Treeview(
            manual, columns=columns, show="headings", height=5, selectmode="browse"
        )
        for column in columns:
            self._manual_tree.heading(column, text=column)
            self._manual_tree.column(column, width=125, minwidth=90, anchor="center")
        self._manual_tree.pack(fill="x", expand=True)
        self._manual_tree.bind("<Double-1>", lambda _event: self._edit_manual_channel())

    def _build_setup(self, parent):
        setup = ttk.Frame(parent)
        setup.pack(fill="x", padx=14, pady=6)
        setup.columnconfigure(0, weight=1, uniform="setup")
        setup.columnconfigure(1, weight=1, uniform="setup")

        pump = ttk.LabelFrame(setup, text="Pump and Port Setup", padding=10)
        pump.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        pump.columnconfigure(1, weight=1)
        self._entry_row(pump, 0, "Flow cell port:", self._flow_cell_port_var)
        self._entry_row(pump, 1, "Waste port:", self._waste_port_var)
        self._entry_row(pump, 2, "Mixing tube port:", self._mix_port_var)
        self._entry_row(pump, 3, "10 mM stock port:", self._stock_port_var)
        self._entry_row(pump, 4, "Buffer port:", self._buffer_port_var)
        self._entry_row(pump, 5, "Air port:", self._air_port_var)
        self._entry_row(
            pump, 6, "Port 4 line volume (µL):", self._mix_line_volume_var
        )
        self._entry_row(
            pump, 7, "Stock air spacer (µL):", self._stock_air_spacer_var
        )
        self._entry_row(
            pump, 8, "Port 4 air push (µL):", self._mix_line_air_push_var
        )
        self._entry_row(pump, 9, "Pump speed (1–40):", self._pump_speed_var)
        self._entry_row(
            pump, 10, "Initial buffer speed (1–40):", self._initial_buffer_speed_var
        )
        self._entry_row(
            pump, 11, "Final cleanup speed (1–40):", self._final_cleanup_speed_var
        )
        self._entry_row(pump, 12, "Syringe capacity (µL):", self._syringe_capacity_var)
        self._entry_row(pump, 13, "Mix volume per cycle (µL):", self._mix_volume_var)
        self._entry_row(pump, 14, "Mix cycles after stock:", self._mix_cycles_var)
        self._entry_row(pump, 15, "Flow-cell equilibration (s):", self._equilibration_var)
        ttk.Label(
            pump,
            text=(
                "Stock delivery below one syringe capacity uses spacer air, then "
                "stock. Larger deliveries omit spacer air. The port-4 air push and "
                "line-volume discard occur before mixing in both cases."
            ),
            wraplength=430,
            justify="left",
            foreground="#666666",
        ).grid(row=16, column=0, columnspan=2, sticky="w", pady=(10, 0))

        plan = ttk.LabelFrame(setup, text="Titration Plan", padding=10)
        plan.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        plan.columnconfigure(1, weight=1)
        self._entry_row(
            plan, 0, "Stock concentration (µM):", self._stock_concentration_var
        )
        self._entry_row(
            plan, 1, "Starting buffer volume (µL):", self._initial_buffer_volume_var
        )
        self._entry_row(
            plan, 2, "Flow-cell aliquot each point (µL):", self._aliquot_volume_var
        )
        ttk.Label(plan, text="Desired concentrations (µM):").grid(
            row=3, column=0, sticky="ne", padx=(0, 8), pady=4
        )
        ttk.Entry(plan, textvariable=self._concentrations_var).grid(
            row=3, column=1, sticky="ew", pady=4
        )
        self._entry_row(plan, 4, "SWV replicates per channel:", self._replicates_var)
        ttk.Checkbutton(
            plan,
            text="Bypass initial buffer pump steps (mixing tube is already filled)",
            variable=self._skip_initial_buffer_var,
        ).grid(row=5, column=0, columnspan=2, sticky="w", pady=(6, 2))
        ttk.Label(
            plan,
            text=(
                "Enter ascending concentrations separated by commas or spaces. "
                "Before every point after the first, the prior flow-cell aliquot is "
                "moved from port 1 to waste at port 2. Bypassing initial buffer still "
                "uses the entered starting volume for concentration calculations."
            ),
            wraplength=430,
            justify="left",
            foreground="#666666",
        ).grid(row=6, column=0, columnspan=2, sticky="w", pady=(10, 8))
        ttk.Button(
            plan, text="Generate Recipe", command=self._generate_recipe
        ).grid(row=7, column=0, columnspan=2, sticky="ew", pady=(6, 0))

    def _build_calculation_preview(self, parent):
        calculation = ttk.LabelFrame(
            parent, text="Calculated Liquid Plan", padding=10
        )
        calculation.pack(fill="x", padx=14, pady=6)
        columns = (
            "Target",
            "Before",
            "Mix Volume Before",
            "Stock Addition",
            "Volume After Stock",
            "Flow-cell Aliquot",
            "Remaining",
        )
        self._calculation_tree = ttk.Treeview(
            calculation, columns=columns, show="headings", height=5
        )
        for column in columns:
            self._calculation_tree.heading(column, text=column)
            self._calculation_tree.column(
                column, width=140, minwidth=100, anchor="center"
            )
        self._calculation_tree.pack(fill="x", expand=True)

    def _build_recipe_preview(self, parent):
        preview = ttk.LabelFrame(parent, text="Generated Recipe Preview", padding=10)
        preview.pack(fill="both", expand=True, padx=14, pady=(6, 14))

        actions = ttk.Frame(preview)
        actions.pack(fill="x", pady=(0, 8))
        self._send_button = ttk.Button(
            actions,
            text="Send to Queue",
            command=self._send_to_queue,
            state="disabled",
        )
        self._send_button.pack(side="left")
        self._lock_bo_button = ttk.Button(
            actions,
            text="Lock Auto Settings for BO",
            command=self._lock_auto_settings_for_bo,
            state="disabled",
        )
        self._lock_bo_button.pack(side="left", padx=6)
        ttk.Button(
            actions, text="Clear Preview", command=self._clear_recipe
        ).pack(side="left", padx=6)
        self._summary_label = ttk.Label(
            actions, text="No recipe generated.", foreground="#666666"
        )
        self._summary_label.pack(side="left", padx=8)

        columns = ("Step", "Type", "Point", "Details")
        host = ttk.Frame(preview)
        host.pack(fill="both", expand=True)
        self._recipe_tree = ttk.Treeview(
            host, columns=columns, show="headings", height=13
        )
        widths = {"Step": 60, "Type": 130, "Point": 120, "Details": 650}
        for column in columns:
            self._recipe_tree.heading(column, text=column)
            self._recipe_tree.column(
                column,
                width=widths[column],
                minwidth=50,
                anchor="center" if column != "Details" else "w",
            )
        self._recipe_tree.pack(side="left", fill="both", expand=True)
        vertical = ttk.Scrollbar(host, orient="vertical", command=self._recipe_tree.yview)
        vertical.pack(side="right", fill="y")
        self._recipe_tree.configure(yscrollcommand=vertical.set)

    def _receive_best_parameters(self):
        if not callable(self._get_best_parameters):
            messagebox.showwarning(
                "Bayesian Optimization", "Bayesian Optimization is not available."
            )
            return
        try:
            groups = self._get_best_parameters()
        except Exception as exc:
            messagebox.showerror("Bayesian Optimization", str(exc))
            return
        if not groups:
            messagebox.showwarning(
                "Bayesian Optimization",
                "The current BO session has no completed observations.",
            )
            return

        self._parameter_groups = copy.deepcopy(groups)
        active_channels = {
            int(channel)
            for group in self._parameter_groups
            for channel in group.get("channels", [])
        }
        existing_manual = self._manual_channel_params
        self._manual_channel_params = {}
        for group in self._parameter_groups:
            for channel in group.get("channels", []):
                channel = int(channel)
                if channel in existing_manual:
                    manual_params = copy.deepcopy(existing_manual[channel])
                else:
                    manual_params = self._default_manual_params(group["params"])
                self._manual_channel_params[channel] = manual_params
        for row in self._parameter_tree.get_children():
            self._parameter_tree.delete(row)
        for group in self._parameter_groups:
            params = group["params"]
            conditioning = (
                f"{float(params['conditioning_potential']):g} V / "
                f"{float(params['conditioning_time']):g} s"
            )
            self._parameter_tree.insert(
                "",
                "end",
                values=(
                    group.get("name", f"Group {group.get('id', '')}"),
                    ", ".join(str(channel) for channel in group.get("channels", [])),
                    f"{float(group.get('score', 0.0)):.5g}",
                    f"{float(params['begin_potential']):g}",
                    f"{float(params['end_potential']):g}",
                    f"{float(params['step_potential']):g}",
                    f"{float(params['amplitude']):g}",
                    f"{float(params['frequency']):g}",
                    conditioning,
                ),
            )
        self._refresh_manual_tree()
        self._status_var.set(
            f"Received {len(self._parameter_groups)} optimized parameter group(s) "
            f"for {len(active_channels)} channel(s). Edit the manual SWV table as needed."
        )

    def prepare_for_bo(self, groups):
        """Prepare editable manual-channel rows before optimized values exist."""
        self._bo_setup_groups = copy.deepcopy(groups or [])
        self._bo_locked_settings = None
        self._bo_locked_plan = None
        self._parameter_groups = []
        self._manual_channel_params = {}
        for group in self._bo_setup_groups:
            for channel in group.get("channels", []):
                self._manual_channel_params[int(channel)] = self._default_manual_params(
                    group["params"]
                )
        for row in self._parameter_tree.get_children():
            self._parameter_tree.delete(row)
        self._refresh_manual_tree()
        self._lock_bo_button.configure(state="normal")
        self._status_var.set(
            "Configure autotitration and manual SWV settings, then click "
            "'Lock Auto Settings for BO'."
        )

    def cancel_bo_autotitration(self):
        self._bo_setup_groups = []
        self._bo_locked_settings = None
        self._bo_locked_plan = None
        self._lock_bo_button.configure(state="disabled")

    def bo_settings_locked(self):
        return self._bo_locked_settings is not None and self._bo_locked_plan is not None

    def _lock_auto_settings_for_bo(self):
        if not self._bo_setup_groups:
            messagebox.showwarning(
                "Automated Titration",
                "Enable automatic titration from the BO Setup tab first.",
            )
            return
        try:
            settings = self._read_settings()
            plan = calculate_titration_plan(
                parse_concentrations(self._concentrations_var.get()),
                stock_concentration_um=settings["stock_concentration"],
                initial_buffer_volume_ul=settings["initial_buffer_volume"],
                aliquot_volume_ul=settings["aliquot_volume"],
            )
            expected_channels = {
                int(channel)
                for group in self._bo_setup_groups
                for channel in group.get("channels", [])
            }
            if expected_channels != set(self._manual_channel_params):
                raise ValueError("Manual SWV settings are required for every BO channel.")
        except Exception as exc:
            messagebox.showerror("Lock Auto Settings for BO", str(exc))
            return
        self._bo_locked_settings = copy.deepcopy(settings)
        self._bo_locked_plan = copy.deepcopy(plan)
        self._lock_bo_button.configure(state="disabled")
        self._status_var.set("Autotitration settings are locked for the active BO run.")
        if callable(self._on_lock_for_bo):
            self._on_lock_for_bo()

    def run_locked_after_bo(self, optimized_groups):
        """Build, enqueue, and immediately start the locked post-BO titration."""
        if not self.bo_settings_locked():
            raise RuntimeError("Autotitration settings were not locked before BO started.")
        if not callable(self._send_queue_item) or not callable(self._run_queue):
            raise RuntimeError("The autotitration queue runner is not available.")
        self._parameter_groups = copy.deepcopy(optimized_groups)
        self._recipe = self._build_recipe(
            copy.deepcopy(self._bo_locked_settings),
            copy.deepcopy(self._bo_locked_plan),
        )
        for item in self._recipe:
            if item.get("type") == "SWV":
                queue_item = self._materialize_swv_item(item)
            else:
                queue_item = {
                    key: copy.deepcopy(value)
                    for key, value in item.items()
                    if not key.startswith("_")
                }
            self._send_queue_item(queue_item)
        self._status_var.set(
            f"BO complete; queued {len(self._recipe)} locked autotitration steps."
        )
        self._run_queue()

    def _refresh_manual_tree(self):
        for row in self._manual_tree.get_children():
            self._manual_tree.delete(row)
        for channel in sorted(self._manual_channel_params):
            params = self._manual_channel_params[channel]
            self._manual_tree.insert(
                "", "end", iid=str(channel),
                values=(
                    channel,
                    f"{float(params['begin_potential']):g}",
                    f"{float(params['end_potential']):g}",
                    f"{float(params['step_potential']):g}",
                    f"{float(params['amplitude']):g}",
                    f"{float(params['frequency']):g}",
                    f"{float(params['conditioning_potential']):g}",
                    f"{float(params['conditioning_time']):g}",
                ),
            )

    @staticmethod
    def _default_manual_params(optimized_params):
        params = copy.deepcopy(optimized_params)
        params.update(
            {
                "amplitude": 0.036,
                "step_potential": 0.002,
                "frequency": 200.0,
            }
        )
        return params

    def _edit_manual_channel(self):
        selection = self._manual_tree.selection()
        if not selection:
            messagebox.showwarning(
                "Manual SWV Settings",
                "Receive optimized parameters, then select a channel to edit.",
            )
            return
        channel = int(selection[0])
        self._show_manual_editor(channel=channel)

    def _edit_all_manual_channels(self):
        if not self._manual_channel_params:
            messagebox.showwarning(
                "Manual SWV Settings",
                "Receive optimized parameters before setting manual channel values.",
            )
            return
        first_channel = min(self._manual_channel_params)
        self._show_manual_editor(
            channel=None,
            initial=self._manual_channel_params[first_channel],
        )

    def _show_manual_editor(self, *, channel=None, initial=None):
        current = initial if initial is not None else self._manual_channel_params[channel]
        title = (
            "Set Manual SWV Settings — All Channels"
            if channel is None
            else f"Manual SWV Settings — Channel {channel}"
        )
        dialog = tk.Toplevel(self._frame)
        dialog.title(title)
        dialog.transient(self._frame.winfo_toplevel())
        dialog.grab_set()
        body = ttk.Frame(dialog, padding=12)
        body.pack(fill="both", expand=True)
        labels = {
            "begin_potential": "Begin potential (V):",
            "end_potential": "End potential (V):",
            "step_potential": "Step potential (V):",
            "amplitude": "Amplitude (V):",
            "frequency": "Frequency (Hz):",
            "conditioning_potential": "Conditioning potential (V):",
            "conditioning_time": "Conditioning time (s):",
        }
        variables = {}
        for row, name in enumerate(PARAMETER_ORDER):
            variables[name] = tk.StringVar(value=f"{float(current[name]):g}")
            self._entry_row(body, row, labels[name], variables[name])

        def save():
            try:
                values = {name: float(var.get()) for name, var in variables.items()}
                if values["step_potential"] <= 0:
                    raise ValueError("Step potential must be greater than zero.")
                if values["amplitude"] <= 0:
                    raise ValueError("Amplitude must be greater than zero.")
                if values["frequency"] <= 0:
                    raise ValueError("Frequency must be greater than zero.")
                if values["conditioning_time"] < 0:
                    raise ValueError("Conditioning time cannot be negative.")
            except ValueError as exc:
                messagebox.showerror("Manual SWV Settings", str(exc), parent=dialog)
                return
            if channel is None:
                for manual_channel in self._manual_channel_params:
                    self._manual_channel_params[manual_channel] = copy.deepcopy(values)
                status = (
                    f"Updated manual SWV settings for "
                    f"{len(self._manual_channel_params)} channel(s)."
                )
            else:
                self._manual_channel_params[channel] = values
                status = f"Updated manual SWV settings for channel {channel}."
            self._refresh_manual_tree()
            self._status_var.set(status)
            dialog.destroy()

        buttons = ttk.Frame(body)
        buttons.grid(row=len(PARAMETER_ORDER), column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(buttons, text="Cancel", command=dialog.destroy).pack(side="right")
        ttk.Button(buttons, text="Save", command=save).pack(side="right", padx=6)

    def _generate_recipe(self):
        try:
            settings = self._read_settings()
            self._plan = calculate_titration_plan(
                parse_concentrations(self._concentrations_var.get()),
                stock_concentration_um=settings["stock_concentration"],
                initial_buffer_volume_ul=settings["initial_buffer_volume"],
                aliquot_volume_ul=settings["aliquot_volume"],
            )
            self._recipe = self._build_recipe(settings, self._plan)
        except Exception as exc:
            messagebox.showerror("Titration Recipe", str(exc))
            return

        self._refresh_calculation_tree()
        self._refresh_recipe_tree()
        total_stock = sum(point.stock_added_ul for point in self._plan)
        summary = (
            f"{len(self._plan)} concentration point(s), {total_stock:.2f} µL total "
            f"stock, {len(self._recipe)} queue step(s)"
        )
        if not self._parameter_groups:
            summary += " — pump-only preview; receive BO parameters to add SWV steps"
        self._summary_label.configure(text=summary)
        self._send_button.configure(
            state="normal" if self._recipe and self._parameter_groups else "disabled"
        )
        self._status_var.set(f"Recipe generated: {summary}.")

    def _refresh_calculation_tree(self):
        for row in self._calculation_tree.get_children():
            self._calculation_tree.delete(row)
        for point in self._plan:
            self._calculation_tree.insert(
                "",
                "end",
                values=(
                    f"{point.target_concentration_um:g} µM",
                    f"{point.concentration_before_um:g} µM",
                    f"{point.volume_before_stock_ul:.3f} µL",
                    f"{point.stock_added_ul:.3f} µL",
                    f"{point.volume_after_stock_ul:.3f} µL",
                    f"{point.aliquot_removed_ul:.3f} µL",
                    f"{point.volume_remaining_ul:.3f} µL",
                ),
            )

    def _read_settings(self):
        settings = {
            "stock_port": self._port(self._stock_port_var.get(), "Stock port"),
            "buffer_port": self._port(self._buffer_port_var.get(), "Buffer port"),
            "mix_port": self._port(self._mix_port_var.get(), "Mixing tube port"),
            "flow_port": self._port(self._flow_cell_port_var.get(), "Flow cell port"),
            "waste_port": self._port(self._waste_port_var.get(), "Waste port"),
            "air_port": self._port(self._air_port_var.get(), "Air port"),
            "mix_line_volume": self._nonnegative(
                self._mix_line_volume_var.get(), "Port 4 line volume"
            ),
            "stock_air_spacer": self._nonnegative(
                self._stock_air_spacer_var.get(), "Stock air spacer"
            ),
            "mix_line_air_push": self._positive(
                self._mix_line_air_push_var.get(), "Port 4 air push"
            ),
            "speed": self._integer(self._pump_speed_var.get(), "Pump speed", minimum=1),
            "initial_buffer_speed": self._integer(
                self._initial_buffer_speed_var.get(), "Initial buffer speed", minimum=1
            ),
            "final_cleanup_speed": self._integer(
                self._final_cleanup_speed_var.get(), "Final cleanup speed", minimum=1
            ),
            "syringe_capacity": self._positive(
                self._syringe_capacity_var.get(), "Syringe capacity"
            ),
            "mix_volume": self._positive(self._mix_volume_var.get(), "Mix volume"),
            "mix_cycles": self._integer(
                self._mix_cycles_var.get(), "Mix cycles", minimum=0
            ),
            "equilibration": self._nonnegative(
                self._equilibration_var.get(), "Equilibration time"
            ),
            "stock_concentration": self._positive(
                self._stock_concentration_var.get(), "Stock concentration"
            ),
            "initial_buffer_volume": self._positive(
                self._initial_buffer_volume_var.get(), "Starting buffer volume"
            ),
            "aliquot_volume": self._positive(
                self._aliquot_volume_var.get(), "Flow-cell aliquot"
            ),
            "replicates": self._integer(
                self._replicates_var.get(), "SWV replicates", minimum=1
            ),
            "skip_initial_buffer": bool(self._skip_initial_buffer_var.get()),
        }
        if any(
            settings[name] > 40
            for name in ("speed", "initial_buffer_speed", "final_cleanup_speed")
        ):
            raise ValueError("Pump speeds must be between 1 and 40.")
        if settings["syringe_capacity"] < 250:
            raise ValueError(
                "Syringe capacity must be at least 250 µL for the final cleanup stroke."
            )
        if settings["mix_volume"] > settings["syringe_capacity"]:
            raise ValueError("Mix volume cannot exceed syringe capacity.")
        if settings["stock_air_spacer"] >= settings["syringe_capacity"]:
            raise ValueError("Stock air spacer must be smaller than syringe capacity.")
        if settings["mix_line_air_push"] > settings["syringe_capacity"]:
            raise ValueError("Port 4 air push cannot exceed syringe capacity.")
        return settings

    def _build_recipe(self, settings, plan):
        recipe = [
            self._pump_item("INIT", details="Pump: Initialize"),
        ]
        if not settings.get("skip_initial_buffer", False):
            self._append_transfer(
                recipe,
                source_port=settings["buffer_port"],
                destination_port=settings["mix_port"],
                volume_ul=settings["initial_buffer_volume"],
                speed=settings.get("initial_buffer_speed", settings["speed"]),
                capacity=settings["syringe_capacity"],
                label="Initial buffer → mixing tube",
                point="Setup",
            )

        for point in plan:
            point_label = f"{point.target_concentration_um:g} µM"
            if point.stock_added_ul > 1e-9:
                if settings["mix_line_volume"] <= 0:
                    raise ValueError(
                        "Port 4 line volume must be greater than zero when the "
                        "recipe includes a stock addition."
                    )
                self._append_air_assisted_stock_delivery(
                    recipe,
                    stock_volume_ul=point.stock_added_ul,
                    settings=settings,
                    point=point_label,
                )
                for cycle in range(1, settings["mix_cycles"] + 1):
                    recipe.extend(
                        [
                            self._pump_item(
                                "VALVE",
                                port=settings["mix_port"],
                                details=f"Mix cycle {cycle}: valve → mixing tube",
                                point=point_label,
                            ),
                            self._pump_item(
                                "ASPIRATE",
                                speed=settings["speed"],
                                volume=settings["mix_volume"],
                                details=(
                                    f"Mix cycle {cycle}: aspirate "
                                    f"{settings['mix_volume']:.2f} µL"
                                ),
                                point=point_label,
                            ),
                            self._pump_item(
                                "DISPENSE",
                                speed=settings["speed"],
                                volume=settings["mix_volume"],
                                details=(
                                    f"Mix cycle {cycle}: dispense "
                                    f"{settings['mix_volume']:.2f} µL"
                                ),
                                point=point_label,
                            ),
                        ]
                    )
                    self._append_port4_air_flush_and_clear(
                        recipe,
                        settings=settings,
                        point=point_label,
                        label=f"After mix cycle {cycle}",
                    )

            if point.index > 1:
                self._append_transfer(
                    recipe,
                    source_port=settings["flow_port"],
                    destination_port=settings["waste_port"],
                    volume_ul=settings["aliquot_volume"],
                    speed=settings["speed"],
                    capacity=settings["syringe_capacity"],
                    label="Previous flow-cell aliquot → waste",
                    point=point_label,
                )

            self._append_transfer(
                recipe,
                source_port=settings["mix_port"],
                destination_port=settings["flow_port"],
                volume_ul=settings["aliquot_volume"],
                speed=settings["speed"],
                capacity=settings["syringe_capacity"],
                label="Mixing tube → flow cell",
                point=point_label,
            )
            if settings["equilibration"] > 0:
                recipe.append(
                    {
                        "type": "PAUSE",
                        "status": "pending",
                        "details": f"Equilibrate flow cell for {settings['equilibration']:g} s",
                        "pause_seconds": settings["equilibration"],
                        "_point": point_label,
                    }
                )

            # Complete one pass across every configured channel before starting
            # the next replicate (for example: 1,2,3,1,2,3).
            for replicate in range(1, settings["replicates"] + 1):
                for group in self._parameter_groups:
                    for channel in group.get("channels", []):
                        recipe.append(
                            {
                                "type": "SWV",
                                "status": "pending",
                                "details": (
                                    f"{point_label} | optimized | {group['name']} | "
                                    f"MUX ch {channel} | "
                                    f"rep {replicate}/{settings['replicates']}"
                                ),
                                "_point": point_label,
                                "_titration_group": copy.deepcopy(group),
                                "_mux_channel": int(channel),
                                "_swv_source": "optimized",
                            }
                        )
                for group in self._parameter_groups:
                    for channel in group.get("channels", []):
                        channel = int(channel)
                        manual_params = getattr(
                            self, "_manual_channel_params", {}
                        ).get(channel)
                        if manual_params is None:
                            continue
                        manual_group = copy.deepcopy(group)
                        manual_group["name"] = f"Manual ch {channel}"
                        manual_group["params"] = copy.deepcopy(manual_params)
                        manual_group.pop("session_id", None)
                        recipe.append(
                            {
                                "type": "SWV",
                                "status": "pending",
                                "details": (
                                    f"{point_label} | manual | MUX ch {channel} | "
                                    f"rep {replicate}/{settings['replicates']}"
                                ),
                                "_point": point_label,
                                "_titration_group": manual_group,
                                "_mux_channel": channel,
                                "_swv_source": "manual",
                            }
                        )

        if plan and plan[-1].volume_remaining_ul > 1e-9:
            self._append_transfer(
                recipe,
                source_port=settings["mix_port"],
                destination_port=settings["waste_port"],
                volume_ul=plan[-1].volume_remaining_ul,
                speed=settings.get("final_cleanup_speed", settings["speed"]),
                capacity=settings["syringe_capacity"],
                label=(
                    "Final cleanup: remaining mixing-tube fluid "
                    f"({plan[-1].volume_remaining_ul:.3f} µL) → waste"
                ),
                point="Final cleanup",
            )
            self._append_transfer(
                recipe,
                source_port=settings["mix_port"],
                destination_port=settings["waste_port"],
                volume_ul=250.0,
                speed=settings.get("final_cleanup_speed", settings["speed"]),
                capacity=settings["syringe_capacity"],
                label="Final cleanup: extra 250 µL clearing stroke",
                point="Final cleanup",
            )
        return recipe

    def _append_air_assisted_stock_delivery(
        self,
        recipe,
        *,
        stock_volume_ul,
        settings,
        point,
    ):
        """Deliver a stock slug, push it through the port-4 line, then clear air."""
        capacity = settings["syringe_capacity"]
        if stock_volume_ul < capacity:
            spacer = min(
                settings["stock_air_spacer"],
                max(0.0, capacity - stock_volume_ul),
            )
            if spacer > 0:
                recipe.extend(
                    [
                        self._pump_item(
                            "VALVE",
                            port=settings["air_port"],
                            details=(
                                f"Stock delivery: valve → air port "
                                f"{settings['air_port']}"
                            ),
                            point=point,
                        ),
                        self._pump_item(
                            "ASPIRATE",
                            speed=settings["speed"],
                            volume=spacer,
                            details=(
                                f"Stock delivery: aspirate {spacer:.3f} µL air"
                            ),
                            point=point,
                        ),
                    ]
                )
            recipe.extend(
                [
                    self._pump_item(
                        "VALVE",
                        port=settings["stock_port"],
                        details=(
                            f"Stock delivery: valve → 10 mM stock port "
                            f"{settings['stock_port']}"
                        ),
                        point=point,
                    ),
                    self._pump_item(
                        "ASPIRATE",
                        speed=settings["speed"],
                        volume=stock_volume_ul,
                        details=(
                            f"Stock delivery: aspirate {stock_volume_ul:.3f} µL "
                            "10 mM stock"
                        ),
                        point=point,
                    ),
                    self._pump_item(
                        "VALVE",
                        port=settings["mix_port"],
                        details=(
                            f"Stock delivery: valve → mixing tube port "
                            f"{settings['mix_port']}"
                        ),
                        point=point,
                    ),
                    self._pump_item(
                        "DISPENSE",
                        speed=settings["speed"],
                        volume=spacer + stock_volume_ul,
                        details=(
                            f"Stock delivery: dispense "
                            f"{spacer + stock_volume_ul:.3f} µL "
                            f"({spacer:.3f} µL air + "
                            f"{stock_volume_ul:.3f} µL stock)"
                        ),
                        point=point,
                    ),
                ]
            )
        else:
            self._append_transfer(
                recipe,
                source_port=settings["stock_port"],
                destination_port=settings["mix_port"],
                volume_ul=stock_volume_ul,
                speed=settings["speed"],
                capacity=capacity,
                label=f"10 mM stock → mixing tube ({stock_volume_ul:.3f} µL; no spacer)",
                point=point,
            )

        self._append_port4_air_flush_and_clear(
            recipe,
            settings=settings,
            point=point,
            label="Stock line push",
        )

    def _append_port4_air_flush_and_clear(
        self,
        recipe,
        *,
        settings,
        point,
        label,
    ):
        """Push liquid out of the port-4 line with air, then discard line air."""
        recipe.extend(
            [
                self._pump_item(
                    "VALVE",
                    port=settings["air_port"],
                    details=f"{label}: valve → air port {settings['air_port']}",
                    point=point,
                ),
                self._pump_item(
                    "ASPIRATE",
                    speed=settings["speed"],
                    volume=settings["mix_line_air_push"],
                    details=(
                        f"{label}: aspirate "
                        f"{settings['mix_line_air_push']:.3f} µL air"
                    ),
                    point=point,
                ),
                self._pump_item(
                    "VALVE",
                    port=settings["mix_port"],
                    details=(
                        f"{label}: valve → mixing tube port "
                        f"{settings['mix_port']}"
                    ),
                    point=point,
                ),
                self._pump_item(
                    "DISPENSE",
                    speed=settings["speed"],
                    volume=settings["mix_line_air_push"],
                    details=(
                        f"{label}: dispense "
                        f"{settings['mix_line_air_push']:.3f} µL air to port "
                        f"{settings['mix_port']}"
                    ),
                    point=point,
                ),
            ]
        )
        self._append_transfer(
            recipe,
            source_port=settings["mix_port"],
            destination_port=settings["waste_port"],
            volume_ul=settings["mix_line_volume"],
            speed=settings["speed"],
            capacity=settings["syringe_capacity"],
            label=(
                f"{label}: clear {settings['mix_line_volume']:.3f} µL port-4 "
                "air-filled line → waste"
            ),
            point=point,
        )

    def _append_transfer(
        self,
        recipe,
        *,
        source_port,
        destination_port,
        volume_ul,
        speed,
        capacity,
        label,
        point,
    ):
        chunks = split_transfer(volume_ul, capacity)
        for index, chunk in enumerate(chunks, 1):
            suffix = f" (stroke {index}/{len(chunks)})" if len(chunks) > 1 else ""
            recipe.extend(
                [
                    self._pump_item(
                        "VALVE",
                        port=source_port,
                        details=f"{label}: valve → source port {source_port}{suffix}",
                        point=point,
                    ),
                    self._pump_item(
                        "ASPIRATE",
                        speed=speed,
                        volume=chunk,
                        details=f"{label}: aspirate {chunk:.3f} µL{suffix}",
                        point=point,
                    ),
                    self._pump_item(
                        "VALVE",
                        port=destination_port,
                        details=(
                            f"{label}: valve → destination port {destination_port}{suffix}"
                        ),
                        point=point,
                    ),
                    self._pump_item(
                        "DISPENSE",
                        speed=speed,
                        volume=chunk,
                        details=f"{label}: dispense {chunk:.3f} µL{suffix}",
                        point=point,
                    ),
                ]
            )

    @staticmethod
    def _pump_item(action, *, speed=None, volume=None, port=None, details="", point=""):
        params = {}
        if speed is not None:
            params["speed"] = int(speed)
        if volume is not None:
            params["volume"] = float(volume)
        if port is not None:
            params["port"] = int(port)
        return {
            "type": f"PUMP_{action}",
            "status": "pending",
            "details": details,
            "pump_action": {"name": action, "params": params},
            "_point": point,
        }

    def _refresh_recipe_tree(self):
        for row in self._recipe_tree.get_children():
            self._recipe_tree.delete(row)
        for index, item in enumerate(self._recipe, 1):
            self._recipe_tree.insert(
                "",
                "end",
                values=(
                    index,
                    item.get("type", ""),
                    item.get("_point", ""),
                    item.get("details", ""),
                ),
            )

    def _materialize_swv_item(self, item):
        group = item["_titration_group"]
        channel = int(item["_mux_channel"])
        source = item.get("_swv_source", "optimized")
        params = copy.deepcopy(group["params"])
        options = copy.deepcopy(group.get("method_options") or {})
        base_script = build_swv_script(params, options)
        script = wrap_mux(base_script, channel)
        params_for_hash = {
            name: params[name] for name in PARAMETER_ORDER if name in params
        }
        params_for_hash["bandwidth"] = str(options.get("bandwidth", "4k"))
        if group.get("session_id"):
            params_for_hash["bo_session_id"] = group["session_id"]
        saved_path, _saved_name = self._session.registry.save_script(
            "SWV",
            script,
            params=params_for_hash,
            mux_channel=channel,
            note=(
                f"Automated titration | {source} | {group['name']} | MUX ch {channel}"
            ),
        )
        try:
            hash_key = self._session.registry.hash_key_for(saved_path)
        except Exception:
            hash_key = "-"
        return {
            "type": "SWV",
            "script_path": str(saved_path),
            "status": "pending",
            "details": item["details"],
            "method_ref": {
                "hash_key": hash_key,
                "technique": "SWV",
                "params": params_for_hash,
                "mux_channel": channel,
            },
        }

    def _send_to_queue(self):
        if not self._recipe or not self._parameter_groups:
            messagebox.showwarning(
                "Automated Titration",
                "Generate a recipe with Bayesian Optimization parameters first.",
            )
            return
        if not callable(self._send_queue_item) or self._session is None:
            messagebox.showwarning("Automated Titration", "Queue is not available.")
            return
        try:
            queued = 0
            for item in self._recipe:
                if item.get("type") == "SWV":
                    queue_item = self._materialize_swv_item(item)
                else:
                    queue_item = {
                        key: copy.deepcopy(value)
                        for key, value in item.items()
                        if not key.startswith("_")
                    }
                self._send_queue_item(queue_item)
                queued += 1
        except Exception as exc:
            messagebox.showerror("Automated Titration", f"Could not build queue: {exc}")
            return
        self._status_var.set(f"Sent {queued} automated titration steps to the queue.")
        messagebox.showinfo(
            "Automated Titration", f"Added {queued} steps to Queue & Execution."
        )

    def _clear_recipe(self):
        self._recipe = []
        self._plan = []
        self._refresh_calculation_tree()
        self._refresh_recipe_tree()
        self._send_button.configure(state="disabled")
        self._summary_label.configure(text="No recipe generated.")
        self._status_var.set("Recipe preview cleared.")

    @staticmethod
    def _entry_row(parent, row, label, variable):
        ttk.Label(parent, text=label).grid(
            row=row, column=0, sticky="e", padx=(0, 8), pady=4
        )
        ttk.Entry(parent, textvariable=variable).grid(
            row=row, column=1, sticky="ew", pady=4
        )

    @staticmethod
    def _port(value, label):
        port = AutomatedTitrationTab._integer(value, label, minimum=1)
        if port > 9:
            raise ValueError(f"{label} must be between 1 and 9.")
        return port

    @staticmethod
    def _integer(value, label, *, minimum):
        try:
            parsed_float = float(value)
            parsed = int(parsed_float)
        except (TypeError, ValueError) as exc:
            raise ValueError(f"{label} must be an integer.") from exc
        if not math.isfinite(parsed_float) or parsed_float != parsed or parsed < minimum:
            raise ValueError(f"{label} must be an integer of at least {minimum}.")
        return parsed

    @staticmethod
    def _positive(value, label):
        parsed = AutomatedTitrationTab._nonnegative(value, label)
        if parsed <= 0:
            raise ValueError(f"{label} must be greater than zero.")
        return parsed

    @staticmethod
    def _nonnegative(value, label):
        try:
            parsed = float(value)
        except (TypeError, ValueError) as exc:
            raise ValueError(f"{label} must be a number.") from exc
        if not math.isfinite(parsed) or parsed < 0:
            raise ValueError(f"{label} must be zero or greater.")
        return parsed
