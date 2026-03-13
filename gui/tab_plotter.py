"""
gui/tab_plotter.py — Plotter tab.

Handles:
  - Static CSV loading and rendering
  - Live streaming voltammogram during an active measurement
  - Column normalisation for various CSV encodings / header spellings
"""

import io
import itertools
import queue
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from gui.widgets import AutoScaleToolbar
from core.session import SessionState


class PlotterTab:
    """Manages the 'Plotter' notebook tab.

    Parameters
    ----------
    parent_frame:
        The ``ttk.Frame`` added to the notebook for this tab.
    session:
        Shared :class:`~core.session.SessionState`.
    notebook:
        The parent ``ttk.Notebook`` — used to auto-switch to this tab when
        a live plot starts.
    """

    def __init__(
        self,
        parent_frame: ttk.Frame,
        session: SessionState,
        notebook: ttk.Notebook,
    ):
        self._frame    = parent_frame
        self._session  = session
        self._notebook = notebook

        # Live-plot state
        self._live_queue:  queue.Queue = queue.Queue(maxsize=10_000)
        self._live_x:      list = []
        self._live_y:      list = []
        self._live_active: bool = False
        self._live_job          = None
        self._plot_line         = None
        self._plotted_files     = []

        # Colour cycle (independent from session — plotter owns it)
        _colors = (
            plt.rcParams.get("axes.prop_cycle", plt.cycler(color=["#1f77b4"]))
            .by_key()
            .get("color", ["#1f77b4"])
        )
        self._colors       = _colors
        self._color_cycle  = itertools.cycle(_colors)

        self._build()

    # ── Build ─────────────────────────────────────────────────────────────────

    def _build(self):
        controls = ttk.Frame(self._frame)
        controls.pack(side="top", fill="x", pady=5, padx=5)
        ttk.Button(controls, text="📂 Load and Plot CSV",
                   command=self._load_and_plot_csv).pack(side="left")
        ttk.Button(controls, text="🗑 Clear Plot",
                   command=self.clear_plot).pack(side="left", padx=5)
        ttk.Button(controls, text="Legend",
                   command=self._toggle_legend).pack(side="left")
        self._overlay_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(controls, text="Overlay",
                        variable=self._overlay_var).pack(side="left", padx=6)
        ttk.Label(controls, text="Plot Y:").pack(side="left", padx=(15, 5))
        self._plot_series_var = tk.StringVar(value="Auto")
        self._plot_series_combo = ttk.Combobox(
            controls,
            textvariable=self._plot_series_var,
            state="readonly",
            width=26,
            values=(
                "Auto",
                "Current (uA)",
                "Current Forward (uA)",
                "Current Reverse (uA)",
                "Current Diff (uA)",
            ),
        )
        self._plot_series_combo.pack(side="left")
        self._plot_series_combo.bind("<<ComboboxSelected>>", self._on_series_change)

        self._fig = Figure(figsize=(8, 6), dpi=100)
        self._ax  = self._fig.add_subplot(111)
        self._reset_axes()

        self._canvas = FigureCanvasTkAgg(self._fig, master=self._frame)
        self._canvas.draw()

        toolbar_frame = ttk.Frame(self._frame)
        toolbar_frame.pack(side="top", fill="x")
        self._toolbar = AutoScaleToolbar(
            self._canvas, toolbar_frame,
            get_bounds=self._get_data_bounds,
        )
        self._toolbar.update()

        self._canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self._legend = None
        self._legend_visible = True

    def _reset_axes(self, title: str = "Voltammogram"):
        self._ax.set_title(title)
        self._ax.set_xlabel("Potential (V)")
        self._ax.set_ylabel("Current (uA)")
        self._ax.grid(visible=True, which="major", linestyle="-")
        self._ax.grid(visible=True, which="minor", linestyle="--", alpha=0.2)
        self._ax.minorticks_on()

    # ── Static CSV plotting ───────────────────────────────────────────────────

    def _load_and_plot_csv(self):
        paths = filedialog.askopenfilenames(
            title="Select a measurement CSV",
            filetypes=(("CSV files", "*.csv"), ("All files", "*.*")),
        )
        for path in paths:
            if path:
                self.plot_data(path)

    def plot_data(self, csv_path, color=None, label=None, track: bool = True, allow_overlay: bool = True):
        """Load a CSV and add it to the plot."""
        try:
            df = self._read_csv(csv_path)
        except Exception as exc:
            self._session.log(f"Plot error: failed to read {csv_path}: {exc}")
            messagebox.showerror("Plot Error", f"Failed to read data:\n{exc}")
            return

        pot_col = self._find_column(
            df,
            ("Potential (V)", "Potential(V)", "E (V)", "E(V)", "Potential"),
        )
        if not pot_col:
            msg = "CSV must contain a potential column (e.g. 'Potential (V)')."
            self._session.log(f"Plot error: {msg}")
            messagebox.showerror("Plot Error", msg)
            return

        series_choice = getattr(self, "_plot_series_var", None)
        series_choice = series_choice.get() if series_choice else "Auto"

        try:
            y_values, y_label, series_label = self._resolve_plot_series(df, series_choice)
        except ValueError as exc:
            msg = f"{exc}"
            self._session.log(f"Plot error: {msg}")
            messagebox.showerror("Plot Error", msg)
            return

        try:
            if not allow_overlay or not self._overlay_var.get():
                self._ax.clear()
                self._reset_axes()
            if color is None:
                color = next(self._color_cycle)
            base_label = label or Path(csv_path).name
            label = base_label
            if series_label and series_label not in ("Current (uA)",):
                label = f"{label} | {series_label}"
            # Remove existing line with same label (replace on re-plot)
            if label:
                for line in list(self._ax.lines):
                    if line.get_label() == label:
                        line.remove()
            self._ax.plot(df[pot_col], y_values, color=color, label=label)
            self._reset_axes()
            self._ax.set_ylabel(y_label)
            self._legend = self._ax.legend(loc="best")
            if self._legend is not None:
                self._legend.set_draggable(True)
                self._legend.set_visible(self._legend_visible)
            self._canvas.draw()
            if track:
                self._remember_plotted_file(csv_path, color, base_label)
        except Exception as exc:
            self._session.log(f"Plot render error: {exc}")
            messagebox.showerror("Plot Error", f"Failed to render plot:\n{exc}")
    def clear_plot(self):
        self._ax.clear()
        self._reset_axes()
        self._color_cycle  = itertools.cycle(self._colors)
        self._plot_line    = None
        self._live_x.clear()
        self._live_y.clear()
        self._plotted_files.clear()
        self._legend = None
        self._session.last_live_plot_color = None
        self._session.last_live_plot_label = None
        self._canvas.draw()

    # ── Live plot ─────────────────────────────────────────────────────────────

    def start_live(self, title: str = None, color: str = None, label: str = None):
        """Begin a live streaming plot for the current measurement."""
        self._live_queue  = queue.Queue(maxsize=10_000)
        self._live_x      = []
        self._live_y      = []
        self._live_active = True
        self._plot_line   = None
        self._ax.clear()

        if color is None:
            color = next(self._color_cycle)
        self._session.last_live_plot_color = color

        if label is None:
            label = title or "Live"
        if label:
            label = f"{label} @ {datetime.now().strftime('%H:%M:%S')}"
        self._session.last_live_plot_label = label

        self._reset_axes(title or "Live Voltammogram")
        (self._plot_line,) = self._ax.plot([], [], lw=1, color=color, label=label)
        self._legend = self._ax.legend(loc="best")
        if self._legend is not None:
            self._legend.set_draggable(True)
            self._legend.set_visible(self._legend_visible)
        self._canvas.draw()

        if self._live_job is None:
            self._live_job = self._frame.after(250, self._poll)

    def stop_live(self):
        """Stop the live streaming plot."""
        self._live_active = False
        if self._live_job is not None:
            self._frame.after_cancel(self._live_job)
            self._live_job = None

    def push_live_point(self, data_point: dict):
        """Thread-safe: push a ``{potential, current}`` dict for live rendering."""
        if not self._live_active:
            return
        series_choice = getattr(self, "_plot_series_var", None)
        series_choice = series_choice.get() if series_choice else "Auto"
        value = self._resolve_live_value(data_point, series_choice)
        if value is None:
            return
        try:
            self._live_queue.put_nowait((data_point["potential"], value))
        except (queue.Full, KeyError):
            pass

    def _poll(self):
        if not self._live_active:
            self._live_job = None
            return

        updated = False
        while True:
            try:
                pot, cur = self._live_queue.get_nowait()
            except queue.Empty:
                break
            self._live_x.append(pot)
            self._live_y.append(cur)
            updated = True

        if updated:
            if self._plot_line is None:
                (self._plot_line,) = self._ax.plot(
                    self._live_x, self._live_y, lw=1
                )
            else:
                self._plot_line.set_data(self._live_x, self._live_y)
            # Autoscale less often to avoid UI slowdowns on dense streams.
            n = len(self._live_x)
            if n <= 50 or (n % 200 == 0):
                self._ax.relim()
                self._ax.autoscale_view()
            self._canvas.draw_idle()

        self._live_job = self._frame.after(250, self._poll)

    # ── CSV helpers ───────────────────────────────────────────────────────────

    @staticmethod
    def _read_csv(csv_path) -> pd.DataFrame:
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                return pd.read_csv(csv_path, encoding=enc)
            except UnicodeDecodeError:
                pass
        with open(csv_path, "r", encoding="utf-8", errors="replace") as fh:
            return pd.read_csv(io.StringIO(fh.read()))

    @staticmethod
    def _normalize(header: str) -> str:
        h = header.strip().lower()
        for old, new in (
            ("\u03bc", "\u00b5"),
            ("\u00b5", "mu"),
            ("\ufffd", "mu"),
            ("âµ", "mu"),
            ("ï¿½", "mu"),
        ):
            h = h.replace(old, new)
        return h

    def _find_column(self, df: pd.DataFrame, candidates: tuple):
        for c in candidates:
            if c in df.columns:
                return c
        norm_map = {self._normalize(col): col for col in df.columns}
        for c in candidates:
            nc = self._normalize(c)
            if nc in norm_map:
                return norm_map[nc]
        return None

    def _resolve_plot_series(self, df, series_choice):
        current_col = self._find_column(
            df,
            ("Current (uA)", "Current (muA)"),
        )
        forward_col = self._find_column(
            df,
            (
                "Current Forward (uA)",
                "Current Forward (muA)",
                "Current Fwd (uA)",
                "I_fwd (uA)",
            ),
        )
        reverse_col = self._find_column(
            df,
            (
                "Current Reverse (uA)",
                "Current Reverse (muA)",
                "Current Rev (uA)",
                "I_rev (uA)",
            ),
        )
        diff_col = self._find_column(
            df,
            (
                "Current Diff (uA)",
                "Current Diff (muA)",
                "Current Difference (uA)",
                "Current Difference (muA)",
                "I_diff (uA)",
            ),
        )

        if series_choice == "Current (uA)":
            if not current_col:
                raise ValueError("CSV must contain 'Current (uA)' to plot that series.")
            return df[current_col], "Current (uA)", "Current (uA)"

        if series_choice == "Current Forward (uA)":
            if not forward_col:
                raise ValueError("CSV must contain 'Current Forward (uA)' to plot that series.")
            return df[forward_col], "Current Forward (uA)", "Current Forward (uA)"

        if series_choice == "Current Reverse (uA)":
            if not reverse_col:
                raise ValueError("CSV must contain 'Current Reverse (uA)' to plot that series.")
            return df[reverse_col], "Current Reverse (uA)", "Current Reverse (uA)"

        if series_choice == "Current Diff (uA)":
            if diff_col:
                return df[diff_col], "Current Diff (uA)", "Current Diff (uA)"
            if forward_col and reverse_col:
                return (
                    df[forward_col] - df[reverse_col],
                    "Current Diff (uA)",
                    "Current Diff (uA) (derived)",
                )
            raise ValueError(
                "CSV must contain 'Current Diff (uA)' or both 'Current Forward (uA)' and 'Current Reverse (uA)'."
            )

        # Auto selection: prefer total current, then diff, then derived diff, then forward/reverse.
        if current_col:
            return df[current_col], "Current (uA)", "Current (uA)"
        if diff_col:
            return df[diff_col], "Current Diff (uA)", "Current Diff (uA)"
        if forward_col and reverse_col:
            return (
                df[forward_col] - df[reverse_col],
                "Current Diff (uA)",
                "Current Diff (uA) (derived)",
            )
        if forward_col:
            return df[forward_col], "Current Forward (uA)", "Current Forward (uA)"
        if reverse_col:
            return df[reverse_col], "Current Reverse (uA)", "Current Reverse (uA)"

        raise ValueError(
            "CSV must contain 'Current (uA)' or SWV columns "
            "('Current Forward (uA)', 'Current Reverse (uA)', or 'Current Diff (uA)')."
        )

    def _on_series_change(self, _event=None):
        y_label = self._plot_series_var.get()
        if y_label in ("Auto", ""):
            y_label = "Current (uA)"
        self._ax.set_ylabel(y_label)
        if self._live_active:
            self._canvas.draw_idle()
            return
        if self._plotted_files:
            self._replot_loaded_csvs()

    def _toggle_legend(self):
        if self._legend is None:
            return
        self._legend_visible = not self._legend_visible
        self._legend.set_visible(self._legend_visible)
        self._canvas.draw_idle()

    def _remember_plotted_file(self, csv_path, color, base_label):
        path_str = str(csv_path)
        for entry in self._plotted_files:
            if entry["path"] == path_str:
                entry["color"] = color
                entry["label"] = base_label
                return
        self._plotted_files.append({"path": path_str, "color": color, "label": base_label})

    def _replot_loaded_csvs(self):
        files = list(self._plotted_files)
        self._ax.clear()
        self._reset_axes()
        for entry in files:
            self.plot_data(entry["path"], color=entry["color"], label=entry["label"], track=False)
        if self._legend is not None:
            self._legend.set_draggable(True)

    @staticmethod
    def _resolve_live_value(data_point: dict, series_choice: str):
        """Pick the correct live Y value based on the current series selection."""
        if series_choice == "Current Forward (uA)":
            return data_point.get("current_forward")
        if series_choice == "Current Reverse (uA)":
            return data_point.get("current_reverse")
        if series_choice == "Current Diff (uA)":
            if "current_diff" in data_point:
                return data_point.get("current_diff")
            if "current_forward" in data_point and "current_reverse" in data_point:
                return data_point["current_forward"] - data_point["current_reverse"]
            return data_point.get("current")

        # Auto or "Current (uA)"
        return data_point.get("current")
    def _get_data_bounds(self):
        """Return (x_min, x_max, y_min, y_max) across all plotted lines, or None."""
        x_min = x_max = y_min = y_max = None
        for line in self._ax.lines:
            try:
                xs = line.get_xdata(orig=False)
                ys = line.get_ydata(orig=False)
            except Exception:
                continue
            for v in xs:
                try:
                    fv = float(v)
                    x_min = fv if x_min is None else min(x_min, fv)
                    x_max = fv if x_max is None else max(x_max, fv)
                except Exception:
                    pass
            for v in ys:
                try:
                    fv = float(v)
                    y_min = fv if y_min is None else min(y_min, fv)
                    y_max = fv if y_max is None else max(y_max, fv)
                except Exception:
                    pass
        if None in (x_min, x_max, y_min, y_max):
            return None
        # Add 5 % margin
        xr = (x_max - x_min) or max(abs(x_min), 1.0) * 0.1
        yr = (y_max - y_min) or max(abs(y_min), 1.0) * 0.1
        return (x_min - xr*0.05, x_max + xr*0.05,
                y_min - yr*0.05, y_max + yr*0.05)
