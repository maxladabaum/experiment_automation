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
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

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

        self._fig = Figure(figsize=(8, 6), dpi=100)
        self._ax  = self._fig.add_subplot(111)
        self._reset_axes()

        self._canvas = FigureCanvasTkAgg(self._fig, master=self._frame)
        self._canvas.draw()

        toolbar_frame = ttk.Frame(self._frame)
        toolbar_frame.pack(side="top", fill="x")
        toolbar = NavigationToolbar2Tk(self._canvas, toolbar_frame)
        toolbar.update()

        self._canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    def _reset_axes(self, title: str = "Voltammogram"):
        self._ax.set_title(title)
        self._ax.set_xlabel("Potential (V)")
        self._ax.set_ylabel("Current (µA)")
        self._ax.grid(visible=True, which="major", linestyle="-")
        self._ax.grid(visible=True, which="minor", linestyle="--", alpha=0.2)
        self._ax.minorticks_on()

    # ── Static CSV plotting ───────────────────────────────────────────────────

    def _load_and_plot_csv(self):
        path = filedialog.askopenfilename(
            title="Select a measurement CSV",
            filetypes=(("CSV files", "*.csv"), ("All files", "*.*")),
        )
        if path:
            self.plot_data(path)

    def plot_data(self, csv_path, color=None, label=None):
        """Load a CSV and add it to the plot."""
        try:
            df = self._read_csv(csv_path)
        except Exception as exc:
            self._session.log(f"Plot error: failed to read {csv_path}: {exc}")
            messagebox.showerror("Plot Error", f"Failed to read data:\n{exc}")
            return

        pot_col = self._find_column(df, ("Potential (V)",))
        cur_col = self._find_column(
            df, ("Current (uA)", "Current (µA)", "Current (μA)", "Current (�A)")
        )

        if not pot_col or not cur_col:
            msg = "CSV must contain 'Potential (V)' and 'Current (uA)' columns."
            self._session.log(f"Plot error: {msg}")
            messagebox.showerror("Plot Error", msg)
            return

        try:
            if color is None:
                color = next(self._color_cycle)
            if label is None:
                label = Path(csv_path).name
            # Remove existing line with same label (replace on re-plot)
            if label:
                for line in list(self._ax.lines):
                    if line.get_label() == label:
                        line.remove()
            self._ax.plot(df[pot_col], df[cur_col], color=color, label=label)
            self._reset_axes()
            self._ax.legend(loc="best")
            self._canvas.draw()
            self._notebook.select(self._frame)
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

        if color is None:
            color = next(self._color_cycle)
        self._session.last_live_plot_color = color

        if label is None:
            label = title or "Live"
        self._session.last_live_plot_label = label

        self._ax.set_title(title or "Live Voltammogram")
        self._ax.set_xlabel("Potential (V)")
        self._ax.set_ylabel("Current (µA)")
        self._ax.grid(visible=True, which="major", linestyle="-")
        self._ax.grid(visible=True, which="minor", linestyle="--", alpha=0.2)
        self._ax.minorticks_on()
        (self._plot_line,) = self._ax.plot([], [], lw=1, color=color, label=label)
        self._canvas.draw()
        self._notebook.select(self._frame)

        if self._live_job is None:
            self._live_job = self._frame.after(100, self._poll)

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
        try:
            self._live_queue.put_nowait(
                (data_point["potential"], data_point["current"])
            )
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
            self._ax.relim()
            self._ax.autoscale_view()
            if self._session.last_live_plot_label:
                self._ax.legend(loc="best")
            self._canvas.draw_idle()

        self._live_job = self._frame.after(100, self._poll)

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
        for old, new in (("\u03bc", "\u00b5"), ("\u00b5", "mu"), ("\ufffd", "mu")):
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
