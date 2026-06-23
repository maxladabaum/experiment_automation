"""
gui/widgets.py — Reusable custom Tkinter / Matplotlib widgets.

Currently contains:

_AutoScaleToolbar
    A subclass of NavigationToolbar2Tk whose "Home" button restores the
    data bounds (with a small margin) rather than matplotlib's default
    full reset.  Left-click-only zoom is also enforced so that right-click
    can be reused for panning by the plotter.

Usage
-----
    from gui.widgets import AutoScaleToolbar

    toolbar = AutoScaleToolbar(
        canvas,
        toolbar_frame,
        get_bounds=plotter_tab._get_plot_data_bounds,
    )
"""

import tkinter as tk
from tkinter import ttk

from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk


class AutoScaleToolbar(NavigationToolbar2Tk):
    """Matplotlib navigation toolbar with smarter Home behaviour.

    Parameters
    ----------
    canvas:
        The ``FigureCanvasTkAgg`` instance.
    window:
        The Tkinter frame that hosts the toolbar.
    get_bounds:
        Optional callable ``() → (x_min, x_max, y_min, y_max) | None``.
        When supplied, pressing *Home* restores exactly these bounds
        (with a 5 % margin already baked in by the plotter).  If it
        returns ``None``, the default autoscale behaviour is used.
    """

    def __init__(self, canvas, window, *, get_bounds=None):
        self._get_bounds = get_bounds
        super().__init__(canvas, window)

    # ── Restrict zoom to left-click only ──────────────────────────────────────

    def press_zoom(self, event):
        if event.button != 1:
            return
        return super().press_zoom(event)

    def release_zoom(self, event):
        if event.button != 1:
            return
        return super().release_zoom(event)

    # ── Smart Home ────────────────────────────────────────────────────────────

    def home(self, *args):
        axes = self.canvas.figure.axes
        if not axes:
            return
        ax = axes[0]

        bounds = None
        if self._get_bounds is not None:
            try:
                bounds = self._get_bounds()
            except Exception:
                bounds = None

        if bounds is not None:
            x_min, x_max, y_min, y_max = bounds
            ax.set_xlim(x_min, x_max)
            ax.set_ylim(y_min, y_max)
            self.canvas.draw_idle()
            return

        # Fallback: standard autoscale with a small margin
        ax.relim()
        ax.autoscale_view(tight=True)
        ax.margins(x=0.05, y=0.05)
        self.canvas.draw_idle()


class ScrollableFrame(ttk.Frame):
    """A reusable frame with vertical and horizontal overflow handling."""

    def __init__(self, parent, *, min_width=0, fit_width=True, **kwargs):
        super().__init__(parent, **kwargs)
        self._min_width = int(min_width or 0)
        self._fit_width = bool(fit_width)

        self._canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0)
        self._vscroll = ttk.Scrollbar(self, orient="vertical", command=self._canvas.yview)
        self._hscroll = ttk.Scrollbar(self, orient="horizontal", command=self._canvas.xview)
        self.content = ttk.Frame(self._canvas)
        self._window = self._canvas.create_window((0, 0), window=self.content, anchor="nw")

        self._canvas.configure(yscrollcommand=self._vscroll.set, xscrollcommand=self._hscroll.set)
        self._canvas.grid(row=0, column=0, sticky="nsew")
        self._vscroll.grid(row=0, column=1, sticky="ns")
        self._hscroll.grid(row=1, column=0, sticky="ew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        self.content.bind("<Configure>", self._on_content_configure)
        self._canvas.bind("<Configure>", self._on_canvas_configure)
        self._bind_mousewheel(self._canvas)
        self._bind_mousewheel(self.content)

    def _on_content_configure(self, _event=None):
        self._canvas.configure(scrollregion=self._canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        if not self._fit_width:
            return
        width = max(int(event.width), self._min_width)
        self._canvas.itemconfigure(self._window, width=width)

    def _bind_mousewheel(self, widget):
        widget.bind("<Enter>", lambda _e: self._canvas.bind_all("<MouseWheel>", self._on_mousewheel))
        widget.bind("<Leave>", lambda _e: self._canvas.unbind_all("<MouseWheel>"))

    def _on_mousewheel(self, event):
        self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


class FlowFrame(ttk.Frame):
    """A simple wrapping container for rows of buttons and compact controls."""

    def __init__(self, parent, *, hgap=4, vgap=4, **kwargs):
        super().__init__(parent, **kwargs)
        self._hgap = hgap
        self._vgap = vgap
        self.bind("<Configure>", lambda _e: self._layout())

    def add(self, widget):
        widget.place(in_=self)
        widget.bind("<Configure>", lambda _e: self._layout())
        self.after_idle(self._layout)
        return widget

    def separator(self):
        sep = ttk.Separator(self, orient="vertical")
        return self.add(sep)

    def _layout(self):
        width = max(1, self.winfo_width())
        x = 0
        y = 0
        row_h = 0
        for child in self.winfo_children():
            req_w = child.winfo_reqwidth()
            req_h = child.winfo_reqheight()
            if x and x + req_w > width:
                x = 0
                y += row_h + self._vgap
                row_h = 0
            child.place(x=x, y=y, width=req_w, height=req_h)
            x += req_w + self._hgap
            row_h = max(row_h, req_h)
        self.configure(height=y + row_h)
