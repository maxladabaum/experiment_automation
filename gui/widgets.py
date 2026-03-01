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
