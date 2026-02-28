"""
core/session.py — Session-wide shared state.

:class:`SessionState` is the single source of truth for all mutable runtime
data (queue, measurement counter, running flag, current runner …).  It is
created once in ``gui/app.py`` and injected into every tab that needs it.

This means tabs never import each other — they communicate exclusively through
the shared ``SessionState`` object.
"""

import itertools
from pathlib import Path
from typing import Callable, Dict, List, Optional

import matplotlib.pyplot as plt

from .method_registry import MethodRegistry
from .runner import SerialMeasurementRunner


class SessionState:
    """Holds all mutable state for one application session.

    Parameters
    ----------
    log_callback:
        Callable ``(str) → None`` wired to the GUI log panel.
    status_callback:
        Callable ``(str) → None`` wired to the GUI status bar.
    """

    def __init__(
        self,
        log_callback:    Callable[[str], None] = print,
        status_callback: Callable[[str], None] = print,
    ):
        self._log    = log_callback
        self._status = status_callback

        # ── Queue ─────────────────────────────────────────────────────────────
        self.measurement_queue: List[dict] = []
        self.is_running  = False
        self.current_runner: Optional[SerialMeasurementRunner] = None

        # ── Measurement tagging ───────────────────────────────────────────────
        self.measurement_counter = 0

        # ── Script registry (deduplication) ───────────────────────────────────
        self.registry = MethodRegistry(log_callback=log_callback)

        # ── Queue clipboard (copy / paste) ────────────────────────────────────
        self.queue_clipboard: List[dict] = []

        # ── Live plot helpers ─────────────────────────────────────────────────
        _colors = (
            plt.rcParams.get("axes.prop_cycle", plt.cycler(color=["#1f77b4"]))
            .by_key()
            .get("color", ["#1f77b4"])
        )
        self._plot_color_cycle = itertools.cycle(_colors)
        self.last_live_plot_color: Optional[str]  = None
        self.last_live_plot_label: Optional[str]  = None

    # ── Measurement tag ───────────────────────────────────────────────────────

    def next_meas_tag(self) -> str:
        """Increment counter and return the next sequential measurement tag.

        Format: ``meas_NNN`` (grows beyond 999 automatically).
        """
        self.measurement_counter += 1
        return f"meas_{self.measurement_counter:03d}"

    def reset_counter(self):
        """Reset measurement counter to zero."""
        self.measurement_counter = 0
        self._log("[Session] Measurement counter reset to 0.")

    # ── Plot colour ───────────────────────────────────────────────────────────

    def next_plot_color(self) -> str:
        """Return the next colour from the matplotlib colour cycle."""
        color = next(self._plot_color_cycle)
        self.last_live_plot_color = color
        return color

    # ── Convenience passthrough ───────────────────────────────────────────────

    def log(self, msg: str):
        self._log(msg)

    def set_status(self, msg: str):
        self._status(msg)

    def stop_current_runner(self):
        """Signal the active runner (if any) to stop."""
        if self.current_runner is not None:
            self.current_runner.stop()
