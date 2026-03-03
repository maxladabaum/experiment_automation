"""
gui/tab_script.py — Script Preview tab.

Shows the last generated MethodSCRIPT.
"""

import tkinter as tk
from tkinter import ttk


class ScriptTab:
    """Manages the 'Script Preview' notebook tab.

    Parameters
    ----------
    parent_frame:
        The ``ttk.Frame`` added to the notebook for this tab.
    """

    def __init__(self, parent_frame: ttk.Frame):
        self._frame = parent_frame
        self._build()

    # ── Build ─────────────────────────────────────────────────────────────────

    def _build(self):
        text_frame = ttk.Frame(self._frame)
        text_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self._text = tk.Text(
            text_frame,
            wrap="none",
            font=("Courier", 11),
        )
        self._text.pack(fill="both", expand=True)

    # ── Public API ────────────────────────────────────────────────────────────

    def update(self, script: str):
        """Replace the displayed script with new content."""
        self._text.delete("1.0", tk.END)
        self._text.insert("1.0", script)

    def get(self) -> str:
        """Return the current text content (useful if user hand-edited it)."""
        return self._text.get("1.0", tk.END)
