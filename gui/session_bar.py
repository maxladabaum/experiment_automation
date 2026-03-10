"""
gui/session_bar.py - Bottom-of-window Session & Experiment control bar.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from core.session_manager import SessionManager


class SessionBar:
    """Bottom-bar widget that drives a SessionManager."""

    def __init__(
        self,
        root: tk.Misc,
        session_manager: SessionManager,
        on_start_session=None,
    ):
        self._root = root
        self._mgr = session_manager
        self._on_start_session_cb = on_start_session

        self._session_name_var = tk.StringVar()
        self._session_user_var = tk.StringVar()
        self._session_chip_id_var = tk.StringVar()
        self._session_notes_var = tk.StringVar()

        self._experiment_name_var = tk.StringVar()
        self._experiment_notes_var = tk.StringVar()

        self._build()

    def _build(self):
        outer = ttk.Frame(self._root)
        outer.pack(fill="both", expand=True, padx=5, pady=(0, 5))

        ttk.Label(
            outer,
            textvariable=self._mgr.status_var,
            foreground="blue",
            anchor="w",
        ).pack(side="bottom", fill="x", padx=4, pady=(2, 0))

        row = ttk.Frame(outer)
        row.pack(fill="x", expand=True)

        sess = ttk.LabelFrame(row, text="Session")
        sess.pack(side="left", fill="x", expand=True, padx=(0, 4))
        sess.columnconfigure(1, weight=1)
        sess.columnconfigure(3, weight=1)

        ttk.Label(sess, text="Session Name:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(sess, textvariable=self._session_name_var, width=20).grid(
            row=0, column=1, sticky="we", padx=5, pady=2
        )
        ttk.Button(sess, text="Start Session", command=self._on_start_session).grid(
            row=0, column=2, padx=4, pady=2
        )
        ttk.Button(sess, text="End Session", command=self._on_end_session).grid(
            row=0, column=3, padx=4, pady=2
        )
        ttk.Button(sess, text="Choose Session", command=self._on_choose_session).grid(
            row=0, column=4, padx=4, pady=2
        )

        ttk.Label(sess, text="User:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(sess, textvariable=self._session_user_var, width=14).grid(
            row=1, column=1, sticky="we", padx=5, pady=2
        )
        ttk.Label(sess, text="Chip ID:").grid(row=1, column=2, sticky="w", padx=5, pady=2)
        ttk.Entry(sess, textvariable=self._session_chip_id_var, width=14).grid(
            row=1, column=3, sticky="we", padx=5, pady=2
        )

        ttk.Label(sess, text="Notes:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(sess, textvariable=self._session_notes_var, width=34).grid(
            row=2, column=1, columnspan=3, sticky="we", padx=5, pady=2
        )

        ttk.Button(sess, text="Update Session Metadata", command=self._on_update_session).grid(
            row=3, column=0, columnspan=2, sticky="w", padx=5, pady=2
        )

        exp = ttk.LabelFrame(row, text="Experiment")
        exp.pack(side="right", fill="x", expand=True, padx=(4, 0))
        exp.columnconfigure(1, weight=1)

        ttk.Label(exp, text="Experiment Name:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(exp, textvariable=self._experiment_name_var, width=22).grid(
            row=0, column=1, sticky="we", padx=5, pady=2
        )

        ttk.Label(exp, text="Notes:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        ttk.Entry(exp, textvariable=self._experiment_notes_var, width=26).grid(
            row=1, column=1, sticky="we", padx=5, pady=2
        )

        btn_row = ttk.Frame(exp)
        btn_row.grid(row=2, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        ttk.Button(btn_row, text="Start Experiment", command=self._on_start_experiment).pack(
            side="left", padx=(0, 4)
        )
        ttk.Button(btn_row, text="End Experiment", command=self._on_end_experiment).pack(side="left")

    def _on_start_session(self):
        started = self._mgr.start_session(
            name=self._session_name_var.get(),
            user=self._session_user_var.get(),
            chip_id=self._session_chip_id_var.get(),
            notes=self._session_notes_var.get(),
        )
        if started and self._on_start_session_cb:
            self._on_start_session_cb()

    def _on_end_session(self):
        self._mgr.end_session()

    def _on_choose_session(self):
        path = filedialog.askdirectory(
            initialdir=str(self._mgr.data_root),
            title="Choose Existing Session Folder",
        )
        if not path:
            return
        opened = self._mgr.open_session(path)
        if opened:
            self._apply_session_metadata(self._mgr.session_metadata())
            if self._on_start_session_cb:
                self._on_start_session_cb()
        else:
            messagebox.showerror("Session Not Opened", "Could not open the selected session.")

    def _apply_session_metadata(self, data: dict):
        if not data:
            return
        self._session_name_var.set(data.get("session_name", ""))
        self._session_user_var.set(data.get("user", ""))
        self._session_chip_id_var.set(data.get("chip_id", ""))
        self._session_notes_var.set(data.get("notes", ""))

    def _on_update_session(self):
        self._mgr.update_session_metadata(
            user=self._session_user_var.get(),
            chip_id=self._session_chip_id_var.get(),
            notes=self._session_notes_var.get(),
        )

    def _on_start_experiment(self):
        self._mgr.start_experiment(
            name=self._experiment_name_var.get(),
            notes=self._experiment_notes_var.get(),
        )

    def _on_end_experiment(self):
        self._mgr.end_experiment()
