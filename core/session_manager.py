"""
core/session_manager.py — Session and Experiment lifecycle management.

Handles the two-level folder hierarchy:
    measurement_data/
        <session_name>_<timestamp>/          ← Session
            session_metadata.json
            session_log.txt
            <experiment_name>_<timestamp>/   ← Experiment
                experiment_metadata.json
                <csv files go here>

Public API
----------
SessionManager.start_session(name, user, chip_id, notes)
SessionManager.end_session()
SessionManager.update_session_metadata(user, chip_id, notes)
SessionManager.start_experiment(name, notes)
SessionManager.end_experiment()
SessionManager.require_session()    → session_path | None  (shows error dialog)
SessionManager.require_experiment() → data_folder  | None  (shows error dialog)
SessionManager.log(message)         → writes to session_log.txt + calls log_callback
SessionManager.status_label_var     → tk.StringVar  "Session: X | Experiment: Y"
"""

import json
from datetime import datetime
from pathlib import Path
from tkinter import messagebox
import tkinter as tk
from typing import Optional

from config import DATA_DIR


class SessionManager:
    """Manages session and experiment folders, metadata, and the session log.

    Parameters
    ----------
    log_callback:
        Callable ``(str) → None`` — called for every log line so it
        appears in the GUI log panel as well as the on-disk session log.
    data_root:
        Base directory for all session folders.  Defaults to ``DATA_DIR``
        from config.
    """

    def __init__(self, log_callback=None, data_root: Path = None):
        self._log_cb   = log_callback or (lambda m: print(m))
        self._data_root = Path(data_root) if data_root else Path(DATA_DIR)
        self._data_root.mkdir(exist_ok=True)

        # ── State ──────────────────────────────────────────────────────────────
        self.current_session_path:    Optional[Path] = None
        self.current_experiment_path: Optional[Path] = None

        self._session_metadata_path:    Optional[Path] = None
        self._experiment_metadata_path: Optional[Path] = None
        self._session_log_path:         Optional[Path] = None
        self._session_started_at:  Optional[str] = None
        self._experiment_started_at: Optional[str] = None


        # Raw field values (kept so update_session_metadata can re-read them)
        self._session_raw: dict = {}

        # ── Tkinter observable ─────────────────────────────────────────────────
        self.status_var = tk.StringVar(
            value="Session: (none)  |  Experiment: (none)"
        )

    # ── Internal helpers ───────────────────────────────────────────────────────

    @staticmethod
    def _sanitize(value: str, fallback: str) -> str:
        """Strip invalid filesystem characters and return a safe folder name."""
        cleaned = (value or "").strip()
        if not cleaned:
            return fallback
        for ch in '<>:"/\\|?*':
            cleaned = cleaned.replace(ch, "_")
        cleaned = cleaned.strip().strip(".")
        return cleaned or fallback

    @staticmethod
    def _unique_path(base: Path) -> Path:
        """Return *base* if it doesn't exist, else *base_02*, *base_03*, …"""
        if not base.exists():
            return base
        for idx in range(2, 1000):
            candidate = Path(f"{base}_{idx:02d}")
            if not candidate.exists():
                return candidate
        return Path(f"{base}_{datetime.now().strftime('%Y%m%d_%H%M%S')}")

    def _write_json(self, path: Path, payload: dict):
        try:
            with open(path, "w", encoding="utf-8") as fh:
                json.dump(payload, fh, indent=2)
        except OSError as exc:
            messagebox.showerror("Save Failed", f"Could not write metadata:\n{exc}")

    def _update_status_var(self):
        s = self.current_session_path.name    if self.current_session_path    else "(none)"
        e = self.current_experiment_path.name if self.current_experiment_path else "(none)"
        self.status_var.set(f"Session: {s}  |  Experiment: {e}")

    # ── Session log ────────────────────────────────────────────────────────────

    def log(self, message: str):
        """Write *message* to the on-disk session log and forward to the GUI."""
        if self._session_log_path:
            ts   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            line = f"[{ts}] {message}"
            try:
                with open(self._session_log_path, "a", encoding="utf-8") as fh:
                    fh.write(line + "\n")
            except OSError:
                pass
        self._log_cb(message)

    # ── Session lifecycle ──────────────────────────────────────────────────────

    def start_session(
        self,
        name:    str,
        user:    str,
        chip_id: str,
        notes:   str = "",
    ) -> bool:
        """Create the session folder and write initial metadata.

        Returns True on success, False if validation fails.
        """
        # Validate required fields
        missing = [
            label for label, val in [
                ("Session Name", name),
                ("User",         user),
                ("Chip ID",      chip_id),
            ]
            if not str(val).strip()
        ]
        if missing:
            messagebox.showerror(
                "Missing Metadata",
                "Fill out all session fields before starting:\n"
                + ", ".join(missing),
            )
            return False

        # If a session is already open, close it first
        if self.current_session_path:
            self.end_session()

        ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback = f"session_{ts}"
        base_name  = self._sanitize(name, fallback)
        folder_name = f"{base_name}_{ts}"
        session_path = self._unique_path(self._data_root / folder_name)
        session_path.mkdir(parents=True, exist_ok=True)

        self.current_session_path      = session_path
        self.current_experiment_path   = None
        self._session_metadata_path    = session_path / "session_metadata.json"
        self._session_log_path         = session_path / "session_log.txt"
        self._session_started_at       = datetime.now().isoformat(timespec="seconds")
        self._experiment_started_at    = None
        self._experiment_metadata_path = None

        self._session_raw = {
            "session_name": name.strip() or folder_name,
            "session_folder": session_path.name,
            "user": user.strip(),
            "chip_id": chip_id.strip(),
            "notes": notes.strip(),
            "started_at": self._session_started_at,
            "ended_at": None,
        }
        self._write_json(self._session_metadata_path, self._session_raw)
        self._update_status_var()
        self.log(f"Session started: {session_path.name}")
        return True

    def end_session(self):
        """Mark the session as ended and update metadata."""
        if not self.current_session_path:
            return
        if self.current_experiment_path:
            self.end_experiment()
        self._session_raw["ended_at"] = datetime.now().isoformat(timespec="seconds")
        if self._session_metadata_path:
            self._write_json(self._session_metadata_path, self._session_raw)
        self.log(f"Session ended: {self.current_session_path.name}")
        self.current_session_path      = None
        self.current_experiment_path   = None
        self._session_metadata_path    = None
        self._session_log_path         = None
        self._session_started_at       = None
        self._update_status_var()

    def update_session_metadata(
        self,
        user:    str,
        chip_id: str,
        notes:   str,
    ):
        """Update mutable session metadata fields without closing the session."""
        if not self._session_metadata_path:
            messagebox.showwarning("No Session", "Start a session first.")
            return
        self._session_raw.update({
            "user":    user.strip(),
            "chip_id": chip_id.strip(),
            "notes":   notes.strip(),
        })
        self._write_json(self._session_metadata_path, self._session_raw)
        self.log("Session metadata updated.")

    # ── Experiment lifecycle ───────────────────────────────────────────────────

    def start_experiment(self, name: str, notes: str = "") -> bool:
        """Create an experiment subfolder inside the current session.

        Returns True on success.
        """
        if not self.current_session_path:
            messagebox.showerror(
                "No Session",
                "Start a session before starting an experiment.",
            )
            return False

        if self.current_experiment_path:
            self.end_experiment()

        ts          = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback    = f"experiment_{ts}"
        base_name   = self._sanitize(name, fallback) if name.strip() else fallback
        folder_name = f"{base_name}_{ts}"
        exp_path    = self._unique_path(self.current_session_path / folder_name)
        exp_path.mkdir(parents=True, exist_ok=True)

        self.current_experiment_path   = exp_path
        self._experiment_metadata_path = exp_path / "experiment_metadata.json"
        self._experiment_started_at    = datetime.now().isoformat(timespec="seconds")

        self._write_json(self._experiment_metadata_path, {
            "experiment_name":   name.strip() or folder_name,
            "experiment_folder": exp_path.name,
            "notes":             notes.strip(),
            "started_at":        self._experiment_started_at,
            "ended_at":          None,
        })
        self._update_status_var()
        self.log(f"Experiment started: {exp_path.name}")
        return True

    def end_experiment(self):
        """Mark the experiment as ended."""
        if not self.current_experiment_path:
            return
        if self._experiment_metadata_path and self._experiment_metadata_path.exists():
            try:
                with open(self._experiment_metadata_path, "r", encoding="utf-8") as fh:
                    data = json.load(fh)
                data["ended_at"] = datetime.now().isoformat(timespec="seconds")
                self._write_json(self._experiment_metadata_path, data)
            except Exception:
                pass
        self.log(f"Experiment ended: {self.current_experiment_path.name}")
        self.current_experiment_path   = None
        self._experiment_metadata_path = None
        self._experiment_started_at    = None
        self._update_status_var()

    # ── Guard helpers ─────────────────────────────────────────────────────────

    def require_session(self) -> Optional[Path]:
        """Return the session path, or show an error dialog and return None."""
        if not self.current_session_path:
            messagebox.showerror(
                "No Active Session",
                "Please start a session before performing this action.",
            )
            return None
        return self.current_session_path

    def require_experiment(self) -> Optional[Path]:
        """Return the experiment data folder, or show an error dialog and return None."""
        if not self.current_session_path:
            messagebox.showerror(
                "No Active Session",
                "Please start a session and an experiment first.",
            )
            return None
        if not self.current_experiment_path:
            messagebox.showerror(
                "No Active Experiment",
                "Please start an experiment before running measurements.",
            )
            return None
        return self.current_experiment_path

    # ── Convenience properties ────────────────────────────────────────────────

    @property
    def has_session(self) -> bool:
        return self.current_session_path is not None

    @property
    def has_experiment(self) -> bool:
        return self.current_experiment_path is not None
