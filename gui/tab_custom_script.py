"""
gui/tab_custom_script.py — Custom MethodSCRIPT file loader panel.

This is not a standalone notebook tab; it is a *panel* rendered inside
the existing ``tab_method.py`` params frame whenever the user clicks
"Custom Script (File)" in the technique selector.

The panel provides:
  - File browser to load any .ms / .txt MethodSCRIPT file
  - MUX16 channel field (same as CV / SWV)
  - "Run Now" and "Add to Queue" buttons
  - Automatic detection of existing MUX headers in the loaded file with
    a confirmation dialog before overriding

Usage (called from MethodTab._show_custom_params)
-------------------------------------------------
    from gui.tab_custom_script import CustomScriptPanel

    panel = CustomScriptPanel(
        params_frame      = self._params_frame,
        session           = self._session,
        on_run_now        = self._run_now,
        on_add_to_queue   = self._add_to_queue,
        on_script_preview = self._script_preview,
        save_script_fn    = self._session.registry.save_script,
        wrap_mux_fn       = self._wrap_mux,
        parse_mux_fn      = self._get_mux_channels,
    )
"""

from pathlib import Path
import hashlib
import re
from typing import Optional
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk

from core.session import SessionState


class CustomScriptPanel:
    """Renders custom-script controls inside *params_frame* and handles
    all interactions for loading, MUX-wrapping, running, and queuing a
    user-supplied MethodSCRIPT file.

    Parameters
    ----------
    params_frame:
        The ``ttk.LabelFrame`` owned by MethodTab that this panel populates.
    session:
        Shared :class:`~core.session.SessionState`.
    on_run_now:
        Callable ``(technique, script, mux_channel) → None``  — same
        signature used by CV/SWV so app.py can dispatch it uniformly.
    on_add_to_queue:
        Callable ``(item: dict) → None`` — appends a queue item.
    on_script_preview:
        Callable ``(script: str) → None`` — pushes text to Script tab.
    save_script_fn:
        Callable ``(technique, script, params, mux_channel=None, note=None)
        → (Path, str)`` — e.g. ``session.registry.save_script``.
    wrap_mux_fn:
        Callable ``(base_script: str, channel: int) → str`` — prepends
        the MUX header for *channel*.
    parse_mux_fn:
        Callable ``(param_dict) → list[int] | None`` — returns parsed
        channel list from the mux_channel entry, or None on error.
    """

    def __init__(
        self,
        params_frame:      ttk.LabelFrame,
        session:           SessionState,
        on_run_now,
        on_add_to_queue,
        on_script_preview,
        save_script_fn,
        wrap_mux_fn,
        parse_mux_fn,
    ):
        self._frame           = params_frame
        self._session         = session
        self._run_now_cb      = on_run_now
        self._add_to_queue_cb = on_add_to_queue
        self._preview_cb      = on_script_preview
        self._save_script     = save_script_fn
        self._wrap_mux        = wrap_mux_fn
        self._parse_mux       = parse_mux_fn

        # State kept across calls (survives re-renders of the param panel)
        self.script_text: str             = ""
        self.script_path: Optional[str]   = None
        self.script_name: Optional[str]   = None
        self._has_mux_header: bool   = False

        # Param dict (mimics the cv/swv_params dict so _parse_mux works)
        self.params: dict = {}

        self._build()

    # ── Build ──────────────────────────────────────────────────────────────────

    def _build(self):
        path_var = tk.StringVar(value=self.script_path or "")
        self.params["script_path_var"] = path_var

        # Row 0 — file path + browse
        ttk.Label(self._frame, text="MethodSCRIPT file:").grid(
            row=0, column=0, sticky="w", pady=2)
        path_entry = ttk.Entry(self._frame, textvariable=path_var, width=46)
        path_entry.grid(row=0, column=1, pady=2, sticky="w")
        ttk.Button(self._frame, text="Browse…",
                   command=self._browse).grid(row=0, column=2, padx=5)

        # Row 1 — MUX channel
        ttk.Label(
            self._frame,
            text="MUX16 Channels (1-16, 0=off, e.g. 1-3,7-9):",
        ).grid(row=1, column=0, sticky="w", pady=2)
        mux_entry = ttk.Entry(self._frame, width=20)
        mux_entry.insert(0, "0")
        mux_entry.grid(row=1, column=1, sticky="w", pady=2)
        self.params["mux_channel"] = mux_entry

        # Row 2 — MUX-header-detected notice (hidden by default)
        self._mux_notice = ttk.Label(
            self._frame,
            text="ℹ Script already contains a MUX header.",
            foreground="orange",
        )
        # only shown when a script with an existing header is loaded

        # Row 3 — optional library note
        ttk.Label(self._frame, text="Library note (optional):").grid(
            row=3, column=0, sticky="w", pady=2)
        self._note_var = tk.StringVar(value="")
        ttk.Entry(self._frame, textvariable=self._note_var, width=40).grid(
            row=3, column=1, sticky="w", pady=2)

        # Row 3 — action buttons
        btn_frame = ttk.Frame(self._frame)
        btn_frame.grid(row=4, column=0, columnspan=3, pady=20)
        ttk.Button(btn_frame, text="Run Now",
                   command=self._run_now).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Add to Queue",
                   command=self._add_to_queue).pack(side="left", padx=5)

    # ── File loading ───────────────────────────────────────────────────────────

    def _browse(self):
        file_path = filedialog.askopenfilename(
            title="Select a MethodSCRIPT file",
            filetypes=(
                ("MethodSCRIPT", "*.ms;*.txt"),
                ("All Files", "*.*"),
            ),
        )
        if not file_path:
            return
        try:
            with open(file_path, "r", encoding="utf-8", errors="replace") as fh:
                script = fh.read()
        except OSError as exc:
            messagebox.showerror("Load Failed", f"Could not read script:\n{exc}")
            return

        self.script_text = script
        self.script_path = file_path
        self.script_name = Path(file_path).name
        self._has_mux_header = self._detect_mux_header(script)

        if self.params.get("script_path_var"):
            self.params["script_path_var"].set(file_path)

        if self._has_mux_header:
            self._mux_notice.grid(row=2, column=0, columnspan=3, sticky="w", padx=4)
        else:
            try:
                self._mux_notice.grid_remove()
            except Exception:
                pass

        self._preview_cb(script)

    # ── MUX header helpers ─────────────────────────────────────────────────────

    @staticmethod
    def _detect_mux_header(script: str) -> bool:
        """Return True if the script already has a set_gpio_cfg / set_gpio header."""
        cfg_found = False
        for line in script.splitlines():
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue
            if stripped == "set_gpio_cfg 0x3FFi 1" and not cfg_found:
                cfg_found = True
                continue
            if cfg_found and stripped.startswith("set_gpio ") \
                    and not stripped.startswith("set_gpio_cfg"):
                return True
        return False

    @staticmethod
    def _strip_mux_header(script: str) -> str:
        """Remove the first set_gpio_cfg … set_gpio pair from *script*."""
        lines = script.splitlines()
        cfg_idx = gpio_idx = None
        for idx, line in enumerate(lines):
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue
            if stripped == "set_gpio_cfg 0x3FFi 1" and cfg_idx is None:
                cfg_idx = idx
                continue
            if cfg_idx is not None and stripped.startswith("set_gpio ") \
                    and not stripped.startswith("set_gpio_cfg"):
                gpio_idx = idx
                break
        if cfg_idx is None or gpio_idx is None:
            return script
        for i in sorted([cfg_idx, gpio_idx], reverse=True):
            del lines[i]
        return "\n".join(lines)

    def _confirm_mux_override(self) -> bool:
        if not self._has_mux_header:
            return True
        return messagebox.askyesno(
            "MUX Header Detected",
            "This script already includes a MUX header.\n"
            "Generate new scripts for the selected channel(s) anyway?",
        )

    # ── Parameter extraction (CV/SWV) ───────────────────────────────────────────

    @staticmethod
    def _parse_si_value(token: str):
        """Parse a MethodSCRIPT SI token to float. Returns None on failure."""
        token = token.strip()
        if not token:
            return None
        # Handle plain integers/floats
        try:
            return float(token)
        except ValueError:
            pass

        m = re.match(r"^([+-]?\d+(?:\.\d+)?)([afpnumkMGTPE])$", token)
        if not m:
            return None
        val = float(m.group(1))
        prefix = m.group(2)
        factors = {
            "a": 1e-18, "f": 1e-15, "p": 1e-12, "n": 1e-9, "u": 1e-6,
            "m": 1e-3, "k": 1e3, "M": 1e6, "G": 1e9, "T": 1e12, "P": 1e15, "E": 1e18,
        }
        return val * factors[prefix]

    @staticmethod
    def _fmt_value(val):
        if val is None:
            return None
        return f"{val:g}"

    def _extract_params(self, script: str):
        """Return (technique, params) if script looks like CV or SWV."""
        lines = [ln.strip() for ln in script.splitlines()]
        for line in lines:
            if not line or line.startswith("#"):
                continue
            if line.startswith("meas_loop_cv"):
                tokens = line.split()
                if len(tokens) < 8:
                    return None, None
                begin_v = self._fmt_value(self._parse_si_value(tokens[3]))
                v1      = self._fmt_value(self._parse_si_value(tokens[4]))
                v2      = self._fmt_value(self._parse_si_value(tokens[5]))
                step    = self._fmt_value(self._parse_si_value(tokens[6]))
                rate    = self._fmt_value(self._parse_si_value(tokens[7]))
                n_scans = "1"
                for tok in tokens[8:]:
                    if tok.startswith("nscans(") and tok.endswith(")"):
                        n_scans = tok[len("nscans("):-1]
                        break
                params = {
                    "begin_potential": begin_v or tokens[3],
                    "vertex1": v1 or tokens[4],
                    "vertex2": v2 or tokens[5],
                    "step_potential": step or tokens[6],
                    "scan_rate": rate or tokens[7],
                    "n_scans": str(n_scans),
                    "cond_potential": "0",
                    "cond_time": "0",
                    "mux_channel": "0",
                }
                return "CV", params

            if line.startswith("meas_loop_swv"):
                tokens = line.split()
                if len(tokens) < 9:
                    return None, None
                begin_v = self._fmt_value(self._parse_si_value(tokens[5]))
                end_v   = self._fmt_value(self._parse_si_value(tokens[6]))
                step    = self._fmt_value(self._parse_si_value(tokens[7]))
                amp     = self._fmt_value(self._parse_si_value(tokens[8]))
                freq    = self._fmt_value(self._parse_si_value(tokens[9])) if len(tokens) > 9 else None
                n_scans = "1"
                for tok in tokens[10:]:
                    if tok.startswith("nscans(") and tok.endswith(")"):
                        n_scans = tok[len("nscans("):-1]
                        break
                params = {
                    "begin_potential": begin_v or tokens[5],
                    "end_potential": end_v or tokens[6],
                    "step_potential": step or tokens[7],
                    "amplitude": amp or tokens[8],
                    "frequency": freq or (tokens[9] if len(tokens) > 9 else ""),
                    "n_scans": str(n_scans),
                    "cycle_delay": "0",
                    "cond_potential": "0",
                    "cond_time": "0",
                    "mux_channel": "0",
                }
                return "SWV", params
        return None, None

    # ── Run Now ────────────────────────────────────────────────────────────────

    def _run_now(self):
        if not self.script_path or not self.script_text:
            messagebox.showerror(
                "No Script Loaded",
                "Browse and select a MethodSCRIPT file before running.",
            )
            return
        channels = self._parse_mux(self.params)
        if channels is None:
            return

        base_script = self.script_text

        if channels:
            if not self._confirm_mux_override():
                return
            base_script = self._strip_mux_header(base_script)
            if len(channels) == 1:
                mux_script = self._wrap_mux(base_script, channels[0])
                self._run_now_cb("Custom", mux_script, channels[0])
            else:
                self._run_now_cb("CV_MUX_SEQ", base_script, channels)
        else:
            self._run_now_cb("Custom", self.script_text, None)

    # ── Add to Queue ───────────────────────────────────────────────────────────

    def _add_to_queue(self):
        if not self.script_path or not self.script_text:
            messagebox.showerror(
                "No Script Loaded",
                "Browse and select a MethodSCRIPT file before adding to the queue.",
            )
            return

        # Require an active session before touching the queue
        if not self._session.require_session():
            return

        channels = self._parse_mux(self.params)
        if channels is None:
            return

        base_script = self.script_text
        note = (getattr(self, "_note_var", None).get() or "").strip()
        detected_technique, detected_params = self._extract_params(base_script)
        script_hash = hashlib.sha1(base_script.encode("utf-8")).hexdigest()[:12]
        base_params = detected_params or {"custom_hash": script_hash}
        save_technique = detected_technique or "Custom"

        if channels:
            if not self._confirm_mux_override():
                return
            base_script = self._strip_mux_header(base_script)
            for ch in channels:
                mux_script = self._wrap_mux(base_script, ch)
                try:
                    filepath, filename = self._save_script(
                        save_technique,
                        mux_script,
                        base_params,
                        ch,
                        note=note,
                    )
                except Exception as exc:
                    messagebox.showerror(
                        "File Error",
                        f"Failed to save Custom script (MUX ch {ch}): {exc}",
                    )
                    return
                label = self.script_name or filename
                self._add_to_queue_cb({
                    "type":        "Custom",
                    "script_path": str(filepath),
                    "status":      "pending",
                    "details":     f"{label} (MUX ch {ch})",
                })
            messagebox.showinfo(
                "Added",
                f"Custom script added for {len(channels)} MUX channel(s).",
            )
            return

        # No MUX — save and queue as-is
        try:
            filepath, filename = self._save_script(
                save_technique,
                self.script_text,
                base_params,
                note=note,
            )
        except Exception as exc:
            messagebox.showerror(
                "File Error", f"Failed to save Custom script: {exc}"
            )
            return

        label   = self.script_name or filename
        details = f"{label} (saved as {filename})" if label != filename else filename
        self._add_to_queue_cb({
            "type":        "Custom",
            "script_path": str(filepath),
            "status":      "pending",
            "details":     details,
        })
        messagebox.showinfo("Added", f"Custom script added to queue.\nSaved as: {filename}")
