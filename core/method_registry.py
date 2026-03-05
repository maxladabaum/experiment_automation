"""
core/method_registry.py — Script deduplication and file management.

Provides :class:`MethodRegistry` which is the **single** place MethodSCRIPT
files are written to disk.  Identical scripts (same technique + parameters +
MUX channel) are stored only once, identified by a short MD5-based hash key.

Usage::

    registry = MethodRegistry(log_callback=some_fn)
    filepath, filename = registry.save_script("SWV", script_text, mux_channel=3)
    # Second call with same args → returns cached path, no new file written
    filepath, filename = registry.save_script("SWV", script_text, mux_channel=3)

Hash key format:  ``<technique_slug>[_ch<N>]_<6hexchars>``
Example:          ``swv_ch3_a3f2c1``
"""

import hashlib
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, Optional, Tuple
from methods.library_map import compute_hash, lookup, register, update_note
from config import METHODS_DIR


class MethodRegistry:
    """Manages saving and deduplication of MethodSCRIPT files.

    Parameters
    ----------
    log_callback:
        Optional callable for log output.  Defaults to ``print``.
    base_path:
        Root directory for saved scripts.  Defaults to ``config.METHODS_DIR``.
        A per-day subfolder is created automatically.
    """

    def __init__(
        self,
        log_callback: Callable[[str], None] = print,
        base_path: Optional[Path] = None,
    ):
        self._log      = log_callback
        self.base_path = Path(base_path) if base_path else Path(METHODS_DIR)
        self.base_path.mkdir(exist_ok=True)

        # hash_key → (filepath, filename)
        self._registry: Dict[str, Tuple[Path, str]] = {}
        # str(filepath) → hash_key  (reverse lookup)
        self._path_to_key: Dict[str, str] = {}

    # ── Public API ────────────────────────────────────────────────────────────

    def save_script(
        self,
        technique:   str,
        script:      str,
        params:      Optional[dict] = None,
        mux_channel: Optional[int] = None,
        note:        Optional[str] = None,
    ) -> Tuple[Path, str]:
        """Save a MethodSCRIPT, checking session cache then library before writing.

        Level 1 — session registry (in-memory, fastest, lost on restart)
        Level 2 — persistent library (methods/library/, survives restarts)
        Level 3 — genuinely new: write to library + dated working copy
        """
        if params is None:
            # Fall back to script-content hashing for ad-hoc scripts.
            params = {
                "_script_hash": hashlib.sha1(script.encode("utf-8")).hexdigest()[:12]
            }
        key = self._make_key(technique, params, mux_channel)

        # Level 1: session cache
        if key in self._registry:
            fp, fn = self._registry[key]
            if note is not None:
                update_note(key, note)
            self._log(f"[Registry] Session hit  '{fn}'  ({key})")
            return fp, fn

        # Level 2: persistent library
        lib_path = lookup(key)
        if lib_path is not None:
            fn = lib_path.name
            self._registry[key]           = (lib_path, fn)
            self._path_to_key[str(lib_path)] = key
            if note is not None:
                update_note(key, note)
            self._log(f"[Library]  Found        '{fn}'  ({key})")
            return lib_path, fn

        # Level 3: new — write to library and a dated working copy
        lib_path = register(key, technique, params, mux_channel, script, note=note)

        date_folder = self.base_path / datetime.now().strftime("%Y-%m-%d")
        date_folder.mkdir(exist_ok=True)
        slug     = technique.lower().replace(" ", "_")
        if mux_channel is not None:
            slug = f"{slug}_ch{mux_channel}"
        existing = len(list(date_folder.glob("*.ms")))
        filename = f"{existing + 1:03d}_{slug}.ms"
        filepath = date_folder / filename
        filepath.write_text(script, encoding="utf-8")

        self._registry[key]              = (filepath, filename)
        self._path_to_key[str(filepath)] = key
        self._log(f"[Library]  Saved new    '{filename}'  ({key})")
        return filepath, filename

    def hash_key_for(self, filepath) -> str:
        """Return the hash key for an already-saved script path (for UI display)."""
        return self._path_to_key.get(str(filepath), "—")

    def clear(self):
        """Clear the in-memory registry (does not delete files from disk).

        Call this to force fresh files to be written (e.g. after a hardware
        config change).
        """
        count = len(self._registry)
        self._registry.clear()
        self._path_to_key.clear()
        self._log(
            f"[Registry] Cleared ({count} entries). "
            "New scripts will be saved as fresh files."
        )

    @property
    def size(self) -> int:
        """Number of unique scripts currently in the registry."""
        return len(self._registry)

    # ── Internal helpers ──────────────────────────────────────────────────────

    @staticmethod
    def _make_key(technique: str, params: dict, mux_channel: Optional[int]) -> str:
        """Delegate to library_map.compute_hash — hash is param-based, not script-based."""
        return compute_hash(technique, params, mux_channel)
