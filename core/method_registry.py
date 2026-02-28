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
        technique: str,
        script: str,
        mux_channel: Optional[int] = None,
    ) -> Tuple[Path, str]:
        """Save a MethodSCRIPT to disk, deduplicating identical scripts.

        If the same (technique, script_content, mux_channel) triple has
        already been saved this session, the existing path is returned and
        **no new file is written**.

        Returns
        -------
        (filepath, filename)
        """
        key = self._make_key(technique, script, mux_channel)

        if key in self._registry:
            fp, fn = self._registry[key]
            self._log(f"[Registry] Reusing '{fn}'  (hash: {key})")
            return fp, fn

        # New script — persist to disk
        date_folder = self.base_path / datetime.now().strftime("%Y-%m-%d")
        date_folder.mkdir(exist_ok=True)

        slug     = technique.lower().replace(" ", "_")
        if mux_channel is not None:
            slug = f"{slug}_ch{mux_channel}"
        existing = len(list(date_folder.glob("*.ms")))
        filename = f"{existing + 1:03d}_{slug}.ms"
        filepath = date_folder / filename

        filepath.write_text(script)

        self._registry[key]           = (filepath, filename)
        self._path_to_key[str(filepath)] = key
        self._log(f"[Registry] Saved '{filename}'  (hash: {key})")
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
    def _make_key(technique: str, script: str, mux_channel: Optional[int]) -> str:
        """Compute a short, stable hash key for a (technique, script, channel)
        triple.

        Format: ``<slug>_<6hexchars>``
        """
        slug = technique.lower().replace(" ", "_")
        if mux_channel is not None:
            slug = f"{slug}_ch{mux_channel}"
        raw = f"{slug}||{script}"
        h   = hashlib.md5(raw.encode("utf-8"), usedforsecurity=False).hexdigest()[:6]
        return f"{slug}_{h}"
