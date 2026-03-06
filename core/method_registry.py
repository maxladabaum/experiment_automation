"""
core/method_registry.py - Script deduplication and file management.

MethodRegistry is the single place MethodSCRIPT files are written to disk.
Identical scripts (same technique + params + MUX channel) are stored once,
identified by a short hash key.
"""

import hashlib
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, Optional, Tuple

from methods.library_map import compute_hash, lookup, register, update_note
from config import METHODS_DIR, SAVE_DATED_METHOD_COPIES


class MethodRegistry:
    """Manages saving and deduplication of MethodSCRIPT files."""

    def __init__(
        self,
        log_callback: Callable[[str], None] = print,
        base_path: Optional[Path] = None,
    ):
        self._log = log_callback
        self.base_path = Path(base_path) if base_path else Path(METHODS_DIR)
        self.base_path.mkdir(exist_ok=True)

        # hash_key -> (filepath, filename)
        self._registry: Dict[str, Tuple[Path, str]] = {}
        # str(filepath) -> hash_key
        self._path_to_key: Dict[str, str] = {}

    def save_script(
        self,
        technique: str,
        script: str,
        params: Optional[dict] = None,
        mux_channel: Optional[int] = None,
        note: Optional[str] = None,
    ) -> Tuple[Path, str]:
        """
        Save a MethodSCRIPT, checking session cache then library before writing.

        Level 1: in-memory session registry
        Level 2: persistent library (methods/library)
        Level 3: new script -> write to library (+ optional dated working copy)
        """
        if params is None:
            # Fall back to script-content hashing for ad-hoc scripts.
            params = {"_script_hash": hashlib.sha1(script.encode("utf-8")).hexdigest()[:12]}
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
            self._registry[key] = (lib_path, fn)
            self._path_to_key[str(lib_path)] = key
            if note is not None:
                update_note(key, note)
            self._log(f"[Library]  Found        '{fn}'  ({key})")
            return lib_path, fn

        # Level 3: genuinely new
        lib_path = register(key, technique, params, mux_channel, script, note=note)

        if SAVE_DATED_METHOD_COPIES:
            date_folder = self.base_path / datetime.now().strftime("%Y-%m-%d")
            date_folder.mkdir(exist_ok=True)
            slug = technique.lower().replace(" ", "_")
            if mux_channel is not None:
                slug = f"{slug}_ch{mux_channel}"
            existing = len(list(date_folder.glob("*.ms")))
            filename = f"{existing + 1:03d}_{slug}.ms"
            filepath = date_folder / filename
            filepath.write_text(script, encoding="utf-8")

            self._registry[key] = (filepath, filename)
            self._path_to_key[str(filepath)] = key
            self._log(f"[Library]  Saved new    '{filename}'  ({key}) [dated copy]")
            return filepath, filename

        filename = lib_path.name
        self._registry[key] = (lib_path, filename)
        self._path_to_key[str(lib_path)] = key
        self._log(f"[Library]  Saved new    '{filename}'  ({key})")
        return lib_path, filename

    def hash_key_for(self, filepath) -> str:
        """Return hash key for an already-saved script path."""
        return self._path_to_key.get(str(filepath), "-")

    def clear(self):
        """Clear in-memory registry (does not delete files from disk)."""
        count = len(self._registry)
        self._registry.clear()
        self._path_to_key.clear()
        self._log(
            f"[Registry] Cleared ({count} entries). "
            "New scripts will be saved as fresh files."
        )

    @property
    def size(self) -> int:
        """Number of unique scripts currently in the in-memory registry."""
        return len(self._registry)

    @staticmethod
    def _make_key(technique: str, params: dict, mux_channel: Optional[int]) -> str:
        """Hash is param-based, not script-text-based."""
        return compute_hash(technique, params, mux_channel)
