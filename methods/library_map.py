# methods/library_map.py
"""
Persistent library of MethodSCRIPT files.

The library_map is a JSON file that survives across sessions. It maps a
parameter-based hash to a canonical .ms file stored in methods/library/.

The hash is computed from the *parameters* (technique + raw param values +
mux channel), NOT from the generated script text — so the same experimental
setup always maps to the same hash, even if the script generator changes
whitespace or comments.

Folder layout
-------------
methods/
  library_map.json          ← persistent { hash_key: entry_dict }
  library/
    swv_a3f2c1.ms           ← canonical files, named by hash
    cv_00ff12.ms
  archive/                  ← old dated folders moved here manually
  YYYY-MM-DD/               ← dated working copies written each session
"""

import json
import hashlib
from pathlib import Path
from datetime import datetime
from typing import Optional

_METHODS_ROOT = Path("methods")
_LIBRARY_DIR  = _METHODS_ROOT / "library"
_ARCHIVE_DIR  = _METHODS_ROOT / "archive"
_MAP_FILE     = _METHODS_ROOT / "library_map.json"

_map: dict = {}   # in-memory cache


def _ensure_dirs():
    _METHODS_ROOT.mkdir(exist_ok=True)
    _LIBRARY_DIR.mkdir(exist_ok=True)
    _ARCHIVE_DIR.mkdir(exist_ok=True)


def load_map() -> dict:
    """Load library_map.json into memory. No-op if already loaded."""
    global _map
    _ensure_dirs()
    if _map:
        return _map
    if _MAP_FILE.exists():
        try:
            _map = json.loads(_MAP_FILE.read_text(encoding="utf-8"))
        except Exception:
            _map = {}
    return _map


def _persist():
    _MAP_FILE.write_text(json.dumps(_map, indent=2), encoding="utf-8")


def compute_hash(technique: str, params: dict, mux_channel: Optional[int]) -> str:
    """Compute a stable hash from *parameters*, not generated script text.

    Parameters
    ----------
    technique:   "CV" or "SWV"
    params:      raw param values as strings, e.g. {"begin_potential": "-0.5", ...}
                 Keys are sorted before hashing so insertion order doesn't matter.
    mux_channel: integer channel or None

    Returns
    -------
    Hash key like ``swv_a3f2c1`` or ``cv_ch3_00ff12``
    """
    slug = technique.lower().replace(" ", "_")
    if mux_channel is not None:
        slug = f"{slug}_ch{mux_channel}"

    # Sort params so key order never affects the hash
    canonical = json.dumps(
        {k: str(v).strip() for k, v in sorted(params.items())},
        separators=(",", ":")
    )
    raw = f"{slug}||{canonical}"
    h = hashlib.md5(raw.encode("utf-8"), usedforsecurity=False).hexdigest()[:6]
    return f"{slug}_{h}"


def lookup(hash_key: str) -> Optional[Path]:
    """Return the library Path if this hash exists and the file is on disk."""
    load_map()
    entry = _map.get(hash_key)
    if entry is None:
        return None
    p = Path(entry["filepath"])
    if p.exists():
        return p
    # Stale entry — file was deleted externally
    del _map[hash_key]
    _persist()
    return None


def register(hash_key: str, technique: str, params: dict,
             mux_channel: Optional[int], script: str) -> Path:
    """Write script into the library and record in the map.

    Only called when lookup() returned None.
    """
    _ensure_dirs()
    lib_path = _LIBRARY_DIR / f"{hash_key}.ms"
    lib_path.write_text(script, encoding="utf-8")

    _map[hash_key] = {
        "technique":   technique,
        "mux_channel": mux_channel,
        "params":      {k: str(v).strip() for k, v in params.items()},
        "added_at":    datetime.now().isoformat(timespec="seconds"),
        "filepath":    str(lib_path),
    }
    _persist()
    return lib_path


def all_entries() -> dict:
    """Return full map — used by the hash-finder UI tool."""
    load_map()
    return dict(_map)


def find_by_technique(technique: str) -> dict:
    """Return all entries for a given technique."""
    load_map()
    t = technique.upper()
    return {k: v for k, v in _map.items()
            if v.get("technique", "").upper() == t}


def reload():
    """Force a full reload from disk."""
    global _map
    _map = {}
    load_map()