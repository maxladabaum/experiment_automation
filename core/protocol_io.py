"""Save protocol script references as relative paths; never bundle script contents."""

import copy
import json
import os
import tempfile
from pathlib import Path, PureWindowsPath

from methods import library_map


def relative_script_path(script_path, destination):
    """Use methods/... for library files, otherwise a path relative to the protocol."""
    normalized = str(script_path).replace("\\", "/")
    path = Path(normalized).expanduser()
    # A configured library can itself live below a directory named methods.
    # Use its actual root before interpreting old, foreign methods/... paths.
    if path.is_absolute():
        for root in (library_map._METHODS_ROOT.resolve(), Path(library_map.__file__).resolve().parent):
            try:
                relative = path.relative_to(root)
                return (Path("methods") / relative).as_posix()
            except ValueError:
                pass
    # Only unavailable legacy absolute library addresses are migrated by suffix.
    # Existing external files, and explicit recipe-relative addresses, retain
    # their location even when an ancestor directory happens to be named methods.
    if not path.exists() and (path.is_absolute() or PureWindowsPath(normalized).is_absolute()):
        relative = library_map.legacy_library_suffix(normalized)
        if relative is not None:
            return (Path("methods") / relative).as_posix()
    if not path.is_absolute():
        return normalized
    try:
        relative = Path(os.path.relpath(path, Path(destination).parent)).as_posix()
        # Reserve bare methods/... for the configured/bundled library namespace.
        return "./" + relative if relative.lower().startswith("methods/") else relative
    except ValueError:
        raise ValueError(
            "Cannot save a relative script reference across different drives. "
            "Place the script in this setup's configured methods library or "
            "on the same drive as the recipe, then save again."
        ) from None


def portable_items(items, destination):
    result = copy.deepcopy(items)
    for item in result:
        if isinstance(item, dict) and item.get("script_path"):
            item["script_path"] = relative_script_path(item["script_path"], destination)
    return result


def write_protocol(path, payload):
    """Write relative references without changing the live queue or recipe."""
    data = copy.deepcopy(payload)
    data["items"] = portable_items(data["items"], path)
    path = Path(path)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=path.parent,
                                         prefix=path.name + ".", suffix=".tmp", delete=False) as fh:
            temporary = Path(fh.name)
            json.dump(data, fh, indent=2)
        os.replace(temporary, path)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)


def loaded_items(payload, source_path):
    """Resolve relative paths using this setup's library or the protocol's folder."""
    items = payload if isinstance(payload, list) else payload.get("items")
    if not isinstance(items, list):
        raise ValueError("Protocol file missing 'items' list")
    result = copy.deepcopy(items)
    for item in result:
        if isinstance(item, dict) and item.get("script_path"):
            path = library_map.resolve_script_path(item["script_path"], Path(source_path).parent)
            if path is not None:
                item["script_path"] = str(path)
    return result
