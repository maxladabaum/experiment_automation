from __future__ import annotations

import csv
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np


CHANNEL_RE = re.compile(r"(?:^|[_\-\s])ch(?:annel)?\s*0*(\d+)(?:\D|$)", re.IGNORECASE)
MEAS_RE = re.compile(r"(?:^|[_\-\s])meas[_\-\s].*?(\d{3,})(?:\D|$)", re.IGNORECASE)


@dataclass(frozen=True)
class SWVFile:
    path: str
    channel: int
    ts: str
    scan: str
    folder_index: int


def collect_swv_csvs_from_folders(folders: List[str]) -> List[SWVFile]:
    files: List[SWVFile] = []
    for folder_index, folder in enumerate(folders):
        root = Path(folder)
        if not root.exists():
            continue
        candidates = root.rglob("*.csv") if root.is_dir() else [root]
        for path in candidates:
            if not path.is_file() or _is_analysis_output(path):
                continue
            if not _looks_like_measurement_csv(path):
                continue
            channel = _infer_channel(path)
            files.append(
                SWVFile(
                    path=str(path),
                    channel=channel,
                    ts=_infer_timestamp(path),
                    scan=_infer_scan_id(path),
                    folder_index=folder_index,
                )
            )
    return files


def filter_finite(v_raw, i_raw) -> Tuple[np.ndarray, np.ndarray]:
    v = np.asarray(v_raw, dtype=float)
    i = np.asarray(i_raw, dtype=float)
    mask = np.isfinite(v) & np.isfinite(i)
    return v[mask], i[mask]


def group_by_channel_and_sort(files: List[SWVFile]) -> Dict[int, List[SWVFile]]:
    grouped: Dict[int, List[SWVFile]] = {}
    for item in files:
        grouped.setdefault(int(item.channel), []).append(item)
    for channel in grouped:
        grouped[channel].sort(key=lambda f: (f.ts, f.scan, os.path.basename(f.path)))
    return grouped


def load_swv_csv(
    filepath: str,
    voltage_col: str = "Potential (V)",
    current_col: Optional[str] = None,
) -> Tuple[np.ndarray, np.ndarray]:
    try:
        import pandas as pd

        df = pd.read_csv(filepath)
        if voltage_col not in df.columns:
            raise ValueError(
                f"Voltage column '{voltage_col}' not found. Columns: {list(df.columns)}"
            )
        if current_col is None:
            if "Current Diff (uA)" in df.columns:
                current_col = "Current Diff (uA)"
            elif "Current (uA)" in df.columns:
                current_col = "Current (uA)"
            else:
                raise ValueError(
                    "Cannot auto-pick current column. Need 'Current Diff (uA)' or 'Current (uA)'. "
                    f"Columns: {list(df.columns)}"
                )
        elif current_col not in df.columns:
            raise ValueError(
                f"Current column '{current_col}' not found. Columns: {list(df.columns)}"
            )
        return (
            df[voltage_col].to_numpy(dtype=float),
            df[current_col].to_numpy(dtype=float),
        )
    except Exception:
        with open(filepath, "r", encoding="utf-8-sig", errors="replace", newline="") as fh:
            reader = csv.reader(fh)
            rows = [row for row in reader if any(str(cell).strip() for cell in row)]
        if not rows:
            raise ValueError(f"CSV is empty: {filepath}")

        header = [str(cell).strip() for cell in rows[0]]
        data_rows = rows[1:]
        voltage_idx = _find_column(header, [voltage_col, "Potential (V)", "potential", "Ewe/V"])
        current_candidates = (
            [current_col]
            if current_col
            else [
                "Current Diff (uA)",
                "current_diff",
                "Current (uA)",
                "Current (ÂµA)",
                "Current (Ã‚ÂµA)",
                "current",
            ]
        )
        current_idx = _find_column(header, current_candidates)
        if voltage_idx is None or current_idx is None:
            raise ValueError(f"CSV lacks potential/current columns: {filepath}")

        voltage = []
        current = []
        for row in data_rows:
            if max(voltage_idx, current_idx) >= len(row):
                continue
            try:
                v_value = float(str(row[voltage_idx]).strip())
                i_value = float(str(row[current_idx]).strip())
            except ValueError:
                continue
            voltage.append(v_value)
            current.append(i_value)
        return np.asarray(voltage, dtype=float), np.asarray(current, dtype=float)


def _find_column(header: List[str], candidates: List[Optional[str]]) -> Optional[int]:
    normalized = [_normalize_name(name) for name in header]
    for candidate in candidates:
        if not candidate:
            continue
        needle = _normalize_name(candidate)
        if needle in normalized:
            return normalized.index(needle)
    for idx, name in enumerate(normalized):
        if "potential" in name or name in {"ewev", "voltage"}:
            if any("potential" in _normalize_name(c or "") for c in candidates):
                return idx
        if "current" in name and any("current" in _normalize_name(c or "") for c in candidates):
            return idx
    return None


def _normalize_name(value: str) -> str:
    value = str(value).replace("Ã‚Âµ", "u").replace("Âµ", "u")
    return "".join(ch.lower() for ch in value if ch.isalnum())


def _looks_like_measurement_csv(path: Path) -> bool:
    if path.name.lower().endswith("_results.csv"):
        return False
    method_path = _method_snapshot_for_csv(path)
    if method_path.exists():
        try:
            if "meas_loop_swv" in method_path.read_text(encoding="utf-8", errors="replace").lower():
                return True
        except OSError:
            pass
    if "swv" in path.name.lower():
        return True
    try:
        with open(path, "r", encoding="utf-8-sig", errors="replace", newline="") as fh:
            header = next(csv.reader(fh), [])
        normalized = [_normalize_name(cell) for cell in header]
        return any("potential" in cell for cell in normalized) and any("current" in cell for cell in normalized)
    except Exception:
        return False


def _is_analysis_output(path: Path) -> bool:
    lowered = [part.lower() for part in path.parts]
    return "bo_analysis" in lowered or "analysis" in lowered or "bo_sessions" in lowered


def _method_snapshot_for_csv(path: Path) -> Path:
    return path.parent / "methods_used" / f"{path.stem}.ms"


def _infer_channel(path: Path) -> int:
    match = CHANNEL_RE.search(path.stem)
    if match:
        return int(match.group(1))
    return 1


def _infer_timestamp(path: Path) -> str:
    stat = path.stat()
    return f"{int(stat.st_mtime_ns):020d}"


def _infer_scan_id(path: Path) -> str:
    match = MEAS_RE.search(path.stem)
    if match:
        return match.group(1)
    return path.stem
