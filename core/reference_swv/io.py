import os
import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

FILENAME_RE = re.compile(
    r"^swv_ch(?P<ch>\d+)_([0-9a-f]+)_meas_(?P<date>\d{8})_(?P<time>\d{4})_(?P<scan>\d+)_ch(?P<ch2>\d+)\.csv$",
    re.IGNORECASE,
)

@dataclass(frozen=True)
class SWVFile:
    scan: int
    ch: int
    ts: int
    path: str
    folder_index: int


def collect_swv_csvs_from_folders(folders: List[str]) -> List[SWVFile]:
    out: List[SWVFile] = []
    for folder_index, folder in enumerate(folders):
        if not os.path.isdir(folder):
            raise ValueError(f"Folder not found: {folder}")
        for fn in os.listdir(folder):
            if not fn.lower().endswith(".csv"):
                continue
            m = FILENAME_RE.match(fn)
            if not m:
                continue
            ch = int(m.group("ch"))
            if int(m.group("ch2")) != ch:
                continue
            ts = int(f"{m.group('date')}{m.group('time')}")
            out.append(
                SWVFile(
                    scan=int(m.group("scan")),
                    ch=ch,
                    ts=ts,
                    path=os.path.join(folder, fn),
                    folder_index=folder_index,
                )
            )
    return out


def group_by_channel_and_sort(files: List[SWVFile]) -> Dict[int, List[SWVFile]]:
    d: Dict[int, List[SWVFile]] = {}
    for f in files:
        d.setdefault(f.ch, []).append(f)
    for ch in d:
        d[ch].sort(key=lambda f: (f.folder_index, f.ts, f.scan))
    return d


def load_swv_csv(
    filepath: str,
    voltage_col: str = "Potential (V)",
    current_col: Optional[str] = None,
) -> Tuple[np.ndarray, np.ndarray]:
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
                f"Cannot auto-pick current column. Need 'Current Diff (uA)' or 'Current (uA)'. "
                f"Columns: {list(df.columns)}"
            )
    elif current_col not in df.columns:
        raise ValueError(
            f"Current column '{current_col}' not found. Columns: {list(df.columns)}"
        )

    return df[voltage_col].to_numpy(dtype=float), df[current_col].to_numpy(dtype=float)


def filter_finite(v: np.ndarray, y: np.ndarray) -> Tuple[np.ndarray, np.ndarray]:
    v, y = np.asarray(v, dtype=float), np.asarray(y, dtype=float)
    mask = np.isfinite(v) & np.isfinite(y)
    return v[mask], y[mask]
