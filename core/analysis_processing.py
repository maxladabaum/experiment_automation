from __future__ import annotations

from typing import Dict

import numpy as np


def apply_smoothing(y, smooth_window: int, smooth_polyorder: int):
    signal = np.asarray(y, dtype=float)
    if signal.size < 3 or smooth_window <= 2:
        return signal.copy()
    window = int(smooth_window)
    if window % 2 == 0:
        window += 1
    window = min(window, signal.size if signal.size % 2 else signal.size - 1)
    if window <= max(2, int(smooth_polyorder)):
        return signal.copy()
    try:
        from scipy.signal import savgol_filter

        return np.asarray(savgol_filter(signal, window, int(smooth_polyorder)), dtype=float)
    except Exception:
        kernel = np.ones(window, dtype=float) / float(window)
        return np.convolve(signal, kernel, mode="same")


def detect_dominant_peak(y, boundary_margin: int = 1) -> int:
    signal = np.asarray(y, dtype=float)
    if signal.size == 0:
        raise ValueError("Cannot detect a peak in an empty signal.")
    margin = max(0, min(int(boundary_margin), max(0, signal.size // 4)))
    if signal.size > margin * 2:
        local = signal[margin:signal.size - margin]
        return int(margin + np.nanargmax(local))
    return int(np.nanargmax(signal))


def rotate_offset_using_bracketing_minima(v, y, peak_idx: int, search_window_v: float) -> Dict[str, object]:
    return _correct_with_bracketing_minima(v, y, peak_idx, search_window_v, mode="argmin_window")


def rotate_offset_using_prominent_bracketing_minima(v, y, peak_idx: int, search_window_v: float) -> Dict[str, object]:
    return _correct_with_bracketing_minima(v, y, peak_idx, search_window_v, mode="prominent_minima")


def _correct_with_bracketing_minima(v, y, peak_idx: int, search_window_v: float, mode: str) -> Dict[str, object]:
    voltage = np.asarray(v, dtype=float)
    signal = np.asarray(y, dtype=float)
    if voltage.size != signal.size or signal.size < 5:
        raise ValueError("Voltage and current arrays must have at least five matching points.")
    peak_idx = int(max(0, min(peak_idx, signal.size - 1)))
    window = abs(float(search_window_v))
    left_candidates = _candidate_indices(voltage, signal, peak_idx, -1, window)
    right_candidates = _candidate_indices(voltage, signal, peak_idx, 1, window)
    if left_candidates.size == 0 or right_candidates.size == 0:
        raise ValueError("Could not find minima bracketing the dominant peak.")

    left_idx = int(left_candidates[np.argmin(signal[left_candidates])])
    right_idx = int(right_candidates[np.argmin(signal[right_candidates])])
    if left_idx >= right_idx:
        raise ValueError("Invalid bracketing minima around peak.")

    baseline = np.interp(voltage, [voltage[left_idx], voltage[right_idx]], [signal[left_idx], signal[right_idx]])
    corrected = signal - baseline
    return {
        "y_corrected": corrected,
        "local_baseline": baseline,
        "left_idx": left_idx,
        "right_idx": right_idx,
        "left_local_min_candidates": left_candidates,
        "right_local_min_candidates": right_candidates,
        "minima_mode": mode,
    }


def _candidate_indices(voltage, signal, peak_idx: int, direction: int, window_v: float):
    peak_v = float(voltage[peak_idx])
    if direction < 0:
        mask = np.arange(signal.size) < peak_idx
    else:
        mask = np.arange(signal.size) > peak_idx
    if window_v > 0:
        mask &= np.abs(voltage - peak_v) <= window_v
    candidates = np.where(mask)[0]
    local = _local_minima(signal, candidates)
    return local if local.size else candidates


def _local_minima(signal, candidates):
    if candidates.size == 0:
        return candidates
    out = []
    for idx in candidates:
        if idx <= 0 or idx >= signal.size - 1:
            continue
        if signal[idx] <= signal[idx - 1] and signal[idx] <= signal[idx + 1]:
            out.append(int(idx))
    return np.asarray(out, dtype=int)
