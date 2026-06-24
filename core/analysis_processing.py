from __future__ import annotations

from typing import Dict, Tuple

import numpy as np


def _normalized_savgol_params(signal_len: int, smooth_window: int, smooth_polyorder: int) -> Tuple[int, int]:
    w = int(smooth_window)
    if signal_len < 3 or w <= 2:
        return 0, 0
    if w >= signal_len:
        w = max(3, signal_len if signal_len % 2 == 1 else signal_len - 1)
    if w % 2 == 0:
        w += 1
    if w > signal_len:
        w = signal_len if signal_len % 2 == 1 else signal_len - 1
    if w < 3:
        return 0, 0
    p = int(min(max(0, smooth_polyorder), w - 1))
    return w, p


def _savgol_coefficients(window_length: int, polyorder: int) -> np.ndarray:
    half = window_length // 2
    x = np.arange(-half, half + 1, dtype=float)
    vand = np.vander(x, N=polyorder + 1, increasing=True)
    return np.linalg.pinv(vand)[0]


def _savgol_numpy(signal: np.ndarray, window_length: int, polyorder: int) -> np.ndarray:
    coeffs = _savgol_coefficients(window_length, polyorder)
    half = window_length // 2
    padded = np.pad(np.asarray(signal, dtype=float), (half, half), mode="edge")
    return np.convolve(padded, coeffs[::-1], mode="valid")


def fallback_find_peaks(
    signal: np.ndarray,
    prominence: float = 0.0,
    distance: int = 1,
) -> Tuple[np.ndarray, Dict[str, np.ndarray]]:
    y = np.asarray(signal, dtype=float)
    n = y.size
    if n < 3:
        return np.array([], dtype=int), {"prominences": np.array([], dtype=float)}

    candidates = []
    idx = 1
    while idx < n - 1:
        if y[idx] > y[idx - 1]:
            end = idx
            while end + 1 < n and y[end + 1] == y[idx]:
                end += 1
            if end < n - 1 and y[end] > y[end + 1]:
                candidates.append((idx + end) // 2)
            idx = end + 1
            continue
        idx += 1

    if not candidates:
        return np.array([], dtype=int), {"prominences": np.array([], dtype=float)}

    peaks = np.asarray(candidates, dtype=int)
    prominences = np.zeros(peaks.size, dtype=float)
    for i, peak_idx in enumerate(peaks):
        peak_val = y[peak_idx]

        left_min = peak_val
        left = peak_idx
        while left > 0:
            left -= 1
            left_min = min(left_min, y[left])
            if y[left] > peak_val:
                break

        right_min = peak_val
        right = peak_idx
        while right < n - 1:
            right += 1
            right_min = min(right_min, y[right])
            if y[right] > peak_val:
                break

        prominences[i] = peak_val - max(left_min, right_min)

    keep = prominences >= float(prominence)
    peaks = peaks[keep]
    prominences = prominences[keep]

    if peaks.size == 0:
        return np.array([], dtype=int), {"prominences": np.array([], dtype=float)}

    if distance > 1 and peaks.size > 1:
        order = np.argsort(prominences)[::-1]
        selected = []
        for order_idx in order:
            peak_idx = int(peaks[order_idx])
            if all(abs(peak_idx - kept_idx) >= distance for kept_idx in selected):
                selected.append(peak_idx)
        selected = np.array(sorted(selected), dtype=int)
        prom_map = {int(idx): float(prom) for idx, prom in zip(peaks, prominences)}
        peaks = selected
        prominences = np.asarray([prom_map[int(idx)] for idx in peaks], dtype=float)

    return peaks.astype(int), {"prominences": prominences.astype(float)}


def apply_smoothing(i: np.ndarray, smooth_window: int, smooth_polyorder: int) -> np.ndarray:
    signal = np.asarray(i, dtype=float)
    w, p = _normalized_savgol_params(signal.size, smooth_window, smooth_polyorder)
    if w == 0:
        return signal.copy()
    try:
        from scipy.signal import savgol_filter

        return np.asarray(savgol_filter(signal, window_length=w, polyorder=p), dtype=float)
    except Exception:
        return np.asarray(_savgol_numpy(signal, w, p), dtype=float)


def find_peak_candidates(
    i_smooth: np.ndarray,
    prominence: float = 0.02,
    distance: int = 5,
    boundary_margin: int = 5,
) -> dict:
    signal = np.asarray(i_smooth, dtype=float)
    try:
        from scipy.signal import find_peaks as scipy_find_peaks
    except Exception:
        scipy_find_peaks = None

    peak_fn = scipy_find_peaks if scipy_find_peaks is not None else fallback_find_peaks

    raw_peaks, _ = peak_fn(signal, distance=distance)
    raw_peaks = raw_peaks.astype(int)
    raw_valid_peaks = np.array(
        [p for p in raw_peaks if boundary_margin < p < len(signal) - boundary_margin],
        dtype=int,
    )

    peaks_by_pass = []
    valid_peaks = []

    for prom in (prominence, 0.005):
        peaks, props = peak_fn(signal, prominence=prom, distance=distance)
        peaks = peaks.astype(int)
        valid = np.array(
            [p for p in peaks if boundary_margin < p < len(signal) - boundary_margin],
            dtype=int,
        )
        peaks_by_pass.append(
            {
                "prominence": prom,
                "all_peaks": peaks,
                "valid_peaks": valid,
                "prominences": props.get("prominences"),
            }
        )
        if valid.size:
            valid_peaks = valid
            break

    if len(valid_peaks):
        dominant_idx = int(valid_peaks[np.argmax(signal[valid_peaks])])
    else:
        idx = int(np.argmax(signal))
        dominant_idx = max(boundary_margin, min(idx, len(signal) - boundary_margin - 1))

    return {
        "raw_peaks": raw_peaks,
        "raw_valid_peaks": raw_valid_peaks,
        "passes": peaks_by_pass,
        "valid_peaks": np.asarray(valid_peaks, dtype=int),
        "dominant_idx": int(dominant_idx),
    }


def detect_dominant_peak(
    i_smooth: np.ndarray,
    prominence: float = 0.02,
    distance: int = 5,
    boundary_margin: int = 5,
) -> int:
    return find_peak_candidates(
        i_smooth,
        prominence=prominence,
        distance=distance,
        boundary_margin=boundary_margin,
    )["dominant_idx"]


def _estimate_point_spacing(voltage: np.ndarray) -> float:
    diffs = np.abs(np.diff(np.asarray(voltage, dtype=float)))
    diffs = diffs[np.isfinite(diffs) & (diffs > 1e-12)]
    return float(np.median(diffs)) if diffs.size else 1.0


def _support_points_for_window(voltage: np.ndarray, window_v: float, fraction: float = 0.10) -> int:
    dv = _estimate_point_spacing(voltage)
    target_v = max(float(window_v) * float(fraction), dv * 2.0)
    pts = int(round(target_v / max(dv, 1e-12)))
    return max(2, pts)


def _window_indices(
    voltage: np.ndarray,
    peak_idx: int,
    search_window_v: float,
) -> Tuple[np.ndarray, np.ndarray]:
    v = np.asarray(voltage, dtype=float)
    v_peak = float(v[int(peak_idx)])

    left_idxs = np.where((v >= v_peak - search_window_v) & (v < v_peak))[0]
    if left_idxs.size == 0:
        left_idxs = np.arange(0, int(peak_idx))

    right_idxs = np.where((v <= v_peak + search_window_v) & (v > v_peak))[0]
    if right_idxs.size == 0:
        right_idxs = np.arange(int(peak_idx) + 1, len(v))

    return left_idxs.astype(int), right_idxs.astype(int)


def _peak_has_expected_flanks(
    y: np.ndarray,
    peak_idx: int,
    flank_points: int,
) -> bool:
    left_start = int(peak_idx) - int(flank_points)
    right_end = int(peak_idx) + int(flank_points)
    if left_start < 0 or right_end >= len(y):
        return False

    left_vals = y[left_start:int(peak_idx) + 1]
    right_vals = y[int(peak_idx):right_end + 1]
    left_diff = np.diff(left_vals)
    right_diff = np.diff(right_vals)
    if left_diff.size == 0 or right_diff.size == 0:
        return False

    left_rising_frac = float(np.mean(left_diff > 0))
    right_falling_frac = float(np.mean(right_diff < 0))
    if left_rising_frac < 0.60 or right_falling_frac < 0.60:
        return False

    eps = 0.05 * max(float(np.max(y) - np.min(y)), 1e-12)
    left_gain = float(y[int(peak_idx)] - left_vals[0])
    right_gain = float(y[int(peak_idx)] - right_vals[-1])
    return bool(left_gain > eps and right_gain > eps)


def _candidate_peak_indices(y: np.ndarray, peak_idx: int) -> np.ndarray:
    candidates = find_peak_candidates(y, boundary_margin=0)
    merged = np.concatenate(
        (
            np.asarray([peak_idx], dtype=int),
            np.asarray(candidates.get("valid_peaks", []), dtype=int),
            np.asarray(candidates.get("raw_valid_peaks", []), dtype=int),
        )
    )
    merged = merged[(merged >= 0) & (merged < len(y))]
    if not merged.size:
        return np.asarray([int(peak_idx)], dtype=int)
    unique = np.unique(merged)
    return unique[np.argsort(y[unique])[::-1]]


def _select_bracketing_peak_idx(
    voltage: np.ndarray,
    y: np.ndarray,
    peak_idx: int,
    search_window_v: float,
) -> int:
    flank_points = _support_points_for_window(voltage, search_window_v)

    for candidate_idx in _candidate_peak_indices(y, peak_idx):
        left_idxs, right_idxs = _window_indices(voltage, int(candidate_idx), search_window_v)
        if left_idxs.size == 0 or right_idxs.size == 0:
            continue
        if left_idxs.size < 2 or right_idxs.size < 2:
            continue
        if not _peak_has_expected_flanks(y, int(candidate_idx), flank_points):
            continue
        return int(candidate_idx)

    return int(peak_idx)


def _linear_baseline_from_indices(
    voltage: np.ndarray,
    y: np.ndarray,
    left_idx: int,
    right_idx: int,
) -> np.ndarray:
    v = np.asarray(voltage, dtype=float)
    y = np.asarray(y, dtype=float)
    v0, v1 = float(v[left_idx]), float(v[right_idx])
    y0, y1 = float(y[left_idx]), float(y[right_idx])

    denom = (v1 - v0) if abs(v1 - v0) > 1e-12 else 1e-12
    slope = (y1 - y0) / denom
    return slope * v + (y0 - slope * v0)


def rotate_offset_using_bracketing_minima(
    voltage: np.ndarray,
    y: np.ndarray,
    peak_idx: int,
    search_window_v: float = 0.12,
    require_local_minima_on_both_sides: bool = False,
) -> Dict[str, object]:
    del require_local_minima_on_both_sides
    v = np.asarray(voltage, dtype=float)
    signal = np.asarray(y, dtype=float)

    if len(v) < 5:
        raise ValueError("Too few points to compute bracketing minima baseline.")
    if peak_idx <= 0 or peak_idx >= len(signal) - 1:
        raise ValueError("Peak index is on/near boundary.")

    peak_idx = _select_bracketing_peak_idx(v, signal, int(peak_idx), search_window_v)
    left_idxs, right_idxs = _window_indices(v, peak_idx, search_window_v)
    left_idx = int(left_idxs[np.argmin(signal[left_idxs])])
    right_idx = int(right_idxs[np.argmin(signal[right_idxs])])

    if right_idx <= left_idx:
        raise ValueError("Failed to find valid left/right minima (indices overlap).")

    baseline = _linear_baseline_from_indices(v, signal, left_idx, right_idx)

    return {
        "y_corrected": signal - baseline,
        "local_baseline": baseline,
        "left_idx": left_idx,
        "right_idx": right_idx,
        "left_local_min_candidates": np.array([], dtype=int),
        "right_local_min_candidates": np.array([], dtype=int),
        "minima_mode": "argmin_window",
    }


def rotate_offset_using_prominent_bracketing_minima(
    voltage: np.ndarray,
    y: np.ndarray,
    peak_idx: int,
    search_window_v: float = 0.12,
    distance: int = 3,
    require_local_minima_on_both_sides: bool = False,
) -> Dict[str, object]:
    del require_local_minima_on_both_sides
    try:
        from scipy.signal import find_peaks as scipy_find_peaks
    except Exception:
        scipy_find_peaks = None

    v = np.asarray(voltage, dtype=float)
    signal = np.asarray(y, dtype=float)

    if len(v) < 5:
        raise ValueError("Too few points to compute bracketing minima baseline.")
    if peak_idx <= 0 or peak_idx >= len(signal) - 1:
        raise ValueError("Peak index is on/near boundary.")

    peak_idx = _select_bracketing_peak_idx(v, signal, int(peak_idx), search_window_v)
    left_window_idxs, right_window_idxs = _window_indices(v, peak_idx, search_window_v)

    y_inv = -signal
    peak_fn = scipy_find_peaks if scipy_find_peaks is not None else fallback_find_peaks
    minima_idxs, props = peak_fn(y_inv, prominence=0, distance=distance)
    minima_idxs = minima_idxs.astype(int)
    prominences = np.asarray(props.get("prominences", np.zeros(len(minima_idxs))), dtype=float)

    if minima_idxs.size == 0:
        fallback = rotate_offset_using_bracketing_minima(v, signal, peak_idx, search_window_v)
        fallback.update(
            {
                "left_local_min_candidates": np.array([], dtype=int),
                "right_local_min_candidates": np.array([], dtype=int),
                "minima_mode": "prominent_local_minima_fallback",
            }
        )
        return fallback

    min_peak_separation_pts = max(distance, _support_points_for_window(v, search_window_v, fraction=0.08))

    def _pick_side(window_idxs: np.ndarray, sign: int) -> Tuple[np.ndarray, np.ndarray]:
        in_window = np.isin(minima_idxs, window_idxs)

        for sep_frac in (1.0, 0.5, 0.0):
            sep = int(round(min_peak_separation_pts * sep_frac))
            dist_ok = (sign * (minima_idxs - peak_idx)) >= sep
            mask = in_window & dist_ok
            candidates = minima_idxs[mask]
            prom = prominences[mask]
            if candidates.size:
                return candidates, prom

        return np.array([], dtype=int), np.array([], dtype=float)

    left_candidates, left_prom = _pick_side(left_window_idxs, sign=-1)
    right_candidates, right_prom = _pick_side(right_window_idxs, sign=+1)

    if left_candidates.size == 0 or right_candidates.size == 0:
        fallback = rotate_offset_using_bracketing_minima(v, signal, peak_idx, search_window_v)
        fallback.update(
            {
                "left_local_min_candidates": left_candidates,
                "right_local_min_candidates": right_candidates,
                "minima_mode": "prominent_local_minima_fallback",
            }
        )
        return fallback

    left_order = np.argsort(-left_prom)
    right_order = np.argsort(-right_prom)
    left_candidates = left_candidates[left_order]
    right_candidates = right_candidates[right_order]
    left_idx = int(left_candidates[0])
    right_idx = int(right_candidates[0])

    if right_idx <= left_idx:
        fallback = rotate_offset_using_bracketing_minima(v, signal, peak_idx, search_window_v)
        fallback.update(
            {
                "left_local_min_candidates": left_candidates,
                "right_local_min_candidates": right_candidates,
                "minima_mode": "prominent_local_minima_fallback",
            }
        )
        return fallback

    baseline = _linear_baseline_from_indices(v, signal, left_idx, right_idx)

    return {
        "y_corrected": signal - baseline,
        "local_baseline": baseline,
        "left_idx": left_idx,
        "right_idx": right_idx,
        "left_local_min_candidates": left_candidates,
        "right_local_min_candidates": right_candidates,
        "minima_mode": "prominent_local_minima",
    }
