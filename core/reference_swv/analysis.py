import os
from typing import Dict, List, Optional, Tuple

import numpy as np
try:
    import pywt
except Exception:
    pywt = None
try:
    from scipy.stats import skew
except Exception:
    skew = None

from .io import (
    SWVFile,
    collect_swv_csvs_from_folders,
    filter_finite,
    group_by_channel_and_sort,
    load_swv_csv,
)
from .processing import (
    apply_smoothing,
    detect_dominant_peak,
    rotate_offset_using_prominent_bracketing_minima,
    rotate_offset_using_bracketing_minima,
)


def _run_correction_pass(
    v: np.ndarray,
    y_for_correction: np.ndarray,
    smooth_window: int,
    smooth_polyorder: int,
    minima_search_window_V: float,
    use_prominent_minima: bool,
    peak_source: Optional[np.ndarray] = None,
    peak_idx: Optional[int] = None,
) -> dict:
    y_corr_input = np.asarray(y_for_correction, dtype=float)
    peak_signal = np.asarray(peak_source if peak_source is not None else y_corr_input, dtype=float)
    selected_peak_idx = int(detect_dominant_peak(peak_signal) if peak_idx is None else peak_idx)

    corr = (
        rotate_offset_using_prominent_bracketing_minima(v, y_corr_input, selected_peak_idx, minima_search_window_V)
        if use_prominent_minima
        else rotate_offset_using_bracketing_minima(v, y_corr_input, selected_peak_idx, minima_search_window_V)
    )
    y_corr = np.asarray(corr["y_corrected"], dtype=float)
    y_corr_smooth = (
        apply_smoothing(y_corr, smooth_window, smooth_polyorder)
        if smooth_window > 0 else y_corr.copy()
    )
    left_idx, right_idx = int(corr["left_idx"]), int(corr["right_idx"])
    segment = y_corr_smooth[left_idx:right_idx + 1]
    peak_idx_corr = left_idx + detect_dominant_peak(segment, boundary_margin=0)

    return {
        "peak_idx": selected_peak_idx,
        "peak_idx_corr": int(peak_idx_corr),
        "corrected_current": y_corr,
        "smoothed_corrected_current": y_corr_smooth,
        "local_baseline": np.asarray(corr["local_baseline"], dtype=float),
        "left_idx": left_idx,
        "right_idx": right_idx,
        "left_local_min_candidates": np.asarray(corr.get("left_local_min_candidates", []), dtype=int),
        "right_local_min_candidates": np.asarray(corr.get("right_local_min_candidates", []), dtype=int),
        "minima_mode": corr.get("minima_mode", "argmin_window"),
    }


def analyze_swv_file(
    filepath: str,
    crop_range: Tuple[float, float] = (-0.6, -0.2),
    voltage_col: str = "Potential (V)",
    current_col: Optional[str] = None,
    smooth_window: int = 9,
    smooth_polyorder: int = 2,
    minima_search_window_V: float = 0.30,
    use_prominent_minima: bool = False,
    use_double_correction: bool = False,
    min_peak_height_uA: Optional[float] = None,
    compute_skew: bool = True,
    compute_wavelet_energy: bool = True,
) -> dict:
    v_raw, i_raw = load_swv_csv(filepath, voltage_col=voltage_col, current_col=current_col)
    v_raw, i_raw = filter_finite(v_raw, i_raw)

    return analyze_swv_arrays(
        v_raw=v_raw,
        i_raw=i_raw,
        crop_range=crop_range,
        smooth_window=smooth_window,
        smooth_polyorder=smooth_polyorder,
        minima_search_window_V=minima_search_window_V,
        use_prominent_minima=use_prominent_minima,
        use_double_correction=use_double_correction,
        min_peak_height_uA=min_peak_height_uA,
        compute_skew=compute_skew,
        compute_wavelet_energy=compute_wavelet_energy,
        file_path=filepath,
    )


def analyze_swv_arrays(
    v_raw: np.ndarray,
    i_raw: np.ndarray,
    crop_range: Tuple[float, float] = (-0.6, -0.2),
    smooth_window: int = 9,
    smooth_polyorder: int = 2,
    minima_search_window_V: float = 0.30,
    use_prominent_minima: bool = False,
    use_double_correction: bool = False,
    min_peak_height_uA: Optional[float] = None,
    compute_skew: bool = True,
    compute_wavelet_energy: bool = True,
    file_path: Optional[str] = None,
) -> dict:
    mask = (v_raw >= crop_range[0]) & (v_raw <= crop_range[1])
    v, i = v_raw[mask], i_raw[mask]

    if len(v) < 5:
        raise ValueError("Too few points after cropping.")

    i_smooth = apply_smoothing(i, smooth_window, smooth_polyorder) if smooth_window > 0 else i.copy()
    first_pass = _run_correction_pass(
        v=v,
        y_for_correction=i_smooth,
        smooth_window=smooth_window,
        smooth_polyorder=smooth_polyorder,
        minima_search_window_V=minima_search_window_V,
        use_prominent_minima=use_prominent_minima,
    )
    final_pass = first_pass
    second_pass = None
    double_correction_error = None
    if use_double_correction:
        try:
            second_pass = _run_correction_pass(
                v=v,
                y_for_correction=first_pass["corrected_current"],
                peak_source=first_pass["smoothed_corrected_current"],
                smooth_window=smooth_window,
                smooth_polyorder=smooth_polyorder,
                minima_search_window_V=minima_search_window_V,
                use_prominent_minima=use_prominent_minima,
            )
            final_pass = second_pass
        except Exception as exc:
            double_correction_error = str(exc)

    y_corr = final_pass["corrected_current"]
    y_corr_smooth = final_pass["smoothed_corrected_current"]
    left_idx, right_idx = int(final_pass["left_idx"]), int(final_pass["right_idx"])
    peak_idx_corr = int(final_pass["peak_idx_corr"])
    peak_height = float(y_corr[peak_idx_corr])

    if min_peak_height_uA is not None and peak_height < float(min_peak_height_uA):
        raise ValueError(f"Peak height {peak_height:.4g} uA below cutoff {min_peak_height_uA:.4g} uA")

    wavelet_energy = np.nan
    if compute_wavelet_energy and pywt is not None:
        coeffs = pywt.wavedec(y_corr, "haar", level=3)
        wavelet_energy = float(sum(np.sum(c**2) for c in coeffs))

    if compute_skew and skew is not None:
        skew_val = float(skew(y_corr))
    elif compute_skew:
        mean = float(np.mean(y_corr))
        std = float(np.std(y_corr))
        skew_val = 0.0 if std <= 1e-12 else float(np.mean(((y_corr - mean) / std) ** 3))
    else:
        skew_val = np.nan
    peak_offset_norm = np.nan
    if compute_skew:
        v_left = float(v[left_idx])
        v_right = float(v[right_idx])
        denom = (v_right - v_left) / 2.0
        if denom != 0:
            peak_offset_norm = float((v[peak_idx_corr] - (v_left + v_right) / 2.0) / denom)

    return {
        "file_path": file_path,
        "voltage": v,
        "raw_current": i,
        "smoothed_current": i_smooth,
        "corrected_current": y_corr,
        "smoothed_corrected_current": y_corr_smooth,
        "local_baseline": first_pass["local_baseline"],
        "first_pass_corrected_current": first_pass["corrected_current"] if use_double_correction else None,
        "first_pass_smoothed_corrected_current": first_pass["smoothed_corrected_current"] if use_double_correction else None,
        "first_pass_local_baseline": first_pass["local_baseline"] if use_double_correction else None,
        # Use corrected-trace peak position for peak voltage (and drift downstream)
        "peak_voltage": float(v[peak_idx_corr]),
        "peak_current": peak_height,
        "peak_current_raw": float(i[first_pass["peak_idx"]]),
        "peak_idx": first_pass["peak_idx"],
        "peak_idx_corr": peak_idx_corr,
        "left_min_idx": left_idx,
        "right_min_idx": right_idx,
        "left_local_min_candidates": np.asarray(final_pass["left_local_min_candidates"], dtype=int),
        "right_local_min_candidates": np.asarray(final_pass["right_local_min_candidates"], dtype=int),
        "minima_mode": final_pass["minima_mode"],
        "first_pass_peak_idx": first_pass["peak_idx"] if use_double_correction else None,
        "first_pass_peak_idx_corr": first_pass["peak_idx_corr"] if use_double_correction else None,
        "first_pass_left_min_idx": first_pass["left_idx"] if use_double_correction else None,
        "first_pass_right_min_idx": first_pass["right_idx"] if use_double_correction else None,
        "first_pass_left_local_min_candidates": (
            np.asarray(first_pass["left_local_min_candidates"], dtype=int) if use_double_correction else np.array([], dtype=int)
        ),
        "first_pass_right_local_min_candidates": (
            np.asarray(first_pass["right_local_min_candidates"], dtype=int) if use_double_correction else np.array([], dtype=int)
        ),
        "first_pass_minima_mode": first_pass["minima_mode"] if use_double_correction else None,
        "second_pass_corrected_current": second_pass["corrected_current"] if second_pass is not None else None,
        "second_pass_smoothed_corrected_current": (
            second_pass["smoothed_corrected_current"] if second_pass is not None else None
        ),
        "second_pass_local_baseline": second_pass["local_baseline"] if second_pass is not None else None,
        "second_pass_peak_idx": second_pass["peak_idx"] if second_pass is not None else None,
        "second_pass_peak_idx_corr": second_pass["peak_idx_corr"] if second_pass is not None else None,
        "second_pass_left_min_idx": second_pass["left_idx"] if second_pass is not None else None,
        "second_pass_right_min_idx": second_pass["right_idx"] if second_pass is not None else None,
        "second_pass_left_local_min_candidates": (
            np.asarray(second_pass["left_local_min_candidates"], dtype=int) if second_pass is not None else np.array([], dtype=int)
        ),
        "second_pass_right_local_min_candidates": (
            np.asarray(second_pass["right_local_min_candidates"], dtype=int) if second_pass is not None else np.array([], dtype=int)
        ),
        "second_pass_minima_mode": second_pass["minima_mode"] if second_pass is not None else None,
        "double_correction_requested": bool(use_double_correction),
        "double_correction_applied": bool(second_pass is not None),
        "double_correction_error": double_correction_error,
        "correction_passes": 2 if second_pass is not None else 1,
        "skew": skew_val,
        "peak_offset_norm": peak_offset_norm,
        "wavelet_energy": wavelet_energy,
        "status": "OK",
    }

def partial_traces_for_failure_arrays(
    v_raw: np.ndarray,
    i_raw: np.ndarray,
    crop_range: Tuple[float, float],
    smooth_window: int,
    smooth_polyorder: int,
    minima_search_window_V: float,
    use_prominent_minima: bool,
    use_double_correction: bool,
) -> dict:
    base = dict(voltage=None, raw_current=None, smoothed_current=None,
                smoothed_corrected_current=None,
                corrected_current=None, local_baseline=None,
                peak_idx=None, peak_idx_corr=None, left_min_idx=None, right_min_idx=None,
                left_local_min_candidates=np.array([], dtype=int),
                right_local_min_candidates=np.array([], dtype=int),
                minima_mode=None,
                first_pass_corrected_current=None,
                first_pass_smoothed_corrected_current=None,
                first_pass_local_baseline=None,
                first_pass_peak_idx=None,
                first_pass_peak_idx_corr=None,
                first_pass_left_min_idx=None,
                first_pass_right_min_idx=None,
                first_pass_left_local_min_candidates=np.array([], dtype=int),
                first_pass_right_local_min_candidates=np.array([], dtype=int),
                first_pass_minima_mode=None,
                second_pass_corrected_current=None,
                second_pass_smoothed_corrected_current=None,
                second_pass_local_baseline=None,
                second_pass_peak_idx=None,
                second_pass_peak_idx_corr=None,
                second_pass_left_min_idx=None,
                second_pass_right_min_idx=None,
                second_pass_left_local_min_candidates=np.array([], dtype=int),
                second_pass_right_local_min_candidates=np.array([], dtype=int),
                second_pass_minima_mode=None,
                double_correction_requested=bool(use_double_correction),
                double_correction_applied=False,
                double_correction_error=None,
                correction_passes=1)
    try:
        mask = (v_raw >= crop_range[0]) & (v_raw <= crop_range[1])
        v, i = v_raw[mask], i_raw[mask]
        base.update(voltage=v, raw_current=i)

        if len(v) < 5:
            return {**base, "partial_error": "Too few points after cropping."}

        i_smooth = apply_smoothing(i, smooth_window, smooth_polyorder) if smooth_window > 0 else i.copy()
        base["smoothed_current"] = i_smooth

        first_pass = _run_correction_pass(
            v=v,
            y_for_correction=i_smooth,
            smooth_window=smooth_window,
            smooth_polyorder=smooth_polyorder,
            minima_search_window_V=minima_search_window_V,
            use_prominent_minima=use_prominent_minima,
        )
        final_pass = first_pass
        second_pass = None
        double_correction_error = None
        if use_double_correction:
            try:
                second_pass = _run_correction_pass(
                    v=v,
                    y_for_correction=first_pass["corrected_current"],
                    peak_source=first_pass["smoothed_corrected_current"],
                    smooth_window=smooth_window,
                    smooth_polyorder=smooth_polyorder,
                    minima_search_window_V=minima_search_window_V,
                    use_prominent_minima=use_prominent_minima,
                )
                final_pass = second_pass
            except Exception as exc:
                double_correction_error = str(exc)

        return {
            **base,
            "corrected_current": final_pass["corrected_current"],
            "smoothed_corrected_current": final_pass["smoothed_corrected_current"],
            "local_baseline": first_pass["local_baseline"],
            "peak_idx": first_pass["peak_idx"],
            "peak_idx_corr": final_pass["peak_idx_corr"],
            "left_min_idx": int(final_pass["left_idx"]),
            "right_min_idx": int(final_pass["right_idx"]),
            "left_local_min_candidates": np.asarray(final_pass["left_local_min_candidates"], dtype=int),
            "right_local_min_candidates": np.asarray(final_pass["right_local_min_candidates"], dtype=int),
            "minima_mode": final_pass["minima_mode"],
            "first_pass_corrected_current": first_pass["corrected_current"] if use_double_correction else None,
            "first_pass_smoothed_corrected_current": first_pass["smoothed_corrected_current"] if use_double_correction else None,
            "first_pass_local_baseline": first_pass["local_baseline"] if use_double_correction else None,
            "first_pass_peak_idx": first_pass["peak_idx"] if use_double_correction else None,
            "first_pass_peak_idx_corr": first_pass["peak_idx_corr"] if use_double_correction else None,
            "first_pass_left_min_idx": first_pass["left_idx"] if use_double_correction else None,
            "first_pass_right_min_idx": first_pass["right_idx"] if use_double_correction else None,
            "first_pass_left_local_min_candidates": (
                np.asarray(first_pass["left_local_min_candidates"], dtype=int) if use_double_correction else np.array([], dtype=int)
            ),
            "first_pass_right_local_min_candidates": (
                np.asarray(first_pass["right_local_min_candidates"], dtype=int) if use_double_correction else np.array([], dtype=int)
            ),
            "first_pass_minima_mode": first_pass["minima_mode"] if use_double_correction else None,
            "second_pass_corrected_current": second_pass["corrected_current"] if second_pass is not None else None,
            "second_pass_smoothed_corrected_current": (
                second_pass["smoothed_corrected_current"] if second_pass is not None else None
            ),
            "second_pass_local_baseline": second_pass["local_baseline"] if second_pass is not None else None,
            "second_pass_peak_idx": second_pass["peak_idx"] if second_pass is not None else None,
            "second_pass_peak_idx_corr": second_pass["peak_idx_corr"] if second_pass is not None else None,
            "second_pass_left_min_idx": second_pass["left_idx"] if second_pass is not None else None,
            "second_pass_right_min_idx": second_pass["right_idx"] if second_pass is not None else None,
            "second_pass_left_local_min_candidates": (
                np.asarray(second_pass["left_local_min_candidates"], dtype=int) if second_pass is not None else np.array([], dtype=int)
            ),
            "second_pass_right_local_min_candidates": (
                np.asarray(second_pass["right_local_min_candidates"], dtype=int) if second_pass is not None else np.array([], dtype=int)
            ),
            "second_pass_minima_mode": second_pass["minima_mode"] if second_pass is not None else None,
            "double_correction_applied": bool(second_pass is not None),
            "double_correction_error": double_correction_error,
            "correction_passes": 2 if second_pass is not None else 1,
            "partial_error": None,
        }
    except Exception as e:
        return {**base, "partial_error": str(e)}


def compute_drift_fields(all_results: List[dict]) -> List[dict]:
    """
    Adds two drift fields to each result (in-place), computed per channel
    relative to each channel's first valid (OK) scan:

      peak_voltage_drift           peak_voltage               - reference peak_voltage  (V)
      skew_drift                   skew                       - reference skew
      peak_offset_norm_drift        peak_offset_norm          - reference peak_offset_norm
    """
    ref: Dict[int, dict] = {}

    # Sort globally so we always pick the lowest scan_number as reference
    sorted_results = sorted(
        all_results, key=lambda r: (str(r["channel"]), r["scan_number"])
    )

    for r in sorted_results:
        ch = r["channel"]
        if r.get("status") != "OK":
            r["peak_voltage_drift"] = np.nan
            r["skew_drift"] = np.nan
            r["peak_offset_norm_drift"] = np.nan
            continue

        if ch not in ref:
            ref[ch] = r  # first OK scan for this channel = reference

        r["peak_voltage_drift"] = r["peak_voltage"] - ref[ch]["peak_voltage"]
        r["skew_drift"]         = r["skew"]         - ref[ch]["skew"]
        r["peak_offset_norm_drift"] = r["peak_offset_norm"] - ref[ch]["peak_offset_norm"]

    return all_results


def run_batch(
    folders: List[str],
    crop_range: Tuple[float, float] = (-0.6, -0.2),
    voltage_col: str = "Potential (V)",
    current_col: Optional[str] = None,
    smooth_window: int = 9,
    smooth_polyorder: int = 2,
    minima_search_window_V: float = 0.30,
    use_prominent_minima: bool = False,
    use_double_correction: bool = False,
    min_peak_height_uA: Optional[float] = None,
    min_start_voltage: float = -0.6,
    scan_range: Optional[Tuple[int, int]] = None,
    compute_skew: bool = True,
    compute_wavelet_energy: bool = True,
    progress_callback=None,
) -> List[dict]:
    files = collect_swv_csvs_from_folders(folders)
    if not files:
        raise ValueError("No SWV CSVs found.")

    by_ch = group_by_channel_and_sort(files)
    all_results: List[dict] = []

    ordered: List[Tuple[int, SWVFile]] = [
        (ch, f)
        for ch, flist in sorted(by_ch.items(), key=lambda item: str(item[0]))
        for f in flist
    ]

    total = len(ordered)
    scan_counters: Dict[object, int] = {}

    for idx, (ch, f) in enumerate(ordered):
        if progress_callback:
            progress_callback(idx + 1, total, os.path.basename(f.path))

        try:
            v_check, i_check = load_swv_csv(f.path, voltage_col=voltage_col, current_col=current_col)
            v_check, i_check = filter_finite(v_check, i_check)
        except Exception:
            continue

        if len(v_check) == 0 or float(v_check[0]) < float(min_start_voltage):
            continue

        # Skip files that have no data points within the crop range (e.g. LSV sweeps
        # that cover a completely different voltage window than the SWV crop range).
        in_crop = (v_check >= crop_range[0]) & (v_check <= crop_range[1])
        if in_crop.sum() < 5:
            continue

        scan_counters[ch] = scan_counters.get(ch, 0) + 1
        scan_number = scan_counters[ch]

        # If a scan_range filter is active, skip analysis+storage for out-of-range
        # scans BUT only after the counter has been incremented so numbering stays
        # consistent with the full dataset.
        if scan_range is not None and not (scan_range[0] <= scan_number <= scan_range[1]):
            continue

        common = dict(
            channel=ch,
            channel_label=f"Ch{ch}",
            timestamp=f.ts,
            scan_id_from_name=f.scan,
            scan_number=scan_number,
            folder_index=f.folder_index,
            file_path=f.path,
            file_name=os.path.basename(f.path),
        )

        try:
            r = analyze_swv_arrays(
                v_raw=v_check,
                i_raw=i_check,
                crop_range=crop_range,
                smooth_window=smooth_window,
                smooth_polyorder=smooth_polyorder,
                minima_search_window_V=minima_search_window_V,
                use_prominent_minima=use_prominent_minima,
                use_double_correction=use_double_correction,
                min_peak_height_uA=min_peak_height_uA,
                compute_skew=compute_skew,
                compute_wavelet_energy=compute_wavelet_energy,
                file_path=f.path,
            )
            r.update(common)
            all_results.append(r)

        except Exception as e:
            partial = partial_traces_for_failure_arrays(
                v_raw=v_check,
                i_raw=i_check,
                crop_range=crop_range,
                smooth_window=smooth_window,
                smooth_polyorder=smooth_polyorder,
                minima_search_window_V=minima_search_window_V,
                use_prominent_minima=use_prominent_minima,
                use_double_correction=use_double_correction,
            )
            all_results.append({
                **common,
                "peak_current": np.nan,
                "peak_current_raw": np.nan,
                "peak_voltage": np.nan,
                "skew": np.nan,
                "peak_offset_norm": np.nan,
                "wavelet_energy": np.nan,
                "status": "FAILED",
                "error": str(e),
                **{k: partial.get(k) for k in (
                    "voltage", "raw_current", "smoothed_current",
                    "corrected_current", "smoothed_corrected_current",
                    "local_baseline", "partial_error",
                    "left_min_idx", "right_min_idx", "peak_idx", "peak_idx_corr",
                    "left_local_min_candidates", "right_local_min_candidates",
                    "minima_mode", "first_pass_corrected_current",
                    "first_pass_smoothed_corrected_current", "first_pass_local_baseline",
                    "first_pass_peak_idx", "first_pass_peak_idx_corr",
                    "first_pass_left_min_idx", "first_pass_right_min_idx",
                    "first_pass_left_local_min_candidates", "first_pass_right_local_min_candidates",
                    "first_pass_minima_mode", "second_pass_corrected_current",
                    "second_pass_smoothed_corrected_current", "second_pass_local_baseline",
                    "second_pass_peak_idx", "second_pass_peak_idx_corr",
                    "second_pass_left_min_idx", "second_pass_right_min_idx",
                    "second_pass_left_local_min_candidates", "second_pass_right_local_min_candidates",
                    "second_pass_minima_mode", "double_correction_requested",
                    "double_correction_applied", "double_correction_error",
                    "correction_passes",
                )},
            })

    # Compute drift relative to each channel's first valid scan
    compute_drift_fields(all_results)

    return all_results
