"""
Headless SWV analysis entry point for Bayesian optimization integration.
"""

from __future__ import annotations

import argparse
import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

from core.analysis import run_batch


def _clip01(value: float) -> float:
    try:
        value = float(value)
    except Exception:
        return 0.0
    if value < 0.0:
        return 0.0
    if value > 1.0:
        return 1.0
    return value


def _finite(values: Iterable[Any]) -> List[float]:
    out: List[float] = []
    for value in values:
        try:
            f = float(value)
        except Exception:
            continue
        if f == f:
            out.append(f)
    return out


def _median(values: Iterable[Any], default: float = 0.0) -> float:
    vals = sorted(_finite(values))
    if not vals:
        return float(default)
    mid = len(vals) // 2
    if len(vals) % 2:
        return float(vals[mid])
    return float((vals[mid - 1] + vals[mid]) / 2.0)


def _mean(values: Iterable[Any], default: float = 0.0) -> float:
    vals = _finite(values)
    if not vals:
        return float(default)
    return float(sum(vals) / len(vals))


def _std(values: Iterable[Any]) -> float:
    vals = _finite(values)
    if len(vals) < 2:
        return 0.0
    mean = sum(vals) / len(vals)
    return float((sum((v - mean) ** 2 for v in vals) / (len(vals) - 1)) ** 0.5)


def _cv(values: Iterable[Any]) -> float:
    vals = [abs(v) for v in _finite(values)]
    if len(vals) < 2:
        return 0.0
    mean = sum(vals) / len(vals)
    if mean <= 1e-12:
        return 0.0
    return _std(vals) / mean


def _score_from_cv(values: Iterable[Any]) -> float:
    return _clip01(1.0 / (1.0 + _cv(values)))


def _float_or_none(value: Any) -> Optional[float]:
    try:
        f = float(value)
    except Exception:
        return None
    return f if f == f else None


def _abs_finite(values: Iterable[Any]) -> List[float]:
    out = []
    for value in values:
        f = _float_or_none(value)
        if f is not None:
            out.append(abs(f))
    return out


def _mean(values: Iterable[Any], default: float = 0.0) -> float:
    vals = _finite(values)
    if not vals:
        return default
    return sum(vals) / len(vals)


def _build_channel_metrics(results: List[dict]) -> Dict[str, dict]:
    grouped: Dict[str, List[dict]] = {}
    for row in results:
        channel = row.get("channel")
        if channel is None:
            continue
        text = str(channel).strip().lower()
        try:
            key = str(int(float(text)))
        except (TypeError, ValueError):
            key = text
        grouped.setdefault(key, []).append(row)

    metrics: Dict[str, dict] = {}
    for channel, rows in grouped.items():
        total = len(rows)
        ok_rows = [r for r in rows if str(r.get("status", "")).upper() == "OK"]
        success_score = len(ok_rows) / total if total else 0.0
        if not ok_rows:
            metrics[channel] = {
                "snr": 0.0,
                "peak_prominence": 0.0,
                "repeat_scan_snr": 0.0,
                "peak_shape_score": 0.0,
                "baseline_stability_score": 0.0,
                "replicate_consistency_score": 0.0,
                "success_score": _clip01(success_score),
                "ok_scan_count": 0,
                "total_scan_count": total,
                "std_peak_current_uA": 0.0,
                "peak_currents_uA": [],
                "std_background_rms_uA": 0.0,
                "repeat_relative_std": 0.0,
            }
            continue

        peak_currents = _abs_finite(r.get("peak_current", 0.0) for r in ok_rows)
        background_rms = _abs_finite(r.get("background_current_rms", 0.0) for r in ok_rows)
        bracket_point_counts = [
            max(1.0, _float_or_none(r.get("bracket_point_count")) or 1.0)
            for r in ok_rows
        ]
        crop_point_counts = [
            max(1.0, _float_or_none(r.get("crop_point_count")) or 1.0)
            for r in ok_rows
        ]
        offset_scores = [
            max(0.0, 1.0 - abs(_float_or_none(r.get("peak_offset_norm")) or 0.0))
            for r in ok_rows
        ]
        width_scores = _score_from_cv(r.get("bracket_width_V", 0.0) for r in ok_rows)
        baseline_scores = _score_from_cv(background_rms)
        replicate_scores = _score_from_cv(peak_currents)
        mean_peak_current = _mean(peak_currents, 0.0)
        mean_background_rms = _mean(background_rms, 0.0)
        std_peak_current = _std(peak_currents)
        std_background_rms = _std(background_rms)
        repeat_relative_std = 0.5 * (
            _cv(peak_currents) + _cv(background_rms)
        )
        mean_bracket_point_count = _mean(bracket_point_counts, 1.0)
        mean_crop_point_count = _mean(crop_point_counts, 1.0)
        snr_unadjusted = mean_peak_current / max(mean_background_rms, 1e-12)
        repeat_scan_snr = (
            mean_peak_current / max(std_peak_current, 1e-12)
            if len(peak_currents) > 1
            else 0.0
        )
        metrics[channel] = {
            "snr": snr_unadjusted,
            "snr_unadjusted": snr_unadjusted,
            "peak_prominence": snr_unadjusted,
            "repeat_scan_snr": repeat_scan_snr,
            "peak_shape_score": _clip01(0.5 * _mean(offset_scores, 0.0) + 0.5 * width_scores),
            "baseline_stability_score": _clip01(baseline_scores),
            "replicate_consistency_score": _clip01(replicate_scores),
            "success_score": _clip01(success_score),
            "ok_scan_count": len(ok_rows),
            "total_scan_count": total,
            "mean_bracket_point_count": mean_bracket_point_count,
            "mean_crop_point_count": mean_crop_point_count,
            "median_peak_current_uA": _median(peak_currents, 0.0),
            "mean_peak_current_uA": mean_peak_current,
            "std_peak_current_uA": std_peak_current,
            "peak_currents_uA": peak_currents,
            "median_background_rms_uA": _median(background_rms, 0.0),
            "mean_background_rms_uA": mean_background_rms,
            "std_background_rms_uA": std_background_rms,
            "repeat_relative_std": repeat_relative_std,
        }
    return metrics


def _normalize_scan_windows(raw: str) -> Optional[Tuple[Tuple[int, int], ...]]:
    text = (raw or "").strip()
    if not text:
        return None
    windows = []
    for token in text.replace(";", ",").split(","):
        token = token.strip()
        if not token:
            continue
        if "-" not in token:
            raise ValueError(f"Invalid scan window token: {token}")
        start_s, end_s = token.split("-", 1)
        start, end = int(start_s.strip()), int(end_s.strip())
        if end <= start:
            raise ValueError(f"Scan window end must be greater than start: {token}")
        windows.append((start, end))
    return tuple(windows) if windows else None


def _analysis_args_from_request(payload: dict) -> dict:
    analysis = dict(payload.get("analysis") or {})
    crop_min = float(analysis.get("crop_min_v", -0.61))
    crop_max = float(analysis.get("crop_max_v", -0.30))
    return {
        "folders": [str(Path(p)) for p in payload.get("folders", [])],
        "crop_range": (crop_min, crop_max),
        "smooth_window": int(analysis.get("smooth_window", 15)),
        "smooth_polyorder": int(analysis.get("smooth_polyorder", 2)),
        "minima_search_window_V": float(analysis.get("minima_search_window_v", 0.30)),
        "use_prominent_minima": bool(analysis.get("use_prominent_minima", False)),
        "use_double_correction": bool(analysis.get("use_double_correction", False)),
        "min_peak_height_uA": (
            0.001 if analysis.get("min_peak_height_ua") in (None, "", "none")
            else float(analysis.get("min_peak_height_ua"))
        ),
        "require_local_minima_on_both_sides": bool(analysis.get("require_local_minima_on_both_sides", False)),
        "peak_voltage_min_V": (
            None if analysis.get("peak_voltage_min_v") in (None, "", "none")
            else float(analysis.get("peak_voltage_min_v"))
        ),
        "peak_voltage_max_V": (
            None if analysis.get("peak_voltage_max_v") in (None, "", "none")
            else float(analysis.get("peak_voltage_max_v"))
        ),
        "left_min_voltage_min_V": (
            None if analysis.get("left_min_voltage_min_v") in (None, "", "none")
            else float(analysis.get("left_min_voltage_min_v"))
        ),
        "left_min_voltage_max_V": (
            None if analysis.get("left_min_voltage_max_v") in (None, "", "none")
            else float(analysis.get("left_min_voltage_max_v"))
        ),
        "right_min_voltage_min_V": (
            None if analysis.get("right_min_voltage_min_v") in (None, "", "none")
            else float(analysis.get("right_min_voltage_min_v"))
        ),
        "right_min_voltage_max_V": (
            None if analysis.get("right_min_voltage_max_v") in (None, "", "none")
            else float(analysis.get("right_min_voltage_max_v"))
        ),
        "min_start_voltage": float(analysis.get("min_start_voltage_v", -0.70)),
        "scan_windows": _normalize_scan_windows(str(analysis.get("scan_windows", ""))),
        "scan_range": None,
        "compute_skew": bool(analysis.get("compute_skew", True)),
        "compute_wavelet_energy": bool(analysis.get("compute_wavelet_energy", True)),
        "compute_wavelet_denoised_trace": bool(analysis.get("compute_wavelet_denoised_trace", False)),
        "use_wavelet_for_correction": bool(analysis.get("use_wavelet_for_correction", False)),
    }


def run_request(payload: dict) -> dict:
    args = _analysis_args_from_request(payload)
    folders = args.pop("folders")
    if not folders:
        raise ValueError("At least one folder is required")
    results = run_batch(folders=folders, **args)
    channel_metrics = _build_channel_metrics(results)
    output_dir = Path(payload.get("output_dir") or Path(folders[0]) / "bo_analysis")
    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    stem = str(payload.get("output_stem") or f"bo_analysis_{timestamp}")
    results_csv = output_dir / f"{stem}_results.csv"
    _write_results_csv(results_csv, results)
    summary = {
        "schema_version": 1,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "folders": folders,
        "analysis": dict(payload.get("analysis") or {}),
        "result_count": len(results),
        "channel_metrics": channel_metrics,
        "results_csv": str(results_csv),
    }
    summary_path = output_dir / f"{stem}.json"
    tmp_summary_path = summary_path.with_name(f".{summary_path.name}.tmp")
    with open(tmp_summary_path, "w", encoding="utf-8") as fh:
        json.dump(summary, fh, indent=2)
        fh.flush()
        os.fsync(fh.fileno())
    os.replace(tmp_summary_path, summary_path)
    summary["summary_path"] = str(summary_path)
    return summary


def _write_results_csv(path: Path, rows: List[dict]) -> None:
    import csv

    fieldnames: List[str] = []
    for row in rows:
        for key in row:
            if key not in fieldnames:
                fieldnames.append(key)
    tmp_path = path.with_name(f".{path.name}.tmp")
    with open(tmp_path, "w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames or ["empty"])
        writer.writeheader()
        for row in rows:
            writer.writerow({key: _csv_value(row.get(key)) for key in fieldnames})
        fh.flush()
        os.fsync(fh.fileno())
    os.replace(tmp_path, path)


def _csv_value(value: Any) -> Any:
    try:
        import numpy as np

        if isinstance(value, np.ndarray):
            return value.tolist()
        if isinstance(value, np.generic):
            return value.item()
    except Exception:
        pass
    return value


def main() -> int:
    parser = argparse.ArgumentParser(description="Headless BO analysis runner.")
    parser.add_argument("--request", required=True, help="Path to request JSON.")
    args = parser.parse_args()
    with open(args.request, "r", encoding="utf-8") as fh:
        payload = json.load(fh)
    summary = run_request(payload)
    print(summary["summary_path"])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
