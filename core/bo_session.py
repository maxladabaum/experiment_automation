"""
core/bo_session.py - Closed-loop SWV Bayesian optimization integration.

The GUI-facing contract is intentionally small:

* load a user-editable JSON configuration
* propose one valid SWV method for a mux batch
* create normal queue items through MethodRegistry
* import external analysis JSON outputs
* retain publication-grade records inside the active experiment folder

The external analysis app is the source of per-channel metrics. This module
only scores those metrics and updates the optimizer state.
"""

from __future__ import annotations

import csv
import hashlib
import itertools
import json
import math
import pickle
import random
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

from config import BO_ANALYSIS_FILE_GLOB, BO_ANALYSIS_OUTPUT_DIR, BO_DEFAULT_CONFIG_PATH
from core.mscript_parser import to_si_string


PARAMETER_ORDER = (
    "begin_potential",
    "end_potential",
    "step_potential",
    "amplitude",
    "frequency",
    "conditioning_potential",
    "conditioning_time",
)

OPTIMIZER_ORDER = (
    "begin_potential",
    "end_potential",
    "step_potential",
    "amplitude",
    "frequency",
    "conditioning_potential",
    "conditioning_time",
)


DEFAULT_INITIAL_METHOD = {
    "begin_potential": -0.7,
    "end_potential": -0.1,
    "step_potential": 0.002,
    "amplitude": 0.036,
    "frequency": 200.0,
    "conditioning_potential": -0.7,
    "conditioning_time": 0.0,
}


@dataclass
class BOSuggestion:
    iteration: int
    method_id: str
    params: Dict[str, float]
    created_at: str
    status: str = "suggested"


def load_bo_config(path: Optional[str | Path] = None) -> dict:
    config_path = Path(path) if path else Path(BO_DEFAULT_CONFIG_PATH)
    with open(config_path, "r", encoding="utf-8") as fh:
        config = json.load(fh)
    if not isinstance(config, dict):
        raise ValueError("BO config must be a JSON object")
    return normalize_bo_config(config)


def normalize_bo_config(config: dict) -> dict:
    cfg = dict(config)
    cfg.setdefault("schema_version", 1)
    cfg.setdefault("name", "SWV mux Bayesian optimization")
    cfg.setdefault("max_iterations", 20)
    cfg.setdefault("n_initial_points", 8)
    cfg.setdefault("random_seed", 42)
    cfg.setdefault("channels", list(range(1, 11)))
    cfg.setdefault("method_options", {})
    cfg["method_options"].setdefault("bandwidth", "4k")
    cfg["method_options"].setdefault(
        "ba_range",
        {"mode": "fixed", "fixed": "100n", "auto_min": "100n", "auto_max": "100n"},
    )
    if "initial_parameters" not in cfg:
        if isinstance(cfg.get("initial_method"), dict):
            cfg["initial_parameters"] = dict(cfg["initial_method"])
        elif isinstance(cfg.get("initial_design"), list) and cfg["initial_design"]:
            first = cfg["initial_design"][0]
            cfg["initial_parameters"] = dict(first) if isinstance(first, dict) else dict(DEFAULT_INITIAL_METHOD)
        else:
            cfg["initial_parameters"] = dict(DEFAULT_INITIAL_METHOD)
    cfg.pop("initial_method", None)
    cfg.pop("initial_design", None)
    cfg.setdefault("parameters", {})
    for name in PARAMETER_ORDER:
        p = dict(cfg["parameters"].get(name) or {})
        p.setdefault("label", name.replace("_", " ").title())
        p.setdefault("mode", "locked")
        p.setdefault("value", cfg["initial_parameters"].get(name, DEFAULT_INITIAL_METHOD[name]))
        p.setdefault("values", [p["value"]])
        if name == "conditioning_potential":
            p.setdefault("tie_to", "begin_potential")
        cfg["parameters"][name] = p
    cfg.setdefault("constraints", {})
    cfg["constraints"].setdefault("min_scan_window", 0.4)
    cfg["constraints"].setdefault("max_effective_scan_rate", 1.0)
    cfg["constraints"].setdefault("require_end_after_begin", True)
    cfg["constraints"].setdefault("conditioning_potential_tied_by_default", True)
    cfg.setdefault("scoring", {})
    cfg["scoring"].setdefault(
        "channel_weights",
        {
            "snr": 0.35,
            "peak_shape": 0.20,
            "baseline": 0.20,
            "replicate_consistency": 0.15,
            "success": 0.10,
            "snr_saturation": 20.0,
        },
    )
    cfg["scoring"].setdefault(
        "run_weights",
        {
            "lambda_variability": 0.20,
            "lambda_failed": 0.40,
            "lambda_low": 0.20,
            "low_channel_threshold": 0.50,
        },
    )
    cfg.setdefault("analysis", {})
    cfg["analysis"].setdefault("file_glob", BO_ANALYSIS_FILE_GLOB)
    cfg["analysis"].setdefault("copy_outputs_into_record", True)
    cfg["analysis"].setdefault("crop_min_v", -0.6)
    cfg["analysis"].setdefault("crop_max_v", -0.1)
    cfg["analysis"].setdefault("smooth_window", 15)
    cfg["analysis"].setdefault("smooth_polyorder", 2)
    cfg["analysis"].setdefault("minima_search_window_v", 0.30)
    cfg["analysis"].setdefault("use_prominent_minima", False)
    cfg["analysis"].setdefault("use_double_correction", True)
    cfg["analysis"].setdefault("min_peak_height_ua", None)
    cfg["analysis"].setdefault("min_start_voltage_v", -0.6)
    cfg["analysis"].setdefault("scan_windows", "")
    cfg["analysis"].setdefault("compute_skew", False)
    cfg["analysis"].setdefault("compute_wavelet_energy", False)
    cfg["analysis"].setdefault("compute_wavelet_denoised_trace", False)
    cfg["analysis"].setdefault("use_wavelet_for_correction", False)
    cfg.setdefault("records", {})
    cfg["records"].setdefault("folder_prefix", "bo_session")
    return cfg


def save_bo_config(config: dict, path: str | Path) -> Path:
    out = Path(path)
    out.parent.mkdir(parents=True, exist_ok=True)
    normalized = normalize_bo_config(config)
    with open(out, "w", encoding="utf-8") as fh:
        json.dump(normalized, fh, indent=2)
    return out


def validate_bo_config(config: dict) -> List[str]:
    cfg = normalize_bo_config(config)
    errors: List[str] = []
    try:
        channels = parse_channels(cfg.get("channels", []))
        if not channels:
            errors.append("At least one mux channel is required")
    except ValueError as exc:
        errors.append(str(exc))
    try:
        candidates = generate_candidates(cfg)
        if not candidates:
            errors.append("No valid candidates remain after constraints")
    except Exception as exc:
        errors.append(f"Candidate generation failed: {exc}")
    try:
        candidate_errors = validate_candidate(resolve_initial_parameters(cfg), cfg)
        if candidate_errors:
            errors.append(f"Initial parameters are invalid: {'; '.join(candidate_errors)}")
    except Exception as exc:
        errors.append(f"Initial parameter validation failed: {exc}")
    return errors


def parse_channels(value: Any) -> List[int]:
    if isinstance(value, str):
        tokens = value.replace(";", ",").split(",")
        channels = [int(tok.strip()) for tok in tokens if tok.strip()]
    else:
        channels = [int(v) for v in value]
    bad = [ch for ch in channels if ch < 1 or ch > 16]
    if bad:
        raise ValueError("Mux channels must be between 1 and 16")
    seen = set()
    ordered = []
    for ch in channels:
        if ch not in seen:
            ordered.append(ch)
            seen.add(ch)
    return ordered


def active_parameters(config: dict) -> List[str]:
    params = normalize_bo_config(config)["parameters"]
    return [
        name for name in PARAMETER_ORDER
        if str(params.get(name, {}).get("mode", "")).lower() == "active"
    ]


def generate_candidates(config: dict) -> List[Dict[str, float]]:
    cfg = normalize_bo_config(config)
    params_cfg = cfg["parameters"]
    names_for_product: List[str] = []
    values_for_product: List[List[float]] = []

    for name in PARAMETER_ORDER:
        p_cfg = params_cfg[name]
        mode = str(p_cfg.get("mode", "locked")).lower()
        if mode == "tied":
            continue
        if mode == "active":
            values = _float_values(p_cfg.get("values", []), f"{name}.values")
        else:
            values = [_float_value(p_cfg.get("value", cfg["initial_parameters"].get(name)), name)]
        names_for_product.append(name)
        values_for_product.append(values)

    candidates: List[Dict[str, float]] = []
    seen = set()
    for combo in itertools.product(*values_for_product):
        candidate = dict(zip(names_for_product, combo))
        for name in PARAMETER_ORDER:
            p_cfg = params_cfg[name]
            if str(p_cfg.get("mode", "locked")).lower() == "tied":
                tie_to = p_cfg.get("tie_to") or "begin_potential"
                candidate[name] = float(candidate[tie_to])
        for name in PARAMETER_ORDER:
            candidate.setdefault(name, float(cfg["initial_parameters"].get(name, DEFAULT_INITIAL_METHOD[name])))
        errors = validate_candidate(candidate, cfg)
        if errors:
            continue
        key = candidate_key(candidate)
        if key not in seen:
            candidates.append(candidate)
            seen.add(key)

    initial = resolve_initial_parameters(cfg)
    if not validate_candidate(initial, cfg):
        key = candidate_key(initial)
        candidates = [c for c in candidates if candidate_key(c) != key]
        candidates.insert(0, initial)
    return candidates


def resolve_initial_parameters(config: dict) -> Dict[str, float]:
    return resolve_method_payload(config, normalize_bo_config(config).get("initial_parameters", {}))


def resolve_initial_method(config: dict) -> Dict[str, float]:
    """Backward-compatible alias for older call sites."""
    return resolve_initial_parameters(config)


def resolve_method_payload(config: dict, payload: dict) -> Dict[str, float]:
    cfg = normalize_bo_config(config)
    result = dict(DEFAULT_INITIAL_METHOD)
    result.update({k: float(v) for k, v in cfg.get("initial_parameters", {}).items() if k in PARAMETER_ORDER})
    result.update({k: float(v) for k, v in (payload or {}).items() if k in PARAMETER_ORDER})
    for name, p_cfg in cfg["parameters"].items():
        mode = str(p_cfg.get("mode", "locked")).lower()
        if mode == "locked":
            result[name] = _float_value(p_cfg.get("value", result.get(name)), name)
        elif mode == "tied":
            tie_to = p_cfg.get("tie_to") or "begin_potential"
            result[name] = float(result[tie_to])
    return result


def resolve_initial_design(config: dict) -> List[Dict[str, float]]:
    """Backward-compatible helper: the current app uses one initial_parameters object."""
    return [resolve_initial_parameters(config)]


def validate_candidate(candidate: Dict[str, float], config: dict) -> List[str]:
    cfg = normalize_bo_config(config)
    constraints = cfg["constraints"]
    errors: List[str] = []
    begin = float(candidate["begin_potential"])
    end = float(candidate["end_potential"])
    step = float(candidate["step_potential"])
    frequency = float(candidate["frequency"])
    cond = float(candidate["conditioning_potential"])
    scan_window = end - begin
    if constraints.get("require_end_after_begin", True) and end <= begin:
        errors.append("end_potential must be greater than begin_potential")
    min_window = float(constraints.get("min_scan_window", 0.4))
    if scan_window < min_window - 1e-12:
        errors.append(f"end_potential - begin_potential must be at least {min_window:g} V")
    max_scan_rate = float(constraints.get("max_effective_scan_rate", 1.0))
    if step * frequency > max_scan_rate + 1e-12:
        errors.append(f"step_potential * frequency must be <= {max_scan_rate:g} V/s")
    cond_cfg = cfg["parameters"].get("conditioning_potential", {})
    if str(cond_cfg.get("mode", "")).lower() == "tied" and abs(cond - begin) > 1e-9:
        errors.append("conditioning_potential is tied to begin_potential")
    return errors


def compute_channel_quality(metrics: dict, scoring: dict) -> dict:
    weights = scoring.get("channel_weights", {})
    snr_saturation = float(weights.get("snr_saturation", 20.0))
    if "snr" in metrics:
        snr_raw = float(metrics.get("snr", 0.0))
    else:
        peak_current = abs(float(metrics.get("peak_current", 0.0)))
        baseline_noise = float(metrics.get("baseline_noise", 0.0))
        snr_raw = peak_current / (baseline_noise + 1e-12)

    component_scores = {
        "normalized_SNR": _clip01(snr_raw / max(snr_saturation, 1e-12)),
        "peak_shape_score": _clip01(metrics.get("peak_shape_score", 0.0)),
        "baseline_stability_score": _clip01(metrics.get("baseline_stability_score", 0.0)),
        "replicate_consistency_score": _clip01(metrics.get("replicate_consistency_score", 0.0)),
        "success_score": _clip01(metrics.get("success_score", 1.0)),
    }
    weighted = (
        float(weights.get("snr", 0.35)) * component_scores["normalized_SNR"]
        + float(weights.get("peak_shape", 0.20)) * component_scores["peak_shape_score"]
        + float(weights.get("baseline", 0.20)) * component_scores["baseline_stability_score"]
        + float(weights.get("replicate_consistency", 0.15)) * component_scores["replicate_consistency_score"]
        + float(weights.get("success", 0.10)) * component_scores["success_score"]
    )
    total_weight = (
        float(weights.get("snr", 0.35))
        + float(weights.get("peak_shape", 0.20))
        + float(weights.get("baseline", 0.20))
        + float(weights.get("replicate_consistency", 0.15))
        + float(weights.get("success", 0.10))
    )
    component_scores["Q_channel"] = _clip01(weighted / max(total_weight, 1e-12))
    component_scores["snr_raw"] = snr_raw
    return component_scores


def compute_run_quality(channel_metrics: dict, scoring: dict) -> dict:
    per_channel = {}
    q_values = []
    for channel, metrics in _channel_items(channel_metrics):
        result = compute_channel_quality(metrics, scoring)
        per_channel[str(channel)] = result
        q_values.append(float(result["Q_channel"]))

    run_weights = scoring.get("run_weights", {})
    if not q_values:
        q_values = [0.0]
    mean_q = sum(q_values) / len(q_values)
    std_q = _std(q_values)
    threshold = float(run_weights.get("low_channel_threshold", 0.5))
    failed_fraction = sum(1 for q in q_values if q <= 0.0) / len(q_values)
    low_fraction = sum(1 for q in q_values if q < threshold) / len(q_values)
    q_run = (
        mean_q
        - float(run_weights.get("lambda_variability", 0.20)) * std_q
        - float(run_weights.get("lambda_failed", 0.40)) * failed_fraction
        - float(run_weights.get("lambda_low", 0.20)) * low_fraction
    )
    return {
        "Q_run": _clip01(q_run),
        "mean_Q_channel": mean_q,
        "std_Q_channel": std_q,
        "failed_channel_fraction": failed_fraction,
        "low_channel_fraction": low_fraction,
        "Q_channels": {ch: data["Q_channel"] for ch, data in per_channel.items()},
        "channel_components": per_channel,
    }


def extract_channel_metrics(payload: Any) -> dict:
    if isinstance(payload, dict):
        for key in ("channel_metrics", "channels", "per_channel_metrics"):
            value = payload.get(key)
            if isinstance(value, dict):
                return value
            if isinstance(value, list):
                return {str(i + 1): item for i, item in enumerate(value) if isinstance(item, dict)}
        if payload and all(isinstance(v, dict) for v in payload.values()):
            return payload
    raise ValueError("Analysis JSON must contain channel_metrics, channels, or a channel-keyed metrics object")


class BOIntegrationSession:
    """Runtime state and record keeping for one BO session."""

    STATE_FILE = "bo_state.json"

    def __init__(
        self,
        config: dict,
        experiment_dir: str | Path,
        config_path: Optional[str | Path] = None,
        analysis_output_dir: Optional[str | Path] = None,
    ):
        self.config = normalize_bo_config(config)
        self.config_path = Path(config_path) if config_path else None
        self.experiment_dir = Path(experiment_dir)
        self.analysis_output_dir = Path(analysis_output_dir) if analysis_output_dir else Path(BO_ANALYSIS_OUTPUT_DIR)
        self.session_id = self._build_session_id()
        folder_prefix = self.config.get("records", {}).get("folder_prefix", "bo_session")
        self.record_dir = self._unique_dir(self.experiment_dir / "bo_sessions" / f"{folder_prefix}_{self.session_id}")
        self.methods_dir = self.record_dir / "methods"
        self.analysis_dir = self.record_dir / "analysis"
        self.queue_dir = self.record_dir / "queue"
        self.surrogate_dir = self.record_dir / "surrogate"
        self.acquisition_dir = self.record_dir / "acquisition"
        self.plots_dir = self.record_dir / "plots"
        self.methods_dir.mkdir(parents=True, exist_ok=True)
        self.analysis_dir.mkdir(parents=True, exist_ok=True)
        self.queue_dir.mkdir(parents=True, exist_ok=True)
        self.surrogate_dir.mkdir(parents=True, exist_ok=True)
        self.acquisition_dir.mkdir(parents=True, exist_ok=True)
        self.plots_dir.mkdir(parents=True, exist_ok=True)
        self.candidates = generate_candidates(self.config)
        self.observations: List[dict] = []
        self.suggestions: List[dict] = []
        self.pending: Optional[dict] = None
        self._rng = random.Random(int(self.config.get("random_seed", 42)))
        self._write_session_start_files()
        self.save_state()

    @classmethod
    def start(
        cls,
        config_path: str | Path,
        experiment_dir: str | Path,
        analysis_output_dir: Optional[str | Path] = None,
    ) -> "BOIntegrationSession":
        config = load_bo_config(config_path)
        errors = validate_bo_config(config)
        if errors:
            raise ValueError("; ".join(errors))
        return cls(config, experiment_dir, config_path=config_path, analysis_output_dir=analysis_output_dir)

    def ask_next(self) -> BOSuggestion:
        if self.pending is not None:
            return BOSuggestion(**self.pending)
        if len(self.observations) >= int(self.config.get("max_iterations", 20)):
            raise RuntimeError("Maximum BO iterations reached")
        tried = {candidate_key(obs["params"]) for obs in self.observations}
        available = [c for c in self.candidates if candidate_key(c) not in tried]
        if not available:
            raise RuntimeError("All valid candidates have been evaluated")
        iteration = len(self.observations) + 1
        params = self._choose_candidate(available)
        method_id = f"{self.session_id}_iter_{iteration:03d}"
        suggestion = {
            "iteration": iteration,
            "method_id": method_id,
            "params": params,
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "status": "suggested",
        }
        self.pending = suggestion
        self.suggestions.append(dict(suggestion))
        self._write_json(self.methods_dir / f"iter_{iteration:03d}_suggested_method.json", suggestion)
        self.save_state()
        return BOSuggestion(**suggestion)

    def build_queue_items(self, registry, suggestion: BOSuggestion) -> List[dict]:
        channels = parse_channels(self.config.get("channels", []))
        base_script = build_swv_script(suggestion.params, self.config.get("method_options", {}))
        items = []
        params_for_hash = self._params_for_method_ref(suggestion.params)
        for channel in channels:
            script = wrap_mux(base_script, channel)
            note = f"BO {suggestion.method_id} | MUX ch {channel}"
            saved_path, saved_name = registry.save_script(
                "SWV",
                script,
                params=params_for_hash,
                mux_channel=channel,
                note=note,
            )
            hash_key = "-"
            try:
                hash_key = registry.hash_key_for(saved_path)
            except Exception:
                pass
            item = {
                "type": "SWV",
                "script_path": str(saved_path),
                "status": "pending",
                "details": f"BO iter {suggestion.iteration:03d} | {saved_name} (MUX ch {channel})",
                "method_ref": {
                    "hash_key": hash_key,
                    "technique": "SWV",
                    "params": dict(params_for_hash),
                    "mux_channel": channel,
                },
                "bo_ref": {
                    "session_id": self.session_id,
                    "iteration": suggestion.iteration,
                    "method_id": suggestion.method_id,
                    "record_dir": str(self.record_dir),
                },
            }
            items.append(item)
        self._write_json(
            self.methods_dir / f"iter_{suggestion.iteration:03d}_queue_items.json",
            {"method_id": suggestion.method_id, "items": items},
        )
        return items

    def record_queued(self, suggestion: BOSuggestion, items: List[dict]) -> None:
        if self.pending and self.pending.get("method_id") == suggestion.method_id:
            self.pending["status"] = "queued"
        for record in self.suggestions:
            if record.get("method_id") == suggestion.method_id:
                record["status"] = "queued"
                record["queued_at"] = datetime.now().isoformat(timespec="seconds")
                record["queue_item_count"] = len(items)
        self.save_state()

    def import_analysis(self, path: str | Path, notes: str = "") -> dict:
        if self.pending is None:
            raise RuntimeError("No pending BO suggestion is waiting for analysis")
        source = Path(path)
        if not source.exists():
            raise FileNotFoundError(source)
        with open(source, "r", encoding="utf-8") as fh:
            payload = json.load(fh)
        channel_metrics = extract_channel_metrics(payload)
        quality = compute_run_quality(channel_metrics, self.config.get("scoring", {}))
        iteration = int(self.pending["iteration"])
        retained_path = self._retain_analysis_file(source, iteration)
        observation = {
            "iteration": iteration,
            "method_id": self.pending["method_id"],
            "params": dict(self.pending["params"]),
            "analysis_source": str(source),
            "analysis_record": str(retained_path),
            "channel_metrics": channel_metrics,
            "quality": quality,
            "Q_run": quality["Q_run"],
            "notes": notes,
            "completed_at": datetime.now().isoformat(timespec="seconds"),
        }
        self.observations.append(observation)
        for record in self.suggestions:
            if record.get("method_id") == self.pending["method_id"]:
                record["status"] = "completed"
                record["completed_at"] = observation["completed_at"]
                record["Q_run"] = observation["Q_run"]
        self.pending = None
        self._write_json(self.analysis_dir / f"iter_{iteration:03d}_quality.json", observation)
        self._write_history_csv()
        self._write_surrogate_and_acquisition_artifacts(iteration)
        self._write_plots(observation)
        self.save_state()
        return observation

    def record_queue_completion(self, summary: dict) -> Optional[Path]:
        """Retain queue completion details for BO-owned queue items."""
        if not isinstance(summary, dict):
            return None
        items = []
        for item in summary.get("items", []):
            ref = item.get("bo_ref") if isinstance(item, dict) else None
            if isinstance(ref, dict) and ref.get("session_id") == self.session_id:
                items.append(dict(item))
        if not items:
            return None
        iterations = sorted({int((item.get("bo_ref") or {}).get("iteration", 0)) for item in items})
        iteration = iterations[0] if len(iterations) == 1 else 0
        payload = {
            "session_id": self.session_id,
            "recorded_at": datetime.now().isoformat(timespec="seconds"),
            "queue_summary": {
                "start_index": summary.get("start_index"),
                "total": summary.get("total"),
                "completed": summary.get("completed"),
                "failed": summary.get("failed"),
                "stopped": summary.get("stopped"),
            },
            "bo_iterations": iterations,
            "items": items,
        }
        name = f"iter_{iteration:03d}_queue_completion.json" if iteration else "queue_completion_mixed.json"
        path = self.queue_dir / name
        self._write_json(path, payload)
        for item in items:
            ref = item.get("bo_ref") or {}
            method_id = ref.get("method_id")
            if not method_id:
                continue
            for record in self.suggestions:
                if record.get("method_id") == method_id:
                    record["queue_completed_at"] = payload["recorded_at"]
                    record["queue_completion_record"] = str(path)
                    record["completed_queue_items"] = payload["queue_summary"].get("completed")
                    record["failed_queue_items"] = payload["queue_summary"].get("failed")
        self.save_state()
        return path

    def latest_analysis_file(self) -> Optional[Path]:
        folder = self.analysis_output_dir
        if not folder.exists() or not folder.is_dir():
            return None
        pattern = self.config.get("analysis", {}).get("file_glob") or BO_ANALYSIS_FILE_GLOB
        files = [p for p in folder.glob(pattern) if p.is_file()]
        if not files:
            return None
        return max(files, key=lambda p: p.stat().st_mtime)

    def best_observation(self) -> Optional[dict]:
        if not self.observations:
            return None
        return max(self.observations, key=lambda obs: float(obs.get("Q_run", 0.0)))

    def should_stop(self) -> bool:
        return len(self.observations) >= int(self.config.get("max_iterations", 20))

    def save_state(self) -> Path:
        payload = {
            "session_id": self.session_id,
            "config_path": str(self.config_path) if self.config_path else None,
            "record_dir": str(self.record_dir),
            "analysis_output_dir": str(self.analysis_output_dir),
            "candidate_count": len(self.candidates),
            "active_parameters": active_parameters(self.config),
            "pending": self.pending,
            "suggestions": self.suggestions,
            "observations": self.observations,
            "best_observation": self.best_observation(),
            "updated_at": datetime.now().isoformat(timespec="seconds"),
        }
        path = self.record_dir / self.STATE_FILE
        self._write_json(path, payload)
        return path

    def _choose_candidate(self, available: List[Dict[str, float]]) -> Dict[str, float]:
        tried = {candidate_key(obs["params"]) for obs in self.observations}
        available_keys = {candidate_key(c) for c in available}
        initial = resolve_initial_parameters(self.config)
        key = candidate_key(initial)
        if key not in tried and key in available_keys:
            return initial
        if len(self.observations) < int(self.config.get("n_initial_points", 8)):
            return self._maximin_candidate(available)
        gp_choice = self._gp_expected_improvement_candidate(available)
        if gp_choice is not None:
            return gp_choice
        return self._distance_surrogate_candidate(available)

    def _maximin_candidate(self, available: List[Dict[str, float]]) -> Dict[str, float]:
        anchors = [resolve_initial_parameters(self.config)] + [obs["params"] for obs in self.observations]
        if not anchors:
            return self._rng.choice(available)
        encoded_anchors = [encode_candidate(a, self.config) for a in anchors]
        return max(
            available,
            key=lambda c: min(_distance(encode_candidate(c, self.config), a) for a in encoded_anchors),
        )

    def _distance_surrogate_candidate(self, available: List[Dict[str, float]]) -> Dict[str, float]:
        observed = [(encode_candidate(obs["params"], self.config), float(obs["Q_run"])) for obs in self.observations]
        best_q = max(q for _, q in observed)

        def score(candidate: dict) -> float:
            encoded = encode_candidate(candidate, self.config)
            distances = [_distance(encoded, point) for point, _ in observed]
            nearest = min(distances) if distances else 1.0
            weighted_sum = 0.0
            weight_total = 0.0
            for point, q in observed:
                weight = 1.0 / (_distance(encoded, point) + 0.05)
                weighted_sum += weight * q
                weight_total += weight
            predicted = weighted_sum / max(weight_total, 1e-12)
            return predicted + 0.15 * nearest + 0.05 * max(0.0, best_q - predicted)

        return max(available, key=score)

    def _gp_expected_improvement_candidate(self, available: List[Dict[str, float]]) -> Optional[Dict[str, float]]:
        try:
            import numpy as np
            from sklearn.gaussian_process import GaussianProcessRegressor
            from sklearn.gaussian_process.kernels import ConstantKernel, Matern, WhiteKernel
        except Exception:
            return None

    def _fit_gp_surrogate(self):
        try:
            import numpy as np
            from sklearn.gaussian_process import GaussianProcessRegressor
            from sklearn.gaussian_process.kernels import ConstantKernel, Matern, WhiteKernel
        except Exception:
            return None, None
        if len(self.observations) < 2:
            return None, None
        x_train = np.asarray([encode_candidate(obs["params"], self.config) for obs in self.observations], dtype=float)
        y_train = np.asarray([float(obs["Q_run"]) for obs in self.observations], dtype=float)
        try:
            kernel = (
                ConstantKernel(1.0, constant_value_bounds="fixed")
                * Matern(length_scale=np.ones(x_train.shape[1]), nu=2.5)
                + WhiteKernel(noise_level=1e-4, noise_level_bounds=(1e-8, 1e-1))
            )
            gp = GaussianProcessRegressor(
                kernel=kernel,
                normalize_y=True,
                random_state=int(self.config.get("random_seed", 42)),
                n_restarts_optimizer=2,
            )
            gp.fit(x_train, y_train)
            return gp, {"x_train": x_train, "y_train": y_train}
        except Exception:
            return None, None
        if len(self.observations) < 2:
            return None
        x_train = np.asarray([encode_candidate(obs["params"], self.config) for obs in self.observations], dtype=float)
        y_train = np.asarray([float(obs["Q_run"]) for obs in self.observations], dtype=float)
        x_available = np.asarray([encode_candidate(c, self.config) for c in available], dtype=float)
        try:
            kernel = (
                ConstantKernel(1.0, constant_value_bounds="fixed")
                * Matern(length_scale=np.ones(x_train.shape[1]), nu=2.5)
                + WhiteKernel(noise_level=1e-4, noise_level_bounds=(1e-8, 1e-1))
            )
            gp = GaussianProcessRegressor(
                kernel=kernel,
                normalize_y=True,
                random_state=int(self.config.get("random_seed", 42)),
                n_restarts_optimizer=2,
            )
            gp.fit(x_train, y_train)
            mean, std = gp.predict(x_available, return_std=True)
            ei = [_expected_improvement(float(m), float(s), float(y_train.max())) for m, s in zip(mean, std)]
            return available[int(max(range(len(ei)), key=lambda i: ei[i]))]
        except Exception:
            return None

    def _params_for_method_ref(self, params: dict) -> dict:
        result = {name: _format_float(params[name]) for name in PARAMETER_ORDER}
        result["bandwidth"] = str(self.config.get("method_options", {}).get("bandwidth", "4k"))
        result["bo_session_id"] = self.session_id
        return result

    def _retain_analysis_file(self, source: Path, iteration: int) -> Path:
        target = self.analysis_dir / f"iter_{iteration:03d}_{source.name}"
        if source.resolve() == target.resolve():
            return target
        if self.config.get("analysis", {}).get("copy_outputs_into_record", True):
            shutil.copy2(source, target)
            return target
        return source

    def _write_session_start_files(self) -> None:
        self._write_json(self.record_dir / "bo_config_snapshot.json", self.config)
        self._write_json(self.record_dir / "search_space.json", self.config.get("parameters", {}))
        self._write_json(self.record_dir / "constraints.json", self.config.get("constraints", {}))
        self._write_json(
            self.record_dir / "initial_parameters_preview.json",
            {
                "initial_parameters": resolve_initial_parameters(self.config),
                "candidate_count": len(self.candidates),
                "initial_candidates": self.candidates[: int(self.config.get("n_initial_points", 8))],
            },
        )

    def _write_history_csv(self) -> None:
        path = self.record_dir / "history.csv"
        rows = []
        for obs in self.observations:
            row = {
                "iteration": obs["iteration"],
                "method_id": obs["method_id"],
                "Q_run": obs["Q_run"],
                "completed_at": obs["completed_at"],
                "analysis_record": obs["analysis_record"],
            }
            row.update({name: obs["params"].get(name) for name in PARAMETER_ORDER})
            for ch, q in obs["quality"].get("Q_channels", {}).items():
                row[f"Q_ch{ch}"] = q
            rows.append(row)
        fieldnames = []
        for row in rows:
            for key in row:
                if key not in fieldnames:
                    fieldnames.append(key)
        with open(path, "w", encoding="utf-8", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=fieldnames or ["iteration"])
            writer.writeheader()
            writer.writerows(rows)

    def _write_surrogate_and_acquisition_artifacts(self, iteration: int) -> None:
        rows, metadata, gp = self._candidate_prediction_rows()
        self._write_json(self.surrogate_dir / f"iter_{iteration:03d}_surrogate_metadata.json", metadata)
        self._write_csv(self.surrogate_dir / f"iter_{iteration:03d}_candidate_predictions.csv", rows)
        self._write_csv(self.acquisition_dir / f"iter_{iteration:03d}_acquisition_values.csv", rows)

        top = [
            row for row in sorted(rows, key=lambda r: float(r.get("acquisition_value", 0.0)), reverse=True)
            if not row.get("already_tested")
        ][:20]
        self._write_csv(self.acquisition_dir / f"iter_{iteration:03d}_top_candidates.csv", top)

        if gp is not None:
            try:
                with open(self.surrogate_dir / f"iter_{iteration:03d}_gp_model.pkl", "wb") as fh:
                    pickle.dump(gp, fh)
            except Exception as exc:
                metadata["gp_pickle_error"] = str(exc)
                self._write_json(self.surrogate_dir / f"iter_{iteration:03d}_surrogate_metadata.json", metadata)

        self._write_surrogate_projection_plot(iteration, rows, value_key="predicted_mean_Q")
        self._write_surrogate_projection_plot(iteration, rows, value_key="acquisition_value")

    def _candidate_prediction_rows(self) -> Tuple[List[dict], dict, Any]:
        gp, train = self._fit_gp_surrogate()
        observed_keys = {candidate_key(obs["params"]) for obs in self.observations}
        best_q = max((float(obs["Q_run"]) for obs in self.observations), default=0.0)
        rows = []
        metadata = {
            "backend": "gaussian_process" if gp is not None else "distance_weighted_fallback",
            "observation_count": len(self.observations),
            "candidate_count": len(self.candidates),
            "active_parameters": active_parameters(self.config),
            "best_Q_run": best_q,
            "created_at": datetime.now().isoformat(timespec="seconds"),
        }

        if gp is not None and train is not None:
            try:
                import numpy as np
                x_all = np.asarray([encode_candidate(c, self.config) for c in self.candidates], dtype=float)
                means, stds = gp.predict(x_all, return_std=True)
            except Exception:
                means = [0.0] * len(self.candidates)
                stds = [0.0] * len(self.candidates)
                metadata["prediction_error"] = "GP prediction failed; wrote zero predictions"
        else:
            means, stds = self._fallback_predictions()

        for idx, candidate in enumerate(self.candidates):
            encoded = encode_candidate(candidate, self.config)
            mean = float(means[idx])
            std = float(stds[idx])
            if gp is not None:
                acquisition = _expected_improvement(mean, std, best_q)
            else:
                acquisition = mean + 0.15 * std
            key = candidate_key(candidate)
            row = {
                "candidate_index": idx,
                "already_tested": key in observed_keys,
                "predicted_mean_Q": mean,
                "predicted_std_Q": std,
                "acquisition_value": acquisition,
                "best_observed_Q": best_q,
            }
            for name, value in zip(OPTIMIZER_ORDER, encoded):
                row[f"encoded_{name}"] = value
            for name in PARAMETER_ORDER:
                row[name] = candidate.get(name)
            rows.append(row)
        return rows, metadata, gp

    def _fallback_predictions(self) -> Tuple[List[float], List[float]]:
        if not self.observations:
            return [0.0] * len(self.candidates), [1.0] * len(self.candidates)
        observed = [(encode_candidate(obs["params"], self.config), float(obs["Q_run"])) for obs in self.observations]
        means = []
        stds = []
        for candidate in self.candidates:
            encoded = encode_candidate(candidate, self.config)
            distances = [_distance(encoded, point) for point, _ in observed]
            nearest = min(distances) if distances else 1.0
            weighted_sum = 0.0
            weight_total = 0.0
            for point, q in observed:
                weight = 1.0 / (_distance(encoded, point) + 0.05)
                weighted_sum += weight * q
                weight_total += weight
            means.append(weighted_sum / max(weight_total, 1e-12))
            stds.append(nearest)
        return means, stds

    def _write_surrogate_projection_plot(self, iteration: int, rows: List[dict], value_key: str) -> None:
        try:
            import matplotlib
            matplotlib.use("Agg")
            import matplotlib.pyplot as plt
        except Exception:
            return
        if not rows:
            return
        params = active_parameters(self.config)
        if len(params) < 2:
            params = [name for name in OPTIMIZER_ORDER if name in PARAMETER_ORDER][:2]
        x_name = params[0]
        y_name = params[1] if len(params) > 1 else params[0]
        x = [float(row[x_name]) for row in rows]
        y = [float(row[y_name]) for row in rows]
        values = [float(row.get(value_key, 0.0)) for row in rows]
        tested_x = [float(row[x_name]) for row in rows if row.get("already_tested")]
        tested_y = [float(row[y_name]) for row in rows if row.get("already_tested")]

        fig, ax = plt.subplots(figsize=(6.8, 4.6))
        scatter = ax.scatter(x, y, c=values, cmap="viridis", s=36, alpha=0.85)
        if tested_x:
            ax.scatter(tested_x, tested_y, facecolors="none", edgecolors="#d67b32", s=90, linewidths=1.5, label="Tested")
            ax.legend(loc="best")
        ax.set_xlabel(x_name)
        ax.set_ylabel(y_name)
        title = "Predicted Q surface" if value_key == "predicted_mean_Q" else "Acquisition surface"
        ax.set_title(f"Iteration {iteration:03d} {title}")
        fig.colorbar(scatter, ax=ax, label=value_key)
        ax.grid(alpha=0.2)
        fig.tight_layout()
        suffix = "surrogate_projection" if value_key == "predicted_mean_Q" else "acquisition_projection"
        fig.savefig(self.plots_dir / f"iter_{iteration:03d}_{suffix}.png", dpi=160)
        plt.close(fig)

    @staticmethod
    def _write_csv(path: Path, rows: List[dict]) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        fieldnames = []
        for row in rows:
            for key in row:
                if key not in fieldnames:
                    fieldnames.append(key)
        with open(path, "w", encoding="utf-8", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=fieldnames or ["empty"])
            writer.writeheader()
            writer.writerows(rows)

    def _write_plots(self, observation: dict) -> None:
        try:
            import matplotlib
            matplotlib.use("Agg")
            import matplotlib.pyplot as plt
        except Exception:
            return

        q_channels = observation.get("quality", {}).get("Q_channels", {})
        if q_channels:
            channels = sorted(q_channels, key=lambda ch: int(ch))
            values = [float(q_channels[ch]) for ch in channels]
            fig, ax = plt.subplots(figsize=(7.0, 3.6))
            ax.bar(channels, values, color="#155e63")
            ax.set_ylim(0.0, 1.0)
            ax.set_xlabel("Mux channel")
            ax.set_ylabel("Q_channel")
            ax.set_title(f"Iteration {observation['iteration']:03d} channel scores")
            ax.grid(axis="y", alpha=0.25)
            fig.tight_layout()
            fig.savefig(self.plots_dir / f"iter_{observation['iteration']:03d}_channel_scores.png", dpi=160)
            plt.close(fig)

        if self.observations:
            iterations = [int(obs["iteration"]) for obs in self.observations]
            scores = [float(obs["Q_run"]) for obs in self.observations]
            best_so_far = []
            running_best = 0.0
            for score in scores:
                running_best = max(running_best, score)
                best_so_far.append(running_best)
            fig, ax = plt.subplots(figsize=(7.0, 3.6))
            ax.plot(iterations, scores, marker="o", color="#155e63", label="Q_run")
            ax.plot(iterations, best_so_far, color="#d67b32", label="Best so far")
            ax.set_ylim(0.0, 1.0)
            ax.set_xlabel("BO iteration")
            ax.set_ylabel("Score")
            ax.set_title("Bayesian optimization history")
            ax.grid(alpha=0.25)
            ax.legend(loc="best")
            fig.tight_layout()
            fig.savefig(self.plots_dir / "bo_history.png", dpi=160)
            plt.close(fig)

    def _build_session_id(self) -> str:
        stem = str(self.config.get("name", "bo")).strip().lower().replace(" ", "_")
        safe = "".join(ch if ch.isalnum() or ch == "_" else "_" for ch in stem).strip("_") or "bo"
        digest = hashlib.sha1(json.dumps(self.config, sort_keys=True).encode("utf-8")).hexdigest()[:6]
        return f"{safe}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{digest}"

    @staticmethod
    def _unique_dir(path: Path) -> Path:
        if not path.exists():
            return path
        for idx in range(2, 1000):
            candidate = Path(f"{path}_{idx:02d}")
            if not candidate.exists():
                return candidate
        return Path(f"{path}_{datetime.now().strftime('%H%M%S')}")

    @staticmethod
    def _write_json(path: Path, payload: Any) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, indent=2)


def build_swv_script(params: dict, method_options: Optional[dict] = None) -> str:
    options = method_options or {}
    begin_v = float(params["begin_potential"])
    end_v = float(params["end_potential"])
    amp_v = float(params["amplitude"])
    cond_time_s = float(params["conditioning_time"])
    freq_hz = float(params["frequency"])

    begin = to_si_string(str(params["begin_potential"]), "V")
    end = to_si_string(str(params["end_potential"]), "V")
    step = to_si_string(str(params["step_potential"]), "V")
    amplitude = to_si_string(str(params["amplitude"]), "V")
    frequency = to_si_string(str(params["frequency"]), "Hz")
    cond_pot = to_si_string(str(params["conditioning_potential"]), "V")
    cond_time = str(params["conditioning_time"])
    bandwidth = str(options.get("bandwidth", "4k")).strip().lower()
    if bandwidth not in ("4k", "8k"):
        bandwidth = "4k"

    min_mv = int((min(begin_v, end_v) - amp_v) * 1000)
    max_mv = int((max(begin_v, end_v) + amp_v) * 1000)
    use_equilibrium_check = cond_time_s > 0
    eq_interval_s = min(0.2, cond_time_s) if use_equilibrium_check else 0.0
    swv_time_step = to_si_string(str(1.0 / freq_hz), "s") if freq_hz > 0 else "0"
    eq_duration = to_si_string(cond_time, "s") if use_equilibrium_check else "0"
    eq_interval = to_si_string(str(eq_interval_s), "s") if use_equilibrium_check else "0"
    ba = dict(options.get("ba_range") or {})
    ba_mode = str(ba.get("mode", "fixed")).lower()
    fixed_range = str(ba.get("fixed", "100n"))
    auto_min = str(ba.get("auto_min", fixed_range))
    auto_max = str(ba.get("auto_max", fixed_range))

    parts = [
        "e", "var c", "var p", "var f", "var r",
        "set_pgstat_chan 1",
        "set_pgstat_mode 0",
        "set_pgstat_chan 0",
        "set_pgstat_mode 3",
        f"set_max_bandwidth {bandwidth}",
        f"set_range_minmax da {min_mv}m {max_mv}m",
    ]
    if use_equilibrium_check:
        parts.insert(5, "var t")
    if ba_mode == "auto":
        parts += [f"set_range ba {auto_max}", f"set_autoranging ba {auto_min} {auto_max}"]
    else:
        parts += [f"set_range ba {fixed_range}", f"set_autoranging ba {fixed_range} {fixed_range}"]
    parts += [f"set_e {cond_pot if use_equilibrium_check else begin}", "cell_on"]
    if use_equilibrium_check:
        parts += [
            f"# Equilibrium check at {cond_pot} for {cond_time}s",
            "store_var t 0 eb",
            f"meas_loop_ca p c {cond_pot} {eq_interval} {eq_duration}",
            "\tpck_start",
            "\t\tpck_add t",
            "\t\tpck_add p",
            "\t\tpck_add c",
            "\tpck_end",
            f"\tadd_var t {eq_interval}",
            "endloop",
            "store_var t 0 eb",
            f"set_e {begin}",
        ]
    else:
        parts += [f"set_e {begin}"]
    parts += [f"meas_loop_swv p c f r {begin} {end} {step} {amplitude} {frequency}"]
    if use_equilibrium_check:
        parts += [
            "\tpck_start",
            "\t\tpck_add p",
            "\t\tpck_add c",
            "\t\tpck_add f",
            "\t\tpck_add r",
            "\t\tpck_add t",
            "\tpck_end",
            f"\tadd_var t {swv_time_step}",
        ]
    else:
        parts += [
            "\tpck_start",
            "\t\tpck_add p",
            "\t\tpck_add c",
            "\t\tpck_add f",
            "\t\tpck_add r",
            "\tpck_end",
        ]
    parts += ["endloop", "on_finished:", "cell_off"]
    return "\n".join(parts)


def mux_channel_address(channel: int) -> int:
    idx = int(channel) - 1
    return (idx << 4) | idx


def wrap_mux(base_script: str, channel: int) -> str:
    lines = base_script.splitlines()
    header = lines[0].strip() if lines and lines[0].strip() in ("e", "l") else "e"
    rest = lines[1:] if lines and lines[0].strip() in ("e", "l") else lines
    addr = mux_channel_address(channel)
    prefix = [
        header,
        "# MUX16 channel select",
        "set_gpio_cfg 0x3FFi 1",
        f"set_gpio {addr}i",
        "# End MUX16 channel select",
    ]
    return "\n".join(prefix + rest)


def encode_candidate(candidate: dict, config: dict) -> List[float]:
    cfg = normalize_bo_config(config)
    encoded = []
    for name in OPTIMIZER_ORDER:
        p_cfg = cfg["parameters"][name]
        values = _float_values(p_cfg.get("values", [candidate[name]]), name)
        value = float(candidate[name])
        if str(p_cfg.get("encoding", "")).lower() == "log10":
            values = [math.log10(max(v, 1e-12)) for v in values]
            value = math.log10(max(value, 1e-12))
        lo = min(values)
        hi = max(values)
        encoded.append((value - lo) / (hi - lo + 1e-12))
    return encoded


def candidate_key(candidate: dict) -> Tuple[float, ...]:
    return tuple(round(float(candidate[name]), 9) for name in PARAMETER_ORDER)


def _float_values(values: Iterable[Any], label: str) -> List[float]:
    result = [_float_value(v, label) for v in values]
    if not result:
        raise ValueError(f"{label} must contain at least one numeric value")
    return result


def _float_value(value: Any, label: str) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        raise ValueError(f"{label} must be numeric")


def _format_float(value: Any) -> str:
    return f"{float(value):.12g}"


def _channel_items(channel_metrics: dict) -> Iterable[Tuple[str, dict]]:
    for channel, metrics in channel_metrics.items():
        if isinstance(metrics, dict):
            yield str(channel), metrics


def _clip01(value: Any) -> float:
    try:
        numeric = float(value)
    except (TypeError, ValueError):
        numeric = 0.0
    return max(0.0, min(1.0, numeric))


def _std(values: List[float]) -> float:
    if not values:
        return 0.0
    mean = sum(values) / len(values)
    return math.sqrt(sum((v - mean) ** 2 for v in values) / len(values))


def _distance(a: List[float], b: List[float]) -> float:
    return math.sqrt(sum((x - y) ** 2 for x, y in zip(a, b)))


def _expected_improvement(mean: float, std: float, best_score: float, xi: float = 0.01) -> float:
    std = max(std, 1e-12)
    improvement = mean - best_score - xi
    z = improvement / std
    return improvement * _normal_cdf(z) + std * _normal_pdf(z)


def _normal_pdf(x: float) -> float:
    return math.exp(-0.5 * x * x) / math.sqrt(2.0 * math.pi)


def _normal_cdf(x: float) -> float:
    return 0.5 * (1.0 + math.erf(x / math.sqrt(2.0)))
