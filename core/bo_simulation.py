"""Synthetic SWV simulation engine for BO tuning.

The engine builds a continuous parameter landscape, generates fake SWV traces,
scores synthetic channel metrics, and lets the existing BO session machinery
walk that landscape.
"""

from __future__ import annotations

import csv
from dataclasses import dataclass
from datetime import datetime
import json
import math
from pathlib import Path
import random
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from core.bo_analysis import run_request
from core.bo_session import (
    BOIntegrationSession,
    PARAMETER_ORDER,
    normalize_bo_config,
    parse_channels,
    resolve_initial_parameters,
)


LANDSCAPE_TYPES = ("gaussian", "ridge", "smooth_random")


@dataclass
class SimulationDimension:
    name: str
    minimum: float
    maximum: float
    optimum: float
    spread: float
    landscape: str = "gaussian"
    weight: float = 1.0
    delta_follow_q: bool = True
    delta_optimum: Optional[float] = None
    delta_spread: Optional[float] = None
    delta_landscape: str = "gaussian"
    delta_weight: float = 1.0

    def normalized(self, value: float) -> float:
        span = max(self.maximum - self.minimum, 1e-12)
        return (float(value) - self.minimum) / span

    def normalized_optimum(self) -> float:
        return self.normalized(self.optimum)


def default_dimensions(config: dict, limit: int = 3) -> List[dict]:
    cfg = normalize_bo_config(config)
    initial = resolve_initial_parameters(cfg)
    active = [
        name for name in PARAMETER_ORDER
        if str(cfg["parameters"].get(name, {}).get("mode", "")).lower() == "active"
    ]
    if not active:
        active = ["step_potential", "amplitude", "frequency"]
    rows = []
    for name in active[: max(1, int(limit))]:
        p_cfg = cfg["parameters"].get(name, {})
        minimum = float(p_cfg.get("min", initial.get(name, 0.0)))
        maximum = float(p_cfg.get("max", initial.get(name, minimum)))
        if maximum <= minimum:
            values = [float(v) for v in p_cfg.get("values", [initial.get(name, minimum)])]
            minimum, maximum = min(values), max(values)
        span = max(maximum - minimum, 1e-12)
        optimum = float(initial.get(name, (minimum + maximum) / 2.0))
        optimum = min(max(optimum, minimum), maximum)
        rows.append(
            {
                "name": name,
                "minimum": minimum,
                "maximum": maximum,
                "optimum": optimum,
                "spread": span * 0.22,
                "landscape": "gaussian",
                "weight": 1.0,
                "delta_follow_q": True,
                "delta_optimum": optimum,
                "delta_spread": span * 0.22,
                "delta_landscape": "gaussian",
                "delta_weight": 1.0,
            }
        )
    return rows


def dimensions_from_config(sim_config: dict) -> List[SimulationDimension]:
    dims = []
    for raw in sim_config.get("dimensions", []):
        if not isinstance(raw, dict):
            continue
        minimum = float(raw.get("minimum", raw.get("min", 0.0)))
        maximum = float(raw.get("maximum", raw.get("max", minimum + 1.0)))
        if maximum <= minimum:
            minimum, maximum = maximum, minimum
        span = max(maximum - minimum, 1e-12)
        optimum = min(max(float(raw.get("optimum", (minimum + maximum) / 2.0)), minimum), maximum)
        spread = max(float(raw.get("spread", span * 0.22)), span * 1e-6)
        landscape = str(raw.get("landscape", "gaussian")).lower()
        if landscape not in LANDSCAPE_TYPES:
            landscape = "gaussian"
        delta_follow_q = _as_bool(raw.get("delta_follow_q", True))
        delta_optimum = min(max(float(raw.get("delta_optimum", optimum)), minimum), maximum)
        delta_spread = max(float(raw.get("delta_spread", spread)), span * 1e-6)
        delta_landscape = str(raw.get("delta_landscape", landscape)).lower()
        if delta_landscape not in LANDSCAPE_TYPES:
            delta_landscape = "gaussian"
        dims.append(
            SimulationDimension(
                name=str(raw.get("name", "")).strip(),
                minimum=minimum,
                maximum=maximum,
                optimum=optimum,
                spread=spread,
                landscape=landscape,
                weight=max(0.0, float(raw.get("weight", 1.0))),
                delta_follow_q=delta_follow_q,
                delta_optimum=delta_optimum,
                delta_spread=delta_spread,
                delta_landscape=delta_landscape,
                delta_weight=max(0.0, float(raw.get("delta_weight", 1.0))),
            )
        )
    return [dim for dim in dims if dim.name]


class SyntheticSWVSimulationEngine:
    def __init__(self, bo_config: dict, sim_config: dict):
        self.bo_config = normalize_bo_config(bo_config)
        self.sim_config = dict(sim_config or {})
        dims = dimensions_from_config(self.sim_config)
        self.dimensions = dims or dimensions_from_config({"dimensions": default_dimensions(self.bo_config)})
        self.seed = int(self.sim_config.get("seed", self.bo_config.get("random_seed", 42)))
        self.measurement_noise = max(0.0, float(self.sim_config.get("measurement_noise", 0.03)))
        self.channel_noise = max(0.0, float(self.sim_config.get("channel_noise", 0.025)))
        self.peak_emphasis = max(0.0, float(self.sim_config.get("peak_emphasis", 0.70)))
        self.base_peak_uA = max(0.0, float(self.sim_config.get("base_peak_uA", 0.45)))
        self.peak_gain_uA = max(0.0, float(self.sim_config.get("peak_gain_uA", 5.0)))
        self.base_noise_uA = max(1e-6, float(self.sim_config.get("base_noise_uA", 0.08)))
        self.noise_gain_uA = max(0.0, float(self.sim_config.get("noise_gain_uA", 0.45)))
        self.target_response_gain_uA = max(0.0, float(self.sim_config.get("target_response_gain_uA", 2.0)))
        self.target_noise_multiplier = max(0.05, float(self.sim_config.get("target_noise_multiplier", 1.05)))
        self.target_response_floor = _clip01(self.sim_config.get("target_response_floor", 0.0))
        self.trace_points = max(41, int(self.sim_config.get("trace_points", 121)))

    def analysis_payload(self, params: dict, iteration: int) -> dict:
        truth = self.evaluate_truth(params)
        metrics = {}
        traces = {}
        for channel in parse_channels(self.bo_config.get("channels", [])):
            channel_metrics, trace = self._channel_measurement(params, truth, iteration, channel)
            metrics[str(channel)] = channel_metrics
            traces[str(channel)] = trace
        return {
            "schema_version": 1,
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "simulation_engine": {
                "version": 1,
                "iteration": int(iteration),
                "parameters": dict(params),
                "dimensions": [dim.__dict__ for dim in self.dimensions],
            },
            "simulation_truth": truth,
            "channel_metrics": metrics,
            "swv_traces": traces,
        }

    def paired_analysis_payload(self, params: dict, iteration: int, phase: str) -> dict:
        phase_label = str(phase or "buffer").strip().lower()
        if phase_label not in ("buffer", "target"):
            raise ValueError(f"Unsupported paired simulation phase: {phase}")
        truth = self.evaluate_truth(params)
        response_score = self.response_score(params)
        metrics = {}
        traces = {}
        for channel in parse_channels(self.bo_config.get("channels", [])):
            channel_metrics, trace = self._channel_measurement(params, truth, iteration, channel)
            if phase_label == "target":
                channel_metrics, trace = self._target_response_measurement(
                    channel_metrics,
                    trace,
                    response_score,
                    iteration,
                    channel,
                )
            metrics[str(channel)] = channel_metrics
            traces[str(channel)] = trace
        paired_truth = dict(truth)
        paired_truth.update(
            {
                "paired_phase": phase_label,
                "response_score": response_score,
                "expected_delta_peak_uA": self.target_response_gain_uA * response_score,
            }
        )
        return {
            "schema_version": 1,
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "simulation_engine": {
                "version": 3,
                "analysis_mode": "paired_response_direct_summary",
                "iteration": int(iteration),
                "phase": phase_label,
                "parameters": dict(params),
                "dimensions": [dim.__dict__ for dim in self.dimensions],
                "target_response_gain_uA": self.target_response_gain_uA,
                "target_response_floor": self.target_response_floor,
                "target_noise_multiplier": self.target_noise_multiplier,
            },
            "simulation_truth": paired_truth,
            "channel_metrics": metrics,
            "swv_traces": traces,
        }

    def evaluate_truth(self, params: dict) -> dict:
        component_scores = []
        weights = []
        for dim in self.dimensions:
            score = self._dimension_score(dim, float(params.get(dim.name, dim.optimum)))
            component_scores.append(score)
            weights.append(max(dim.weight, 0.0))
        if not component_scores:
            landscape_q = 0.0
        else:
            total_weight = sum(weights) or float(len(component_scores))
            log_sum = 0.0
            for score, weight in zip(component_scores, weights):
                w = weight if sum(weights) > 0.0 else 1.0
                log_sum += w * math.log(max(score, 1e-9))
            landscape_q = _clip01(math.exp(log_sum / max(total_weight, 1e-12)))
        distance = self.normalized_distance(params)
        peak_score = _clip01(0.10 + 0.90 * landscape_q)
        noise_score = _clip01(1.0 - 0.80 * distance)
        shape_score = _clip01(0.25 + 0.75 * landscape_q)
        traditional_success_score = _clip01((peak_score ** 1.8) * (0.60 + 0.40 * noise_score))
        true_q = _clip01(
            self.peak_emphasis * peak_score
            + 0.15 * noise_score
            + 0.10 * shape_score
            + max(0.0, 0.05 - max(0.0, self.peak_emphasis - 0.70) * 0.05)
        )
        delta_peak_score = self.response_score(params)
        expected_delta_peak_uA = self.target_response_gain_uA * delta_peak_score
        q_trad_buffer = true_q
        q_trad_target = true_q
        paired_q = q_trad_buffer + q_trad_target + delta_peak_score
        paired_q_score = _clip01(paired_q / 3.0)
        success_score = paired_q_score if bool(self.sim_config.get("paired_response", False)) else traditional_success_score
        return {
            "true_Q": true_q,
            "success_score": success_score,
            "traditional_success_score": traditional_success_score,
            "landscape_Q": landscape_q,
            "normalized_distance": distance,
            "peak_score": peak_score,
            "noise_score": noise_score,
            "shape_score": shape_score,
            "component_scores": {
                dim.name: score for dim, score in zip(self.dimensions, component_scores)
            },
            "optimum": {dim.name: dim.optimum for dim in self.dimensions},
            "delta_peak_score": delta_peak_score,
            "expected_delta_peak_uA": expected_delta_peak_uA,
            "Q_trad_buffer": q_trad_buffer,
            "Q_trad_target": q_trad_target,
            "paired_Q": paired_q,
            "paired_Q_score": paired_q_score,
        }

    def response_score(self, params: dict) -> float:
        """Target response landscape used for the normalized delta-peak term."""
        component_scores = []
        weights = []
        for dim in self.dimensions:
            value = float(params.get(dim.name, dim.delta_optimum if dim.delta_optimum is not None else dim.optimum))
            if dim.delta_follow_q:
                score = self._dimension_score(dim, value)
            else:
                score = self._delta_dimension_score(dim, value)
            component_scores.append(score)
            weights.append(max(dim.delta_weight, 0.0))
        if not component_scores:
            response = 0.0
        else:
            total_weight = sum(weights) or float(len(component_scores))
            log_sum = 0.0
            for score, weight in zip(component_scores, weights):
                w = weight if sum(weights) > 0.0 else 1.0
                log_sum += w * math.log(max(score, 1e-9))
            response = _clip01(math.exp(log_sum / max(total_weight, 1e-12)))
        return _clip01(self.target_response_floor + (1.0 - self.target_response_floor) * response)

    def evaluate_truth_without_response(self, params: dict) -> dict:
        component_scores = []
        weights = []
        for dim in self.dimensions:
            score = self._dimension_score(dim, float(params.get(dim.name, dim.optimum)))
            component_scores.append(score)
            weights.append(max(dim.weight, 0.0))
        if not component_scores:
            landscape_q = 0.0
        else:
            total_weight = sum(weights) or float(len(component_scores))
            log_sum = 0.0
            for score, weight in zip(component_scores, weights):
                w = weight if sum(weights) > 0.0 else 1.0
                log_sum += w * math.log(max(score, 1e-9))
            landscape_q = _clip01(math.exp(log_sum / max(total_weight, 1e-12)))
        return {"landscape_Q": landscape_q}

    def normalized_distance(self, params: dict) -> float:
        if not self.dimensions:
            return 0.0
        terms = []
        for dim in self.dimensions:
            span = max(dim.maximum - dim.minimum, 1e-12)
            terms.append(((float(params.get(dim.name, dim.optimum)) - dim.optimum) / span) ** 2)
        return min(1.0, math.sqrt(sum(terms) / len(terms)))

    def sample_landscape(self, grid_size: int = 25) -> dict:
        dims = self.dimensions[:3]
        n = max(5, min(45, int(grid_size)))
        axes = [_linspace(dim.minimum, dim.maximum, n) for dim in dims]
        points = []
        if not dims:
            return {"dimensions": [], "points": []}
        for values in _product(axes):
            params = resolve_initial_parameters(self.bo_config)
            for dim, value in zip(dims, values):
                params[dim.name] = value
            truth = self.evaluate_truth(params)
            point = {dim.name: value for dim, value in zip(dims, values)}
            point.update(
                {
                    "true_Q": truth["true_Q"],
                    "success_score": truth["success_score"],
                    "traditional_success_score": truth["traditional_success_score"],
                    "paired_Q": truth["paired_Q"],
                    "paired_Q_score": truth["paired_Q_score"],
                    "landscape_Q": truth["landscape_Q"],
                    "distance": truth["normalized_distance"],
                    "peak_score": truth["peak_score"],
                    "noise_score": truth["noise_score"],
                    "delta_peak": truth["expected_delta_peak_uA"],
                    "delta_peak_score": truth["delta_peak_score"],
                }
            )
            points.append(point)
        return {
            "dimensions": [dim.__dict__ for dim in dims],
            "points": points,
            "paired_response": bool(self.sim_config.get("paired_response", False)),
        }

    def dimension_distributions(self, grid_size: int = 61) -> dict:
        base = resolve_initial_parameters(self.bo_config)
        rows = []
        for dim in self.dimensions[:3]:
            values = _linspace(dim.minimum, dim.maximum, max(11, min(201, int(grid_size))))
            curve = []
            for value in values:
                params = dict(base)
                params[dim.name] = value
                truth = self.evaluate_truth(params)
                curve.append(
                    {
                        "value": value,
                        "true_Q": truth["true_Q"],
                        "success_score": truth["success_score"],
                        "traditional_success_score": truth["traditional_success_score"],
                        "paired_Q": truth["paired_Q"],
                        "paired_Q_score": truth["paired_Q_score"],
                        "peak_score": truth["peak_score"],
                        "noise_score": truth["noise_score"],
                        "delta_peak": truth["expected_delta_peak_uA"],
                        "delta_peak_score": truth["delta_peak_score"],
                    }
                )
            rows.append({"name": dim.name, "curve": curve})
        return {"dimensions": rows}

    def _dimension_score(self, dim: SimulationDimension, value: float) -> float:
        z = (value - dim.optimum) / max(dim.spread, 1e-12)
        if dim.landscape == "ridge":
            return _clip01(math.exp(-0.5 * abs(z)))
        if dim.landscape == "smooth_random":
            x = dim.normalized(value)
            o = dim.normalized_optimum()
            base = math.exp(-0.5 * ((x - o) / max(dim.spread / max(dim.maximum - dim.minimum, 1e-12), 1e-6)) ** 2)
            phase = (self.seed % 997) / 997.0
            smooth = 0.5 + 0.5 * math.sin(2.0 * math.pi * (1.7 * x + phase))
            smooth *= 0.65 + 0.35 * math.cos(2.0 * math.pi * (0.6 * x + phase)) ** 2
            return _clip01(0.68 * base + 0.32 * smooth)
        return _clip01(math.exp(-0.5 * z * z))

    def _delta_dimension_score(self, dim: SimulationDimension, value: float) -> float:
        optimum = dim.delta_optimum if dim.delta_optimum is not None else dim.optimum
        spread = dim.delta_spread if dim.delta_spread is not None else dim.spread
        z = (value - optimum) / max(spread, 1e-12)
        shape = dim.delta_landscape if dim.delta_landscape in LANDSCAPE_TYPES else "gaussian"
        if shape == "ridge":
            return _clip01(math.exp(-0.5 * abs(z)))
        if shape == "smooth_random":
            x = dim.normalized(value)
            o = dim.normalized(optimum)
            base = math.exp(-0.5 * ((x - o) / max(spread / max(dim.maximum - dim.minimum, 1e-12), 1e-6)) ** 2)
            phase = ((self.seed + 311) % 997) / 997.0
            smooth = 0.5 + 0.5 * math.sin(2.0 * math.pi * (1.7 * x + phase))
            smooth *= 0.65 + 0.35 * math.cos(2.0 * math.pi * (0.6 * x + phase)) ** 2
            return _clip01(0.68 * base + 0.32 * smooth)
        return _clip01(math.exp(-0.5 * z * z))

    def _channel_measurement(self, params: dict, truth: dict, iteration: int, channel: int, scan: int = 1) -> Tuple[dict, dict]:
        rng = random.Random(self.seed + int(iteration) * 1009 + int(channel) * 131 + int(scan) * 17)
        true_q = float(truth["true_Q"])
        channel_q = _clip01(true_q + rng.gauss(0.0, self.channel_noise))
        peak_current = max(0.0, self.base_peak_uA + self.peak_gain_uA * channel_q + rng.gauss(0.0, self.measurement_noise))
        background_rms = max(
            1e-6,
            self.base_noise_uA + self.noise_gain_uA * (1.0 - channel_q) + rng.gauss(0.0, self.measurement_noise * 0.25),
        )
        snr = peak_current / max(background_rms, 1e-12)
        metrics = {
            "snr": snr,
            "peak_shape_score": _clip01(0.25 + 0.75 * channel_q + rng.gauss(0.0, self.channel_noise)),
            "baseline_stability_score": _clip01(1.0 - background_rms / max(self.base_noise_uA + self.noise_gain_uA, 1e-12)),
            "replicate_consistency_score": _clip01(0.35 + 0.65 * channel_q + rng.gauss(0.0, self.channel_noise)),
            "success_score": _clip01(
                (max(0.0, peak_current - self.base_peak_uA) / max(self.peak_gain_uA, 1e-12)) ** 1.6
                * (
                    0.65
                    + 0.35
                    * _clip01(1.0 - background_rms / max(self.base_noise_uA + self.noise_gain_uA, 1e-12))
                )
            ),
            "ok_scan_count": 3,
            "total_scan_count": 3,
            "mean_peak_current_uA": peak_current,
            "median_peak_current_uA": peak_current,
            "mean_background_rms_uA": background_rms,
            "median_background_rms_uA": background_rms,
        }
        return metrics, self._swv_trace(params, peak_current, background_rms, rng)

    def _swv_trace(self, params: dict, peak_current: float, background_rms: float, rng: random.Random) -> dict:
        begin = float(params.get("begin_potential", -0.6))
        end = float(params.get("end_potential", -0.1))
        if end <= begin:
            begin, end = -0.6, -0.1
        voltages = _linspace(begin, end, self.trace_points)
        center = begin + 0.58 * (end - begin)
        width = max((end - begin) * 0.075, 1e-4)
        slope = rng.uniform(-0.08, 0.08)
        baseline = rng.uniform(-0.05, 0.05)
        currents = []
        for voltage in voltages:
            peak = peak_current * math.exp(-0.5 * ((voltage - center) / width) ** 2)
            noise = rng.gauss(0.0, background_rms)
            currents.append(baseline + slope * (voltage - begin) + peak + noise)
        return {
            "voltage_v": [round(v, 6) for v in voltages],
            "current_uA": [round(i, 6) for i in currents],
        }

    def _target_response_measurement(
        self,
        metrics: dict,
        trace: dict,
        response_score: float,
        iteration: int,
        channel: int,
    ) -> Tuple[dict, dict]:
        rng = random.Random(self.seed + int(iteration) * 1543 + int(channel) * 211 + 50021)
        response_score = _clip01(response_score + rng.gauss(0.0, self.channel_noise * 0.6))
        delta_peak = max(0.0, self.target_response_gain_uA * response_score + rng.gauss(0.0, self.measurement_noise))
        target_metrics = dict(metrics)
        peak = max(0.0, float(target_metrics.get("mean_peak_current_uA", 0.0) or 0.0) + delta_peak)
        background = max(
            1e-6,
            float(target_metrics.get("mean_background_rms_uA", 0.0) or self.base_noise_uA) * self.target_noise_multiplier
            + rng.gauss(0.0, self.measurement_noise * 0.15),
        )
        target_metrics.update(
            {
                "snr": peak / max(background, 1e-12),
                "mean_peak_current_uA": peak,
                "median_peak_current_uA": peak,
                "mean_background_rms_uA": background,
                "median_background_rms_uA": background,
                "peak_shape_score": _clip01(float(target_metrics.get("peak_shape_score", 0.0) or 0.0) + 0.05 * response_score),
                "success_score": _clip01(float(target_metrics.get("success_score", 1.0) or 0.0) * (0.95 + 0.05 * response_score)),
                "simulated_delta_peak_uA": delta_peak,
            }
        )
        target_trace = {
            "voltage_v": list(trace.get("voltage_v") or []),
            "current_uA": list(trace.get("current_uA") or []),
        }
        currents = [float(value) for value in target_trace.get("current_uA") or []]
        voltages = [float(value) for value in target_trace.get("voltage_v") or []]
        if currents and voltages:
            center = voltages[0] + 0.58 * (voltages[-1] - voltages[0])
            width = max((voltages[-1] - voltages[0]) * 0.075, 1e-4)
            adjusted = []
            for voltage, current in zip(voltages, currents):
                response_peak = delta_peak * math.exp(-0.5 * ((voltage - center) / width) ** 2)
                adjusted.append(round(current + response_peak + rng.gauss(0.0, background * 0.05), 6))
            target_trace["current_uA"] = adjusted
        return target_metrics, target_trace


def run_optimizer_simulation(
    bo_config: dict,
    sim_config: dict,
    output_root: str | Path,
    iterations: int,
    analysis_output_dir: str | Path | None = None,
    progress_callback=None,
) -> dict:
    cfg = normalize_bo_config(bo_config)
    engine = SyntheticSWVSimulationEngine(cfg, sim_config)
    output_root = Path(output_root)
    output_root.mkdir(parents=True, exist_ok=True)
    analysis_output = Path(analysis_output_dir) if analysis_output_dir else output_root / "bo_analysis"
    analysis_output.mkdir(parents=True, exist_ok=True)
    session = BOIntegrationSession(
        cfg,
        output_root,
        config_path=None,
        analysis_output_dir=analysis_output,
    )
    rows = []
    total_iterations = max(1, int(iterations))
    for _idx in range(total_iterations):
        suggestion = session.ask_next()
        if progress_callback:
            progress_callback(_idx, total_iterations, f"Analyzing simulated iteration {suggestion.iteration}")
        raw_dir, simulation_payload = _write_simulated_raw_measurements(
            engine,
            output_root,
            suggestion.params,
            suggestion.iteration,
        )
        summary = run_request(
            {
                "folders": [str(raw_dir)],
                "output_dir": str(analysis_output),
                "output_stem": f"bo_iter_{suggestion.iteration:03d}",
                "analysis": dict(cfg.get("analysis") or {}),
            }
        )
        path = _augment_analysis_summary(Path(summary["summary_path"]), simulation_payload)
        obs = session.import_analysis(path, notes="Simulation engine")
        obs["simulation_truth"] = simulation_payload["simulation_truth"]
        obs["swv_trace_preview"] = simulation_payload.get("swv_traces", {})
        session.observations[-1] = obs
        session._write_json(session.analysis_dir / f"iter_{suggestion.iteration:03d}_quality.json", obs)
        session.save_state()
        rows.append(_row_from_observation(obs))
        if progress_callback:
            progress_callback(_idx + 1, total_iterations, f"Completed simulated iteration {suggestion.iteration}")
    return {
        "session": session,
        "engine": engine,
        "rows": rows,
        "landscape": engine.sample_landscape(int(sim_config.get("grid_size", 25))),
        "distributions": engine.dimension_distributions(),
    }


def run_paired_response_optimizer_simulation(
    bo_config: dict,
    sim_config: dict,
    output_root: str | Path,
    cycles: int,
    batch_size: int,
    analysis_output_dir: str | Path | None = None,
    progress_callback=None,
) -> dict:
    cfg = normalize_bo_config(bo_config)
    cfg["objective"] = "paired_response"
    engine = SyntheticSWVSimulationEngine(cfg, sim_config)
    output_root = Path(output_root)
    output_root.mkdir(parents=True, exist_ok=True)
    analysis_output = Path(analysis_output_dir) if analysis_output_dir else output_root / "bo_analysis"
    analysis_output.mkdir(parents=True, exist_ok=True)
    session = BOIntegrationSession(
        cfg,
        output_root,
        config_path=None,
        analysis_output_dir=analysis_output,
    )
    rows = []
    total_cycles = max(1, int(cycles))
    batch_size = max(1, int(batch_size))
    total_observations = total_cycles * batch_size
    total_traces = total_observations * 2
    completed_observations = 0
    completed_traces = 0
    trace_schedule = []
    for cycle_idx in range(total_cycles):
        suggestions = session.ask_batch(batch_size)
        cycle_number = cycle_idx + 1
        if progress_callback:
            progress_callback(
                completed_traces,
                total_traces,
                f"Simulating paired cycle {cycle_number}/{total_cycles}: buffer batch",
            )
        buffered = []
        for batch_idx, suggestion in enumerate(suggestions, start=1):
            buffer_trace_number = (cycle_idx * batch_size * 2) + batch_idx
            buffer_payload = engine.paired_analysis_payload(suggestion.params, suggestion.iteration, "buffer")
            _annotate_paired_payload(buffer_payload, cycle_number, batch_idx, buffer_trace_number, "buffer")
            buffer_path = _write_simulated_analysis_summary(
                analysis_output / f"bo_iter_{suggestion.iteration:03d}_buffer_simulated.json",
                buffer_payload,
            )
            buffered.append((batch_idx, suggestion, buffer_payload, buffer_path, buffer_trace_number))
            trace_schedule.append(
                {
                    "trace_number": buffer_trace_number,
                    "cycle": cycle_number,
                    "parameter_set": batch_idx,
                    "phase": "buffer",
                    "bo_iteration": suggestion.iteration,
                    "path": str(buffer_path),
                }
            )
            completed_traces += 1
            if progress_callback:
                progress_callback(
                    completed_traces,
                    total_traces,
                    f"Simulated trace {completed_traces}/{total_traces}: cycle {cycle_number}, set {batch_idx}, buffer",
                )
        targeted = []
        if progress_callback:
            progress_callback(
                completed_traces,
                total_traces,
                f"Simulating paired cycle {cycle_number}/{total_cycles}: target batch",
            )
        for batch_idx, suggestion, buffer_payload, buffer_path, buffer_trace_number in buffered:
            target_trace_number = (cycle_idx * batch_size * 2) + batch_size + batch_idx
            target_payload = engine.paired_analysis_payload(suggestion.params, suggestion.iteration, "target")
            _annotate_paired_payload(target_payload, cycle_number, batch_idx, target_trace_number, "target")
            target_path = _write_simulated_analysis_summary(
                analysis_output / f"bo_iter_{suggestion.iteration:03d}_target_simulated.json",
                target_payload,
            )
            targeted.append((batch_idx, suggestion, buffer_payload, buffer_path, target_payload, target_path, buffer_trace_number, target_trace_number))
            trace_schedule.append(
                {
                    "trace_number": target_trace_number,
                    "cycle": cycle_number,
                    "parameter_set": batch_idx,
                    "phase": "target",
                    "bo_iteration": suggestion.iteration,
                    "path": str(target_path),
                }
            )
            completed_traces += 1
            if progress_callback:
                progress_callback(
                    completed_traces,
                    total_traces,
                    f"Simulated trace {completed_traces}/{total_traces}: cycle {cycle_number}, set {batch_idx}, target",
                )
        for batch_idx, suggestion, buffer_payload, buffer_path, target_payload, target_path, buffer_trace_number, target_trace_number in targeted:
            obs = session.import_paired_analysis(
                suggestion,
                buffer_path,
                target_path,
                notes="Paired response simulation engine",
            )
            truth = dict(target_payload.get("simulation_truth") or {})
            obs["simulation_truth"] = truth
            obs["buffer_swv_trace_preview"] = buffer_payload.get("swv_traces", {})
            obs["target_swv_trace_preview"] = target_payload.get("swv_traces", {})
            obs["swv_trace_preview"] = target_payload.get("swv_traces", {})
            obs["paired_cycle"] = cycle_number
            obs["paired_batch_index"] = batch_idx
            obs["buffer_trace_number"] = buffer_trace_number
            obs["target_trace_number"] = target_trace_number
            session.observations[-1] = obs
            session._write_json(session.analysis_dir / f"iter_{suggestion.iteration:03d}_paired_quality.json", obs)
            session._write_history_csv()
            session.save_state()
            rows.append(_row_from_observation(obs))
            completed_observations += 1
            if progress_callback:
                progress_callback(
                    completed_traces,
                    total_traces,
                    f"Compared cycle {cycle_number}, parameter set {batch_idx}",
                )
    return {
        "session": session,
        "engine": engine,
        "rows": rows,
        "landscape": engine.sample_landscape(int(sim_config.get("grid_size", 25))),
        "distributions": engine.dimension_distributions(),
        "paired_response": True,
        "cycles": total_cycles,
        "batch_size": batch_size,
        "total_swv_traces": total_traces,
        "trace_schedule": trace_schedule,
    }


def _write_simulated_raw_measurements(
    engine: SyntheticSWVSimulationEngine,
    experiment_dir: Path,
    params: dict,
    iteration: int,
) -> Tuple[Path, dict]:
    """Write measurement-like CSVs; headless analysis computes all derived outputs."""
    legacy_dir = experiment_dir / "legacy" / f"iter_{int(iteration):03d}"
    legacy_dir.mkdir(parents=True, exist_ok=True)
    truth = engine.evaluate_truth(params)
    traces = {}
    synthetic_metrics = {}
    scan_count = 1
    for channel in parse_channels(engine.bo_config.get("channels", [])):
        for scan in range(1, scan_count + 1):
            metrics, trace = engine._channel_measurement(params, truth, iteration, channel, scan=scan)
            if scan == 1:
                traces[str(channel)] = trace
                synthetic_metrics[str(channel)] = metrics
            voltage = [float(value) for value in trace.get("voltage_v") or []]
            current = [float(value) for value in trace.get("current_uA") or []]
            raw_path = legacy_dir / f"ch{int(channel):03d}_meas_{scan:03d}_simulated_swv.csv"
            _write_raw_swv_csv(raw_path, voltage, current)
    payload = {
        "schema_version": 1,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "simulation_engine": {
            "version": 2,
            "analysis_mode": "headless_from_raw_csv",
            "iteration": int(iteration),
            "parameters": dict(params),
            "dimensions": [dim.__dict__ for dim in engine.dimensions],
            "scan_count_per_channel": scan_count,
        },
        "simulation_truth": truth,
        "synthetic_channel_metrics_preview": synthetic_metrics,
        "swv_traces": traces,
    }
    return legacy_dir, payload


def _write_simulated_analysis_summary(path: Path, payload: dict) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    summary = dict(payload)
    summary["headless_analysis_from_simulated_raw"] = False
    summary["direct_simulated_analysis_summary"] = True
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(summary, fh, indent=2)
    return path


def _annotate_paired_payload(payload: dict, cycle: int, parameter_set: int, trace_number: int, phase: str) -> None:
    engine_info = payload.setdefault("simulation_engine", {})
    truth = payload.setdefault("simulation_truth", {})
    metadata = {
        "cycle": int(cycle),
        "parameter_set": int(parameter_set),
        "trace_number": int(trace_number),
        "phase": str(phase),
    }
    engine_info.update(metadata)
    truth.update(
        {
            "paired_cycle": int(cycle),
            "paired_batch_index": int(parameter_set),
            "trace_number": int(trace_number),
            "paired_phase": str(phase),
        }
    )


def _augment_analysis_summary(summary_path: Path, simulation_payload: dict) -> Path:
    with open(summary_path, "r", encoding="utf-8") as fh:
        summary = json.load(fh)
    summary.update(simulation_payload)
    summary["headless_analysis_from_simulated_raw"] = True
    with open(summary_path, "w", encoding="utf-8") as fh:
        json.dump(summary, fh, indent=2)
    return summary_path


def _write_raw_swv_csv(path: Path, voltage: List[float], current: List[float]) -> None:
    with open(path, "w", encoding="utf-8", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(["Potential (V)", "Current (uA)"])
        writer.writerows(zip(voltage, current))


def _row_from_observation(obs: dict) -> dict:
    params = dict(obs.get("params") or {})
    truth = dict(obs.get("simulation_truth") or {})
    quality = dict(obs.get("quality") or {})
    row = {
        "iteration": int(obs.get("iteration", 0)),
        "paired_cycle": int(obs.get("paired_cycle", truth.get("paired_cycle", 0)) or 0),
        "paired_batch_index": int(obs.get("paired_batch_index", truth.get("paired_batch_index", 0)) or 0),
        "buffer_trace_number": int(obs.get("buffer_trace_number", 0) or 0),
        "target_trace_number": int(obs.get("target_trace_number", 0) or 0),
        "Q_run": float(obs.get("Q_run", 0.0)),
        "true_Q": float(truth.get("true_Q", 0.0)),
        "paired_Q": float(truth.get("paired_Q", 0.0) or 0.0),
        "paired_Q_score": float(truth.get("paired_Q_score", 0.0) or 0.0),
        "distance": float(truth.get("normalized_distance", 0.0)),
        "delta_peak": float(quality.get("mean_abs_delta_peak_height_uA", truth.get("expected_delta_peak_uA", 0.0)) or 0.0),
        "expected_delta_peak_uA": float(truth.get("expected_delta_peak_uA", 0.0) or 0.0),
    }
    row.update({name: params.get(name) for name in PARAMETER_ORDER})
    return row


def _clip01(value: Any) -> float:
    try:
        return max(0.0, min(1.0, float(value)))
    except (TypeError, ValueError):
        return 0.0


def _as_bool(value: Any) -> bool:
    if isinstance(value, str):
        return value.strip().lower() not in {"0", "false", "no", "off"}
    return bool(value)


def _linspace(start: float, stop: float, count: int) -> List[float]:
    count = max(2, int(count))
    step = (float(stop) - float(start)) / (count - 1)
    return [float(start) + idx * step for idx in range(count)]


def _product(axes: Sequence[Sequence[float]]) -> Iterable[Tuple[float, ...]]:
    if not axes:
        yield ()
        return
    if len(axes) == 1:
        for a in axes[0]:
            yield (a,)
        return
    for head in axes[0]:
        for tail in _product(axes[1:]):
            yield (head,) + tail
