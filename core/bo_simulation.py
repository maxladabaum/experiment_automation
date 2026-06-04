"""Synthetic SWV simulation engine for BO tuning.

The engine builds a continuous parameter landscape, generates fake SWV traces,
scores synthetic channel metrics, and lets the existing BO session machinery
walk that landscape.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import json
import math
from pathlib import Path
import random
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

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
        dims.append(
            SimulationDimension(
                name=str(raw.get("name", "")).strip(),
                minimum=minimum,
                maximum=maximum,
                optimum=optimum,
                spread=spread,
                landscape=landscape,
                weight=max(0.0, float(raw.get("weight", 1.0))),
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
        self.trace_points = max(41, int(self.sim_config.get("trace_points", 121)))

    def analysis_payload(self, params: dict, iteration: int) -> dict:
        truth = self.evaluate_truth(params)
        metrics = {}
        traces = {}
        for channel in parse_channels(self.bo_config.get("channels", [])):
            channel_metrics, trace = self._channel_measurement(params, truth, iteration, channel)
            metrics[str(channel)] = channel_metrics
            if len(traces) < 3:
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
        true_q = _clip01(
            self.peak_emphasis * peak_score
            + 0.15 * noise_score
            + 0.10 * shape_score
            + max(0.0, 0.05 - max(0.0, self.peak_emphasis - 0.70) * 0.05)
        )
        return {
            "true_Q": true_q,
            "landscape_Q": landscape_q,
            "normalized_distance": distance,
            "component_scores": {
                dim.name: score for dim, score in zip(self.dimensions, component_scores)
            },
            "optimum": {dim.name: dim.optimum for dim in self.dimensions},
        }

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
                    "landscape_Q": truth["landscape_Q"],
                    "distance": truth["normalized_distance"],
                }
            )
            points.append(point)
        return {"dimensions": [dim.__dict__ for dim in dims], "points": points}

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

    def _channel_measurement(self, params: dict, truth: dict, iteration: int, channel: int) -> Tuple[dict, dict]:
        rng = random.Random(self.seed + int(iteration) * 1009 + int(channel) * 131)
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
            "success_score": _clip01(0.60 + 0.40 * channel_q),
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


def run_optimizer_simulation(
    bo_config: dict,
    sim_config: dict,
    output_root: str | Path,
    iterations: int,
) -> dict:
    cfg = normalize_bo_config(bo_config)
    engine = SyntheticSWVSimulationEngine(cfg, sim_config)
    output_root = Path(output_root)
    output_root.mkdir(parents=True, exist_ok=True)
    session = BOIntegrationSession(
        cfg,
        output_root,
        config_path=None,
        analysis_output_dir=output_root / "analysis_outputs",
    )
    rows = []
    for _idx in range(max(1, int(iterations))):
        suggestion = session.ask_next()
        payload = engine.analysis_payload(suggestion.params, suggestion.iteration)
        path = session.analysis_dir / f"iter_{suggestion.iteration:03d}_simulation_engine.json"
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, indent=2)
        obs = session.import_analysis(path, notes="Simulation engine")
        obs["simulation_truth"] = payload["simulation_truth"]
        obs["swv_trace_preview"] = payload.get("swv_traces", {})
        session.observations[-1] = obs
        session._write_json(session.analysis_dir / f"iter_{suggestion.iteration:03d}_quality.json", obs)
        session.save_state()
        rows.append(_row_from_observation(obs))
    return {
        "session": session,
        "engine": engine,
        "rows": rows,
        "landscape": engine.sample_landscape(int(sim_config.get("grid_size", 25))),
    }


def _row_from_observation(obs: dict) -> dict:
    params = dict(obs.get("params") or {})
    truth = dict(obs.get("simulation_truth") or {})
    row = {
        "iteration": int(obs.get("iteration", 0)),
        "Q_run": float(obs.get("Q_run", 0.0)),
        "true_Q": float(truth.get("true_Q", 0.0)),
        "distance": float(truth.get("normalized_distance", 0.0)),
    }
    row.update({name: params.get(name) for name in PARAMETER_ORDER})
    return row


def _clip01(value: Any) -> float:
    try:
        return max(0.0, min(1.0, float(value)))
    except (TypeError, ValueError):
        return 0.0


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
