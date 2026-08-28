"""
core/queue_eta.py - Optional queue ETA helpers.

This module is intentionally standalone. Nothing imports it by default, so
adding this file does not change runtime behavior unless you call it.
"""

# TODO: Calibrate ETA accuracy against real runs (per-measurement overhead,
# device handshake time, and additional MethodSCRIPT loop types).

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
import re
from typing import Dict, Iterable, List, Mapping, Optional, Tuple


# Conservative defaults; tune these to match your lab timing.
DEFAULT_ITEM_SECONDS: Dict[str, float] = {
    "CV": 120.0,
    "SWV": 90.0,
    "CUSTOM": 120.0,
    "MANUAL": 60.0,
    "PUMP_INIT": 8.0,
    "PUMP_SET_SPEED": 2.0,
    "PUMP_VALVE": 3.0,
    "PUMP_ASPIRATE": 8.0,
    "PUMP_DISPENSE": 8.0,
}


@dataclass(frozen=True)
class QueueETA:
    total_seconds: float
    known_seconds: float
    unknown_item_count: int
    excluded_alert_count: int
    per_item_seconds: List[float]

    @property
    def has_unknowns(self) -> bool:
        return self.unknown_item_count > 0


@dataclass(frozen=True)
class RunningQueueETA:
    current_step_remaining_seconds: Optional[float]
    remaining_after_current_seconds: float
    total_remaining_seconds: Optional[float]
    known_remaining_seconds: float
    unknown_item_count: int
    excluded_alert_count: int
    current_step_predictable: bool

    @property
    def has_unknowns(self) -> bool:
        return self.unknown_item_count > 0


_SCRIPT_ETA_CACHE: Dict[str, Tuple[float, float]] = {}


def _parse_si_value(token: str) -> Optional[float]:
    """Parse MethodSCRIPT SI tokens (e.g., 100m, 2k). Returns None on failure."""
    if token is None:
        return None
    token = str(token).strip()
    if not token:
        return None
    try:
        return float(token)
    except ValueError:
        pass
    m = re.match(r"^([+-]?\d+(?:\.\d+)?)([afpnumkMGTPE])$", token)
    if not m:
        return None
    val = float(m.group(1))
    prefix = m.group(2)
    factors = {
        "a": 1e-18, "f": 1e-15, "p": 1e-12, "n": 1e-9, "u": 1e-6,
        "m": 1e-3, "k": 1e3, "M": 1e6, "G": 1e9, "T": 1e12, "P": 1e15, "E": 1e18,
    }
    return val * factors[prefix]


def _sum_wait_seconds(script_text: str) -> float:
    total = 0.0
    for raw in script_text.splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("wait "):
            token = line.split(maxsplit=1)[1]
            secs = _parse_si_value(token)
            if secs is not None:
                total += max(0.0, float(secs))
    return total


def _estimate_from_script(script_text: str) -> Optional[float]:
    wait_seconds = _sum_wait_seconds(script_text)
    for raw in script_text.splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("meas_loop_cv"):
            tokens = line.split()
            if len(tokens) < 8:
                return None
            begin = _parse_si_value(tokens[3])
            v1 = _parse_si_value(tokens[4])
            v2 = _parse_si_value(tokens[5])
            scan_rate = _parse_si_value(tokens[7])
            if None in (begin, v1, v2, scan_rate):
                return None
            if scan_rate <= 0:
                return None
            n_scans = 1
            for tok in tokens[8:]:
                if tok.startswith("nscans(") and tok.endswith(")"):
                    try:
                        n_scans = max(1, int(tok[len("nscans("):-1]))
                    except ValueError:
                        n_scans = 1
                    break
            # Path length: begin -> v1 -> v2 -> begin
            path = abs(v1 - begin) + abs(v2 - v1) + abs(begin - v2)
            per_scan = path / scan_rate
            return max(0.0, wait_seconds + (per_scan * n_scans))

        if line.startswith("meas_loop_swv"):
            tokens = line.split()
            if len(tokens) < 10:
                return None
            begin = _parse_si_value(tokens[5])
            end = _parse_si_value(tokens[6])
            step = _parse_si_value(tokens[7])
            freq = _parse_si_value(tokens[9]) if len(tokens) > 9 else None
            if None in (begin, end, step, freq):
                return None
            if step <= 0 or freq <= 0:
                return None
            n_scans = 1
            for tok in tokens[10:]:
                if tok.startswith("nscans(") and tok.endswith(")"):
                    try:
                        n_scans = max(1, int(tok[len("nscans("):-1]))
                    except ValueError:
                        n_scans = 1
                    break
            steps = int(abs(end - begin) / step) + 1
            per_scan = steps / freq
            return max(0.0, wait_seconds + (per_scan * n_scans))

        if line.startswith("meas_loop_lsv"):
            tokens = line.split()
            if len(tokens) < 7:
                return None
            begin = _parse_si_value(tokens[3])
            end = _parse_si_value(tokens[4])
            scan_rate = _parse_si_value(tokens[6])
            if None in (begin, end, scan_rate):
                return None
            if scan_rate <= 0:
                return None
            path = abs(end - begin)
            total = path / scan_rate
            return max(0.0, wait_seconds + total)
    return None


def _estimate_from_script_path(script_path: str) -> Optional[float]:
    try:
        path = Path(script_path)
    except Exception:
        return None
    if not path.exists():
        return None
    try:
        mtime = path.stat().st_mtime
    except OSError:
        mtime = None
    cache_key = str(path.resolve())
    if mtime is not None:
        cached = _SCRIPT_ETA_CACHE.get(cache_key)
        if cached and cached[0] == mtime:
            return cached[1]
    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return None
    est = _estimate_from_script(text)
    if est is not None and mtime is not None:
        _SCRIPT_ETA_CACHE[cache_key] = (mtime, est)
    return est


def estimate_item_seconds(
    item: Mapping[str, object],
    default_item_seconds: Optional[Mapping[str, float]] = None,
    include_alert_waits: bool = False,
) -> Optional[float]:
    """
    Return estimated seconds for one queue item.

    Returns:
      - float seconds when estimated
      - 0.0 for ALERT when include_alert_waits=False
      - None when item type is unknown
    """
    if default_item_seconds is None:
        default_item_seconds = DEFAULT_ITEM_SECONDS

    item_type = str(item.get("type", "")).strip().upper()
    if not item_type:
        return None

    if item_type == "PAUSE":
        try:
            return max(0.0, float(item.get("pause_seconds", 0.0)))
        except (TypeError, ValueError):
            return None

    if item_type == "ALERT":
        if include_alert_waits:
            # Alert pauses are user-dependent; treat as unknown unless caller
            # chooses to model this separately.
            return None
        return 0.0

    if item_type.startswith("PUMP_"):
        return default_item_seconds.get(item_type)

    # Prefer script-based ETA when available (CV/SWV/Custom).
    script_path = item.get("script_path")
    if script_path:
        est = _estimate_from_script_path(str(script_path))
        if est is not None:
            return est

    # Measurement techniques: CV/SWV/etc.
    return default_item_seconds.get(item_type)


def estimate_queue_eta(
    queue_items: Iterable[Mapping[str, object]],
    start_index: int = 0,
    step_delay_seconds: float = 0.0,
    default_item_seconds: Optional[Mapping[str, float]] = None,
    include_alert_waits: bool = False,
) -> QueueETA:
    """
    Estimate queue duration from queue items.

    step_delay_seconds is applied between estimated items (not after the last).
    """
    if default_item_seconds is None:
        default_item_seconds = DEFAULT_ITEM_SECONDS

    items = list(queue_items)
    if start_index < 0:
        start_index = 0
    if start_index >= len(items):
        return QueueETA(0.0, 0.0, 0, 0, [])

    total_seconds = 0.0
    known_seconds = 0.0
    unknown_item_count = 0
    excluded_alert_count = 0
    per_item_seconds: List[float] = []

    pending = items[start_index:]
    for index, item in enumerate(pending):
        est = estimate_item_seconds(
            item,
            default_item_seconds=default_item_seconds,
            include_alert_waits=include_alert_waits,
        )

        item_type = str(item.get("type", "")).strip().upper()
        if item_type == "ALERT" and not include_alert_waits:
            excluded_alert_count += 1

        if est is None:
            unknown_item_count += 1
            est = 0.0
        else:
            known_seconds += est

        per_item_seconds.append(est)
        total_seconds += est

        # Add inter-step delay between items.
        if index < len(pending) - 1:
            total_seconds += max(0.0, float(step_delay_seconds))
            known_seconds += max(0.0, float(step_delay_seconds))

    return QueueETA(
        total_seconds=total_seconds,
        known_seconds=known_seconds,
        unknown_item_count=unknown_item_count,
        excluded_alert_count=excluded_alert_count,
        per_item_seconds=per_item_seconds,
    )


def estimate_running_queue_eta(
    queue_items: Iterable[Mapping[str, object]],
    next_index: int,
    current_step_elapsed_seconds: float = 0.0,
    current_step_estimated_seconds: Optional[float] = None,
    step_delay_seconds: float = 0.0,
    default_item_seconds: Optional[Mapping[str, float]] = None,
    include_alert_waits: bool = False,
    include_next_step_delay: bool = True,
) -> RunningQueueETA:
    """
    Estimate remaining time for an active queue step plus everything after it.

    next_index is the absolute queue index of the next queued item that would
    start after the current active step completes.
    """
    if default_item_seconds is None:
        default_item_seconds = DEFAULT_ITEM_SECONDS

    items = list(queue_items)
    if next_index < 0:
        next_index = 0

    after_eta = estimate_queue_eta(
        items,
        start_index=next_index,
        step_delay_seconds=step_delay_seconds,
        default_item_seconds=default_item_seconds,
        include_alert_waits=include_alert_waits,
    )

    extra_delay = 0.0
    if include_next_step_delay and next_index < len(items):
        extra_delay = max(0.0, float(step_delay_seconds))

    remaining_after_current_seconds = after_eta.total_seconds + extra_delay
    known_remaining_seconds = after_eta.known_seconds + extra_delay

    current_step_predictable = current_step_estimated_seconds is not None
    current_step_remaining_seconds: Optional[float]
    total_remaining_seconds: Optional[float]
    unknown_item_count = after_eta.unknown_item_count

    if current_step_predictable:
        current_step_remaining_seconds = max(
            0.0,
            float(current_step_estimated_seconds) - max(0.0, float(current_step_elapsed_seconds)),
        )
        total_remaining_seconds = current_step_remaining_seconds + remaining_after_current_seconds
        known_remaining_seconds += current_step_remaining_seconds
    else:
        current_step_remaining_seconds = None
        total_remaining_seconds = None
        unknown_item_count += 1

    return RunningQueueETA(
        current_step_remaining_seconds=current_step_remaining_seconds,
        remaining_after_current_seconds=remaining_after_current_seconds,
        total_remaining_seconds=total_remaining_seconds,
        known_remaining_seconds=known_remaining_seconds,
        unknown_item_count=unknown_item_count,
        excluded_alert_count=after_eta.excluded_alert_count,
        current_step_predictable=current_step_predictable,
    )


def eta_finish_time(total_seconds: float, now: Optional[datetime] = None) -> datetime:
    """Return estimated finish datetime."""
    if now is None:
        now = datetime.now()
    return now + timedelta(seconds=max(0.0, float(total_seconds)))


def format_duration(total_seconds: float) -> str:
    """Format seconds as Hh Mm Ss."""
    total = int(round(max(0.0, float(total_seconds))))
    h, rem = divmod(total, 3600)
    m, s = divmod(rem, 60)
    if h > 0:
        return f"{h}h {m}m {s}s"
    if m > 0:
        return f"{m}m {s}s"
    return f"{s}s"
