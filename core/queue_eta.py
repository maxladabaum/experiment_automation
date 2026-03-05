"""
core/queue_eta.py - Optional queue ETA helpers.

This module is intentionally standalone. Nothing imports it by default, so
adding this file does not change runtime behavior unless you call it.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Dict, Iterable, List, Mapping, Optional


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
