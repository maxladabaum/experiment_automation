"""Mass-balance calculations for automated serial titrations."""

from __future__ import annotations

from dataclasses import dataclass
import math
import re
from typing import Iterable, List


@dataclass(frozen=True)
class TitrationPoint:
    """Calculated liquid state for one requested concentration."""

    index: int
    target_concentration_um: float
    concentration_before_um: float
    volume_before_stock_ul: float
    stock_added_ul: float
    volume_after_stock_ul: float
    bubble_clear_loss_ul: float
    aliquot_removed_ul: float
    volume_remaining_ul: float


def parse_concentrations(text: str) -> List[float]:
    """Parse comma, whitespace, or semicolon separated concentrations in µM."""
    tokens = [token for token in re.split(r"[\s,;]+", str(text or "").strip()) if token]
    if not tokens:
        raise ValueError("Enter at least one desired concentration.")
    try:
        values = [float(token) for token in tokens]
    except ValueError as exc:
        raise ValueError("Concentrations must be numbers separated by commas or spaces.") from exc
    if any(not math.isfinite(value) for value in values):
        raise ValueError("Concentrations must be finite numbers.")
    return values


def calculate_titration_plan(
    desired_concentrations_um: Iterable[float],
    *,
    stock_concentration_um: float,
    initial_buffer_volume_ul: float,
    aliquot_volume_ul: float,
    bubble_liquid_loss_per_clear_ul: float = 0.0,
    clears_per_stock_addition: int = 0,
) -> List[TitrationPoint]:
    """Calculate exact stock additions for an ascending serial titration.

    The mixing tube begins with analyte-free buffer. At each point, concentrated
    stock is added to the well-mixed liquid, then one aliquot is removed. The
    removal preserves concentration while reducing both volume and analyte mass.
    """
    stock = _positive_finite(stock_concentration_um, "Stock concentration")
    volume = _positive_finite(initial_buffer_volume_ul, "Initial buffer volume")
    aliquot = _positive_finite(aliquot_volume_ul, "Aliquot volume")
    try:
        bubble_loss_per_clear = float(bubble_liquid_loss_per_clear_ul)
    except (TypeError, ValueError) as exc:
        raise ValueError("Bubble liquid loss must be a number.") from exc
    if not math.isfinite(bubble_loss_per_clear) or bubble_loss_per_clear < 0:
        raise ValueError("Bubble liquid loss must be zero or greater.")
    try:
        clear_count = int(clears_per_stock_addition)
    except (TypeError, ValueError) as exc:
        raise ValueError("Clear count must be an integer.") from exc
    if clear_count < 0:
        raise ValueError("Clear count must be zero or greater.")
    targets = [float(value) for value in desired_concentrations_um]
    if not targets:
        raise ValueError("Enter at least one desired concentration.")

    analyte_amount = 0.0  # µM·µL; unit conversion cancels in the balance.
    current_concentration = 0.0
    points: List[TitrationPoint] = []

    for index, target in enumerate(targets, 1):
        if not math.isfinite(target) or target < 0:
            raise ValueError(f"Concentration {index} must be a finite value at or above 0 µM.")
        if target >= stock:
            raise ValueError(
                f"Concentration {index} ({target:g} µM) must be below the "
                f"stock concentration ({stock:g} µM)."
            )
        if target + 1e-12 < current_concentration:
            raise ValueError(
                "Desired concentrations must be nondecreasing; a stock-only "
                "serial titration cannot lower concentration."
            )

        numerator = target * volume - analyte_amount
        stock_added = max(0.0, numerator / (stock - target))
        volume_after_stock = volume + stock_added
        analyte_after_stock = analyte_amount + stock * stock_added
        achieved = analyte_after_stock / volume_after_stock
        bubble_clear_loss = (
            bubble_loss_per_clear * clear_count
            if stock_added > 1e-9
            else 0.0
        )
        total_removed = bubble_clear_loss + aliquot
        if total_removed > volume_after_stock + 1e-9:
            raise ValueError(
                f"Point {index} has only {volume_after_stock:g} µL available, "
                f"less than the {total_removed:g} µL combined bubble-clear "
                "loss and flow-cell aliquot."
            )

        remaining = max(0.0, volume_after_stock - total_removed)
        points.append(
            TitrationPoint(
                index=index,
                target_concentration_um=target,
                concentration_before_um=current_concentration,
                volume_before_stock_ul=volume,
                stock_added_ul=stock_added,
                volume_after_stock_ul=volume_after_stock,
                bubble_clear_loss_ul=bubble_clear_loss,
                aliquot_removed_ul=aliquot,
                volume_remaining_ul=remaining,
            )
        )

        volume = remaining
        current_concentration = achieved
        analyte_amount = achieved * remaining

    return points


def split_transfer(volume_ul: float, max_stroke_ul: float) -> List[float]:
    """Split a liquid transfer into syringe-safe strokes."""
    volume = _positive_finite(volume_ul, "Transfer volume")
    capacity = _positive_finite(max_stroke_ul, "Syringe capacity")
    full_strokes = int(math.floor(volume / capacity))
    chunks = [capacity] * full_strokes
    remainder = volume - full_strokes * capacity
    if remainder > 1e-9:
        chunks.append(remainder)
    return chunks


def _positive_finite(value: float, label: str) -> float:
    try:
        parsed = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{label} must be a number.") from exc
    if not math.isfinite(parsed) or parsed <= 0:
        raise ValueError(f"{label} must be greater than zero.")
    return parsed
