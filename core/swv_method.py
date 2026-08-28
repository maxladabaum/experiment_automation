"""
BO-facing SWV MethodSCRIPT helper that mirrors the original Methods-tab SWV builder.
"""

from __future__ import annotations

import re
from typing import Optional

from core.mscript_parser import to_si_string


EMSTAT_PICO_HIGH_SPEED_BA_RANGES = (
    ("100 nA", "59n"),
    ("1 uA", "590n"),
    ("6 uA", "3687500p"),
    ("13 uA", "7375n"),
    ("25 uA", "14750n"),
    ("50 uA", "29500n"),
    ("100 uA", "59u"),
    ("200 uA", "118u"),
    ("1 mA", "590u"),
    ("5 mA", "2950u"),
)


def format_swv_frequency_hz(value: float) -> str:
    """MethodSCRIPT `meas_loop_swv` expects an integer Hz token."""
    if value <= 0:
        return "0"
    return str(int(round(value)))


def range_labels(profile):
    return [label for label, _selector in profile]


def range_index(profile, label: str) -> int:
    for idx, (option_label, _selector) in enumerate(profile):
        if option_label == label:
            return idx
    raise ValueError(f"Unsupported current range selection: {label}")


def range_selector(profile, label: str) -> str:
    for option_label, selector in profile:
        if option_label == label:
            return selector
    raise ValueError(f"Unsupported current range selection: {label}")


def range_label_value(text: str) -> float:
    raw = (text or "").strip()
    if not raw:
        raise ValueError("Current range label is empty")
    match = re.fullmatch(r"([0-9]*\.?[0-9]+)\s*([pnum]?)\s*(?:A)?", raw, flags=re.IGNORECASE)
    if not match:
        raise ValueError(f"Unsupported current range label: {text}")
    value = float(match.group(1))
    prefix = match.group(2).lower()
    scale = {
        "": 1.0,
        "m": 1e-3,
        "u": 1e-6,
        "n": 1e-9,
        "p": 1e-12,
    }[prefix]
    return value * scale


def normalize_range_label(profile, label: str, direction: str) -> str:
    labels = range_labels(profile)
    if label in labels:
        return label

    stripped = (label or "").strip()
    for option_label, selector in profile:
        if stripped == selector:
            return option_label

    target = range_label_value(stripped)
    choices = [(option_label, range_label_value(option_label)) for option_label in labels]
    if direction == "down":
        eligible = [option_label for option_label, value in choices if value <= target]
        return eligible[-1] if eligible else labels[0]
    if direction == "up":
        eligible = [option_label for option_label, value in choices if value >= target]
        return eligible[0] if eligible else labels[-1]
    return min(choices, key=lambda item: abs(item[1] - target))[0]


def normalize_swv_ba_range_options(method_options: Optional[dict] = None) -> dict:
    options = dict(method_options or {})
    ba = dict(options.get("ba_range") or {})
    mode = str(ba.get("mode", "fixed") or "fixed").strip().lower()
    fixed_label = normalize_range_label(
        EMSTAT_PICO_HIGH_SPEED_BA_RANGES,
        str(ba.get("fixed", "100 nA")),
        "up",
    )
    auto_min_label = normalize_range_label(
        EMSTAT_PICO_HIGH_SPEED_BA_RANGES,
        str(ba.get("auto_min", fixed_label)),
        "down",
    )
    auto_max_label = normalize_range_label(
        EMSTAT_PICO_HIGH_SPEED_BA_RANGES,
        str(ba.get("auto_max", fixed_label)),
        "up",
    )
    if range_index(EMSTAT_PICO_HIGH_SPEED_BA_RANGES, auto_min_label) > range_index(
        EMSTAT_PICO_HIGH_SPEED_BA_RANGES, auto_max_label
    ):
        raise ValueError("Autorange minimum must be less than or equal to autorange maximum.")
    return {
        "mode": "auto" if mode == "auto" else "fixed",
        "fixed_label": fixed_label,
        "auto_min_label": auto_min_label,
        "auto_max_label": auto_max_label,
        "fixed_selector": range_selector(EMSTAT_PICO_HIGH_SPEED_BA_RANGES, fixed_label),
        "auto_min_selector": range_selector(EMSTAT_PICO_HIGH_SPEED_BA_RANGES, auto_min_label),
        "auto_max_selector": range_selector(EMSTAT_PICO_HIGH_SPEED_BA_RANGES, auto_max_label),
    }


def build_swv_methodscript(params: dict, method_options: Optional[dict] = None) -> str:
    # Keep this in lockstep with gui.tab_method.MethodTab._build_swv_script.
    options = dict(method_options or {})
    begin_v = float(params["begin_potential"])
    end_v = float(params["end_potential"])
    amp_v = float(params["amplitude"])
    cond_time_s = float(params["conditioning_time"])
    freq_hz = float(params["frequency"])

    begin = to_si_string(str(params["begin_potential"]), "V")
    end = to_si_string(str(params["end_potential"]), "V")
    step = to_si_string(str(params["step_potential"]), "V")
    amplitude = to_si_string(str(params["amplitude"]), "V")
    frequency = format_swv_frequency_hz(freq_hz)
    cond_pot = to_si_string(str(params["conditioning_potential"]), "V")
    cond_time = str(params["conditioning_time"])
    bandwidth = str(options.get("bandwidth", "4k")).strip().lower()
    if bandwidth not in ("4k", "8k"):
        raise ValueError(f"Unsupported SWV bandwidth: {bandwidth}")
    ba_cfg = normalize_swv_ba_range_options(options)

    min_mv = int((min(begin_v, end_v) - amp_v) * 1000)
    max_mv = int((max(begin_v, end_v) + amp_v) * 1000)
    use_equilibrium_check = cond_time_s > 0
    eq_interval_s = min(0.2, cond_time_s) if use_equilibrium_check else 0.0
    swv_time_step = to_si_string(str(1.0 / freq_hz), "s") if freq_hz > 0 else "0"
    eq_duration = to_si_string(cond_time, "s") if use_equilibrium_check else "0"
    eq_interval = to_si_string(str(eq_interval_s), "s") if use_equilibrium_check else "0"

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
    if ba_cfg["mode"] == "auto":
        parts += [
            f"set_range ba {ba_cfg['auto_max_selector']}",
            f"set_autoranging ba {ba_cfg['auto_min_selector']} {ba_cfg['auto_max_selector']}",
        ]
    else:
        parts += [
            f"set_range ba {ba_cfg['fixed_selector']}",
            f"set_autoranging ba {ba_cfg['fixed_selector']} {ba_cfg['fixed_selector']}",
        ]
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
