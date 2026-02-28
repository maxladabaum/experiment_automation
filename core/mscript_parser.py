"""
core/mscript_parser.py — PalmSens MethodSCRIPT parser.

Pure Python, no GUI imports.  Adapted from the official mscript.py sample.
Provides:
  - VarType namedtuple
  - SI_PREFIX_FACTOR dict
  - MSCRIPT_VAR_TYPES_DICT
  - MScriptVar   — parses a single variable token from a data packet
  - parse_mscript_data_package(line) → list[MScriptVar] | None
  - to_si_string(value_str, unit) → str   (float → device SI string)
"""

import collections
import math
import warnings
from typing import Dict, List, Optional

# ── Named tuple for variable metadata ────────────────────────────────────────
VarType = collections.namedtuple("VarType", ["id", "name", "unit"])

# ── SI prefix → multiplier ────────────────────────────────────────────────────
SI_PREFIX_FACTOR: Dict[str, float] = {
    "a": 1e-18, "f": 1e-15, "p": 1e-12, "n": 1e-9,  "u": 1e-6,
    "m": 1e-3,  " ": 1e0,   "k": 1e3,   "M": 1e6,   "G": 1e9,
    "T": 1e12,  "P": 1e15,  "E": 1e18,  "i": 1e0,
}

# ── Variable-type registry ────────────────────────────────────────────────────
MSCRIPT_VAR_TYPES_LIST: List[VarType] = [
    VarType("aa", "unknown",                 ""),
    VarType("ab", "WE vs RE potential",      "V"),
    VarType("ac", "CE vs GND potential",     "V"),
    VarType("ad", "SE vs GND potential",     "V"),
    VarType("ae", "RE vs GND potential",     "V"),
    VarType("af", "WE vs GND potential",     "V"),
    VarType("ag", "WE vs CE potential",      "V"),
    VarType("as", "AIN0 potential",          "V"),
    VarType("at", "AIN1 potential",          "V"),
    VarType("au", "AIN2 potential",          "V"),
    VarType("av", "AIN3 potential",          "V"),
    VarType("aw", "AIN4 potential",          "V"),
    VarType("ax", "AIN5 potential",          "V"),
    VarType("ay", "AIN6 potential",          "V"),
    VarType("az", "AIN7 potential",          "V"),
    VarType("ba", "WE current",              "A"),
    VarType("ca", "Phase",                   "degrees"),
    VarType("cb", "Impedance",               "\u2126"),
    VarType("cc", "Z_real",                  "\u2126"),
    VarType("cd", "Z_imag",                  "\u2126"),
    VarType("ce", "EIS E TDD",               "V"),
    VarType("cf", "EIS I TDD",               "A"),
    VarType("cg", "EIS sampling frequency",  "Hz"),
    VarType("ch", "EIS E AC",                "Vrms"),
    VarType("ci", "EIS E DC",                "V"),
    VarType("cj", "EIS I AC",                "Arms"),
    VarType("ck", "EIS I DC",                "A"),
    VarType("da", "Applied potential",        "V"),
    VarType("db", "Applied current",          "A"),
    VarType("dc", "Applied frequency",        "Hz"),
    VarType("dd", "Applied AC amplitude",     "Vrms"),
    VarType("ea", "Channel",                  ""),
    VarType("eb", "Time",                     "s"),
    VarType("ec", "Pin mask",                 ""),
    VarType("ed", "Temperature",              "\u00B0 Celsius"),
    VarType("ee", "Count",                    ""),
    VarType("ha", "Generic current 1",        "A"),
    VarType("hb", "Generic current 2",        "A"),
    VarType("hc", "Generic current 3",        "A"),
    VarType("hd", "Generic current 4",        "A"),
    VarType("ia", "Generic potential 1",      "V"),
    VarType("ib", "Generic potential 2",      "V"),
    VarType("ic", "Generic potential 3",      "V"),
    VarType("id", "Generic potential 4",      "V"),
    VarType("ja", "Misc. generic 1",          ""),
    VarType("jb", "Misc. generic 2",          ""),
    VarType("jc", "Misc. generic 3",          ""),
    VarType("jd", "Misc. generic 4",          ""),
]

MSCRIPT_VAR_TYPES_DICT: Dict[str, VarType] = {v.id: v for v in MSCRIPT_VAR_TYPES_LIST}


# ── Public helpers ────────────────────────────────────────────────────────────

def get_variable_type(var_id: str) -> VarType:
    """Look up a VarType by its two-character id.  Returns an 'unknown' entry
    and emits a warning if the id is not recognised."""
    if var_id in MSCRIPT_VAR_TYPES_DICT:
        return MSCRIPT_VAR_TYPES_DICT[var_id]
    warnings.warn(f'Unsupported VarType id "{var_id}"!')
    return VarType(var_id, "unknown", "")


class MScriptVar:
    """Parse and hold a single variable token from a MethodSCRIPT data line."""

    def __init__(self, data: str):
        assert len(data) >= 10, f"Token too short: {data!r}"
        self.data        = data[:]
        self.id          = data[0:2]
        if data[2:10] == "     nan":
            self.raw_value  = math.nan
            self.si_prefix  = " "
        else:
            self.raw_value  = self._decode_value(data[2:9])
            self.si_prefix  = data[9]
        self.raw_metadata = data.split(",")[1:]
        self.metadata     = self._parse_metadata(self.raw_metadata)

    # ── Properties ──────────────────────────────────────────────────────────

    @property
    def type(self) -> VarType:
        return get_variable_type(self.id)

    @property
    def si_prefix_factor(self) -> float:
        return SI_PREFIX_FACTOR[self.si_prefix]

    @property
    def value(self) -> float:
        return self.raw_value * self.si_prefix_factor

    # ── Static helpers ───────────────────────────────────────────────────────

    @staticmethod
    def _decode_value(var: str) -> int:
        assert len(var) == 7
        return int(var, 16) - (2 ** 27)

    @staticmethod
    def _parse_metadata(tokens: List[str]) -> Dict[str, int]:
        metadata: Dict[str, int] = {}
        for token in tokens:
            if len(token) == 2 and token[0] == "1":
                metadata["status"] = int(token[1], 16)
            if len(token) == 3 and token[0] == "2":
                metadata["cr"] = int(token[1:], 16)
        return metadata


def parse_mscript_data_package(line: str) -> Optional[List[MScriptVar]]:
    """Parse a complete MethodSCRIPT data line (e.g. ``Pab...;ba...\\n``).

    Returns a list of :class:`MScriptVar` objects, or ``None`` if the line
    does not look like a data packet.
    """
    if not (line.startswith("P") and line.endswith("\n")):
        return None

    vars_out: List[MScriptVar] = []
    for var in line[1:-1].split(";"):
        if len(var) < 10:
            continue
        value_token = var[2:9]
        if any(ch not in "0123456789ABCDEFabcdef" for ch in value_token):
            continue
        if var[9] not in SI_PREFIX_FACTOR:
            continue
        try:
            vars_out.append(MScriptVar(var))
        except Exception:
            continue
    return vars_out


def to_si_string(value_str: str, unit: str = "V") -> str:
    """Convert a plain float string to the SI-prefix string the device expects.

    Examples::

        to_si_string("0.5",  "V")   → "500m"
        to_si_string("15",   "Hz")  → "15"
        to_si_string("0.02", "V")   → "20m"
    """
    try:
        val = float(value_str)
    except (ValueError, TypeError):
        return value_str

    if unit in ("V", "V/s"):
        if val == 0:
            return "0"
        milli = val * 1_000.0
        formatted = f"{milli:.12f}".rstrip("0").rstrip(".")
        if formatted in ("", "-0", "+0"):
            formatted = "0"
        return f"{formatted}m"

    if unit == "Hz":
        return f"{int(val)}" if float(val).is_integer() else f"{val:g}"

    return value_str  # fallback
