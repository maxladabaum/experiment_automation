"""
config.py — Application-wide constants and defaults.

Machine-specific hardware ports and data paths can be overridden by a local
JSON config file outside this repository. See local_config.example.json.
All other modules import from here — never hardcode constants elsewhere.
"""

from pathlib import Path
import json
import os


def _default_local_config_path() -> Path:
    configured = os.getenv("EA_LOCAL_CONFIG_PATH", "").strip()
    if configured:
        return Path(configured).expanduser()

    local_app_data = os.getenv("LOCALAPPDATA", "").strip()
    if local_app_data:
        return Path(local_app_data) / "ExperimentAutomation" / "local_config.json"

    return Path.home() / ".experiment_automation" / "local_config.json"


LOCAL_CONFIG_PATH = _default_local_config_path()


def _load_local_config() -> dict:
    try:
        with open(LOCAL_CONFIG_PATH, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


_LOCAL_CONFIG = _load_local_config()


def _local_value(key: str, default=None):
    return _LOCAL_CONFIG.get(key, default)


def _env_or_local(env_name: str, local_key: str, default=""):
    value = os.getenv(env_name)
    if value is not None:
        return value
    return _local_value(local_key, default)


def _as_int(value, default: int) -> int:
    try:
        if value in (None, ""):
            return default
        return int(value)
    except (TypeError, ValueError):
        return default


def _as_com_port_int(value, default: int) -> int:
    if isinstance(value, str) and value.strip().upper().startswith("COM"):
        value = value.strip()[3:].strip()
    return _as_int(value, default)


def _as_float(value, default: float) -> float:
    try:
        if value in (None, ""):
            return default
        return float(value)
    except (TypeError, ValueError):
        return default


def _as_path(value, default) -> Path:
    raw = default if value in (None, "") else value
    return Path(os.path.expandvars(str(raw))).expanduser()


def _as_string_list(value, default):
    if value in (None, ""):
        return list(default)
    if isinstance(value, str):
        return [part.strip() for part in value.split(",") if part.strip()]
    if isinstance(value, (list, tuple)):
        return [str(part).strip() for part in value if str(part).strip()]
    return list(default)


# ── Version ──────────────────────────────────────────────────────────────────
APP_VERSION = "2.1.2"

# ── Pump hardware defaults ────────────────────────────────────────────────────
PUMP_DEFAULT_COM_PORT = _as_com_port_int(
    _env_or_local("EA_PUMP_COM_PORT", "pump_com_port", 8), 8
)
PUMP_DEFAULT_BAUD = _as_int(
    _env_or_local("EA_PUMP_BAUD", "pump_baud", 9600), 9600
)
PUMP_DEFAULT_DEV = _as_int(
    _env_or_local("EA_PUMP_DEV", "pump_dev", 1), 1
)
PUMP_DEFAULT_STEPS = _as_int(
    _env_or_local("EA_PUMP_STEPS", "pump_steps", 100_000), 100_000
)   # steps / stroke (generic default)
PUMP_DEFAULT_SYRINGE = _as_float(
    _env_or_local("EA_PUMP_SYRINGE_UL", "pump_syringe_ul", 1_250.0), 1_250.0
)   # µL (generic default)
PUMP_SPEED_MIN          = 1
PUMP_SPEED_MAX          = 40

# Calibrated values for the Cavro Centris w/ 250 µL syringe
PREFERRED_STEPS_PER_STROKE = _as_int(
    _env_or_local(
        "EA_PREFERRED_STEPS_PER_STROKE",
        "preferred_steps_per_stroke",
        181_490,
    ),
    181_490,
)
PREFERRED_SYRINGE_UL = _as_float(
    _env_or_local("EA_PREFERRED_SYRINGE_UL", "preferred_syringe_ul", 250.0),
    250.0,
)

# ── File / folder paths ───────────────────────────────────────────────────────
_DEFAULT_DATA_DIR = Path.home() / "Documents" / "Experiment Automation Data"

DATA_DIR = _as_path(  # where measurement CSVs land
    _env_or_local(
        "EA_DATA_DIR",
        "data_dir",
        _DEFAULT_DATA_DIR,
    ),
    _DEFAULT_DATA_DIR,
)
METHODS_DIR = _as_path(  # where user-created .ms scripts and library_map.json are saved
    _env_or_local("EA_METHODS_DIR", "methods_dir", DATA_DIR / "methods"),
    DATA_DIR / "methods",
)
RECIPE_DIR = _as_path(  # where user-created recipes and custom blocks are saved
    _env_or_local("EA_RECIPE_DIR", "recipe_dir", DATA_DIR / "recipe_maker"),
    DATA_DIR / "recipe_maker",
)
_SESSION_ARCHIVE_RAW = _env_or_local(
    "EA_SESSION_ARCHIVE_DIR", "session_archive_dir", ""
)
SESSION_ARCHIVE_DIR = (
    _as_path(_SESSION_ARCHIVE_RAW, _SESSION_ARCHIVE_RAW)
    if _SESSION_ARCHIVE_RAW not in (None, "")
    else None
)
BLOCKS_DIR      = Path("recipe_maker") / "default_blocks"  # bundled default block definitions
SAVE_DATED_METHOD_COPIES = False            # if True, also write methods/YYYY-MM-DD/*.ms working copies
# Bayesian optimization integration (optional)
BO_CONFIG_DIR = Path(os.getenv("EA_BO_CONFIG_DIR", str(Path("optimizer") / "bo_configs")))
BO_DEFAULT_CONFIG_PATH = Path(
    os.getenv("EA_BO_DEFAULT_CONFIG_PATH", str(BO_CONFIG_DIR / "default_swv_bo.json"))
)
BO_LOCAL_PATHS_CONFIG = Path(os.getenv("EA_BO_LOCAL_PATHS_CONFIG", str(BO_CONFIG_DIR / "local_paths.json")))


def _load_bo_local_paths() -> dict:
    try:
        with open(BO_LOCAL_PATHS_CONFIG, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


_BO_LOCAL_PATHS = _load_bo_local_paths()
BO_ANALYSIS_OUTPUT_DIR = Path(
    os.getenv(
        "EA_BO_ANALYSIS_OUTPUT_DIR",
        str(_BO_LOCAL_PATHS.get("analysis_output_dir", "analysis_outputs")),
    )
)
BO_ANALYSIS_FILE_GLOB = os.getenv(
    "EA_BO_ANALYSIS_FILE_GLOB",
    str(_BO_LOCAL_PATHS.get("analysis_file_glob", "*.json")),
)
_DEFAULT_ANALYSIS_PROJECT = Path(__file__).resolve().parent
BO_EXTERNAL_ANALYSIS_PROJECT = Path(
    os.getenv(
        "EA_BO_ANALYSIS_PROJECT",
        str(_BO_LOCAL_PATHS.get("analysis_project", _DEFAULT_ANALYSIS_PROJECT)),
    )
)
BO_EXTERNAL_ANALYSIS_SCRIPT = Path(
    os.getenv(
        "EA_BO_ANALYSIS_SCRIPT",
        str(
            _BO_LOCAL_PATHS.get(
                "analysis_script",
                BO_EXTERNAL_ANALYSIS_PROJECT / "analysis_worker" / "bo_headless.py",
            )
        ),
    )
)
BO_EXTERNAL_ANALYSIS_PYTHON = os.getenv(
    "EA_BO_ANALYSIS_PYTHON",
    str(_BO_LOCAL_PATHS.get("analysis_python", "")),
).strip()
BO_EXTERNAL_ANALYSIS_TIMEOUT_SECONDS = float(
    os.getenv(
        "EA_BO_ANALYSIS_TIMEOUT_SECONDS",
        str(_BO_LOCAL_PATHS.get("analysis_timeout_seconds", 900)),
    )
)
BO_EXTERNAL_ANALYSIS_MODE = os.getenv(
    "EA_BO_ANALYSIS_MODE",
    str(_BO_LOCAL_PATHS.get("analysis_mode", "external")),
).strip().lower()
#keep in mind that methods are already double saved under library and the experiments where they are used

# ── Serial device detection keywords ─────────────────────────────────────────
DEVICE_KEYWORDS = _as_string_list(
    _env_or_local(
        "EA_DEVICE_KEYWORDS",
        "device_keywords",
        ["ESPicoDev", "EmStat", "USB Serial Port", "FTDI"],
    ),
    ["ESPicoDev", "EmStat", "USB Serial Port", "FTDI"],
)
DEVICE_BAUDRATE = _as_int(
    _env_or_local("EA_DEVICE_BAUDRATE", "device_baudrate", 230_400),
    230_400,
)
DEVICE_DEFAULT_PORT = str(
    _env_or_local("EA_POTENTIOSTAT_PORT", "potentiostat_port", "") or ""
).strip()

# ── GUI geometry ──────────────────────────────────────────────────────────────
WINDOW_GEOMETRY = "1400x900"
WINDOW_TITLE    = f"Electrochemistry Automation System  v{APP_VERSION}"

# Slack integration (optional)
# Set these via environment variables on the machine running the GUI.
SLACK_ENABLE         = os.getenv("EA_SLACK_ENABLE", "0").strip().lower() in ("1", "true", "yes", "on")
SLACK_BOT_TOKEN      = os.getenv("EA_SLACK_BOT_TOKEN", "").strip()
SLACK_SIGNING_SECRET = os.getenv("EA_SLACK_SIGNING_SECRET", "").strip()
SLACK_TARGET         = os.getenv("EA_SLACK_TARGET", "").strip()  # channel ID (C/G) or DM ID (D)
SLACK_PORT           = int(os.getenv("EA_SLACK_PORT", "8765"))
SLACK_ONLY_WHEN_EXPERIMENT = os.getenv("EA_SLACK_ONLY_WHEN_EXPERIMENT", "0").strip().lower() in (
    "1", "true", "yes", "on"
)

# ngrok integration (optional, for Slack Events API on local machines)
NGROK_AUTOSTART = os.getenv("EA_NGROK_AUTOSTART", "1").strip().lower() in (
    "1", "true", "yes", "on"
)
NGROK_PATH = os.getenv("EA_NGROK_PATH", "").strip()
NGROK_DOMAIN = os.getenv("EA_NGROK_DOMAIN", "").strip()
