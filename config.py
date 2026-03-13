"""
config.py — Application-wide constants and defaults.

Edit this file to change hardware defaults, paths, and version info.
All other modules import from here — never hardcode constants elsewhere.
"""

from pathlib import Path
import os

# ── Version ──────────────────────────────────────────────────────────────────
APP_VERSION = "2.1.2"

# ── Pump hardware defaults ────────────────────────────────────────────────────
PUMP_DEFAULT_COM_PORT   = 8
PUMP_DEFAULT_BAUD       = 9600
PUMP_DEFAULT_DEV        = 1
PUMP_DEFAULT_STEPS      = 100_000   # steps / stroke (generic default)
PUMP_DEFAULT_SYRINGE    = 1_250.0   # µL (generic default)
PUMP_SPEED_MIN          = 1
PUMP_SPEED_MAX          = 40

# Calibrated values for the Cavro Centris w/ 250 µL syringe
PREFERRED_STEPS_PER_STROKE  = 181_490
PREFERRED_SYRINGE_UL        = 250.0

# ── File / folder paths ───────────────────────────────────────────────────────
METHODS_DIR     = Path("methods")           # where .ms scripts are saved
DATA_DIR        = Path(r"C:\Users\Chien Lab\Desktop\Data_Drive\unc(master)")  # where measurement CSVs land
#DATA_DIR        = Path("measurement_data") #for local testing purposes
BLOCKS_DIR      = Path("recipe_maker") / "default_blocks"  # where block definitions are saved
SAVE_DATED_METHOD_COPIES = False            # if True, also write methods/YYYY-MM-DD/*.ms working copies
#keep in mind that methods are already double saved under library and the experiments where they are used

# ── Serial device detection keywords ─────────────────────────────────────────
DEVICE_KEYWORDS = ["ESPicoDev", "EmStat", "USB Serial Port", "FTDI"]
DEVICE_BAUDRATE = 230_400

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
