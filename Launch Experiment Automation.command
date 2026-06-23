#!/bin/bash

set -u

clear
echo "========================================"
echo "[INFO] Experiment Automation Launcher"

HERE="$(cd "$(dirname "$0")" && pwd)"
APPDIR=""

if [ -f "$HERE/main.py" ]; then
    APPDIR="$HERE"
elif [ -f "$HERE/experiment_automation/main.py" ]; then
    APPDIR="$HERE/experiment_automation"
fi

if [ -z "$APPDIR" ]; then
    echo "[ERROR] Could not find the experiment_automation app folder."
    echo
    echo "Put this launcher either:"
    echo "  1. inside the experiment_automation folder, or"
    echo "  2. one folder above experiment_automation."
    echo
    read -r -p "Press Enter to close..."
    exit 1
fi

cd "$APPDIR" || exit 1

echo "[INFO] App folder: $APPDIR"

if [ -f "venv_gui/bin/activate" ]; then
    echo "[INFO] Virtual environment detected: venv_gui"
    # shellcheck disable=SC1091
    source "venv_gui/bin/activate"
elif [ -f ".venv/bin/activate" ]; then
    echo "[INFO] Virtual environment detected: .venv"
    # shellcheck disable=SC1091
    source ".venv/bin/activate"
elif [ -f "venv/bin/activate" ]; then
    echo "[INFO] Virtual environment detected: venv"
    # shellcheck disable=SC1091
    source "venv/bin/activate"
else
    echo "[INFO] No repo virtual environment detected. Using system Python."
fi

PYTHON_BIN=""
for candidate in python python3; do
    if command -v "$candidate" >/dev/null 2>&1 && "$candidate" -c "import tkinter" >/dev/null 2>&1; then
        PYTHON_BIN="$candidate"
        break
    fi
done

if [ -z "$PYTHON_BIN" ]; then
    echo "[ERROR] Could not find a Python interpreter with tkinter support."
    echo
    echo "Install a Python build that includes Tk, or activate a venv that can import tkinter."
    echo
    read -r -p "Press Enter to close..."
    exit 1
fi

echo "[INFO] Python: $($PYTHON_BIN -c 'import sys; print(sys.executable)')"
echo "[INFO] Launching app in 2 seconds..."
echo "========================================"
sleep 2

$PYTHON_BIN -m main
EC=$?

if [ "$EC" -ne 0 ]; then
    echo
    echo "Launcher detected an error (exit $EC)."
    read -r -p "Press Enter to close..."
fi

exit "$EC"
