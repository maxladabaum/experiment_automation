"""
main.py — Application entry point.

Run with:
    python main.py

On Windows with Cavro FUSION installed (32-bit Python):
    C:\\pump32_py38_new\\Scripts\\activate.bat
    python main.py

On 64-bit Python (no pump hardware):
    python main.py
    (pump tab will be hidden / disabled automatically)
"""
#!/usr/bin/env python3
# electrochemistry_automation_gui.py
# For bash use -> source venv_gui/Scripts/activate
# For powershell use -> venv_gui\Scripts\Activate.ps1
# To run -> python -m main

import sys
import os
import tkinter as tk

# Make sure the package root is on sys.path so all imports resolve
ROOT = os.path.dirname(os.path.abspath(__file__))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from gui.app import ElectrochemGUI


def main():
    root = tk.Tk()
    _app = ElectrochemGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
