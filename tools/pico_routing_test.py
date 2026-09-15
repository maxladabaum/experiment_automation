"""Open the live routing-test dialog without restarting the main application."""
import argparse
from pathlib import Path
import sys
import tkinter as tk

sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from gui.pico_routing_dialog import PicoRoutingDialog


if __name__ == '__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--port',default='',help='Pico COM port; verified by device identity')
    args=parser.parse_args()
    root=tk.Tk()
    root.withdraw()
    dialog=PicoRoutingDialog(root,port=args.port)
    root.wait_window(dialog)
    root.destroy()
