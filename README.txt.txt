Cavro Centris Pump — Python Control (Windows)

This bundle contains Python scripts and a small GUI to control a Tecan Cavro Centris pump via the Cavro FUSION COM drivers.

It uses the PumpComm COM server that ships with Cavro FUSION and talks to the pump with Cavro command strings (e.g., ZR, A…R, D…R, I#R).

CONTENTS

pump_gui.py — Windows GUI (connect, initialize, valve, aspirate/dispense)

centris_pure.py — minimal driver class (pure command style to dev=1)

sample_to_waste_pure.py — example: sample -> waste transfer

pump_ad_pure_fixed.py — minimal A/D script

valve_i_sweep_pure.py — step through valve ports 1..9

robust_sample_to_waste.py — “bullet-proof” scripted sample -> waste

requirements.txt — Python dependency list (pywin32)

README.txt — this document

WHY 32-BIT PYTHON?

Because the Cavro COM servers are 32-bit (installed under C:\Program Files (x86)\…). COM servers must be loaded by a process of the same bitness. Therefore:

Use 32-bit Python.

Register the DLLs with the 32-bit regsvr32.exe from C:\Windows\SysWOW64\.

REQUIREMENTS (one-time setup)

Install Cavro FUSION (standard Tecan installer).
Default path:
C:\Program Files (x86)\Tecan\Cavro FUSION Software Vx.x.x\

Register the PumpComm COM server (32-bit). Open Command Prompt as Administrator and run:

C:\Windows\SysWOW64\regsvr32.exe "C:\Program Files (x86)\Tecan\Cavro FUSION Software Vx.x.x\PumpCommServer.dll"


You should see: “DllRegisterServer … succeeded.”

Install 32-bit Python (3.13 or 3.11, x86 build).
Example path:
C:\Users\<you>\AppData\Local\Programs\Python\Python313-32\python.exe

Create and activate a 32-bit virtual environment:

"C:\Users\<you>\AppData\Local\Programs\Python\Python313-32\python.exe" -m venv C:\pump32
C:\pump32\Scripts\activate


Install Python libraries:

pip install -r requirements.txt


(Only pywin32 is required; tkinter ships with the official Python installer.)

DEFAULT CONNECTION SETTINGS (proven working)

COM port: COM8 (confirm in Device Manager → Ports (COM & LPT))

Baud: 9600

Device address: 1

Valve command family: I#R (not V#R)

Syringe: 1250 µL

Steps per stroke: 100,000 (“100K”) → 80 steps/µL

HOW TO RUN THE GUI

Close Cavro FUSION (it holds the COM port).

Activate the 32-bit venv:

C:\pump32\Scripts\activate


Run:

python pump_gui.py


In the GUI:

COM port: 8, Baud: 9600, Device #: 1

Click Connect, then Initialize (ZR)

Use Valve quick buttons (1–9) to move the valve (I#R)

Enter a volume and click Aspirate / Dispense

The GUI runs pump actions on worker threads and calls pythoncom.CoInitialize() per thread (required for COM).

HOW TO RUN A SCRIPT (console)

Activate the venv, then for example:

python pump_ad_pure_fixed.py

COMMON PITFALLS & FIXES

“CoInitialize has not been called.”
You used COM in a thread without COM init. In the GUI this is handled; if writing your own threaded code, call pythoncom.CoInitialize() / pythoncom.CoUninitialize() in that thread.

“Class not registered.”
Register PumpComm with the 32-bit registrar (admin):

C:\Windows\SysWOW64\regsvr32.exe "...\PumpCommServer.dll"


Port opens but no reply.
Close FUSION, confirm COM8, use Device 1, stick to pure command style:
PumpSendCommand(cmd, dev, ""). For valve moves use I#R.

Valve strains / stalls.
Wrong port number for your distributor or a blocked line. Try a different port. Keep 0.8–1.2 s delays after valve moves.

A/D fails after a bit.
Run ZR to re-reference, try smaller volumes (start 10–50 µL), then scale.

MINIMAL CODE PATTERN
from win32com.client import gencache
import time

PROGID, COM, DEV = "PumpCommServer.PumpComm", 8, 1
p = gencache.EnsureDispatch(PROGID)
p.PumpInitComm(COM)

# Pure style + 3-arg PumpSendCommand; short sleeps between commands
p.PumpSendCommand("ZR", DEV, ""); time.sleep(1.5)

STEPS, SYR = 100000, 1250  # 80 steps/µL
def ul_to_steps(ul): return int(round(STEPS * (ul / SYR)))

p.PumpSendCommand("I1R", DEV, ""); time.sleep(0.9)                 # valve to port 1
p.PumpSendCommand(f"A{ul_to_steps(50)}R", DEV, ""); time.sleep(1.0) # aspirate 50 µL
p.PumpSendCommand("I9R", DEV, ""); time.sleep(0.9)                 # valve to port 9 (waste)
p.PumpSendCommand(f"D{ul_to_steps(50)}R", DEV, ""); time.sleep(1.0) # dispense 50 µL

p.PumpExitComm()

UNINSTALL / CLEANUP

Unregister PumpComm (Admin):

C:\Windows\SysWOW64\regsvr32.exe /u "...\PumpCommServer.dll"


Remove venv:

rmdir /s /q C:\pump32