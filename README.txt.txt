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

Install Python libraries:

pip install -r requirements.txt

Create and activate a 32-bit virtual environment:

C:\Users\Chien Lab>C:\pump32_py38_new\Scripts\activate.bat

(pump32_py38_new) C:\Users\Chien Lab>python "C:\Users\Chien Lab\Documents\GitHub\experiment_automation\gui_script.py


or run C:\Users\Chien Lab>python "C:\Users\Chien Lab\Documents\GitHub\experiment_automation\gui_script.py"



