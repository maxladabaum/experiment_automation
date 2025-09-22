# centris_pure.py
from win32com.client import gencache
import pythoncom, time

class CentrisPumpPure:
    def __init__(self, com_port=8, dev=1, baud=9600, progid="PumpCommServer.PumpComm"):
        self.com_port = int(com_port)
        self.dev = int(dev)
        self.baud = int(baud)
        self.progid = progid
        self.pump = None
        self.steps_per_stroke = 100000  # “100K”
        self.syringe_ul = 1250          # your syringe

    def open(self):
        self.pump = gencache.EnsureDispatch(self.progid)
        try:
            self.pump.EnableLog = True
            self.pump.LogComPort = True
            self.pump.CommandAckTimeout = 18
            self.pump.CommandRetryCount = 3
            self.pump.BaudRate = self.baud
        except Exception:
            pass
        self.pump.PumpInitComm(self.com_port)

    def close(self):
        if self.pump:
            try: self.pump.PumpExitComm()
            finally: self.pump = None

    def set_steps_per_stroke(self, steps:int): self.steps_per_stroke = int(steps)
    def set_syringe_ul(self, ul:float):        self.syringe_ul = float(ul)

    def _send(self, cmd:str, wait_s:float=1.0) -> str:
        try:
            ans = self.pump.PumpSendCommand(cmd, self.dev, "")
            time.sleep(wait_s)
            return ans or ""
        except pythoncom.com_error:
            self.pump.PumpSendNoWait(cmd, self.dev)
            time.sleep(wait_s)
            try: return self.pump.PumpGetLastAnswer(self.dev) or ""
            except pythoncom.com_error: return ""

    # motions / queries
    def initialize(self): return self._send("ZR", 1.5)

    def valve_to(self, port:int):
        # >>> changed from V{port}R to I{port}R <<<
        return self._send(f"I{int(port)}R", 1.0)

    def _ul_to_steps(self, ul:float) -> int:
        return max(0, int(round(self.steps_per_stroke * (float(ul)/float(self.syringe_ul)))))

    def aspirate_ul(self, ul:float):  return self._send(f"A{self._ul_to_steps(ul)}R", 1.0)
    def dispense_ul(self, ul:float):  return self._send(f"D{self._ul_to_steps(ul)}R", 1.0)

