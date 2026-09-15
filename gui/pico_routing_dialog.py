"""Opt-in flow-cell routing diagnostics. Normal methods remain unchanged."""
from datetime import datetime
from pathlib import Path
import queue
import threading
import tkinter as tk
from tkinter import ttk

from core.pico_routing_test import run_flowcell_test


class PicoRoutingDialog(tk.Toplevel):
    def __init__(self, parent, port='', session=None):
        super().__init__(parent)
        self.title('Pico electrode routing test')
        self.geometry('700x480')
        self.minsize(570,400)
        self.session = session
        self.events = queue.Queue()
        self.cancel = threading.Event()
        self.running = False
        self.closing = False
        self.port = tk.StringVar(value=port or '')
        self.route = tk.StringVar(value='RE0/CE0 (current configuration)')
        self.mux = tk.StringVar(value='Keep current')
        body=ttk.Frame(self,padding=12)
        body.pack(fill='both',expand=True)
        ttk.Label(body,text='WE0 MUX stays on physical WE0.',font=('Segoe UI',11,'bold')).pack(anchor='w')
        ttk.Label(body,text='Short OCP and +/-10 mV step tests. Raw packets are always saved.\n'
                  'RE1/CE1 is an experimental reversed-bipot test; physical routing is unverified.\n'
                  'Mixed RE/CE pairs are unavailable. Both cells turn off when the test finishes.',
                  wraplength=650,justify='left').pack(anchor='w',pady=8)
        choices=ttk.Frame(body)
        choices.pack(fill='x')
        self.controls=[]
        for label,var,values in [('Device port',self.port,None),('Route to test',self.route,
                 ['RE0/CE0 (current configuration)','RE1/CE1 candidate']),
                 ('WE MUX position',self.mux,['Keep current']+[str(n) for n in range(1,17)])]:
            row=ttk.Frame(choices);row.pack(fill='x',pady=3)
            ttk.Label(row,text=label,width=20).pack(side='left')
            widget=(ttk.Combobox(row,textvariable=var,values=values,state='readonly',width=38)
                    if values else ttk.Entry(row,textvariable=var,width=20))
            widget.pack(side='left');self.controls.append(widget)
        buttons=ttk.Frame(body);buttons.pack(fill='x',pady=8)
        self.run_button=ttk.Button(buttons,text='Run live test',command=self.start)
        self.run_button.pack(side='left')
        self.stop_button=ttk.Button(buttons,text='Stop test',command=self.cancel.set,state='disabled')
        self.stop_button.pack(side='left',padx=8)
        self.output=tk.Text(body,height=10,state='disabled',wrap='word')
        self.output.pack(fill='both',expand=True)
        self.protocol('WM_DELETE_WINDOW',self.close)
        self.after(100,self.poll)

    def append(self,message):
        self.output.configure(state='normal')
        self.output.insert('end',message+'\n');self.output.see('end')
        self.output.configure(state='disabled')

    def start(self):
        if self.running:
            return
        if self.session is not None and self.session.is_running:
            self.append('An experiment is running. Wait for it to finish.');return
        port=self.port.get().strip()
        if not port:
            self.append('Select the Pico COM port in the Methods tab or enter it here.');return
        mux=None if self.mux.get()=='Keep current' else int(self.mux.get())
        candidate=self.route.get()=='RE1/CE1 candidate'
        # Machine-local output only, never a setup-specific path in a method.
        from config import DATA_DIR
        destination=Path(DATA_DIR)/'routing_tests'/datetime.now().strftime('%Y%m%d_%H%M%S_%f')
        self.append('Results: '+str(destination))
        self.cancel.clear();self.running=True
        if self.session is not None:
            self.session.is_running=True
        self.run_button.configure(state='disabled');self.stop_button.configure(state='normal')
        for widget in self.controls:widget.configure(state='disabled')
        def worker():
            try:
                result=run_flowcell_test(port,destination,log=lambda s:self.events.put(('log',s)),
                                        cancelled=self.cancel.is_set,candidate=candidate,mux_position=mux)
            except Exception as exc:
                result={'completed':False,'error':str(exc),'restore_confirmed':False}
            self.events.put(('done',result))
        threading.Thread(target=worker,daemon=True).start()

    def poll(self):
        while not self.events.empty():
            kind,value=self.events.get_nowait()
            if kind=='log':self.append(value)
            else:
                self.running=False
                if self.session is not None:self.session.is_running=False
                self.run_button.configure(state='normal');self.stop_button.configure(state='disabled')
                for widget in self.controls:
                    widget.configure(state='readonly' if isinstance(widget,ttk.Combobox) else 'normal')
                self.append('Completed; routing remains unverified.' if value['completed'] else 'Stopped: '+value.get('error','unknown error'))
                self.append('Channel 0 restored, both cells off.' if value.get('restore_confirmed') else 'Restoration was not confirmed; see raw output.')
        if self.closing and not self.running:self.destroy();return
        self.after(100,self.poll)

    def close(self):
        if self.running:
            self.closing=True;self.cancel.set()
            self.append('Stopping and restoring the instrument before closing...')
        else:self.destroy()
