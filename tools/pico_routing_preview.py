"""Local routing investigation UI. No serial access or measurement execution.

Run with .venv64/Scripts/python.exe tools/pico_routing_preview.py.
Selections describe physical Pico terminals, not MUX RE/CE positions.
"""

import tkinter as tk
from tkinter import ttk


def assess_route(we_mux, reference, counter):
    if we_mux not in range(1, 17) or reference not in (0, 1) or counter not in (0, 1):
        raise ValueError('Choose WE MUX position 1-16 and physical RE/CE terminal 0 or 1.')
    route = f'WE0 via MUX position {we_mux} + RE{reference} + CE{counter}'
    if reference == counter == 0:
        status = 'Existing channel-0 terminal configuration'
        detail = ('The application already uses Pico channel 0. Its MUX addressing '
                  'selects matching WE and RE/CE MUX positions. Actual electrode '
                  'connections still depend on the board wiring.')
    elif reference == counter == 1:
        status = 'Unverified: no executable route supplied'
        detail = ('Investigate channel 1 as the main potentiostat and WE0 as the '
                  'secondary bipot working electrode. Reverse-role support, waveform '
                  'tracking, current acquisition, and firmware/mode compatibility '
                  'must be established before connecting this route to execution. '
                  'Changing set_pgstat_chan alone does not implement this request.')
    else:
        status = 'Unverified: no executable route supplied'
        detail = ('Independent physical RE/CE cross-routing is not established by '
                  'the documented channel selector or shared-reference bipot mode. '
                  'Requires manufacturer-confirmed internal routing or an external '
                  'electrode switch. This is not a selection of MUX RE/CE positions.')
    return route, status, detail


class RoutingPreview(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent, padding=16)
        self.pack(fill='both', expand=True)
        self.we = tk.IntVar(value=1)
        self.reference = tk.IntVar(value=0)
        self.counter = tk.IntVar(value=0)
        ttk.Label(self, text='Experimental Pico terminal routing',
                  font=('Segoe UI', 12, 'bold')).pack(anchor='w')
        ttk.Label(self, text='Preview only. No settings are applied to the instrument.').pack(anchor='w', pady=8)
        for label, variable, values in (
            ('WE0 MUX position', self.we, tuple(range(1, 17))),
            ('Physical reference terminal: RE', self.reference, (0, 1)),
            ('Physical counter terminal: CE', self.counter, (0, 1)),
        ):
            row = ttk.Frame(self)
            row.pack(fill='x', pady=4)
            ttk.Label(row, text=label, width=32).pack(side='left')
            widget = ttk.Combobox(row, textvariable=variable, values=values,
                                  state='readonly', width=6)
            widget.pack(side='left')
            widget.bind('<<ComboboxSelected>>', self.refresh)
        self.output = tk.StringVar()
        ttk.Label(self, textvariable=self.output, wraplength=580,
                  justify='left').pack(fill='x', pady=16)
        ttk.Label(self, text='CE1/RE1 connected: UNKNOWN — requires physical inspection.',
                  wraplength=580).pack(anchor='w')
        ttk.Button(self, text='Run on instrument (unavailable)', state='disabled').pack(anchor='w', pady=12)
        self.refresh()

    def refresh(self, _event=None):
        self.output.set('\n\n'.join(assess_route(
            self.we.get(), self.reference.get(), self.counter.get())))


if __name__ == '__main__':
    root = tk.Tk()
    root.title('Local Pico routing investigation — preview only')
    RoutingPreview(root)
    root.mainloop()
