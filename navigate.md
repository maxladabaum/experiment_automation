experiment_automation/
│
├── main.py                        # Entry point — root + ElectrochemGUI + mainloop()
├── config.py                      # All constants (syringe, baud, steps, paths, etc.)
├── requirements.txt               # Already exists — keep as-is
├── README_txt.txt                 # Already exists — rename to README.md eventually
├── navigate.md                    # This file for folder navigation
│
├── gui/
│   ├── __init__.py
│   ├── app.py                     # Thin shell — creates notebook, wires tabs together
│   ├── tab_method.py              # CV/SWV param forms, generate/run/add-to-queue
│   ├── tab_queue.py               # Queue tree, copy/paste, run/stop, save/load
│   ├── tab_pump.py                # Pump Control tab UI
│   ├── tab_script.py              # Script Preview tab
│   └── tab_plotter.py             # Plotter tab (matplotlib, live plot, load CSV)
│   └── tab_recipe_maker.py        TODO: # Recipe maker UI 
│
├── core/
│   ├── __init__.py
│   ├── runner.py                  # SerialMeasurementRunner (serial comms + data parsing)
│   ├── method_registry.py         # Hash registry, save_script_file, deduplication
│   ├── session.py                 # Shared state — queue, counter, is_running, runner ref
│   └── mscript_parser.py          # VarType, SI prefixes, parse_mscript_data_package
│
├── tecancavro/
│   ├── __init__.py
│   ├── pump_gui.py                # PumpCtrl class (moved + cleaned from pump_gui.py)
│   ├── centris_pure.py            # Already exists — minimal driver, keep here
│   ├── tecanapi.py                # Already exists — Tecan/Cavro API wrapper
│   ├── transport.py               # Already exists — low-level serial transport
│   └── models.py                  # Already exists — pump models/enums
│
│
├── methods/                       # MethodSCRIPT .ms files saved at runtime
│   └── YYYY-MM-DD/                # Auto-created per day by method_registry.py, to be safe
│       ├── 001_cv.ms
│       ├── 002_swv_ch3.ms
│       └── ...
│   └── archive
│       └── ...
│   └── library_map.py             # Hashmap, also include the method finder tool
│   └── library                    # Methods library
│       └── ...
│
├── measurement_data/              # CSV output saved at runtime by runner.py
│   └── YYYY-MM-DD/                TODO: # Auto-created per day, add day tag for saving (bug fix needed here)
│       ├── 001_cv_meas_001.csv
│       ├── 002_swv_ch3_meas_002.csv
│       └── ...
├── tests/                         #old tests
│
│
├── queues/                        # Saved queue .json files (user-facing save/load)
│   ├── my_experiment.json
│   └── ...
│
└── recipe_maker/                  TODO # Not implemented yet
    └── default blocks
        └── flush.json
        └── add_flush.json
        └── add_c6.json
        └── add_aptamer.json
        └── add_ec4.json
        └── add_ec3.json
    └── queue_reference.json       #Including all pump actions, path to methods
