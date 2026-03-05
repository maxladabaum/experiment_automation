experiment_automation/
â”‚ # Session-gated UI: tabs stay hidden until a session is started.
â”œâ”€â”€ main.py                        # Entry point - root + ElectrochemGUI + mainloop()
â”œâ”€â”€ config.py                      # All constants (syringe, baud, steps, paths, etc.)
â”œâ”€â”€ requirements.txt               # Keep as-is
â”œâ”€â”€ README.txt                 
â”œâ”€â”€ navigate.md                    # This file
â”‚
â”œâ”€â”€ gui/
â”‚   â”œâ”€â”€ __init__.py
â”‚   â”œâ”€â”€ app.py                     # Thin shell â€” creates notebook, SessionManager,
â”‚   â”‚                              #   SessionBar, wires all tabs together
â”‚   â”œâ”€â”€ tab_method.py              # CV/SWV/Custom param forms, generate/run/add-to-queue,
â”‚   â”‚                              #   PStrace SWV preset button
â”‚   â”œâ”€â”€ tab_queue.py               # Queue tree, copy/paste, run/stop, save/load,
â”‚   â”‚                              #   routes data_folder to active experiment
â”‚   â”œâ”€â”€ tab_pump.py                # Pump Control tab UI + session autoconnect
â”‚   â”œâ”€â”€ tab_script.py              # Script Preview tab + execution options (raw packet save, step delay)
â”‚   â”œâ”€â”€ tab_plotter.py             # Plotter tab â€” matplotlib, live plot, load CSV,
â”‚   â”‚                              #   uses AutoScaleToolbar for smart Home button
â”‚   â”œâ”€â”€ tab_custom_script.py       # NEW â€” Custom .ms file loader panel (rendered
â”‚   â”‚                              #   inside tab_method params frame)
â”‚   â”œâ”€â”€ session_bar.py             # NEW â€” Bottom-of-window Session + Experiment bar
â”‚   â”‚                              #   (Start/End session, Start/End experiment,
â”‚   â”‚                              #    user/chip-ID/notes fields, status label)
â”‚   â”œâ”€â”€ widgets.py                 # NEW â€” Shared custom widgets:
â”‚   â”‚                              #   AutoScaleToolbar (smart Home, left-click zoom)
â”‚   â””â”€â”€ tab_recipe_maker.py        # Recipe maker UI — block library + recipe builder
â”‚
â”œâ”€â”€ core/
â”‚   â”œâ”€â”€ __init__.py
â”‚   â”œâ”€â”€ runner.py                  # SerialMeasurementRunner â€” serial comms, data
â”‚   â”‚                              #   parsing, CSV save; accepts data_folder arg
â”‚   â”‚                              #   to route output into experiment subfolder
â”‚   â”œâ”€â”€ method_registry.py         # Hash registry, save_script_file, deduplication
â”‚   â”œâ”€â”€ session.py                 # Shared state â€” measurement_queue, counter,
â”‚   â”‚                              #   is_running, runner ref, session_manager slot
â”‚   â”œâ”€â”€ session_manager.py         # NEW â€” Session/Experiment lifecycle:
â”‚   â”‚                              #   folder creation, metadata JSON, session_log.txt,
â”‚   â”‚                              #   require_session() / require_experiment() guards
â”‚   â””â”€â”€ mscript_parser.py         # VarType, SI prefixes, parse_mscript_data_package
â”‚
â”œâ”€â”€ tecancavro/
â”‚   â”œâ”€â”€ __init__.py
â”‚   â”œâ”€â”€ pump_gui.py                # PumpCtrl class
â”‚   â”œâ”€â”€ centris_pure.py            # Minimal Cavro Centris driver
â”‚   â”œâ”€â”€ tecanapi.py                # Tecan/Cavro API wrapper
â”‚   â”œâ”€â”€ transport.py               # Low-level serial transport
â”‚   â””â”€â”€ models.py                  # Pump models / enums
â”‚
â”œâ”€â”€ methods/                       # MethodSCRIPT .ms files saved at runtime
â”‚   â”œâ”€â”€ YYYY-MM-DD/                # Auto-created per day by method_registry.py
â”‚   â”‚   â”œâ”€â”€ 001_cv.ms
â”‚   â”‚   â”œâ”€â”€ 002_swv_ch3.ms
â”‚   â”‚   â””â”€â”€ ...
â”‚   â”œâ”€â”€ archive/
â”‚   â”œâ”€â”€ library_map.py             # Hashmap + method finder tool
â”‚   â””â”€â”€ library/                   # Curated methods library
â”‚       â””â”€â”€ ...
â”‚
â”œâ”€â”€ measurement_data/              # CSV output â€” now organised by session/experiment
â”‚   â””â”€â”€ <session_name>_<timestamp>/          # Created on "Start Session"
â”‚       â”œâ”€â”€ session_metadata.json            # name, user, chip_id, notes, timestamps
â”‚       â”œâ”€â”€ session_log.txt                  # timestamped log of every run in session
â”‚       â””â”€â”€ <experiment_name>_<timestamp>/   # Created on "Start Experiment"
â”‚           â”œâ”€â”€ experiment_metadata.json     # name, notes, timestamps
â”‚           â”œâ”€â”€ 001_cv_143022.csv
â”‚           â”œâ”€â”€ 002_swv_143145.csv
â”‚           â””â”€â”€ ...
â”‚
â”‚   # NOTE: runner.py falls back to a flat YYYY-MM-DD/ subfolder if no
â”‚   # active experiment is set (e.g. during direct "Run Now" without a session).
â”‚
â”œâ”€â”€ tests/                         # Old tests
â”‚
â”œâ”€â”€ queues/                        # Saved queue .json files (user-facing save/load)
â”‚   â”œâ”€â”€ my_experiment.json
â”‚   â””â”€â”€ ...
â”‚
â””â”€â”€ recipe_maker/                  # Recipe maker presets and blocks
    â”œâ”€â”€ default_blocks/
    â”‚   â”œâ”€â”€ flush.json
    â”‚   â”œâ”€â”€ add_c6.json
    â”‚   â”œâ”€â”€ add_aptamer.json
    â”‚   â”œâ”€â”€ add_ec4.json
    â”‚   â””â”€â”€ add_ec3.json
    â””â”€â”€ queue_reference.json

Pump actions (queue types) and handlers:
- PUMP_INIT -> gui/tab_pump.py:_do_init + _queue_init, gui/tab_queue.py:_exec_pump -> PumpCtrl.initialize
- PUMP_SET_SPEED -> gui/tab_pump.py:_do_set_speed + _queue_set_speed, gui/tab_queue.py:_exec_pump -> PumpCtrl.set_speed
- PUMP_VALVE -> gui/tab_pump.py:_do_valve + _queue_valve, gui/tab_queue.py:_exec_pump -> PumpCtrl.valve_to
- PUMP_ASPIRATE -> gui/tab_pump.py:_do_aspirate + _queue_aspirate, gui/tab_queue.py:_exec_pump -> PumpCtrl.aspirate_ul
- PUMP_DISPENSE -> gui/tab_pump.py:_do_dispense + _queue_dispense, gui/tab_queue.py:_exec_pump -> PumpCtrl.dispense_ul

Method paths:
- config.py: METHODS_DIR = methods/
- core/method_registry.py saves to methods/YYYY-MM-DD/ and methods/library/
- methods/library_map.py handles register/lookup for methods/library/


