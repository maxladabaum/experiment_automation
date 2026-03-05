experiment_automation/
| # Session-gated UI: tabs stay hidden until a session is started.
|-- main.py                        # Entry point - root + ElectrochemGUI + mainloop()
|-- config.py                      # All constants (syringe, baud, steps, paths, etc.)
|-- requirements.txt               # Keep as-is
|-- README.txt
|-- navigate.md                    # This file
|
|-- gui/
|   |-- __init__.py
|   |-- app.py                     # Thin shell - creates notebook, SessionManager,
|   |                              #   SessionBar, wires all tabs together
|   |-- tab_method.py              # CV/SWV/Custom param forms, generate/run/add-to-queue,
|   |                              #   PStrace SWV preset button
|   |-- tab_queue.py               # Queue tree, copy/paste, run/stop, save/load,
|   |                              #   routes data_folder to active experiment
|   |-- tab_pump.py                # Pump Control tab UI + session autoconnect
|   |-- tab_script.py              # Script Preview tab + execution options (raw packet save, step delay)
|   |-- tab_plotter.py             # Plotter tab - matplotlib, live plot, load CSV,
|   |                              #   uses AutoScaleToolbar for smart Home button
|   |-- tab_custom_script.py       # NEW - Custom .ms file loader panel (rendered
|   |                              #   inside tab_method params frame)
|   |-- session_bar.py             # NEW - Bottom-of-window Session + Experiment bar
|   |                              #   (Start/End session, Start/End experiment,
|   |                              #    user/chip-ID/notes fields, status label)
|   |-- widgets.py                 # NEW - Shared custom widgets:
|   |                              #   AutoScaleToolbar (smart Home, left-click zoom)
|   `-- tab_recipe_maker.py        # Recipe maker UI - block library + recipe builder
|
|-- core/
|   |-- __init__.py
|   |-- runner.py                  # SerialMeasurementRunner - serial comms, data
|   |                              #   parsing, CSV save; accepts data_folder arg
|   |                              #   to route output into experiment subfolder
|   |-- method_registry.py         # Hash registry, save_script_file, deduplication
|   |-- session.py                 # Shared state - measurement_queue, counter,
|   |                              #   is_running, runner ref, session_manager slot
|   |-- session_manager.py         # NEW - Session/Experiment lifecycle:
|   |                              #   folder creation, metadata JSON, session_log.txt,
|   |                              #   require_session() / require_experiment() guards
|   `-- mscript_parser.py          # VarType, SI prefixes, parse_mscript_data_package
|
|-- tecancavro/
|   |-- __init__.py
|   |-- pump_gui.py                # PumpCtrl class
|   |-- centris_pure.py            # Minimal Cavro Centris driver
|   |-- tecanapi.py                # Tecan/Cavro API wrapper
|   |-- transport.py               # Low-level serial transport
|   `-- models.py                  # Pump models / enums
|
|-- methods/                       # MethodSCRIPT .ms files saved at runtime
|   |-- YYYY-MM-DD/                # Auto-created per day by method_registry.py
|   |   |-- 001_cv.ms
|   |   |-- 002_swv_ch3.ms
|   |   `-- ...
|   |-- archive/
|   |-- library_map.py             # Hashmap + method finder tool
|   `-- library/                   # Curated methods library
|       `-- ...
|
|-- measurement_data/              # CSV output - now organized by session/experiment
|   `-- <session_name>_<timestamp>/          # Created on "Start Session"
|       |-- session_metadata.json            # name, user, chip_id, notes, timestamps
|       |-- session_log.txt                  # timestamped log of every run in session
|       `-- <experiment_name>_<timestamp>/   # Created on "Start Experiment"
|           |-- experiment_metadata.json     # name, notes, timestamps
|           |-- 001_cv_143022.csv
|           |-- 002_swv_143145.csv
|           `-- ...
|
|   # NOTE: runner.py falls back to a flat YYYY-MM-DD/ subfolder if no
|   # active experiment is set (e.g. during direct "Run Now" without a session).
|
|-- tests/                         # Old tests
|
|-- queues/                        # Saved queue .json files (user-facing save/load)
|   |-- my_experiment.json
|   `-- ...
|
`-- recipe_maker/                  # Recipe maker presets and blocks
    |-- default_blocks/
    |   |-- flush.json
    |   |-- add_c6.json
    |   |-- add_aptamer.json
    |   |-- add_ec4.json
    |   `-- add_ec3.json
    `-- queue_reference.json

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
