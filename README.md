Cavro Centris Pump — Python Control (Windows)

This bundle contains Python scripts and a small GUI to control a Tecan Cavro Centris pump via the Cavro FUSION COM drivers.

It uses the PumpComm COM server that ships with Cavro FUSION and talks to the pump with Cavro command strings (e.g., ZR, A…R, D…R, I#R).

CONTENTS

pump_gui.py — Windows GUI (connect, initialize, valve, aspirate/dispense)

tools/legacy/pump/centris_pure.py — legacy minimal driver class (pure command style to dev=1)

sample_to_waste_pure.py — example: sample -> waste transfer

pump_ad_pure_fixed.py — minimal A/D script

valve_i_sweep_pure.py — step through valve ports 1..9

robust_sample_to_waste.py — “bullet-proof” scripted sample -> waste

requirements.txt — main app dependency list

requirements-pump-32bit.txt — pinned 32-bit pump environment, including pywin32

README.md — this document

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

Install Python libraries for the 32-bit pump environment:

pip install -r requirements-pump-32bit.txt

Create and activate a 32-bit virtual environment:

C:\Users\Chien Lab>C:\pump32_py38_new\Scripts\activate.bat

(pump32_py38_new) C:\Users\Chien Lab>python "C:\Users\Chien Lab\Documents\GitHub\experiment_automation\gui_script.py


or run C:\Users\Chien Lab>python "C:\Users\Chien Lab\Documents\GitHub\experiment_automation\gui_script.py"


RUNNING THE PROJECT (IMPORTANT ON WINDOWS)

If you see steps being "skipped" when running in different shells, it is almost always because
the current working directory (CWD) or Python environment is different.

Always do both of these before running:
1) cd to the project folder
2) activate the correct venv (venv_gui)

CMD.EXE

cd "C:\Users\Chien Lab\Desktop\experiment_automation"
"C:\Users\Chien Lab\Desktop\experiment_automation\venv_gui\Scripts\activate.bat"
python -m main

POWERSHELL

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
cd "C:\Users\Chien Lab\Desktop\experiment_automation"
& "C:\Users\Chien Lab\Desktop\experiment_automation\venv_gui\Scripts\Activate.ps1"
python -m main

GIT BASH

cd "/c/Users/Chien Lab/Desktop/experiment_automation"
source "/c/Users/Chien Lab/Desktop/experiment_automation/venv_gui/Scripts/activate"
python -m main

SANITY CHECKS (IF BEHAVIOR DIFFERS)

pwd
which python
python -c "import sys,os; print(sys.executable); print(os.getcwd())"

If these differ between shells (or VS Code vs. external Git Bash), your run will differ too.

OPTIONAL BAYESIAN OPTIMIZATION TAB

The app includes an optional Bayesian Optimization tab for closed-loop SWV method
optimization across a mux batch. Normal measurement, queue, recipe, and plotting
workflows do not require this feature.

The working BO integration lives in:

core/bo_session.py
  BO session state, constrained SWV candidate generation, scoring, records, and
  in-repo analysis import.

core/analysis.py and core/bo_analysis.py
  SWV peak analysis and BO summary generation.

gui/tab_bayesian_optimization.py
  Optional GUI tab for configuration, suggestions, queueing, auto-loop control,
  and analysis import.

optimizer/bo_configs/default_swv_bo.json
  User-editable default SWV mux BO search space.

config.py
  BO path defaults. Environment variables are optional.

optimizer/bo_configs/local_paths.example.json
  Template for machine-local BO paths such as the analysis output folder. Copy
  it to optimizer/bo_configs/local_paths.json or use the BO Setup tab's Save Paths button.

BO TAB LAYOUT

The Bayesian Optimization tab has three internal subtabs:

Setup
  Edit the BO config path, machine-local analysis paths, mux channels,
  active/locked/tied parameter space, and initial parameters.

Run
  Suggest the next method, send the mux batch to the existing queue, preview the
  generated script, import analysis manually, or run the automated loop.

Results & Records
  Review per-channel scores, BO history, best method so far, surrogate and
  acquisition artifacts, queue manifests, analysis records, plots, and export
  files.

BO PARAMETER MODES

active
  BO is allowed to optimize this parameter using the configured values.

locked
  BO keeps this parameter fixed at its configured value.

tied
  BO derives this parameter from another parameter. By default,
  conditioning_potential is tied to begin_potential.

Parameter spaces can be discrete or continuous. Discrete parameters use the
configured values list. Continuous parameters use min/max/scale and an optional
step. Leave step blank/null for smoother sampling; set it when the generated
method should snap to instrument-friendly increments. Each continuous parameter
also has proposal_sigma, which controls the width of local Gaussian proposals
around the best observed method.

The acquisition exploration setting is a 0..1 explore/exploit blend. Lower
values favor high predicted Q_run; higher values favor uncertainty and sparse
regions of the parameter space.

Simulation mode can run fake BO iterations without hardware, or replay old
analysis JSON files from a folder. Fake simulation is useful for quick algorithm
tests; replay is useful for checking record/import behavior against old data.

BO CONFIG AND INITIAL DESIGN

The editable BO config is optimizer/bo_configs/default_swv_bo.json. It controls:

channels
  The mux channels used for one BO iteration. By default this is channels 1-10.

initial_parameters
  The one starting method BO runs first. It should be valid and conservative.
  After this first method, any remaining early exploration is generated
  automatically from the valid search space until n_initial_points is reached.

  The Setup subtab has an Initial Parameters editor, so these can be changed
  without hand-editing JSON. Save the BO config after editing to persist the
  changes.

parameters
  The active/locked/tied mode, units, and allowed values for each SWV parameter.

constraints
  Hard rules that must be true before a method can be suggested or sent to the
  queue. The defaults are:

  end_potential > begin_potential
  end_potential - begin_potential >= 0.4 V
  step_potential * frequency <= 1.0 V/s
  conditioning_potential = begin_potential unless it is unlocked

scoring
  Weights for Q_channel and Q_run.

analysis
  The in-repo SWV analysis settings and retention behavior for analysis JSON outputs.

BO PATH SETTINGS WITHOUT ENVIRONMENT FILES

Environment variables still work, but they are not required. The app can read
an ignored optimizer/bo_configs/local_paths.json for machine-specific path settings:

{
  "analysis_output_dir": "analysis_outputs",
  "analysis_file_glob": "*.json",
  "analysis_mode": "external",
  "analysis_project": ".",
  "analysis_script": "analysis_worker/bo_headless.py",
  "analysis_python": "C:\\Path\\To\\64-bit-Python\\python.exe",
  "analysis_timeout_seconds": 900
}

The Setup subtab can create or update this local file with the Save Paths
button. The tracked optimizer/bo_configs/local_paths.example.json file documents the
expected shape without forcing one machine's paths onto everyone else.

Real BO analysis runs as a blocking external worker. The 32-bit controller
writes a request, launches `bo_headless.py` with the configured 64-bit Python,
and waits for the returned summary and full trace-results JSON before the next
BO measurement begins. Set `analysis_mode` to `local` only for development.

The headless worker is included in this repository under `analysis_worker/`;
the separate Electrochemistry-Analysis-Scripts checkout is no longer required.
The process boundary remains unchanged: run the application with its 32-bit
environment and install the worker dependencies into a separate 64-bit
environment:

    C:\Path\To\64-bit-Python\python.exe -m pip install -r requirements-analysis-64bit.txt

Set `analysis_python` to that 64-bit interpreter. A `.venv64` environment in
the repository root is also detected automatically.

BAYESIAN OPTIMIZATION INTUITION

Bayesian optimization is useful here because each mux batch is expensive. The
goal is to learn from a small number of real experiments instead of sweeping
every combination.

The optimizer keeps a surrogate model:

  method parameters -> predicted Q_run

When scikit-learn is available, the surrogate is a Gaussian process (GP). The GP
predicts both a mean score and an uncertainty for each valid candidate:

  mu(x)    = predicted score for method x
  sigma(x) = uncertainty in that prediction

The acquisition function chooses what to try next. It balances exploitation and
exploration:

  exploitation: try methods with high predicted Q_run
  exploration: try methods where uncertainty is high

The current model-guided acquisition is Expected Improvement (EI):

  EI(x) = expected amount by which candidate x improves over the best observed
          Q_run so far, accounting for both mu(x) and sigma(x)

In plain language: BO does not just choose what looks best right now. It chooses
what has the best chance of teaching us something useful or improving the best
method.

Frequency is encoded on a log10 scale because frequency values grow
multiplicatively rather than additively. A change from 50 Hz to 100 Hz is more
like the change from 400 Hz to 800 Hz than it is like a 50 Hz linear step near
800 Hz.

If scikit-learn is not installed or too few observations exist, the system falls
back to deterministic space-covering suggestions so the GUI still works. For a
publication BO run, install scikit-learn and record surrogate/acquisition
outputs.

QUALITY SCORE

The in-repo analysis runner supplies per-channel metrics. The BO module then
computes:

  Q_channel =
      w_snr        * normalized_SNR
    + w_shape      * peak_shape_score
    + w_baseline   * baseline_stability_score
    + w_replicates * replicate_consistency_score
    + w_success    * success_score

Then one mux batch becomes one BO objective:

  Q_run =
      mean(Q_channel)
    - lambda_variability * std(Q_channel)
    - lambda_failed      * failed_channel_fraction
    - lambda_low         * fraction(Q_channel < threshold)

The optimizer maximizes Q_run. Every Q_channel is still retained so we can show
that optimization improves the mux array, not just one good channel.

BO NEXT STEPS

The current BO objective is still a measurement-quality objective. It rewards
large, clean, stable, internally consistent SWV signals, but it does not yet
optimize directly for target responsiveness.

Important future direction:

  optimize for response to target, not just absolute signal

In practice this likely means one BO iteration should eventually support a
paired workflow such as:

  1. measure baseline / no-target condition
  2. run one or more pump / incubation / target-addition steps
  3. measure post-target condition
  4. score the candidate from the paired response

Candidate future BO response metrics include:

  delta_peak = peak_with_target - peak_without_target
  fractional_change = delta_peak / baseline_peak
  response_SNR
  response_consistency across channels or replicates

This matters especially for KDM-style sensing, where one frequency may maximize
absolute signal but not maximize target-dependent response.

Recipe-level future work:

  add a BO batch block that can live inside a larger recipe

That would allow BO to be embedded in a higher-level automation recipe with
fluidics and conditioning steps around it, for example:

  pre-measurement preparation
  BO baseline measurement batch
  pump target or reagent
  wait / equilibration
  BO post-target measurement batch
  wash / reset / repeat

If this is implemented, the BO / paired-response block should be visually
distinct in the Recipe Maker and queue views so it is easy to distinguish from
ordinary method, pump, pause, and alert items.

BO RUNTIME FLOW

Bayesian Optimization tab
  -> load editable BO config
  -> optionally edit initial_parameters in Setup
  -> start BO session inside the active experiment folder
  -> ask core BO session for the next valid SWV method
  -> save the method through MethodRegistry
  -> add ordinary SWV queue items with BO metadata
  -> instrument queue runs the mux batch
  -> in-repo analysis scores the just-completed BO CSV files
  -> BO tab imports the generated analysis JSON
  -> core computes Q_channel and Q_run
  -> records are retained under experiment/bo_sessions/
  -> next method is suggested

The tab supports both assisted-manual operation and an Auto Loop mode. Auto Loop
starts only from an empty queue, submits one BO mux batch at a time, starts the
existing queue, waits for queue completion, runs in-repo analysis immediately,
records the scores, and repeats until the requested number of completed BO
iterations is reached.

ANALYSIS JSON CONTRACT

The normal path writes summaries under the active experiment folder:

  <active_experiment>/bo_analysis/

The BO session also retains analysis records under:

  <active_experiment>/bo_sessions/<session>/analysis/

Manual imports and simulation replay still accept this JSON shape:

Accepted JSON shape:

{
  "channel_metrics": {
    "1": {
      "snr": 12.4,
      "peak_shape_score": 0.82,
      "baseline_stability_score": 0.76,
      "replicate_consistency_score": 0.69,
      "success_score": 1.0
    }
  }
}

The channel-keyed object may also be supplied directly:

{
  "1": {"snr": 12.4, "success_score": 1.0},
  "2": {"snr": 9.1, "success_score": 1.0}
}

The BO module computes one Q_channel per channel and one Q_run per iteration.
The optimizer only sees Q_run, while records retain every channel metric and
score for reproducibility and publication.

BO RECORDS

Each BO session writes records inside the active experiment folder:

bo_sessions/<bo_session_id>/
  bo_config_snapshot.json
  search_space.json
  constraints.json
  initial_parameters_preview.json
  bo_state.json
  history.csv
  methods/
  queue/
  analysis/
  surrogate/
  acquisition/
  plots/

Current records include suggested methods, queued BO items, imported external
analysis outputs, per-channel metrics, Q_channel, Q_run, the best method so far,
queue completion manifests with measurement tags and CSV paths when available,
channel-score plots, BO history plots, candidate prediction tables, acquisition
value tables, top-candidate tables, and surrogate/acquisition projection plots.

When scikit-learn is available and at least two completed BO observations exist,
the BO session also saves the fitted Gaussian-process model:

surrogate/iter_XXX_gp_model.pkl

If scikit-learn is unavailable, the app still writes deterministic fallback
prediction and acquisition tables so the record remains complete, but those
tables should be labeled as fallback rather than GP-based in a publication.
