# Bayesian Optimization Session Run Guide

This is the operator procedure for running Bayesian optimization (BO) on a
Windows station, including a paired buffer/target run and the optional automatic
handoff to Automated Titration. It is written for a station operated locally or
through AnyDesk.

The most important rule is simple: **do not start from a copied configuration
without reviewing the hardware, fluid, analysis, and scoring settings on the
machine that will perform the experiment.** A JSON file can be valid while
still being scientifically wrong for the new experiment.

## 1. Know which files you are working with

The application uses three layers of BO state—scientific configuration,
machine-local settings, and run/session records—represented by the files below.
They are not interchangeable.

| File type | Normal Windows location | Purpose | Transfer to another station? |
|---|---|---|---|
| Scientific BO config | `%USERPROFILE%\Documents\Experiment Automation Data\bo_configs\default_swv_bo.json` | Channels, parameter space, constraints, optimizer behavior, scoring, and analysis settings | Yes, as a starting point; review every experiment-dependent value |
| BO machine paths | `%USERPROFILE%\Documents\Experiment Automation Data\bo_configs\local_paths.json` | Analysis project, output folder, 64-bit Python, glob, and timeout | Recreate on the destination with **Save Paths**; do not blindly copy absolute paths |
| Application machine config | `%LOCALAPPDATA%\ExperimentAutomation\local_config.json` | Data directories, hardware ports, pump calibration, archive locations, and related station settings | Use only as a reference; redetect hardware and update paths |
| Last BO setup | `%USERPROFILE%\Documents\Experiment Automation Data\bo_configs\last_bo_setup_metadata.json` | Last config plus GUI fields such as target count, paired exchange blocks, and equilibration settings | Optional convenience file; all referenced paths still need review |
| BO session record | `<active experiment>\bo_sessions\<bo_session>` | Immutable config snapshot, state, observations, methods, queue records, analysis, surrogate, and acquisition artifacts | Transfer the complete experiment/session tree only when a real resume or audit is required |

Environment variables named `EA_*` override values in the machine-local files.
If the GUI shows an unexpected old username or directory, inspect both the JSON
files above and the environment variables before changing source code.

### Moving a config to the server

1. On the source computer, press `Win+R` and open:

   ```text
   %USERPROFILE%\Documents\Experiment Automation Data\bo_configs
   ```

2. Copy the desired BO config JSON to the same directory on the server. Give a
   new experiment-specific copy a descriptive name rather than overwriting the
   only known-good configuration.
3. In **Bayesian Optimization → Setup**, use **Browse** beside **BO config**,
   choose the copied JSON, and press **Load**.
4. Recreate the server's analysis paths with **Save Paths**.
5. Press **Save** only after completing the scientific review in this guide.

If only a `bo_config_snapshot.json` is available, it can be copied into the
server's `bo_configs` directory, renamed to something descriptive, and loaded
as a config for a **fresh** run. Do not treat its accompanying `bo_state.json`
as a new run.

## 2. Understand the four similarly named actions

- **Load** reads the BO config shown in the **BO config** field into the Setup
  screen.
- **Save** writes the current scientific setup to that BO config and updates
  `last_bo_setup_metadata.json`.
- **Save Paths** writes machine-specific analysis locations and the analysis
  Python to `local_paths.json`. It does not save the scientific search space.
- **Start BO Session** creates a new session record and snapshots the current
  config. **Load BO Session** opens an existing record and its observations.

For the primary paired Auto Loop workflow, pressing **Start Auto Loop** creates
the paired BO queue item, and the queue creates the session it owns. Do not make
an extra empty session with **Start BO Session** first. Use **Start BO Session**
for the assisted manual/Classic workflow or when explicitly instructed during
recovery.

## 3. Prepare the Windows station

Complete this once after installing or moving the application, and repeat any
check affected by a hardware or software change.

1. Update or copy the application repository to the server.
2. Configure `%LOCALAPPDATA%\ExperimentAutomation\local_config.json` from
   `local_config.example.json` if the station does not already have a working
   local configuration.
3. Connect the pump and potentiostat. Use Windows Device Manager to identify
   them; never assume that another station's COM numbers apply.
4. Confirm the configured pump address, baud, syringe capacity, and steps per
   stroke match the installed hardware.
5. Prepare a separate 64-bit analysis Python and install:

   ```text
   numpy
   pandas
   scipy
   pywavelets
   ```

   From PowerShell, verify the exact interpreter that will be entered in the
   GUI:

   ```powershell
   & "<analysis-python>" -c "import struct,numpy,pandas,scipy,pywt; assert struct.calcsize('P')*8 == 64; print('64-bit BO analysis OK')"
   ```

6. In **Bayesian Optimization → Setup**, set:

   - **Analysis output** to a writable local directory. A started BO session
     uses `<active experiment>\bo_analysis` for its active analysis output.
   - **64-bit analysis Python** to the tested executable.
   - **Application project** to the repository root containing
     `analysis_worker\bo_headless.py`.
   - **Analysis glob** to `*.json` unless the analysis pipeline intentionally
     uses a narrower pattern.

7. Press **Save Paths**.
8. Perform one harmless pump initialization/test and one known-safe
   potentiostat measurement before committing unattended hardware time.
9. Confirm the server will remain awake, AnyDesk will remain reachable, and
   Windows Update or other scheduled restarts cannot interrupt the run.

## 4. Do-not-start gate

Do not begin a production BO run until every box below is true.

- [ ] The pump connects and physically responds using the correct station-local
      COM port, address, calibration, syringe, and speed limits.
- [ ] The potentiostat connects on the verified station-local port.
- [ ] The intended buffer, target, stock, and wash solutions are on the correct
      physical ports.
- [ ] Tubing is primed, bubbles have been addressed, waste capacity is
      sufficient for the complete BO plus titration, and reagent volumes include
      a safety margin.
- [ ] A normal application **Session** and **Experiment** are active.
- [ ] **Queue & Execution** is empty before starting Auto Loop.
- [ ] The configured external analysis interpreter is 64-bit and imports all
      required packages.
- [ ] Channel groups match the physical MUX wiring, contain no duplicate
      channels, and represent the intended independent optimization problems.
- [ ] At least one parameter is **Active**. A generated default config may pass
      **Validate** while reporting `Active parameters: (none)`; that is a fixed
      method, not a useful optimization.
- [ ] Locked values, tied parameters, active ranges, steps, scales, and the
      initial point are physically safe.
- [ ] `end_potential > begin_potential`, the scan window is appropriate, and
      `step_potential × frequency` stays inside the configured scan-rate limit.
- [ ] Bandwidth and fixed/autorange limits are appropriate for the expected
      current and will not clip the signal.
- [ ] Crop, peak, and left/right-minimum voltage windows match the expected
      trace. Smoothing and correction options have been tested on representative
      data.
- [ ] The selected BO type and Q equation reward the scientific outcome actually
      wanted. Inspect the displayed equation, not only the weight names.
- [ ] Repeats, warmup count, batch sizes, optimizer direction, number of groups,
      and target iterations have been converted into an estimated measurement,
      fluid, and time budget.
- [ ] For paired BO, both exchange blocks and both equilibration times have been
      checked against the real cell and tubing.
- [ ] If post-BO titration is enabled, the titration preview has been checked and
      **Lock Auto Settings for BO** has been pressed.

**Validate is necessary but not sufficient.** It checks the BO configuration
and whether valid candidates can be generated. It does not prove that COM
ports, Python packages, solution identities, tubing volumes, or scientific
scoring choices are correct.

## 5. Configure the scientific BO experiment

### 5.1 Start the normal Session and Experiment

Use the bottom Session/Experiment bar:

1. Enter a meaningful **Session Name**, operator, and notes, then press
   **Start Session**.
2. Enter an experiment name, chip ID, aptamer type, and notes, then press
   **Start Experiment**.
3. Confirm the status bar shows both the intended session and experiment.

All measurement, analysis, and BO record paths for this run will be rooted in
that active experiment. Do not switch experiments while BO is running.

### 5.2 Load and immediately identify the config

1. Open **Bayesian Optimization → Setup**.
2. Use **Browse**, select the copied config, and press **Load**.
3. Confirm the **BO config** field points to the server's external data tree,
   not another user's profile and not a temporary extracted ZIP directory.
4. Record the config filename in the experiment notes.

### 5.3 Choose objective and channel groups

For the primary workflow select **Paired-response batched BO optimization**.
Paired Q compares target and buffer phases; Classic Q scores one condition.

For each channel group:

- List only channels connected to comparable electrodes that should share one
  suggested method.
- Do not place a channel in more than one group.
- Use separate groups when channels need independent parameter suggestions.
- Verify the group's optimization direction. `maximize_and_minimize` creates
  two independent streams and ultimately two optimized methods for that group.
  Both can be passed to Automated Titration, so choose this only intentionally.

### 5.4 Review method and optimizer settings

Check **Method Settings**:

- **Bandwidth**: `4k` or `8k` must be compatible with the intended method.
- **BA range**: fixed or autorange, with suitable minimum and maximum.
- **Measurements per channel / point**: repeats acquired for every suggested
  method and phase. More repeats improve repeat statistics but multiply runtime.

Check **Optimizer Behavior by Group**:

- Exploit/explore balance.
- Global and local candidate-pool sizes.
- GP warmup iterations.
- Specific versus random start.
- Optimization direction.
- GP falloff fractions for all active parameters.

Warmup parameter sets count toward the Auto Loop target. In paired mode they
run first in warmup-sized buffer and target batches. Remaining parameter sets
use the regular batch size.

### 5.5 Review parameter space and constraints

For every SWV parameter, confirm its mode:

- **Active**: BO may change it within the declared space.
- **Locked**: the fixed `value` is used.
- **Tied**: its value follows another parameter; conditioning potential is
  commonly tied to begin potential.

For active continuous parameters, review min, max, scale, optional step, and
proposal sigma. For discrete parameters, review every permitted value. Then
open **Initial Parameters** and confirm the starting method is both safe and
inside the effective constraints.

Do not widen a range merely because BO can search it. The constraints prevent
some invalid combinations, but they do not encode every electrode, chemistry,
or instrument limitation.

### 5.6 Review analysis and Q scoring

Use representative historical traces when possible.

1. Confirm crop min/max includes the feature of interest and excludes irrelevant
   regions.
2. Confirm peak and left/right-minimum bounds describe the expected peak.
3. Check Savitzky–Golay window and polynomial order against the point count.
4. Review minimum peak height and prominent/minima requirements.
5. Enable double correction or wavelet processing only when that pipeline has
   been verified for this experiment.
6. Select the scoring tab that matches the BO type.
7. Read the displayed Q equation and predict how failed, flat, noisy, clipped,
   and inverted traces will score.

The optimizer follows `Q_run`. If Q is scientifically wrong, BO can operate
perfectly and still optimize the wrong behavior.

## 6. Configure paired fluid exchange

In **Paired BO Fluid Exchange**:

1. Set **Batch size**: the number of parameter sets measured before the next GP
   update during the post-warmup phase.
2. Set **Warmup batch size**, or deliberately choose **Use all warmup iterations
   as one batch**.
3. Set **Buffer → target block** to the reviewed recipe block that replaces or
   exposes the cell to target solution.
4. Set **Target → buffer block** to the reviewed block that restores the buffer
   condition.
5. Set **Target equilibration (s)** and **Buffer equilibration (s)** from the
   actual transport and binding requirements.
6. Open both JSON blocks in Recipe Maker or a text viewer and calculate their
   total aspirations, dispenses, waste volume, and final fluid state.

The paired cycle runs buffer measurements, the buffer-to-target block, target
equilibration, target measurements, the target-to-buffer block, and buffer
equilibration before the next cycle. Never infer fluid direction from a filename
alone.

## 7. Configure automatic post-BO titration

This must be completed **before** BO starts.

1. In BO Setup, enable **Run autotitration when BO finishes**. The application
   opens **Automated Titration** and prepares rows for all BO channels.
2. In **Pump and Port Setup**, verify every physical port and all line, bubble,
   clearing, mixing, speed, and syringe values.
3. In the concentration plan, enter and verify:

   - Stock concentration.
   - Initial buffer volume.
   - Flow-cell aliquot volume.
   - Plain-buffer aliquot behavior.
   - Desired concentrations in the intended order and below the stock
     concentration.
   - SWV replicates per channel.
   - Initial-buffer and between-point options.

4. Review **Manual SWV Settings by Channel**. The automatic BO handoff requires
   at least one manual setting for every BO channel. These manual settings are
   run as comparison methods in addition to the optimized methods; they are not
   placeholders to ignore.
5. Press **Generate Recipe**.
6. Inspect **Calculated Liquid Plan** row by row. Check stock additions, volume
   remaining, bubble-clear losses, aliquots, and cleanup volume.
7. Inspect **Generated Recipe Preview** from first pump action through final
   cleanup. Confirm the number and order of optimized and manual SWVs.
8. Press **Lock Auto Settings for BO**. The button becomes disabled, the status
   reports that the settings are locked, and the application returns to BO
   Setup.

If any titration value changes, unlock/reconfigure by disabling and re-enabling
the BO autotitration option, regenerate the preview, and lock the new settings
before starting BO.

## 8. Validate safely before production

### 8.1 Save and validate

1. Press **Save** in BO Setup.
2. Press **Validate**.
3. Require `Config valid` and a nonempty list after `Active parameters:`.
4. Recheck that **Save Paths** was performed on the server.
5. Capture a screenshot or copy of the Setup page for the experiment record.

### 8.2 Use the correct smoke test

For a paired-response experiment, do **not** use the Classic manual
Suggest → Send → Analyze sequence as proof that pairing works. Instead:

1. Create a separate throwaway Session/Experiment with safe fluid quantities.
2. Disable automatic post-BO titration for this test.
3. Use the intended paired exchange blocks and set a very small paired Auto Loop
   target, normally one complete parameter set if the protocol permits.
4. Verify that both buffer and target measurements are tagged, analysis
   completes, paired Q is recorded, and the cell returns to the expected fluid
   state.
5. End the throwaway experiment. Start a new production experiment so test
   observations never train the production optimizer.

For Classic BO, the assisted manual smoke test is:

1. Press **Start BO Session**.
2. Press **Suggest Next Method**.
3. Press **Preview Script** and inspect the generated MethodSCRIPT.
4. Press **Send Batch to Queue**.
5. Run the items from **Queue & Execution**.
6. After data exists, press **Run Analysis**. Use **Use Latest Analysis** or
   **Import Analysis JSON** only when intentionally importing an existing
   analysis result.
7. Inspect the trace, per-channel metrics, `Q_channel`, and `Q_run` before
   requesting another method.

## 9. Run paired Auto Loop

1. Start or return to the clean production Session and Experiment.
2. Confirm the queue is empty and nothing is currently running.
3. Load the production config, verify paths, and press **Validate** again.
4. Confirm automatic titration is enabled and locked if it should follow BO.
5. In the BO **Run** subtab, enter **Total target iterations**.
6. Calculate the real work represented by that number using the next section.
7. Press **Start Auto Loop** and confirm the prompts.
8. Watch **Queue & Execution**. Paired Auto Loop is represented by a BO queue
   item and starts the queue automatically.
9. Record the created BO record-folder path shown by the application.

Do not add unrelated queue items, switch experiments, edit the live config, or
change fluid connections while Auto Loop is running.

### What “Total target iterations” means

It is the final target number of parameter sets, including warmups—not “this
many more after the observations already recorded.” Batch size changes how
those sets are physically grouped; it does not add observations to the target.

Each channel group has its own optimizer. A group configured for
`maximize_and_minimize` has two optimizer streams. Therefore:

```text
optimizer streams = sum(1 per normal group, 2 per maximize_and_minimize group)
BO observations = target parameter sets × optimizer streams
```

Physical SWV count is larger again because it includes buffer and target phases,
channels in each group, and measurements per channel. Estimate it before the
run:

```text
paired SWVs ≈ target parameter sets
              × optimizer streams
              × channels represented per stream
              × measurements per channel
              × 2 phases
```

When group sizes differ or comparison methods are added later, calculate from
the actual queue preview rather than relying only on this approximation.

## 10. Monitor the run

Use three views together:

- **Queue & Execution**: current hardware action, completed/failed/stopped
  items, and the live BO queue detail.
- **Bayesian Optimization → Results & Records**: observations, Q trend, best
  methods, surrogate, acquisition, and session record files.
- **Plotter**: raw trace shape, clipping, baseline, peak location, noise, and
  drift.

On disk, confirm files continue to appear in:

```text
<active experiment>\bo_analysis\
<active experiment>\bo_sessions\<bo_session>\
```

Investigate before continuing if any of these occur:

- Flat, clipped, empty, badly shifted, or physically implausible traces.
- Peaks outside configured analysis windows.
- Repeated analysis failures or zero scores.
- A Q trend inconsistent with visibly better/worse traces.
- Unexpectedly repeated parameter suggestions.
- Fluid volumes, bubbles, waste level, leaks, or pressure that differ from the
  plan.
- Any failed or stopped queue item.

### Stopping safely

- For Classic BO Auto Loop, **Stop Auto** prevents another BO cycle from being
  submitted after current work reaches a stopping point.
- For the paired queue-owned Auto Loop, use **Queue & Execution → Stop**. This
  sets the shared queue stop flag and asks the current runner to stop. Expect the
  current physical action to take time to reach a safe stopping point.
- Do not power-cycle the pump or potentiostat unless required for immediate
  physical safety.
- After any stop, inspect the fluid state and pending BO state before resuming or
  starting over.

## 11. Verify the BO-to-titration handoff

When paired BO completes cleanly, the application loads the completed paired BO
session and selects the best recorded result for each group/direction. If
automatic titration was enabled and locked, it then:

1. Opens **Automated Titration**.
2. Builds the titration from the locked pump/concentration settings.
3. Adds the optimized methods plus the locked manual comparison settings.
4. Appends the complete titration recipe to the queue.
5. Starts running from the first newly appended titration item.

Confirm the status says the locked titration steps were queued, and verify the
first titration actions in **Queue & Execution**. With
`maximize_and_minimize`, expect separate optimized max and min methods for each
group; verify that this doubled method set is scientifically intended.

If BO completes but titration does not start, do not reconstruct it from memory.
Review the troubleshooting table, then use **Receive from Bayesian Optimization**
in Automated Titration to inspect the completed best methods before manually
generating or queueing anything.

## 12. Finish and archive

1. Wait for all queue items to reach a final status.
2. Review the final traces, BO history, best method per group/direction, and
   titration output.
3. Copy or archive the complete active experiment, including raw measurements,
   `bo_analysis`, and `bo_sessions`.
4. Record any intervention, stop, restart, manual import, fluid anomaly, or
   configuration change in the experiment/session notes.
5. Press **End Experiment**, then **End Session** when no more experiments will
   be run in that session.
6. Verify the local data remains present and any configured archive ZIP is
   readable before considering the transfer complete.

## 13. Recovery and alternate workflows

### Recovering an interrupted session

Preserve the complete experiment directory before attempting recovery.

1. Reopen the original application Session and Experiment.
2. In BO Setup, press **Load BO Session**.
3. Select the specific folder containing `bo_state.json`, normally:

   ```text
   <experiment>\bo_sessions\<bo_session>
   ```

4. Inspect completed observations, the pending suggestion, queue manifests, and
   analysis files before taking action.
5. If a measurement completed but its analysis was not imported, use the
   matching analysis action only after confirming its tags and iteration.

For Classic BO, a loaded session can continue the assisted manual/Auto Loop
flow. Paired Auto Loop is queue-owned in this version: pressing **Start Auto
Loop** creates a new paired queue session rather than transparently continuing
the old queue-owned run. Do not assume it resumes an interrupted paired loop.
After inspecting with **Load BO Session**, either complete a clearly recoverable
pending result under supervision or start a fresh paired run in a new experiment
and preserve the interrupted record for audit.

### Classic BO

Choose **Classic BO optimization** when one measurement condition directly
produces the objective. Configure Classic Q scoring, then use either the manual
sequence in section 8.2 or **Start Auto Loop** from an empty queue. No paired
exchange blocks are used.

### Paired BO inside a recipe

Recipe Maker can insert a BO batch block into a larger liquid-handling protocol.
Configure its BO config path, target iterations, batch/warmup settings, exchange
blocks, equilibration, channels, scoring, and analysis overrides. Preview the
entire recipe and make sure no surrounding recipe block leaves the cell in a
state that contradicts the BO block's assumed starting condition. The queue
creates a new BO session when it executes the block.

## 14. Supplied example archive

The supplied `example bo session config.zip` is useful for understanding a
previous experiment, but its scientific choices are not defaults for a new one.

Observed configuration:

- Six independent single-channel groups: 2, 4, 5, 7, 8, and 9.
- Three measurements per channel and point.
- Active step potential, amplitude, and frequency; other method parameters are
  locked or tied.
- Paired-response scoring.
- `maximize_and_minimize` for each group, producing 12 optimizer streams.
- 8 kHz bandwidth and autorange from 100 nA to 100 µA.
- A paired score dominated by pairwise repeat-scan SNR.
- 600 recorded observations and one pending item in `bo_state.json`.

Before reusing its `bo_config_snapshot.json`, recheck channel wiring, voltage
bounds, waveform safety, scoring, repeat count, bandwidth/range, analysis
windows, optimizer direction, exchange blocks, equilibration, and the complete
fluid protocol. The snapshot alone does not describe all of those physical
choices.

The ZIP is **not a complete resumable BO session**. It lacks the normal methods,
queue, analysis, search-space, surrogate, and acquisition directories, and its
state contains absolute paths from another Windows profile. Do not load its old
`bo_state.json` on the server as if it were a new experiment. A real resume or
complete audit requires the entire original experiment directory.

## 15. Troubleshooting

| Symptom | Likely cause | Safe response |
|---|---|---|
| Config validates but BO never changes the method | Every parameter is locked/tied, or the search space collapses to one candidate | Require a nonempty **Active parameters** list and inspect candidate ranges before starting |
| Analysis worker will not launch | Wrong/stale Python path, non-64-bit interpreter, missing package, or wrong project/script path | Re-run the PowerShell import check, correct the fields, press **Save Paths**, and analyze a known completed measurement |
| Path contains another username | Copied `local_paths.json`, `local_config.json`, metadata, or state has absolute source-machine paths | Recreate paths on the server; do not edit `config.py` to embed another machine's path |
| Auto Loop refuses to start | Queue is running/nonempty, target is invalid, BO setup is missing, or requested autotitration is not locked | Clear only known disposable queue entries, correct setup, validate, and lock titration before retrying |
| Paired BO cannot load an exchange block | Block path is empty, stale, or the custom block was not transferred | Copy the required custom block, browse to its server-local path, inspect it, save, and validate again |
| Paired run needs to be stopped | **Stop Auto** does not control the queue-owned paired loop | Use **Queue & Execution → Stop**, wait for the runner, then inspect fluid and session state |
| Q is zero or contradicts the trace | Peak/minimum windows, correction, thresholds, scoring mode, or weights do not match the signal | Stop automatic progression, inspect raw and analysis JSON, test revised analysis/scoring on recorded data, and document any rescoring |
| BO completed but titration did not start | Autotitration was not enabled/locked, BO ended with failed/stopped items, or no completed best result exists for every group/stream | Inspect BO status and best groups; receive parameters manually only after confirming completeness |
| Application or network disconnects | AnyDesk/network loss may be harmless, but application/power loss can interrupt a physical action | Determine whether the local application and queue are still running before reconnecting or issuing commands |
| Interrupted paired session was loaded, but Start Auto Loop proposes a new run | Paired Auto Loop creates a new queue-owned session | Do not overwrite the old record; inspect pending data and choose supervised recovery or a fresh experiment |

## 16. One-page production checklist

### Station

- [ ] Correct repository/version is running.
- [ ] Pump and potentiostat ports and physical response verified.
- [ ] 64-bit analysis Python passes the package/import test.
- [ ] AnyDesk, power, sleep, disk space, and restart policy checked.
- [ ] Reagents, priming, tubing, waste, and volumes checked.

### Application

- [ ] Correct normal Session and Experiment are active.
- [ ] Correct BO config loaded from the server-local data tree.
- [ ] Server-local analysis settings saved with **Save Paths**.
- [ ] Queue is empty.

### Scientific BO setup

- [ ] Objective and Q equation are correct.
- [ ] Channels/groups and optimization directions are correct.
- [ ] At least one parameter is active.
- [ ] Initial method, ranges, steps, ties, and constraints are safe.
- [ ] Bandwidth, autorange/fixed range, and repeats are correct.
- [ ] Analysis crop, peak/minima windows, smoothing, and correction are correct.
- [ ] Warmup, batch sizes, target, optimizer streams, total SWVs, duration, and
      fluid consumption have been calculated.

### Paired protocol and titration

- [ ] Buffer → target and target → buffer blocks inspected.
- [ ] Both equilibration times confirmed.
- [ ] Short paired test passed in a separate throwaway experiment.
- [ ] **Run autotitration when BO finishes** is enabled if required.
- [ ] Titration ports, concentrations, volume calculations, cleanup, optimized
      methods, and manual comparison methods reviewed.
- [ ] **Lock Auto Settings for BO** completed.

### Launch and monitor

- [ ] BO config saved and **Validate** reports a nonempty active-parameter list.
- [ ] Production experiment is clean; no throwaway observations are included.
- [ ] **Total target iterations** and resulting observation/SWV counts confirmed.
- [ ] **Start Auto Loop** pressed once.
- [ ] Record-folder path captured.
- [ ] Queue, traces, Q trend, analysis output, fluids, and waste monitored.
- [ ] Paired stop, if needed, will use **Queue & Execution → Stop**.
- [ ] BO-to-titration handoff and first titration steps verified.
- [ ] Complete experiment archived before ending the Session.
