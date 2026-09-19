# Pause, pump reconnection, and interrupted BO recovery

## Pause and resume

**Queue & Execution > Pause** finishes the current pump command or measurement
before pausing. Wait for **Paused between actions** before intervening. **Resume**
continues the same worker and optimizer. The BO tab also has Pause and Resume
buttons for the running queue. A scheduled waiting step freezes its remaining
timer while manually paused. Stop or closing the app ends this in-memory resume
path; use the recovery controls below instead.

## Pump connection loss

Real pump connection failures automatically retry the configured COM port and
address up to three times, with two seconds between attempts. Each attempt can
also take the driver's acknowledgement timeout. Startup port discovery remains
available. Keep the USB cable in its original socket when reconnecting.

Reconnection queries status; it does not initialize the syringe. Status and speed
commands can be retried. A failed motion command might already have moved liquid,
so it is never automatically replayed. The queue stops, leaves later items pending,
and marks the pump position uncertain. Reconcile the physical state before further
motion; a successful, explicitly requested Initialize clears this uncertainty.
Initialization itself can move liquid: use the pump's appropriate recovery procedure,
not blind initialization of a loaded syringe. Automatic retries are bounded; an
unplugged pump cannot keep the experiment advancing as if delivery succeeded.

## Recover a paired BO loop

1. Stop the old queue and wait for the active action and worker to finish. After
   installing this update, close and relaunch the app using the normal pump-enabled
   launcher. These changes cannot repair the worker already running in memory.
2. Use **Choose Session**, then **Choose Experiment**, to reopen the original
   folders. Do not start a new BO session or change the iteration total to a
   remaining-count estimate.
3. With BO stopped, reconnect the pump and reconcile syringe position. Restore the
   flow cell to buffer using the correct exchange for your plumbing, leaving the
   syringe empty. Complete the configured buffer equilibration wait. Do not assume
   a completed CSV or queue row proves that liquid was delivered.
4. In **Queue & Execution**, select **Recover BO** and choose
   `<experiment>/bo_sessions/<saved BO folder>` containing `bo_state.json`.
5. The app asks whether previously completed iterations also had failed exchanges.
   Choose **No** only if their fluid conditions were valid. Otherwise choose **Yes**
   and enter the first invalid per-optimizer iteration, not the displayed queue row.
   That iteration and all later observations are excluded from learning. Earlier
   observations remain; the original data and excluded history are preserved.
   Cancel if you cannot yet identify where reliable measurements ended.
6. Recovery adds and selects a new BO parent row. Click **From Selected** on that
   row. Child progress rows are records, not executable restart points.
7. If all pending traces exist, the app offers **Yes: repeat** or **No: reuse**.
   After a failed or uncertain exchange, choose **Yes**. Reuse only when both
   buffer and target were actually measured in the correct fluids. Incomplete
   pending batches are remeasured. Confirm the restored buffer/empty syringe state
   only after step 3 is complete.
8. Recovery loads the existing optimizer, pending methods, saved configuration,
   and original schedule. It continues toward the original total. It backs up
   state, queue records, and pending CSVs under `recovery_backups` before changing
   records. Remeasurement detaches suspect input records without deleting CSVs;
   excluded observations and derived analysis are retained under
   `excluded_observations`.

For example, with a total of 50 and 11 valid iterations, recovery repeats pending
iteration 12 and continues to 50. If iteration 11 was already invalid, exclude from
11; iterations 1–10 remain in learning. Previously pending suggestions after the
cutoff may change because the optimizer no longer learns from invalid observations.

The supplied `buffer400uL.json` and `kana400uL.json` blocks each transfer 400 µL in
two 200 µL strokes at speed 15. They assume buffer at valve 6, target at valve 7,
and flow-cell delivery at valve 1. They include no equilibration pause. Check this
mapping against your actual tubing before using either block on another setup.

## General queue progress

New runs save `<experiment>/queue_progress.json` atomically before and after each
top-level action, including the measurement counter. BO state and successful scan
records are saved within the BO session during execution. **Load Progress** loads
the checkpoint into the original experiment and selects its restart row; **From
Selected** starts it. It does not begin automatically.

If a crash occurred during a non-BO action, its completion may be uncertain. The
app asks before replaying it; inspect the physical state first. A checkpoint cannot
reconstruct how much liquid moved during a broken connection. Pre-update runs have
no queue checkpoint, but **Recover BO** can find their paired schedule in an
unambiguous saved queue snapshot. Recovery refuses ambiguous schedules.

Paired recovery currently supports the same experiment directory. It does not
automatically relocate paths after moving data to another computer. Classic BO can
still be loaded through its existing assisted workflow; **Recover BO** is for paired
queue-owned loops. **Start Auto Loop** creates a new run, even after **Load BO
Session**; use Recover BO to continue an interrupted paired run.

Tests use saved-state fixtures and simulated/mocked hardware. They verify recovery
and retry behavior, not actual liquid delivery or validity of experimental traces.
