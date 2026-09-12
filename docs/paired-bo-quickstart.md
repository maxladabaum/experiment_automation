# Beginner workflow: paired BO followed by titration

This reproduces the supplied example. Confirm its chemistry, voltage limits,
ports, and volumes match the actual station before using them. Start with a
separate test experiment; production BO should use one consistent analysis and
scoring definition throughout.

## Files and where to select them

- BO Setup > BO config > Browse: `docs/examples/paired_bo/paired_bo_config.json`,
  then **Load**. This supplies scientific settings, not all operational settings.
- BO Setup > Paired BO Fluid Exchange > Buffer -> target block > Browse:
  `recipe_maker/default_blocks/kana_400uL.json`.
- Target -> buffer block > Browse:
  `recipe_maker/default_blocks/buffer_400uL.json`.
- The two exchanges also appear in Recipe Maker > Blocks > Default (or All).
  Press **Refresh Blocks** if the app was already open when they were added.
- Recipe Maker > **Load** starts in the repository's `recipe_maker` folder.
  **Save** continues to use the configured local recipe directory.

The exchange files are pump recipes, not electrochemical methods or a previous
experiment to resume. Their names are arbitrary. BO executes their saved ports,
speeds, and volumes; it does not resize or generate these exchanges. An empty BO
exchange field currently skips that exchange. A missing selected file fails.
Browse to existing files on the station that will run the experiment.

## Run in this order

1. Connect the pump and potentiostat. Confirm simulation is OFF, the syringe is
   calibrated, and the installed syringe supports the planned strokes.
2. Verify the example port map: 1 flow cell, 2 waste, 4 mixing tube, 5 titration
   stock, 6 buffer, 7 BO target, 9 air. Check reagent and waste capacity.
3. Start a test Session and Experiment using the bottom bar.
4. In Recipe Maker, load `recipe_maker/custom_blocks/prime_lines.json` (or
   `primn.json`). These corrected local recipes contain 24 steps: two 250 uL
   transfers each from ports 7, 6, and 5 to waste port 2 at speed 15. Start with
   an empty syringe, Send to Queue, inspect, and run. Both end empty. Reload an
   already queued old copy; editing a file does not update queued steps. These
   custom files must also be present on any destination station.
5. Inspect the supply lines for air. The recipe's 500 uL per line is not a
   guarantee for every tubing length. It primes supply lines, not the cell.
6. Load/run the buffer exchange block to fill the flow-cell line: two 200 uL
   transfers from port 6 to port 1 at speed 15. Confirm the cell is filled and
   bubble-free and allow equilibration. Finish preparation with an empty,
   stopped queue.
7. Open BO Setup and load the configuration listed above.
8. Configure this station's analysis project (repository root), working 64-bit
   analysis Python, writable analysis output, and `*.json` glob. **Save Paths**.
   The interpreter needs numpy, pandas, scipy, and pywavelets. See the full run
   guide for the interpreter check.
9. Select **Paired-response batched BO optimization**. The example has channels
   2, 4, 5, 7, 8, 9 in separate groups; 3 repeats per channel/point; 8k bandwidth;
   autorange 100 nA to 100 uA.
10. Check the search: begin/end locked at -0.6/0 V; active step 1-10 mV in 1 mV
    increments, amplitude 10-200 mV in 10 mV increments, frequency 10-500 Hz in
    1 Hz increments; conditioning potential tied to begin; time locked at 0.2 s.
11. Collect repeated buffer and target traces using the same fixed SWV method
    on every intended channel. The example comparison is -0.6 to 0 V, step
    2 mV, amplitude 36 mV, frequency 200 Hz, conditioning -0.6 V for 0.2 s, 8k.
    Check the current range, peak identity, clipping, and repeatability.
12. Review the analysis windows using those traces. Example crop: -0.55 to 0 V;
    peak: -0.45 to -0.1 V; left minimum: -0.54 to -0.1 V; right minimum:
    -0.45 to -0.01 V. Keep the intended peak and baseline on both sides; inspect
    corrected traces too. Double correction is enabled and minima on both sides
    are required. Adjust before production, not midway through a scored run.
13. Choose the search direction. The example's pairwise-SNR score uses target
    peak minus buffer peak divided by a variability measure. Maximize seeks
    positive response; minimize seeks negative response. Both runs independent
    searches for both directions. The example uses both for all six groups:
    12 streams. Paired means buffer/target measurements, not max/min searches.
14. Resolve warmup explicitly: the archive says 8 warmup cycles in one field,
    but the saved queue/group values give 0 effective warmup iterations. Use
    per-group warmup 0 and random start to reproduce that recorded schedule;
    use 8 per group only if eight warmup points are intended. Warmups count
    within the total iteration target.
15. Select the two exchange blocks above, batch size 1, warmup batch size 1,
    all-warmups-in-one-batch OFF, target and buffer equilibration 60 s each.
    Confirm 400 uL actually exchanges the cell solution on this setup.
16. Leave automatic titration OFF for the test. Restore buffer after preliminary
    measurements, empty the queue, **Save**, **Validate**, and set Run > Total
    target iterations to 1. **Start Auto Loop**. The queue creates the BO session;
    do not create an extra session with Start BO Session first.
17. Verify buffer measurement -> target exchange -> target equilibration ->
    target measurement -> buffer exchange -> buffer equilibration, and correct
    paired analysis. Confirm detected peaks and Q agree with the traces.
18. Start a fresh production Session/Experiment. Preserve test records, restore
    the starting buffer condition, and finish preparation with an empty queue.
19. Enable **Run autotitration when BO finishes**. This opens Automated Titration.
    Configure all settings in the tables below; the BO config does not set them.
20. Enter a manual comparison SWV setting for every BO channel using the values
    from step 11 to reproduce the example. Check current range too. Optimized
    settings come from the new BO automatically; comparisons are additional.
21. **Generate Recipe**. Review Calculated Liquid Plan and Generated Recipe
    Preview through final cleanup. **Lock Auto Settings for BO**. Do not send the
    preview separately to the queue. Reconfigure/regenerate/relock after changes.
22. Return to BO, **Save**, **Validate**, confirm active parameters and locked
    titration settings. Set Total target iterations to 50 and **Start Auto Loop**.
23. Observe the first full cycle. With 50 batches, the example exchanges use
    20 mL target and 20 mL buffer, excluding preparation, tests, and titration.
24. After clean BO completion, verify the app appends and starts locked titration
    with optimized methods and manual comparisons. Both search directions supply
    separate optimized methods when configured.
25. After titration/cleanup finish, inspect results, end Experiment and Session,
    and verify saved data/archive. To stop paired execution use Queue & Execution
    > Stop. Inspect the fluid state before restarting; do not assume a new Auto
    Loop transparently resumes an interrupted paired run.

## Example titration settings

| Pump setting | Value |
|---|---|
| Ports | Same mapping as step 2 |
| Syringe capacity | 250 uL |
| General speed | 15 |
| Initial-buffer / final-cleanup speed | 12 |
| Mixing-line volume | 110 uL |
| Bubble volume / liquid loss per clear | 50 / 50 uL |
| Stock air spacer / line air push | 100 / 250 uL |
| Mixing | 1 cycle of 250 uL |
| Equilibration | 180 s |

| Concentration plan | Value |
|---|---|
| Stock | 10,000 uM (10 mM) |
| Starting buffer | 6,000 uL |
| Cell aliquot / plain-buffer aliquot | 500 / 1,000 uL |
| Concentrations (uM) | 20, 40, 80, 160, 320, 640, 1280, 2560, 5120 |
| Replicates | 10 |
| Skip initial buffer preparation | OFF |
| Measure buffer between concentrations | ON |
| Manual-only | OFF |

Confirm the plumbing-dependent values rather than assuming defaults. The recipe
adds starting buffer to the mixing tube; do not independently preload another
6 mL when using this preparation option.

## Testing all channels

The example selection is not a BO limit. To test channels 1-10, first acquire the
same fixed SWV on all connected working electrodes for direct comparison. Then
update both BO Channels and Channel Groups: one group for a shared optimized
method, or one group per channel for independent searches. Cover every selected
channel in the titration comparisons, regenerate, and lock again. More channels,
repeats, and directions increase run time.

## Queue ETA

Queue & Execution > Queue ETA estimates pending work or the active run. Pump
estimates learn successful queued action durations per action and speed during
the app session. Exact-volume observations take priority; other volumes use a
fit for transfer rate/overhead when available. Before sampling, volume-scaled
defaults are provisional, not a calibrated speed table. Timings are discarded
when the pump connection/calibration/simulation context changes and on app exit.

BO planning counts groups, directions, repeats, representative initial SWV
duration, exchanges, equilibration, and an approximate analysis allowance.
Paired live ETA uses complete batch timings (including analysis/optimizer work)
to scale the remaining plan. Future method choices and analysis cost can change;
the estimate is not a deadline. The display tracks the whole BO loop rather
than only its current nested pump/measurement step. Autotitration is excluded
until its generated steps are appended to the queue. User-controlled alert
waits and unknown items can prevent a complete ETA.

See also the [full BO session run guide](bo-session-run-guide.md).
