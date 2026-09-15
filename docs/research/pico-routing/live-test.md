> **Withdrawn from normal execution (2026-09-14):** The queue-level reversed-bipot option was removed because it changes method settings and does not establish transparent physical CE/RE switching. Normal queue, runner and session code were restored to their pre-experiment versions. Instructions below describing an enabled queue option are historical and must not be followed. Live evidence remains bounded CA only. The requested unchanged-method, unchanged-wiring switch is unresolved.

# Live flow-cell routing test

The actual connected Pico was tested on 14 September 2026. **The reversed-bipot configuration was accepted with the cell off, but its first energized measurement produced an overload flag on WE0's secondary current.** The test was stopped, and restoration to channel 0 with both cells off was confirmed. This does not validate physical RE1/CE1 routing or compatibility with CV/SWV.

The test used COM5, identified by the instrument as `tespico1304#Oct 22 2021 14:38:26` (firmware 1.3.04). The latest session log had completed its queue. The chip was left in the flow cell, in buffer as reported by the operator, with both reference/counter alternatives connected. No pump actions, firmware updates, calibration writes, or commits were performed.

After an initial test with no explicit MUX selection, MUX position 10 was explicitly selected, matching the last measurement in the existing session log. The primary/reference route was the existing physical WE0/RE0/CE0. OCP was measured for two seconds, then four 0.5-second current measurements applied the rounded OCP center, center +10 mV, center -10 mV, and center again. Sampling interval was 100 ms. Current selector was `1u`; device metadata reported low-speed range index 01. The test is a small-signal CA diagnostic, not a full CV or SWV scan.

| Stage | Observed result |
|---|---|
| RE0/CE0 OCP | 20 samples; center estimated from last five: -12.656 mV |
| RE0/CE0 step sequence | 20 samples; current +0.109 to +0.328 nA, underload flagged |
| Reversed-bipot configuration, cell off | Completed without firmware error after correcting the test cleanup |
| Reversed-bipot energized test | First packet: main WE1 current 0; secondary WE0 reported -2.459 microamps with overload flag |
| Abort and restoration | Completed; channel 1 mode off, channel 0 restored to high-speed mode with cell off |
| Subsequent RE0/CE0 OCP | 20 samples; center -9.699 mV |
| Subsequent RE0/CE0 sequences | Two completed 20-sample sequences; absolute current approximately 0.109–0.328 nA |

The overloaded number is the value transmitted by the Pico, **not a reliable current measurement**. Its secondary field contains metadata `12`: status-type 1 with value 2, meaning overload. The primary field contains `14`, meaning underload. Underloaded baseline readings and their small response do not establish sensor health or identify which physical RE/CE was controlling the cell. They also cannot distinguish unsupported reversed routing from connection/control-loop issues. [MethodSCRIPT manual, §5.2 metadata status definitions](https://dev.palmsens.com/methodscript/latest/methodscript/methodscript_main.html)

The first attempt exposed a bug in the new test harness's cleanup: `cell_off` is invalid on the selected secondary bipot channel on this firmware. Cleanup was changed to switch both channels to mode 0 before restoring channel 0's configuration. A separate restoration script then completed successfully. Subsequent tests used that correction. The original raw logs remain available; the first failed cleanup is not represented as a routing rejection.

The application control is now in **Queue & Execution**, directly below the queue buttons, labelled **Electrode routing for queue execution**. It is a single queue-level control. The default is WE0 MUX with RE0/CE0; the three unavailable routes are disabled. The diagnostic button was removed from Methods. Existing items, saved recipes, CV/SWV builders and the normal serial runner are unchanged. No routing command is injected into existing scripts. The new control appears after restarting the application; the running process was not restarted automatically.

Verification of the queue UI change: 44 selected tests passed, including default/legacy sessions, rejection of unavailable routes, queue lifecycle and appended items, CV voltage windows, SWV generation, and device selection. The measurement builders and normal runner were compared with the committed version and are identical. The actual Queue tab was instantiated at 800 x 600; the selector and queue tree fit. This verifies the UI/default behavior, not the unproven alternative physical routes.

The standalone diagnostic tool remains available for deliberate investigations. It is separate from normal queue execution.

To open the diagnostic dialog independently from the repository:

```powershell
.venv64/Scripts/python.exe tools/pico_routing_test.py --port COM5
```

The port is an example for this machine, not a saved default. Future dialog runs save results under the locally configured data directory's `routing_tests` folder. Do not run a separate controller concurrently with a queue: the application-integrated dialog checks the shared running state, while the standalone dialog cannot inspect another process's queue.

These live records are local, uncommitted experiment files:

- [Initial attempt](../../../local_routing_results/20260914_124518_341293/result.json)
- [Corrected baseline packets](../../../local_routing_results/20260914_124721_119146/02_baseline_raw.txt)
- [Accepted reversed configuration](../../../local_routing_results/20260914_124721_119146/03_reverse_configuration.ms)
- [Overloaded candidate packet](../../../local_routing_results/20260914_124721_119146/04_reverse_raw.txt)
- [Candidate cleanup confirmation](../../../local_routing_results/20260914_124721_119146/05_restore_raw.txt)
- [Final default-route check and restoration](../../../local_routing_results/postcheck_20260914_124911/result.json)

The first two result summaries predate the explicit `restore_confirmed` field; their raw restoration logs are the evidence. The final check has `restore_confirmed: true`. Seventeen automated tests passed, covering the diagnostic harness plus existing SWV generation and device selection. Tests include the exact overloaded packet captured above, default-route behavior, unsupported combinations, cancellation/cleanup behavior, and preservation of the original error if restoration itself fails.

**Update, 13:13:** a subsequent retry with a `10u` current-range request completed all 20 reversed-bipot samples without overload and showed a voltage-dependent secondary current. This supersedes the overload-only outcome above. The queue UI no longer claims the route failed its live test. See [the retry results and wiring workaround](workaround.md). Physical terminal identification remains unverified.

**Recommendation:** keep RE0/CE0 as the normal measurement route. Reversed bipot is not ready to enable in a production CV/SWV method. The next discriminating experiment is terminal identification with a dummy fixture or a manufacturer-confirmed reverse configuration; widening a current range alone would not prove the route is correct.
