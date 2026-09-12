# Paired BO example files

These three JSON files are exact copies extracted from
`docs/example bo session config.zip`. They are example inputs, not an active
session or a validated configuration for another station.

| File | Select in the application |
| --- | --- |
| `paired_bo_config.json` | Bayesian Optimization > Setup > BO config > Browse, then Load |
| `kana_400uL.json` | Paired BO Fluid Exchange > Buffer -> target block > Browse |
| `buffer_400uL.json` | Paired BO Fluid Exchange > Target -> buffer block > Browse |

The target block transfers port 7 to port 1; the buffer block transfers port 6
to port 1. Each performs two 200 microlitre transfers at speed 15. Check these
ports, syringe capacity, and exchange volumes against the actual plumbing.
The BO runner does not automatically rescale these blocks.

The configuration retains six single-channel groups (2, 4, 5, 7, 8, 9), three
measurement repeats, both maximize and minimize directions, and the original
analysis/search settings. Its warmup fields disagree: group `n_initial_points`
values are zero while `paired_warmup_cycles` is eight. Explicitly review the
per-group warmup settings before a new run.

The saved example queue used 50 total iterations, batch size 1, warmup batch
size 1, and 60 seconds equilibration in each fluid. Those operational choices
and the post-BO titration settings must be configured in the GUI; loading this
scientific configuration does not configure the entire workflow.

Browse to files on the station running the experiment. Configure that station's
analysis Python, application project, and output folder using Save Paths.
These files do not supply machine-local paths or credentials.

For the full operator procedure, see [BO session run guide](../../bo-session-run-guide.md).
Check representative buffer and target SWV traces before fixing the analysis
windows and optimization direction. Do not change the scoring definition
mid-run and compare the resulting scores as if they used one objective.
