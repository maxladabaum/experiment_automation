> **Withdrawn from normal execution (2026-09-14):** The queue-level reversed-bipot option was removed because it changes method settings and does not establish transparent physical CE/RE switching. Normal queue, runner and session code were restored to their pre-experiment versions. Instructions below describing an enabled queue option are historical and must not be followed. Live evidence remains bounded CA only. The requested unchanged-method, unchanged-wiring switch is unresolved.

# Routing workaround investigation, 2026-09-14

## Enabled for operator testing

Restart the application after the current run has finished. In **Queue & Execution**,
below the queue buttons, open **Electrode routing for queue execution** and choose
**WE0 MUX | RE1 + CE1 candidate (experimental)**. Load a supported measurement
or BO block into the queue and start it normally. The menu choice
applies to all measurements in that queue run; it cannot be changed while running.
Return to **WE0 MUX | RE0 + CE0 (default)** for normal operation.

This enables the reversed-bipot candidate with the existing wiring, not the
external-MUX rewiring proposal below. It remains an experiment: the physical
CE/RE path has not been independently identified. No full CV/LSV has been live
validated by this code change; the earlier live result was small-signal CA.

Experimental execution preserves requested voltage steps, scan count and speed,
but uses a fixed `10u` current-range request on both channels. Bandwidth is 40 Hz
for CV/LSV/CA and 100 Hz for scripts containing SWV/DPV/NPV/PAD. Original high-speed
or max-range configurations are converted to low-speed bipot when parameters fit.
The original current variable is populated from `poly_we(0 ...)`, so plotted and
saved current is WE0 secondary current, not unused WE1 current. Original files
are unchanged; `methods_used` records the converted script and raw packets are
always saved. Overload/timing/nonfinite data or current above 50 microamps stops
the run. This is sample-based protection, not a hardware clamp. Both channels
are disabled and the default channel-0 mode restored, including after stop/error;
failure to receive restoration confirmation is logged and fails the run.

Supported techniques are CV, LSV, CA, SWV, DPV, NPV and PAD (including generated
SWV conditioning). BO uses the same per-measurement converter. Methods require
literal parameters, at most 100 points/s and 1 V/s for sweeps, and a voltage
window within +/-1.6 V with span <=1.6 V. SWV requires integer frequency 1..100 Hz;
pulse methods also validate pulse/interval timing. Unsupported parameters are
rejected, never silently slowed or clipped. The window limits
are conservative routing checks, not a recommendation for sensor exposure. OCP
and EIS are excluded by Pico bipot support; they require the default route.
Unsupported custom control flow/cleanup is also rejected. Static queue methods
are validated before queue start; BO-generated methods are checked before serial
execution, so a BO search containing out-of-limit candidates can stop with an
explicit error. BO parameter ranges are not rewritten.
Items appended during execution are also checked before hardware execution;
experimental queues stop on failure instead of proceeding to later items.

The default remains `we0_re0_ce0` for new and legacy sessions. The choice is not
persisted across application restarts. Default scripts pass through unchanged,
with no added restoration commands or altered parsing. Regression coverage:
77 tests passed, including every current library script passing through the
default converter unchanged, queue lifecycle, SWV/CV generation, and device
selection, generated SWV conditioning, secondary-current sign/columns, BO route
propagation, and a simulated SWV save. An actual 800-pixel Tk layout verified the enabled menu and return
to default. No firmware changes, commits or pushes were made.

SWV's legacy `poly_we` result is secondary forward-minus-reverse current. The
converter removes main-channel forward/reverse packet fields, and the runner
saves the WE0 result as `current` and `current_diff` without a second sign flip.
Separate WE0 forward/reverse components are unavailable from this command and
are not fabricated. Technique markers distinguish SWV packets from preceding CA
conditioning packets. This implementation follows MethodSCRIPT sections 9.1 and
14.13.17. Pico's brochure page 4 excludes OCP/EIS from bipot mode:
https://dev.palmsens.com/methodscript/latest/methodscript/methodscript_main.html
https://assets.palmsens.com/app/uploads/2026/02/EmStat-Pico-Brochure.pdf

The added technique handlers have automated/offline verification only. The live
hardware evidence remains the bounded CA experiment below. In particular, valid
SWV packets alone would not establish correct pulse settling or physical CE/RE
identity; compare against a known dummy cell before treating results as validated.

## Live current-range retry: promising, physical terminals still unverified

The earlier reversed-bipot overload does not establish that routing is impossible.
On firmware 1.3.04, COM5, MUX position 10, the same bounded test completed after
changing only the current range requests from `set_range ba 1u` to `set_range ba 10u`.
The instrument reported range metadata `205` on both current fields. The test
retained overload/timing aborts and the 50 microamp software threshold. This is
sample-based protection, not a hardware current clamp.

OCP center was -0.594 mV, rounded to -1 mV for the steps. Each step lasted 0.5 s
with five samples. Mean secondary current, requested via `poly_we(0 b)`:

| Applied setpoint | Secondary current |
|---|---:|
| -1 mV | -1.161294 microamps |
| +9 mV | +7.381467 microamps |
| -11 mV | -8.844810 microamps |
| -1 mV | -1.159194 microamps |

All 20 candidate samples had status `10` (no warning flags). The primary field
was approximately zero and underloaded. Default-route measurements before and
after were also underloaded and zero at this range. Both cells were switched
off, then channel 0's usual high-speed configuration restored successfully.

Raw scripts, packets and result are in the ignored local directory
`local_routing_results/range10u_20260914_131351_491614/`.
This establishes a reproducible voltage-dependent secondary signal in one run;
it does not independently establish physical RE1 feedback or CE1 drive. The
near-zero default response also needs explanation. No full CV or SWV was tested.
Reversed bipot remains a candidate, not a production queue route.

The discriminating next test uses a known dummy load, measuring WE0-to-RE1
potential independently and identifying CE1 drive with a suitable meter/scope.
Repeat with deliberately distinct reference/counter fixtures, reconfiguring only
with cells off. Check WE0-to-RE0 as a control. Then compare slow CV against a
normal channel-1 measurement on the same fixture. Do not infer physical routing
solely from MethodSCRIPT packet labels. Low-speed bipot technique/bandwidth
limits still apply even if this succeeds.

## Existing MUX alternative, requiring a wiring change

The manufacturer's MUX16 guide, section 2 page 1, explicitly supports separate
WE and RE/CE-bank addressing. Section 6 page 9 identifies the electrode headers.
Source: https://assets.palmsens.com/app/uploads/2021/06/Getting-Started-with-the-Emstat-Pico-MUX16.pdf

Engineering proposal: leave all working electrodes on their existing WE MUX.
Move the two reference/counter alternatives to the external MUX's RE/CE inputs.
Call the physical electrodes A and B to avoid confusing external positions with
the Pico's physical terminals. Use four external pair positions if mixed choices
are needed:

| External pair position | Reference electrode | Counter electrode |
|---|---|---|
| 1 | A | A |
| 2 | B | B |
| 3 | B | A |
| 4 | A | B |

For two complete pairs alone, only positions 1 and 2 are needed. Four profiles
require passive fanout of each electrode lead to the corresponding two MUX
inputs. The shared RE/CE address then selects one prewired combination. This
retains all 16 WE positions and uses Pico RE0/CE0 for every profile; it does NOT
internally cross-route physical RE1/CE1 to WE0.

Check the actual board revision, header pin order, harness common connections,
and any CE/RE shorting jumper before wiring. Disconnect the alternative electrode
leads from physical Pico RE1/CE1 before moving them to the MUX. Never bridge Pico
counter outputs together. Board schematic source:
https://assets.palmsens.com/app/uploads/2021/06/EmStat-Pico-MUX16-Schematics_V3.pdf

Address formula, for one-based external positions:
`((pair_position - 1) << 4) | (we_position - 1)`.
Offline checked all 256 WE/pair addresses: unique, with correct nibbles and both
MUX enable bits clear. For WE position 10, profiles 1..4 are 0x09, 0x19, 0x29,
0x39. These independent addresses have not been applied to the current harness.

The app currently couples the two positions in its normal MUX wrappers. A future
opt-in queue-level pair selector can replace the upper nibble while retaining
the WE selection. Preserve existing addressing as the default; apply the option
consistently to normal and BO execution. Switch with the cell off between scans,
allow settling, and record the selected pair in results. Validate known dummy
loads through every profile before using the sensor. This option has not been
enabled because the current wiring has not been converted or identified.

## Firmware lead

PalmSens' firmware changelog describes improved current-range selection and
simplified bipot commands in 1.5, plus later improvements. It does not establish
arbitrary physical CE/RE cross-routing. A firmware upgrade is not needed for the
successful bounded retry or for the documented direct-GPIO MUX addressing above.
No firmware or calibration was changed.
Source: https://www.palmsens.com/knowledgebase-article/emstat-pico-firmware-1-6-what-has-changed/
