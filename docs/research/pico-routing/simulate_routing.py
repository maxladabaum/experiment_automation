"""Offline dummy-cell thought experiment; no instrument/app imports or I/O.

Run with the repository Python. Outputs stay beside this file. This is not an
EmStat firmware emulator, transistor model, or evidence of accepted commands.
"""
from pathlib import Path
import csv
import json
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

OUT = Path(__file__).resolve().parent


def save_csv(name, headers, rows):
    with (OUT / name).open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)


def main():
    # Analyst-selected ideal parallel RC cell; reference offset is synthetic.
    resistance, capacitance = 10_000.0, 1e-6
    t = np.linspace(0, 4, 2001)
    e = np.where(t <= 2, -0.1 + 0.1*t, 0.3 - 0.1*t)
    slope = np.where(t < 2, 0.1, -0.1)
    rows, currents, drives = [], {}, {}
    combos = [(0, 0), (1, 1), (1, 0), (0, 1)]
    for re, ce in combos:
        name = f"WE0_RE{re}_CE{ce}"
        effective = e + re*0.05
        current = effective/resistance + capacitance*slope
        # WE is mathematical zero here, NOT Pico AGND. Signed CE drive is
        # determined by the solution voltage and counter branch voltage drop.
        drive = -effective - current*(1000 if ce == 0 else 2000)
        currents[name], drives[name] = current, drive
        rows.extend(zip([name]*len(t), t, e, effective, current, drive))
    save_csv("cv.csv", ["hypothetical_route", "time_s", "command_V",
                           "effective_cell_V", "WE0_current_A", "CE_relative_WE_V"], rows)

    # Conditional reversed bipot model: main WE1 has a different dummy resistor.
    # WE0 at zero offset produces target WE0 current IF the firmware supports it.
    secondary = e/resistance + capacitance*slope
    primary = e/20_000.0
    save_csv("bipot_streams.csv", ["time_s", "command_V", "WE1_main_A", "WE0_secondary_A"],
             zip(t, e, primary, secondary))

    # An isolated 50 mV reference-change transient at a held command of 0 V.
    # The 1 ms transition time is assumed, not a Pico specification.
    tr = np.linspace(0, .02, 2001)
    tau, delta = .001, .05
    v = delta*(1-np.exp(-tr/tau))
    ic = capacitance*delta/tau*np.exp(-tr/tau)
    itotal = v/resistance + ic
    save_csv("reference_switch.csv", ["time_after_switch_s", "effective_cell_V",
                                     "capacitive_A", "total_A"], zip(tr, v, ic, itotal))

    # Exact periodic response of an illustrative one-pole measurement filter
    # to a +/-25 mV square wave on a purely resistive cell. End-half-cycle
    # sampling; no staircase or electrochemical kinetics in this submodel.
    freq = np.array([10., 25., 50., 100., 200., 500.])
    amp = .025
    swv_rows = []
    for bandwidth in [100., 4000.]:
        filter_tau = 1/(2*np.pi*bandwidth)
        for f in freq:
            ratio = np.tanh((1/(2*f))/(2*filter_tau))
            swv_rows.append((f, bandwidth, 2*amp/resistance, ratio,
                             2*amp/resistance*ratio))
    save_csv("swv_filter.csv", ["frequency_Hz", "assumed_one_pole_bandwidth_Hz",
                               "unfiltered_difference_A", "retained_fraction",
                               "filtered_difference_A"], swv_rows)

    # Analytic consistency checks, not hardware validation.
    np.testing.assert_allclose(currents["WE0_RE0_CE0"], currents["WE0_RE0_CE1"])
    np.testing.assert_allclose(currents["WE0_RE1_CE0"]-currents["WE0_RE0_CE0"], .05/resistance)
    np.testing.assert_allclose(secondary, currents["WE0_RE0_CE0"])
    integrate = np.trapezoid if hasattr(np, "trapezoid") else np.trapz
    charge = float(integrate(ic, tr))
    assert abs(charge/(capacitance*delta)-1) < 1e-4
    assert abs(primary[0]) < abs(secondary[0])
    assert all(0 <= row[3] <= 1 for row in swv_rows)
    assert all(np.max(np.abs(d)) < .2 for d in drives.values())

    fig, axes = plt.subplots(2, 2, figsize=(11, 8), constrained_layout=True)
    for re, color in [(0, "#222222"), (1, "#197e93")]:
        axes[0, 0].plot(e, currents[f"WE0_RE{re}_CE0"]*1e6, color=color,
                        label=f"RE{re}; CE0 and CE1 overlap")
    axes[0, 0].set(xlabel="Command vs selected RE (V)", ylabel="WE0 current (uA)",
                   title="Synthetic CV: reference offset changes current")
    axes[0, 0].legend(fontsize=8)
    for ce, style in [(0, "-"), (1, "--")]:
        axes[0, 1].plot(t, drives[f"WE0_RE0_CE{ce}"]*1000, style,
                        label=f"CE{ce}: {ce+1} kohm branch")
    axes[0, 1].set(xlabel="Time (s)", ylabel="CE relative to WE (mV)",
                   title="Different CE drive, identical ideal WE current")
    axes[0, 1].legend(fontsize=8)
    axes[1, 0].plot(tr*1000, itotal*1e6, color="#222222")
    axes[1, 0].axhline(5, linestyle="--", color="#197e93", label="New steady current")
    axes[1, 0].set(xlabel="Time after reference change (ms)", ylabel="Current (uA)",
                   title="Assumed 1 ms settling: 50 mV reference change")
    axes[1, 0].legend(fontsize=8)
    for b, style in [(100., "o-"), (4000., "s--")]:
        data = [r for r in swv_rows if r[1] == b]
        axes[1, 1].semilogx([r[0] for r in data], [r[3]*100 for r in data], style,
                           label=f"Assumed {b:g} Hz filter")
    axes[1, 1].set(xlabel="SWV frequency (Hz)", ylabel="Difference retained (%)",
                   title="Illustrative filter, not measured Pico response", ylim=(0, 105))
    axes[1, 1].legend(fontsize=8)
    for ax in axes.flat:
        ax.grid(alpha=.18)
    fig.suptitle("Hypothetical routing and dummy-cell simulation — no hardware tested", fontsize=13)
    fig.savefig(OUT / "simulation.png", dpi=160)
    plt.close(fig)
    summary = {
        "hardware_tested": False,
        "firmware_emulated": False,
        "assumptions": {"Rct_ohm": resistance, "C_F": capacitance,
                        "reference_offset_V": delta, "settling_tau_s": tau,
                        "CE_branch_ohm": [1000, 2000]},
        "reference_change_steady_current_uA": delta/resistance*1e6,
        "reference_change_capacitive_charge_nC": charge*1e9,
        "reference_change_initial_current_uA": float(itotal[0]*1e6),
        "max_CE_drive_relative_WE_V": {k: float(np.max(np.abs(v))) for k,v in drives.items()},
        "SWV_200Hz_100Hz_filter_retained_fraction": next(r[3] for r in swv_rows if r[:2] == (200.,100.)),
        "analytic_checks_passed": True,
    }
    (OUT / "simulation_summary.json").write_text(json.dumps(summary, indent=2)+"\n", encoding="utf-8")
    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    main()
