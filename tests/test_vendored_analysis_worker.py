import csv
import json
import math
import shutil
import subprocess
import sys
from pathlib import Path

from core.reference_swv.io import collect_swv_csvs_from_folders


def test_direction_qualified_filename_is_a_distinct_analysis_channel(tmp_path):
    (tmp_path / "swv_ch1_ab_meas_20260101_1200_1_ch1_max.csv").touch()
    (tmp_path / "swv_ch1_cd_meas_20260101_1201_2_ch1_min.csv").touch()

    files = collect_swv_csvs_from_folders([str(tmp_path)])

    assert sorted(file.ch for file in files) == ["1_max", "1_min"]


def test_vendored_worker_runs_as_subprocess_and_excludes_dc_from_noise(tmp_path):
    data_dir = tmp_path / "data"
    data_dir.mkdir()
    csv_path = data_dir / "swv_ch1_ab_meas_20260101_1200_1_ch1.csv"
    with open(csv_path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        writer.writerow(["Potential (V)", "Current (uA)"])
        for index in range(301):
            voltage = -0.6 + index * 0.002
            peak = 0.8 * math.exp(-((voltage + 0.30) / 0.055) ** 2)
            fluctuation = 0.007 * math.sin(index * 1.73)
            writer.writerow([voltage, 1.3 + peak + fluctuation])
    shutil.copy2(
        csv_path,
        data_dir / "swv_ch1_cd_meas_20260101_1201_2_ch1_max.csv",
    )
    shutil.copy2(
        csv_path,
        data_dir / "swv_ch1_ef_meas_20260101_1202_3_ch1_min.csv",
    )

    output_dir = tmp_path / "output"
    request = {
        "folders": [str(data_dir)],
        "output_dir": str(output_dir),
        "output_stem": "vendored_worker_test",
        "analysis": {
            "crop_min_v": -0.45,
            "crop_max_v": 0.0,
            "smooth_window": 5,
            "smooth_polyorder": 2,
            "minima_search_window_v": 0.2,
            "min_peak_height_ua": None,
            "min_start_voltage_v": -0.7,
        },
    }
    request_path = tmp_path / "request.json"
    request_path.write_text(json.dumps(request), encoding="utf-8")
    project = Path(__file__).resolve().parents[1]
    completed = subprocess.run(
        [sys.executable, str(project / "analysis_worker" / "bo_headless.py"), "--request", str(request_path)],
        cwd=project,
        capture_output=True,
        text=True,
        check=False,
    )

    assert completed.returncode == 0, completed.stderr
    summary_path = Path(completed.stdout.strip().splitlines()[-1])
    summary = json.loads(summary_path.read_text(encoding="utf-8"))
    metrics = summary["channel_metrics"]["1"]
    assert summary["channel_metrics"]["1_max"]["ok_scan_count"] == 1
    assert summary["channel_metrics"]["1_min"]["ok_scan_count"] == 1
    assert metrics["ok_scan_count"] == 1
    assert metrics["median_peak_current_uA"] > 0.5
    assert 0.0 < metrics["median_background_rms_uA"] < 0.1
