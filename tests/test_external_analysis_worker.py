import json
import sys
from pathlib import Path

import pytest

from core.analysis_worker import ExternalAnalysisError, run_external_analysis


def _write_worker(path: Path, body: str):
    path.write_text(body, encoding="utf-8")


def test_external_analysis_launches_and_validates_full_trace_response(tmp_path):
    project = tmp_path / "analysis_project"
    project.mkdir()
    request = tmp_path / "request.json"
    request.write_text(json.dumps({"output_dir": str(tmp_path)}), encoding="utf-8")
    worker = project / "bo_headless.py"
    _write_worker(
        worker,
        """
import argparse, json
from pathlib import Path
parser = argparse.ArgumentParser()
parser.add_argument("--request", required=True)
args = parser.parse_args()
request = json.loads(Path(args.request).read_text(encoding="utf-8"))
root = Path(request["output_dir"])
results = root / "results.json"
results.write_text(json.dumps([{
    "channel": 1,
    "peak_current": 2.5,
    "voltage": [-0.5, -0.4],
    "smoothed_corrected_current": [0.0, 2.5]
}]), encoding="utf-8")
summary = root / "summary.json"
summary.write_text(json.dumps({
    "result_count": 1,
    "channel_metrics": {"1": {"mean_peak_current_uA": 2.5}},
    "results_json": str(results),
    "analysis_engine": "test-worker"
}), encoding="utf-8")
print(summary)
""",
    )

    summary = run_external_analysis(
        request,
        project=project,
        script=worker,
        python_command=sys.executable,
        timeout_seconds=10,
    )

    assert summary["analysis_engine"] == "test-worker"
    assert summary["channel_metrics"]["1"]["mean_peak_current_uA"] == 2.5
    assert Path(summary["results_json"]).exists()


def test_external_analysis_failure_stops_bo(tmp_path):
    project = tmp_path / "analysis_project"
    project.mkdir()
    request = tmp_path / "request.json"
    request.write_text("{}", encoding="utf-8")
    worker = project / "bo_headless.py"
    _write_worker(worker, "raise RuntimeError('analysis failed')\n")

    with pytest.raises(ExternalAnalysisError, match="analysis failed"):
        run_external_analysis(
            request,
            project=project,
            script=worker,
            python_command=sys.executable,
            timeout_seconds=10,
        )
