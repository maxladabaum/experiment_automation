"""Launch the 64-bit electrochemistry analysis worker and validate its response."""

from __future__ import annotations

import json
import os
import shlex
import struct
import subprocess
import sys
from pathlib import Path
from typing import Any, Dict, List

from config import (
    BO_EXTERNAL_ANALYSIS_MODE,
    BO_EXTERNAL_ANALYSIS_PROJECT,
    BO_EXTERNAL_ANALYSIS_PYTHON,
    BO_EXTERNAL_ANALYSIS_SCRIPT,
    BO_EXTERNAL_ANALYSIS_TIMEOUT_SECONDS,
    BO_LOCAL_PATHS_CONFIG,
)


class ExternalAnalysisError(RuntimeError):
    pass


def _python_command(project: Path, configured: str = "") -> List[str]:
    if configured:
        raw = str(configured).strip()
        unquoted = raw
        if len(unquoted) >= 2 and unquoted[0] == unquoted[-1] and unquoted[0] in ("'", '"'):
            unquoted = unquoted[1:-1]
        if Path(unquoted).expanduser().exists():
            return [str(Path(unquoted).expanduser())]
        parts = shlex.split(raw, posix=os.name != "nt")
        if os.name == "nt":
            parts = [
                part[1:-1]
                if len(part) >= 2 and part[0] == part[-1] and part[0] in ("'", '"')
                else part
                for part in parts
            ]
        return parts
    candidates = (
        project / ".venv64" / "Scripts" / "python.exe",
        project / ".venv" / "Scripts" / "python.exe",
        project / ".venv64" / "bin" / "python",
        project / ".venv" / "bin" / "python",
    )
    for candidate in candidates:
        if candidate.exists():
            return [str(candidate)]
    if struct.calcsize("P") * 8 == 64:
        return [sys.executable]
    if os.name == "nt":
        return ["py", "-3-64"]
    raise ExternalAnalysisError(
        "No 64-bit analysis Python was found. Configure analysis_python in "
        "optimizer/bo_configs/local_paths.json or set EA_BO_ANALYSIS_PYTHON."
    )


def _validate_summary(summary_path: Path) -> Dict[str, Any]:
    if not summary_path.exists():
        raise ExternalAnalysisError(f"Analysis worker did not create {summary_path}")
    try:
        with open(summary_path, "r", encoding="utf-8") as fh:
            summary = json.load(fh)
    except Exception as exc:
        raise ExternalAnalysisError(f"Invalid analysis response {summary_path}: {exc}") from exc
    if not isinstance(summary, dict) or not isinstance(summary.get("channel_metrics"), dict):
        raise ExternalAnalysisError("Analysis response is missing channel_metrics")
    if int(summary.get("result_count", 0) or 0) < 1:
        raise ExternalAnalysisError("Analysis worker returned no trace results")
    results_json = summary.get("results_json")
    if not results_json or not Path(results_json).exists():
        raise ExternalAnalysisError("Analysis response is missing the full results JSON")
    summary["summary_path"] = str(summary_path)
    return summary


def run_external_analysis(
    request_path: str | Path,
    *,
    project: str | Path | None = None,
    script: str | Path | None = None,
    python_command: str = "",
    timeout_seconds: float | None = None,
) -> Dict[str, Any]:
    request_path = Path(request_path).resolve()
    project_path = Path(project or BO_EXTERNAL_ANALYSIS_PROJECT).resolve()
    script_path = Path(script or BO_EXTERNAL_ANALYSIS_SCRIPT).resolve()
    if not project_path.is_dir():
        raise ExternalAnalysisError(f"Analysis project not found: {project_path}")
    if not script_path.is_file():
        raise ExternalAnalysisError(f"Analysis worker not found: {script_path}")
    command = _python_command(
        project_path,
        configured=python_command or BO_EXTERNAL_ANALYSIS_PYTHON,
    ) + [str(script_path), "--request", str(request_path)]
    try:
        completed = subprocess.run(
            command,
            cwd=str(project_path),
            capture_output=True,
            text=True,
            timeout=float(timeout_seconds or BO_EXTERNAL_ANALYSIS_TIMEOUT_SECONDS),
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise ExternalAnalysisError(
            f"Analysis worker timed out after {exc.timeout:g} seconds"
        ) from exc
    except OSError as exc:
        raise ExternalAnalysisError(
            f"Could not launch 64-bit analysis worker ({' '.join(command)}): {exc}"
        ) from exc
    if completed.returncode != 0:
        detail = (completed.stderr or completed.stdout or "").strip()
        raise ExternalAnalysisError(
            f"Analysis worker exited with code {completed.returncode}: {detail}"
        )
    output_lines = [line.strip() for line in completed.stdout.splitlines() if line.strip()]
    if not output_lines:
        raise ExternalAnalysisError("Analysis worker returned no response path")
    summary_path = Path(output_lines[-1])
    if not summary_path.is_absolute():
        summary_path = project_path / summary_path
    return _validate_summary(summary_path)


def run_analysis(request_path: str | Path, request: dict) -> Dict[str, Any]:
    settings = {}
    try:
        with open(BO_LOCAL_PATHS_CONFIG, "r", encoding="utf-8") as fh:
            settings = json.load(fh)
        if not isinstance(settings, dict):
            settings = {}
    except Exception:
        settings = {}
    mode = str(settings.get("analysis_mode", BO_EXTERNAL_ANALYSIS_MODE)).strip().lower()
    if mode == "local":
        from core.bo_analysis import run_request

        return run_request(request)
    project = Path(settings.get("analysis_project") or BO_EXTERNAL_ANALYSIS_PROJECT)
    script = Path(
        settings.get("analysis_script")
        or (project / "analysis_worker" / "bo_headless.py")
    )
    return run_external_analysis(
        request_path,
        project=project,
        script=script,
        python_command=str(settings.get("analysis_python") or BO_EXTERNAL_ANALYSIS_PYTHON),
        timeout_seconds=float(
            settings.get("analysis_timeout_seconds", BO_EXTERNAL_ANALYSIS_TIMEOUT_SECONDS)
        ),
    )
