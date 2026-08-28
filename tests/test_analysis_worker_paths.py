from pathlib import Path

from core.analysis_worker import _python_command


def test_configured_python_path_with_spaces_is_single_command_part(tmp_path):
    python_path = tmp_path / "Path With Spaces" / "python.exe"
    python_path.parent.mkdir()
    python_path.write_text("", encoding="utf-8")

    assert _python_command(Path("."), str(python_path)) == [str(python_path)]
    assert _python_command(Path("."), f'"{python_path}"') == [str(python_path)]
