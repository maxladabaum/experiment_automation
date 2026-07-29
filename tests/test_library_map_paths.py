from pathlib import Path

from config import METHODS_DIR
from methods import library_map


def test_library_map_root_is_resolved_from_configured_methods_dir():
    expected = Path(METHODS_DIR).expanduser()

    assert library_map._METHODS_ROOT == expected
    assert library_map._MAP_FILE == expected / "library_map.json"
