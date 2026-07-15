from pathlib import Path

from methods import library_map


def test_library_map_root_is_resolved_from_module_location():
    expected = Path(library_map.__file__).resolve().parent

    assert library_map._METHODS_ROOT == expected
    assert library_map._MAP_FILE == expected / "library_map.json"
