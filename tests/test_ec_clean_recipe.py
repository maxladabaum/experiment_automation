import json
from pathlib import Path
from types import SimpleNamespace

from gui.tab_queue import QueueTab
from methods import library_map


def test_ec_clean_resolves_corrected_scripts_without_local_library(monkeypatch, tmp_path):
    repo = Path(__file__).resolve().parents[1]
    recipe = repo / 'recipe_maker/custom_blocks/ec_clean.json'
    payload = json.loads(recipe.read_text())
    monkeypatch.setattr(library_map, '_METHODS_ROOT', tmp_path / 'empty_methods')
    monkeypatch.setattr(library_map, 'lookup', lambda key: None)
    monkeypatch.chdir(tmp_path)
    tab = QueueTab.__new__(QueueTab)
    items = [tab._deserialize(raw) for raw in payload['items']]
    assert all(item is not None for item in items)
    tab._session = SimpleNamespace(measurement_queue=items)
    assert tab._validate_queue_scripts()
    cv_items = [item for item in items if item['type'] == 'CV']
    assert len(cv_items) == 20
    assert [item['method_ref']['mux_channel'] for item in cv_items] == list(range(1, 11)) * 2
    for index, item in enumerate(cv_items):
        script = Path(item['script_path']).read_text()
        assert 'set_range_minmax da -200m 1500m' in script
        assert script.index('set_pgstat_mode 2') < script.index('set_range_minmax da') < script.index('cell_on')
        assert tab._extract_mux_from_script(script) == item['method_ref']['mux_channel']
        loop = ('meas_loop_cv p c -200m -200m 1500m 10m 1000m nscans(25)' if index < 10
                else 'meas_loop_cv p c -200m -200m 1500m 1m 100m nscans(3)')
        assert loop in script
