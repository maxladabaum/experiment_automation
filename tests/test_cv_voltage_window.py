from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from gui.tab_method import MethodTab
from core.method_registry import MethodRegistry
from methods import library_map


def make_tab(**overrides):
    values = dict(begin_potential='-0.2', vertex1='-0.2', vertex2='1.6',
                  step_potential='0.001', scan_rate='0.1', n_scans='3',
                  cond_potential='0', cond_time='0')
    values.update(overrides)
    tab = MethodTab.__new__(MethodTab)
    tab._cv_configure_voltage_window = Mock(get=Mock(return_value=True))
    tab.cv_params = {key: Mock(get=Mock(return_value=value)) for key, value in values.items()}
    tab._get_ba_range_config = Mock(return_value=dict(mode='auto', set_range_value='100u',
        auto_min_value='1n', auto_max_value='100u', fixed_label='125 uA',
        auto_min_label='1 nA', auto_max_label='100 uA'))
    return tab


@pytest.mark.parametrize('overrides,expected', [
    ({}, 'set_range_minmax da -200m 1600m'),
    ({'vertex1': '1.6', 'vertex2': '-0.2'}, 'set_range_minmax da -200m 1600m'),
    ({'begin_potential': '-0.4'}, 'set_range_minmax da -400m 1600m'),
    ({'cond_potential': '-0.5', 'cond_time': '2'}, 'set_range_minmax da -500m 1600m'),
    ({'cond_potential': '-10', 'cond_time': '0'}, 'set_range_minmax da -200m 1600m'),
    ({'vertex2': '1.6005'}, 'set_range_minmax da -200m 1600.5m'),
])
def test_cv_configures_full_voltage_window_before_turning_cell_on(overrides, expected):
    script = make_tab(**overrides)._build_cv_script()
    assert script.count('set_range_minmax da') == 1
    assert script.index('set_pgstat_mode 2') < script.index(expected) < script.index('set_e ')
    assert script.index(expected) < script.index('cell_on')
    assert 'set_autoranging ba 1n 100u' in script
    assert 'nscans(3)' in script


@pytest.mark.parametrize('overrides', [
    {'vertex2': '2.1'}, {'vertex1': '-1.3'},
    {'vertex1': '-1', 'vertex2': '1.6'}, {'vertex2': 'nan'},
    {'cond_time': '1', 'cond_potential': '-2'},
])
def test_invalid_low_speed_voltage_window_is_rejected(overrides):
    with pytest.raises(ValueError):
        make_tab(**overrides)._build_cv_script()


def test_regenerating_cv_does_not_reuse_old_script(monkeypatch, tmp_path):
    monkeypatch.setattr(library_map, '_METHODS_ROOT', tmp_path)
    monkeypatch.setattr(library_map, '_LIBRARY_DIR', tmp_path / 'library')
    monkeypatch.setattr(library_map, '_MUX_LIBRARY_DIR', tmp_path / 'library/mux_methods')
    monkeypatch.setattr(library_map, '_ARCHIVE_DIR', tmp_path / 'archive')
    monkeypatch.setattr(library_map, '_MAP_FILE', tmp_path / 'library_map.json')
    monkeypatch.setattr(library_map, '_map', {})
    monkeypatch.setattr('core.method_registry.SAVE_DATED_METHOD_COPIES', False)
    tab = make_tab()
    new_params = {key: value.get() for key, value in tab.cv_params.items()}
    new_params.update(tab._serialize_ba_range_config('CV'))
    old_params = dict(new_params)
    old_params.pop('cv_voltage_window_version')
    script = tab._build_cv_script()
    old_script = '\n'.join(line for line in script.splitlines() if not line.startswith('set_range_minmax da'))
    registry = MethodRegistry(base_path=tmp_path)
    old, _ = registry.save_script('CV', old_script, old_params, 1)
    new, _ = registry.save_script('CV', script, new_params, 1)
    assert old != new
    assert 'set_range_minmax da -200m 1600m' in new.read_text()
    assert 'set_range_minmax da' not in old.read_text()
    assert 'cv_voltage_window_version' not in tab._serialize_ba_range_config('SWV')

    tab._cv_configure_voltage_window.get.return_value = False
    disabled_params = {key: value.get() for key, value in tab.cv_params.items()}
    disabled_params.update(tab._serialize_ba_range_config('CV'))
    disabled_script = tab._build_cv_script()
    assert disabled_script == old_script
    disabled, _ = registry.save_script('CV', disabled_script, disabled_params, 1)
    assert disabled != new
    assert 'set_range_minmax da' not in disabled.read_text()
    assert disabled_params['cv_voltage_window_version'] == '0'
