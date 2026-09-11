from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from gui.tab_queue import QueueTab
from methods import library_map


@pytest.fixture
def local_library(monkeypatch, tmp_path):
    monkeypatch.setattr(library_map, '_METHODS_ROOT', tmp_path / 'methods')
    path = tmp_path / 'methods/library/mux_methods/cv_ch1_exact.ms'
    path.parent.mkdir(parents=True)
    path.write_text('exact original script')
    return path


def test_old_windows_path_resolves_to_exact_local_method(local_library):
    old = r'C:\Users\Other Lab\Documents\Experiment Automation Data\methods\library\mux_methods\cv_ch1_exact.ms'
    assert library_map.resolve_script_path(old) == local_library
    assert library_map.resolve_script_path(r'methods\library\mux_methods\cv_ch1_exact.ms') == local_library
    assert library_map.resolve_script_path(str(local_library)) == local_library
    assert library_map.resolve_script_path(old.replace('exact', 'missing')) is None


@pytest.mark.parametrize('start_index', [None, 0])
def test_missing_measurement_blocks_start_before_pump_steps(monkeypatch, start_index):
    tab = QueueTab.__new__(QueueTab)
    tab._session = SimpleNamespace(is_running=False, measurement_queue=[
        {'type': 'PUMP_ASPIRATE'},
        {'type': 'CV', 'script_path': 'missing-original-cleaning-script.ms'},
    ])
    tab._reset_reorder = Mock()
    tab.log = Mock()
    error = Mock()
    monkeypatch.setattr('gui.tab_queue.messagebox.showerror', error)
    if start_index is None:
        tab.run_queue()
    else:
        tab.run_from_index(start_index)
    assert tab._session.is_running is False
    error.assert_called_once()
    assert 'Item 2' in error.call_args.args[1]


def test_resume_checks_only_remaining_items_and_resolves_paths(monkeypatch, local_library):
    tab = QueueTab.__new__(QueueTab)
    tab._session = SimpleNamespace(measurement_queue=[
        {'type': 'CV', 'script_path': 'missing-already-completed.ms'},
        {'type': 'CV', 'script_path': r'C:\old\methods\library\mux_methods\cv_ch1_exact.ms'},
    ])
    assert tab._validate_queue_scripts(1)
    assert tab._session.measurement_queue[1]['script_path'] == str(local_library)


def test_missing_hash_reference_cannot_silently_remove_measurement_from_loaded_queue(monkeypatch, tmp_path):
    import json
    path = tmp_path / 'recipe.json'
    path.write_text(json.dumps({'items': [
        {'type': 'PUMP_ASPIRATE', 'pump_action': {'name': 'ASPIRATE', 'params': {'volume': 10}}},
        {'type': 'CV', 'method_ref': {'hash_key': 'unavailable'}},
    ]}))
    tab = QueueTab.__new__(QueueTab)
    original = [{'type': 'PAUSE', 'pause_seconds': 1}]
    tab._session = SimpleNamespace(is_running=False, measurement_queue=original)
    monkeypatch.setattr(library_map, 'lookup', lambda key: None)
    monkeypatch.setattr('gui.tab_queue.filedialog.askopenfilename', lambda **kwargs: str(path))
    error = Mock()
    monkeypatch.setattr('gui.tab_queue.messagebox.showerror', error)
    tab.load_queue()
    assert tab._session.measurement_queue is original
    assert 'Item 2' in error.call_args.args[1]
