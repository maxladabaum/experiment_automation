import copy
import json
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

from core.protocol_io import loaded_items, write_protocol
from gui.tab_queue import QueueTab
from gui.tab_recipe_maker import RecipeMakerTab
from methods import library_map


def test_protocol_moves_between_setups_using_relative_library_addresses(monkeypatch, tmp_path):
    roots = [tmp_path / 'setup_a' / 'custom_methods', tmp_path / 'setup_b' / 'custom_methods']
    scripts = []
    for root in roots:
        script = root / 'library/mux_methods/cv_ch1_test.ms'
        script.parent.mkdir(parents=True)
        script.write_text('e\n# original method\n')
        scripts.append(script)
    monkeypatch.setattr(library_map, '_METHODS_ROOT', roots[0])
    items = [{'type': 'CV', 'script_path': str(scripts[0]), 'details': 'MUX ch 1'}]
    original = copy.deepcopy(items)
    path = tmp_path / 'protocol.json'
    write_protocol(path, {'items': items})
    saved = json.loads(path.read_text())
    assert saved['items'][0]['script_path'] == 'methods/library/mux_methods/cv_ch1_test.ms'
    assert 'script_content' not in saved['items'][0]
    assert items == original
    monkeypatch.setattr(library_map, '_METHODS_ROOT', roots[1])
    loaded = loaded_items(saved, path)
    assert loaded[0]['script_path'] == str(scripts[1])
    tab = QueueTab.__new__(QueueTab)
    assert tab._deserialize(loaded[0])['script_path'] == str(scripts[1])


def test_legacy_foreign_paths_save_without_old_username(tmp_path):
    path = tmp_path / 'protocol.json'
    write_protocol(path, {'items': [{'type': 'CV', 'script_path':
        r'C:\Users\Other Lab\Documents\Data\methods\library\mux_methods\cv_missing.ms'}]})
    saved = json.loads(path.read_text())
    assert saved['items'][0]['script_path'] == 'methods/library/mux_methods/cv_missing.ms'
    assert list(tmp_path.iterdir()) == [path]  # No script files are created or bundled.


def test_custom_library_nested_below_methods_directory(monkeypatch, tmp_path):
    root = tmp_path / 'methods' / 'station_library'
    monkeypatch.setattr(library_map, '_METHODS_ROOT', root)
    path = tmp_path / 'recipe.json'
    write_protocol(path, {'items': [{'type': 'CV', 'script_path': str(root / 'library/cv.ms')}]})
    assert json.loads(path.read_text())['items'][0]['script_path'] == 'methods/library/cv.ms'


def test_external_methods_folder_does_not_alias_local_library(monkeypatch, tmp_path):
    local = tmp_path / 'local_methods/library/custom.ms'
    external = tmp_path / 'other_project/methods/library/custom.ms'
    for script, contents in ((local, 'LOCAL'), (external, 'EXTERNAL')):
        script.parent.mkdir(parents=True)
        script.write_text(contents)
    monkeypatch.setattr(library_map, '_METHODS_ROOT', local.parents[1])
    destination = tmp_path / 'recipe.json'
    write_protocol(destination, {'items': [{'type': 'CV', 'script_path': str(external)}]})
    saved = json.loads(destination.read_text())
    assert saved['items'][0]['script_path'] == 'other_project/methods/library/custom.ms'
    assert Path(loaded_items(saved, destination)[0]['script_path']).read_text() == 'EXTERNAL'
    external.unlink()
    assert library_map.resolve_script_path(saved['items'][0]['script_path'], tmp_path) is None


def test_explicit_relative_custom_path_is_preserved(monkeypatch, tmp_path):
    monkeypatch.setattr(library_map, '_METHODS_ROOT', tmp_path / 'local_methods')
    address = '../other_project/methods/custom.ms'
    path = tmp_path / 'recipe.json'
    write_protocol(path, {'items': [{'type': 'CV', 'script_path': address}]})
    assert json.loads(path.read_text())['items'][0]['script_path'] == address


def test_external_methods_folder_beside_recipe_uses_explicit_relative_address(monkeypatch, tmp_path):
    local = tmp_path / 'local_library/library/custom.ms'
    external = tmp_path / 'project/methods/library/custom.ms'
    for script in (local, external):
        script.parent.mkdir(parents=True)
        script.write_text('e\n')
    monkeypatch.setattr(library_map, '_METHODS_ROOT', tmp_path / 'local_library')
    destination = tmp_path / 'project/recipe.json'
    write_protocol(destination, {'items': [{'type': 'CV', 'script_path': str(external)}]})
    saved = json.loads(destination.read_text())
    assert saved['items'][0]['script_path'] == './methods/library/custom.ms'
    assert Path(loaded_items(saved, destination)[0]['script_path']) == external
    external.unlink()
    assert library_map.resolve_script_path('./methods/library/custom.ms', destination.parent) is None


def test_missing_relative_script_does_not_fall_back_to_working_directory(monkeypatch, tmp_path):
    recipe_folder = tmp_path / 'recipe'
    recipe_folder.mkdir()
    unrelated = tmp_path / 'scripts/custom.ms'
    unrelated.parent.mkdir()
    unrelated.write_text('WRONG FILE')
    monkeypatch.chdir(tmp_path)
    assert library_map.resolve_script_path('scripts/custom.ms', recipe_folder) is None


def test_bundled_ec_clean_script_saves_as_portable_library_path(tmp_path):
    script = next((Path(__file__).resolve().parents[1] / 'methods/library/ec_clean').glob('*.ms'))
    path = tmp_path / 'recipe.json'
    write_protocol(path, {'items': [{'type': 'CV', 'script_path': str(script)}]})
    assert json.loads(path.read_text())['items'][0]['script_path'] == f'methods/library/ec_clean/{script.name}'


def test_custom_script_resolves_relative_to_protocol_not_working_directory(monkeypatch, tmp_path):
    folder = tmp_path / 'protocol'
    folder.mkdir()
    script = folder / 'scripts/custom.ms'
    script.parent.mkdir()
    script.write_text('e\n')
    path = folder / 'recipe.json'
    write_protocol(path, {'items': [{'type': 'CUSTOM', 'script_path': str(script)}]})
    saved = json.loads(path.read_text())
    assert saved['items'][0]['script_path'] == 'scripts/custom.ms'
    monkeypatch.chdir(tmp_path)
    assert loaded_items(saved, path)[0]['script_path'] == str(script)


def test_local_methods_take_precedence_over_bundled_copy(monkeypatch, tmp_path):
    local = tmp_path / 'local/library/test.ms'
    bundled = tmp_path / 'repo/methods/library/test.ms'
    for path in (local, bundled):
        path.parent.mkdir(parents=True)
        path.write_text('e\n')
    monkeypatch.setattr(library_map, '_METHODS_ROOT', tmp_path / 'local')
    monkeypatch.chdir(tmp_path / 'repo')
    assert library_map.resolve_script_path('methods/library/test.ms') == local


def test_queue_and_recipe_save_buttons_write_relative_paths(monkeypatch, tmp_path):
    item = {'type': 'CV', 'status': 'pending', 'details': 'test',
            'script_path': r'C:\old\methods\library\cv_test.ms'}
    queue_path = tmp_path / 'queue.json'
    recipe_path = tmp_path / 'recipe.json'
    monkeypatch.setattr('gui.tab_queue.filedialog.asksaveasfilename', Mock(return_value=str(queue_path)))
    monkeypatch.setattr('gui.tab_queue.messagebox.showinfo', Mock())
    queue = QueueTab.__new__(QueueTab)
    queue._session = SimpleNamespace(measurement_queue=[item], is_running=False)
    queue.log = Mock()
    queue.save_queue()
    monkeypatch.setattr('gui.tab_recipe_maker.filedialog.asksaveasfilename', Mock(return_value=str(recipe_path)))
    recipe = RecipeMakerTab.__new__(RecipeMakerTab)
    recipe._recipe = [item]
    recipe._recipe_root = tmp_path
    recipe._save_recipe()
    for path in (queue_path, recipe_path):
        assert json.loads(path.read_text())['items'][0]['script_path'] == 'methods/library/cv_test.ms'


def test_existing_absolute_script_and_method_reference_still_load(monkeypatch, tmp_path):
    script = tmp_path / 'original.ms'
    script.write_text('e\n# original experiment\n')
    monkeypatch.setattr(library_map, 'lookup', lambda key: script if key == 'original' else None)
    tab = QueueTab.__new__(QueueTab)
    assert tab._deserialize({'type': 'CV', 'script_path': str(script)})['script_path'] == str(script)
    ref = {'hash_key': 'original', 'technique': 'CV', 'params': {'cycles': 5}, 'mux_channel': 0}
    loaded = tab._deserialize({'type': 'CV', 'method_ref': ref})
    assert loaded['script_path'] == str(script)
    assert loaded['method_ref'] == ref


def test_pump_steps_and_measurement_settings_are_preserved(tmp_path):
    items = [
        {'type': 'PUMP_ASPIRATE', 'pump_action': {'name': 'ASPIRATE', 'params': {'volume': 250, 'speed': 13}}},
        {'type': 'PAUSE', 'pause_seconds': 30},
        {'type': 'CV', 'method_ref': {'hash_key': 'cv_test', 'params': {'cycles': 5}, 'mux_channel': 3}},
    ]
    path = tmp_path / 'recipe.json'
    write_protocol(path, {'items': items})
    assert loaded_items(json.loads(path.read_text()), path) == items
