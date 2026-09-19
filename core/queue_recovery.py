"""BO recovery assessment, backups, and atomic queue progress checkpoints."""
import copy
import csv
import json
import math
import os
from pathlib import Path
import tempfile
import shutil
from datetime import datetime


def atomic_json(path, payload):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(mode='w', encoding='utf8', dir=path.parent,
                                         suffix='.tmp', delete=False) as handle:
            temporary = Path(handle.name)
            json.dump(payload, handle, indent=2)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
    finally:
        if temporary is not None and temporary.exists():
            temporary.unlink()


def assess_pending(record_dir):
    """Require all expected phase/channel/repeat CSVs, never infer delivery."""
    root = Path(record_dir)
    state = json.loads((root/'bo_state.json').read_text(encoding='utf8'))
    config = json.loads((root/'bo_config_snapshot.json').read_text(encoding='utf8'))
    pending = state.get('pending_batch') or ([state['pending']] if state.get('pending') else [])
    expected = set()
    repeats = max(1, int(config.get('measurements_per_channel', 1)))
    for suggestion in pending:
        for phase in ('buffer', 'target'):
            for channel in suggestion.get('channels', []):
                for repeat in range(1, repeats + 1):
                    expected.add((suggestion['method_id'], phase, int(channel), repeat))
    found = {}
    for record in sorted((root/'queue').glob('*queue_completion*.json')):
        data = json.loads(record.read_text(encoding='utf8'))
        if data.get('session_id') != state['session_id']:
            continue
        for item in data.get('items', []):
            ref = item.get('bo_ref') or {}
            key = (ref.get('method_id'), ref.get('phase'),
                   int((item.get('method_ref') or {}).get('mux_channel') or
                       str(ref.get('channel_label', '0')).split('_')[0]),
                   int(ref.get('measurement_repeat_index', 1)))
            if key not in expected or item.get('status') != 'completed':
                continue
            path = Path(item.get('csv_path') or '')
            if not path.is_absolute():
                path = root.parent.parent/path
            if not path.is_file():
                continue
            try:
                with path.open(encoding='utf-8-sig', newline='') as handle:
                    rows = csv.DictReader(handle)
                    valid = sum(1 for row in rows if
                        math.isfinite(float(row['Potential (V)'])) and
                        math.isfinite(float(row['Current (uA)'])))
                if valid >= 2:
                    found[key] = str(path)
            except (OSError, ValueError, KeyError):
                continue
    return {'session_id': state['session_id'], 'observations': len(state.get('observations', [])),
            'pending': len(pending), 'expected_files': len(expected),
            'valid_files': len(found), 'analysis_ready': bool(expected) and expected <= found.keys(),
            'files': list(found.values())}


def recover_bo_item(record_dir):
    """Recover the original schedule, including legacy sessions without a plan."""
    root = Path(record_dir).resolve()
    state = json.loads((root/'bo_state.json').read_text(encoding='utf8'))
    plan = root/'execution_plan.json'
    if plan.exists():
        block = json.loads(plan.read_text(encoding='utf8'))['bo_block']
    else:
        matches = []
        for snapshot in (root.parent.parent/'queue_files').glob('*'):
            if not snapshot.is_file():
                continue
            try:
                data = json.loads(snapshot.read_text(encoding='utf-8-sig'))
            except (ValueError, OSError):
                continue
            for item in data.get('items', []):
                candidate = item.get('bo_block') or {}
                if (item.get('type') == 'BO_AUTO_LOOP' and
                    str(candidate.get('bo_config_path', '')).replace('\\','/').casefold() ==
                    str(state.get('config_path', '')).replace('\\','/').casefold()):
                    matches.append(candidate)
        unique = {json.dumps(x, sort_keys=True): x for x in matches}
        if len(unique) != 1:
            raise ValueError('Cannot identify one original BO schedule. Load the original saved queue and select its BO parent row for recovery.')
        block = next(iter(unique.values()))
    if block.get('objective') != 'paired_response':
        raise ValueError('Recover BO currently requires a paired-response BO session.')
    return {'type': 'BO_AUTO_LOOP', 'status': 'pending', 'bo_block': copy.deepcopy(block),
            'bo_record_dir': str(root), 'bo_resume_record_dir': str(root),
            'details': 'Resume saved paired BO: ' + state['session_id']}


def backup_recovery(record_dir):
    """Snapshot metadata and pending input CSVs before recovery mutates records."""
    root = Path(record_dir)
    target = root/'recovery_backups'/datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    target.mkdir(parents=True, exist_ok=False)
    for name in ('bo_state.json', 'bo_config_snapshot.json', 'execution_plan.json', 'history.csv'):
        if (root/name).is_file():
            shutil.copy2(root/name, target/name)
    if (root/'queue').exists():
        shutil.copytree(root/'queue', target/'queue')
    report = assess_pending(root)
    (target/'pending_csv').mkdir()
    for path in report['files']:
        shutil.copy2(path, target/'pending_csv'/Path(path).name)
    atomic_json(target/'assessment.json', report)
    return target


def exclude_from_iteration(session, first_invalid):
    """Explicit rollback after backup; preserve traces and the rejected history."""
    from dataclasses import fields
    from core.bo_session import BOSuggestion
    first_invalid = int(first_invalid)
    if first_invalid < 1:
        raise ValueError('First invalid iteration must be positive.')
    rejected = [x for x in session.observations if int(x['iteration']) >= first_invalid]
    originals = [x for x in session.suggestions if int(x['iteration']) == first_invalid]
    if not rejected and any(int(x['iteration']) == first_invalid for x in session.pending_batch):
        return 0  # A restart after the rollback was saved must not apply it twice.
    if not rejected or not originals:
        raise ValueError('No completed observations/suggestions found at that iteration.')
    allowed = {x.name for x in fields(BOSuggestion)}
    pending = [{k:v for k,v in x.items() if k in allowed} for x in originals]
    for record in pending:
        record['status'] = 'suggested'
    audit = session.record_dir/'excluded_observations'/datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    audit.mkdir(parents=True, exist_ok=False)
    atomic_json(audit/'observations.json', {'reason':'Operator reported failed/uncertain fluid exchanges',
                'first_invalid_iteration':first_invalid, 'observations':rejected})
    if (session.record_dir/'analysis').exists():
        shutil.copytree(session.record_dir/'analysis', audit/'analysis')
    session.observations = [x for x in session.observations if int(x['iteration']) < first_invalid]
    session.suggestions = [x for x in session.suggestions if int(x['iteration']) < first_invalid] + copy.deepcopy(pending)
    session.pending = None
    session.pending_batch = pending
    session.save_state()
    session._write_history_csv()
    return len(rejected)


def invalidate_pending_measurements(session):
    """After backup, detach suspect traces so a later restart cannot reuse them."""
    pending_ids = {x['method_id'] for x in session.pending_batch}
    for path in session.queue_dir.glob('*queue_completion*.json'):
        payload = json.loads(path.read_text(encoding='utf8'))
        if payload.get('session_id') != session.session_id:
            continue
        items = payload.get('items', [])
        retained = [x for x in items if (x.get('bo_ref') or {}).get('method_id') not in pending_ids]
        if len(retained) != len(items):
            payload['items'] = retained
            payload['recovery_invalidated_pending'] = sorted(pending_ids)
            atomic_json(path, payload)
    for record in session.suggestions:
        if record['method_id'] in pending_ids:
            for key in ('queue_completion_record', 'queue_completion_records',
                        'queue_completed_at', 'completed_queue_items', 'failed_queue_items'):
                record.pop(key, None)
    session.save_state()
