import json
import threading
import time
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from core.queue_recovery import atomic_json, assess_pending, recover_bo_item, backup_recovery
from gui.tab_queue import QueueTab


def saved_run(tmp_path):
    root = tmp_path/'bo_sessions'/'bo_saved'
    root.mkdir(parents=True)
    pending = dict(method_id='m12', iteration=12, channels=[1], group_id=1,
                   group_name='One', optimization_direction='maximize', params={}, created_at='today', status='suggested')
    atomic_json(root/'bo_state.json', dict(session_id='saved', config_path='original.json',
                observations=[{'iteration':i} for i in range(1,12)], pending_batch=[pending]))
    atomic_json(root/'bo_config_snapshot.json', {'measurements_per_channel':2})
    atomic_json(root/'execution_plan.json', {'bo_block':{'objective':'paired_response', 'target_iterations':50}})
    for phase in ('buffer','target'):
        items=[]
        for repeat in (1,2):
            path = tmp_path/f'{phase}_{repeat}.csv'
            path.write_text('Potential (V),Current (uA)\n0,1\n0.1,2\n')
            items.append(dict(status='completed',csv_path=str(path),method_ref={'mux_channel':1},
                bo_ref={'session_id':'saved','method_id':'m12','phase':phase,'measurement_repeat_index':repeat}))
        atomic_json(root/'queue'/f'iter_012_queue_completion_{phase}.json',{'session_id':'saved','items':items})
    return root


def test_readonly_assessment_checks_both_phases_and_all_repeats(tmp_path):
    root=saved_run(tmp_path)
    original=(root/'bo_state.json').read_bytes()
    report=assess_pending(root)
    assert report['observations']==11 and report['pending']==1
    assert report['analysis_ready'] and report['valid_files']==4
    (tmp_path/'target_2.csv').unlink()
    assert not assess_pending(root)['analysis_ready']
    assert (root/'bo_state.json').read_bytes()==original


def test_failed_or_foreign_records_never_count_as_completed(tmp_path):
    root=saved_run(tmp_path)
    path=root/'queue/iter_012_queue_completion_buffer.json'
    data=json.loads(path.read_text()); data['session_id']='another run'
    atomic_json(path,data)
    assert not assess_pending(root)['analysis_ready']


def test_recover_preserves_total_and_backs_up_pending_data(tmp_path):
    root=saved_run(tmp_path)
    item=recover_bo_item(root)
    assert item['bo_block']['target_iterations']==50
    assert item['bo_resume_record_dir']==str(root)
    backup=backup_recovery(root)
    assert (backup/'bo_state.json').read_bytes()==(root/'bo_state.json').read_bytes()
    assert len(list((backup/'pending_csv').glob('*.csv')))==4


def test_legacy_schedule_recovery_rejects_ambiguity(tmp_path):
    root=saved_run(tmp_path); (root/'execution_plan.json').unlink()
    block={'bo_config_path':'original.json','objective':'paired_response','target_iterations':50}
    atomic_json(tmp_path/'queue_files'/'old.json', {'items':[{'type':'BO_AUTO_LOOP','bo_block':block}]})
    assert recover_bo_item(root)['bo_block']['target_iterations']==50
    atomic_json(tmp_path/'queue_files'/'other.json', {'items':[{'type':'BO_AUTO_LOOP','bo_block':dict(block,target_iterations=40)}]})
    with pytest.raises(ValueError,match='one original'):
        recover_bo_item(root)


def paused_tab():
    tab=QueueTab.__new__(QueueTab)
    tab._session=SimpleNamespace(is_running=True,pause_requested=True,update_queue_status=Mock())
    tab._root=Mock(); tab.log=Mock(); tab.set_status=Mock()
    return tab


@pytest.mark.parametrize('stop', [True,False])
def test_pause_gate_does_not_advance_and_stop_releases_it(stop):
    tab=paused_tab(); result=[]
    thread=threading.Thread(target=lambda:result.append(tab._wait_if_paused()))
    thread.start()
    try:
        deadline=time.monotonic()+2
        while not getattr(tab._session,'queue_paused',False) and time.monotonic()<deadline:
            time.sleep(.01)
        assert thread.is_alive() and not result
        if stop: tab._session.is_running=False
        else: tab.resume_queue()
        thread.join(2)
        assert result==[not stop]
    finally:
        tab._session.is_running=False
        thread.join(2)


def test_checkpoint_does_not_persist_physical_state_confirmation(tmp_path):
    tab=paused_tab()
    tab._session.session_manager=SimpleNamespace(current_experiment_path=tmp_path)
    tab._session.measurement_counter=13474
    tab._session.measurement_queue=[{'type':'BO_AUTO_LOOP','bo_recovery_confirmed':True}]
    assert tab._save_progress(0,'before')
    data=json.loads((tmp_path/'queue_progress.json').read_text())
    assert 'bo_recovery_confirmed' not in data['items'][0]
    assert data['measurement_counter']==13474


def test_checkpoint_error_stops_before_more_actions(tmp_path,monkeypatch):
    tab=paused_tab()
    tab._session.session_manager=SimpleNamespace(current_experiment_path=tmp_path)
    tab._session.measurement_counter=1;tab._session.measurement_queue=[]
    monkeypatch.setattr('gui.tab_queue.atomic_json',Mock(side_effect=OSError('full disk')))
    assert not tab._save_progress(0,'before')
    assert not tab._session.is_running


def test_completed_pending_batch_runs_only_analysis(tmp_path):
    root=saved_run(tmp_path)
    state=json.loads((root/'bo_state.json').read_text())
    session=SimpleNamespace(record_dir=root,pending_batch=state['pending_batch'],
        import_paired_analysis=Mock(return_value={'Q_run':1.0}))
    tab=paused_tab();tab._session.pause_requested=False
    tab._run_bo_analysis=Mock(return_value=tmp_path/'analysis.json')
    tab._execute_measurement_item=Mock();tab._exec_pump=Mock()
    tab._recover_pending_analysis(session,{})
    assert tab._run_bo_analysis.call_count==2
    session.import_paired_analysis.assert_called_once()
    tab._execute_measurement_item.assert_not_called();tab._exec_pump.assert_not_called()


def test_pump_failure_is_not_retried_or_swallowed():
    from pump_gui import PumpCtrl
    pump=PumpCtrl(use_sim=True)
    pump.connected=True
    pump._backend=Mock()
    pump._backend.PumpSendCommand.side_effect=RuntimeError('USB lost')
    pump._connect_backend=Mock()
    with pytest.raises(RuntimeError,match='position is uncertain'):
        pump._send('D200R')
    pump._backend.PumpSendNoWait.assert_not_called()
    pump._connect_backend.assert_called_once()


def test_speed_retries_only_after_reconnect(monkeypatch):
    from pump_gui import PumpCtrl
    monkeypatch.setattr('pump_gui.time.sleep',lambda _:None)
    pump=PumpCtrl(use_sim=True);pump.connected=True
    pump._backend=Mock();pump._backend.PumpSendCommand.side_effect=[RuntimeError('lost'),None]
    pump._connect_backend=Mock();pump._get_last_answer=Mock(return_value='ok')
    assert pump._send('S15R')=='ok'
    pump._connect_backend.assert_called_once()
    assert pump._backend.PumpSendCommand.call_count==2


def test_reconnection_bounded_and_only_queries_status(monkeypatch):
    import pump_gui
    from pump_gui import PumpCtrl
    monkeypatch.setattr(pump_gui.time,'sleep',lambda _:None)
    backend=Mock();backend.PumpInitComm.side_effect=[RuntimeError('unplugged'),None]
    monkeypatch.setattr(pump_gui,'gencache',SimpleNamespace(EnsureDispatch=Mock(return_value=backend)),raising=False)
    pump=PumpCtrl(use_sim=True);pump.com_port=4;pump.dev=0;pump.baud=9600
    pump._connect_backend()
    assert pump.connected and backend.PumpInitComm.call_count==2
    backend.PumpSendCommand.assert_called_once_with('Q',0,'')
    assert backend.CommandRetryCount==0


def test_exclude_bad_observations_preserves_earlier_learning_and_audit(tmp_path):
    from core.queue_recovery import exclude_from_iteration
    root=saved_run(tmp_path)
    state=json.loads((root/'bo_state.json').read_text())
    suggestion=state['pending_batch'][0]
    session=SimpleNamespace(record_dir=root,observations=[{'iteration':10},{'iteration':11}],
        suggestions=[dict(suggestion,iteration=11,method_id='m11')],
        save_state=Mock(),_write_history_csv=Mock())
    assert exclude_from_iteration(session,11)==1
    assert session.observations==[{'iteration':10}]
    assert session.pending_batch[0]['method_id']=='m11'
    assert list((root/'excluded_observations').glob('*/observations.json'))
    assert exclude_from_iteration(session,11)==0


def test_remeasurement_invalidates_old_traces_without_deleting_csvs(tmp_path):
    from core.queue_recovery import invalidate_pending_measurements
    root=saved_run(tmp_path)
    state=json.loads((root/'bo_state.json').read_text())
    session=SimpleNamespace(queue_dir=root/'queue',session_id='saved',
        pending_batch=state['pending_batch'], suggestions=[dict(state['pending_batch'][0],
            queue_completion_records=['old.json'])], save_state=Mock())
    backup=backup_recovery(root)
    invalidate_pending_measurements(session)
    assert assess_pending(root)['valid_files']==0
    assert 'queue_completion_records' not in session.suggestions[0]
    assert (tmp_path/'target_1.csv').exists()
    assert len(list((backup/'pending_csv').glob('*.csv')))==4


@pytest.mark.parametrize('batch_size', [1, 2])
def test_saved_bo_continues_pending_batch_without_new_session(tmp_path, monkeypatch, batch_size):
    from dataclasses import asdict
    from core.bo_session import BOIntegrationSession, normalize_bo_config
    config=normalize_bo_config({'objective':'paired_response','channels':[1],
        'channel_groups':[{'name':'One','channels':[1]}], 'n_initial_points':0,
        'parameters':{'amplitude':{'mode':'active','space':'continuous',
            'min':0.01,'max':0.08,'step':None,'value':0.04}},
        'acquisition':{'use_gp':False}})
    original=BOIntegrationSession(config,tmp_path,config_path='no-longer-present.json')
    suggestions=original.ask_batch(2)
    first=asdict(suggestions[0])
    original.observations=[dict(first,Q_run=1.0)]
    original.pending_batch=[asdict(suggestions[1])]
    original.save_state()
    loaded=BOIntegrationSession.load(original.record_dir)
    factory=Mock()
    factory.load.return_value=loaded
    monkeypatch.setattr('gui.tab_queue.BOIntegrationSession',factory)
    monkeypatch.setattr('gui.tab_queue.estimate_item_seconds',lambda _:None)
    loaded.build_queue_items=Mock(side_effect=lambda registry,suggestion,phase:
        [{'type':'SWV','phase':phase,'method_id':suggestion.method_id}])
    loaded.record_queued=Mock();loaded.record_queue_completion=Mock()
    loaded.best_observation=Mock(return_value=None)
    def import_result(suggestion,*args,**kwargs):
        obs=dict(asdict(suggestion),Q_run=2.0,quality={})
        loaded.observations.append(obs);loaded.pending_batch=[];loaded.save_state()
        return obs
    loaded.import_paired_analysis=Mock(side_effect=import_result)
    block={'objective':'paired_response','bo_config_path':'no-longer-present.json',
           'target_iterations':2,'batch_size':batch_size,'warmup_iterations':0}
    item={'type':'BO_AUTO_LOOP','bo_block':block,'bo_record_dir':str(original.record_dir),
          'bo_recovery_confirmed':True,'bo_recovery_mode':'remeasure'}
    tab=paused_tab();tab._session.pause_requested=False
    tab._session.measurement_queue=[item];tab._session.measurement_counter=10
    tab._session.registry=Mock()
    tab._session.session_manager=SimpleNamespace(current_experiment_path=tmp_path,
        require_experiment=lambda:tmp_path,notify_slack=Mock())
    for name in ('refresh','_append_bo_progress','_update_bo_progress','_set_bo_live_details'):
        setattr(tab,name,Mock())
    tab._load_bo_exchange_items=Mock(return_value=[{'type':'PUMP_VALVE'}])
    tab._execute_bo_operational_items=Mock(return_value=True)
    tab._execute_bo_equilibration_pause=Mock(return_value=True)
    tab._run_bo_queue_items=Mock(side_effect=lambda items,*args,**kwargs:(len(items),0,0,items))
    tab._run_bo_analysis=Mock(return_value=tmp_path/'analysis.json')
    assert tab._exec_bo_auto_loop(item)
    factory.assert_not_called()
    factory.load.assert_called_once_with(str(original.record_dir))
    assert loaded.observations[0]==dict(first,Q_run=1.0)
    assert len(loaded.observations)==2 and not loaded.pending_batch
    assert loaded.import_paired_analysis.call_args.args[0].method_id==suggestions[1].method_id
    assert tab._run_bo_queue_items.call_count==2
    assert tab._execute_bo_operational_items.call_count==2
    assert item['bo_session_id']==original.session_id
    assert list((original.record_dir/'recovery_backups').glob('*/bo_state.json'))


def test_pump_ambiguous_motion_does_not_advance_position_or_allow_next_move(monkeypatch):
    from pump_gui import PumpCtrl
    pump=PumpCtrl(use_sim=True);pump.connected=True
    pump._backend=Mock();pump._backend.PumpSendCommand.side_effect=RuntimeError('USB lost')
    pump._connect_backend=Mock()
    with pytest.raises(RuntimeError):
        pump._send('A200R')
    assert pump.position_uncertain and pump._plunger_steps==0
    with pytest.raises(RuntimeError,match='reconcile'):
        pump._send('D200R')
    assert pump._backend.PumpSendCommand.call_count==1


def test_pump_reconnect_exhaustion_is_bounded(monkeypatch):
    import pump_gui
    monkeypatch.setattr(pump_gui.time,'sleep',lambda _:None)
    backend=Mock();backend.PumpInitComm.side_effect=RuntimeError('unplugged')
    monkeypatch.setattr(pump_gui,'gencache',SimpleNamespace(EnsureDispatch=Mock(return_value=backend)),raising=False)
    pump=pump_gui.PumpCtrl(use_sim=True)
    with pytest.raises(RuntimeError,match='exhausted'):
        pump._connect_backend()
    assert not pump.connected
    assert backend.PumpInitComm.call_count==3
    backend.PumpSendCommand.assert_not_called()
