from types import SimpleNamespace
from unittest.mock import Mock

import pytest

import pump_gui


@pytest.mark.parametrize('attribute', ['CLSIDToClassMap', 'CLSIDToPackageMap'])
def test_incomplete_generated_wrapper_uses_direct_binding(monkeypatch, attribute):
    generated = Mock(side_effect=AttributeError(
        f"module 'win32com.gen_py.pump_typelib' has no attribute '{attribute}'"))
    backend = Mock()
    direct = Mock(return_value=backend)
    monkeypatch.setattr(pump_gui, 'gencache', SimpleNamespace(EnsureDispatch=generated), raising=False)
    monkeypatch.setattr(pump_gui, 'dynamic', SimpleNamespace(Dispatch=direct), raising=False)
    pump = pump_gui.PumpCtrl(use_sim=True)

    assert pump._create_com_backend() is backend
    generated.assert_called_once_with(pump_gui.PROGID)
    direct.assert_called_once_with(pump_gui.PROGID)
    backend.PumpInitComm.assert_not_called()
    backend.PumpSendCommand.assert_not_called()


def test_healthy_generated_binding_stays_unchanged(monkeypatch):
    backend = Mock()
    generated = Mock(return_value=backend)
    direct = Mock()
    monkeypatch.setattr(pump_gui, 'gencache', SimpleNamespace(EnsureDispatch=generated), raising=False)
    monkeypatch.setattr(pump_gui, 'dynamic', SimpleNamespace(Dispatch=direct), raising=False)
    assert pump_gui.PumpCtrl(use_sim=True)._create_com_backend() is backend
    direct.assert_not_called()


@pytest.mark.parametrize('error', [AttributeError('Unrelated driver property'),
                                  RuntimeError('Class not registered')])
def test_other_driver_failures_are_not_hidden(monkeypatch, error):
    generated = Mock(side_effect=error)
    direct = Mock()
    monkeypatch.setattr(pump_gui, 'gencache', SimpleNamespace(EnsureDispatch=generated), raising=False)
    monkeypatch.setattr(pump_gui, 'dynamic', SimpleNamespace(Dispatch=direct), raising=False)
    with pytest.raises(type(error), match=str(error)):
        pump_gui.PumpCtrl(use_sim=True)._create_com_backend()
    direct.assert_not_called()
