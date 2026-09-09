from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from gui.app import ElectrochemGUI


def make_app():
    app = ElectrochemGUI.__new__(ElectrochemGUI)
    app._session_mgr = SimpleNamespace(
        _slack=SimpleNamespace(send_image=Mock(return_value=True)),
        notify_slack=Mock(return_value=True), log=Mock(),
    )
    return app


def test_text_report_does_not_upload_image():
    app = make_app()
    assert app._deliver_bug_report("description") == (True, False)
    app._session_mgr._slack.send_image.assert_not_called()
    app._session_mgr.notify_slack.assert_called_once_with("description")


def test_screenshot_contains_report_without_duplicate_text_message():
    app = make_app()
    assert app._deliver_bug_report("description", b"png") == (True, False)
    app._session_mgr._slack.send_image.assert_called_once_with(
        b"png", "bug_report.png", title="Bug report screenshot", comment="description"
    )
    app._session_mgr.notify_slack.assert_not_called()


@pytest.mark.parametrize("raises", [False, True])
@pytest.mark.parametrize("text_ok", [False, True])
def test_failed_screenshot_falls_back_to_text(raises, text_ok):
    app = make_app()
    upload = app._session_mgr._slack.send_image
    upload.return_value = False
    if raises:
        upload.side_effect = RuntimeError("upload failed")
    app._session_mgr.notify_slack.return_value = text_ok
    assert app._deliver_bug_report("description", b"png") == (text_ok, True)
    app._session_mgr.notify_slack.assert_called_once_with("description")
