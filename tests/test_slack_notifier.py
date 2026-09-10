import io
import json
from unittest.mock import patch

import pytest

from core.slack_notifier import SlackNotifier


@pytest.mark.parametrize("station,expected", [
    ("el infanto (setup 5)", "[el infanto (setup 5)]\nTest message"),
    ("", "Test message"),
])
def test_message_station_header_and_target(station, expected):
    notifier = SlackNotifier("fake-token", "C-default", station_name=station)
    with patch("core.slack_notifier.request.urlopen", return_value=io.BytesIO(b'{"ok":true}')) as send:
        assert notifier.send_message("Test message", target="C-reply")
    payload = json.loads(send.call_args.args[0].data)
    assert payload == {"channel": "C-reply", "text": expected}


@pytest.mark.parametrize("station,expected", [
    ("el infanto (setup 5)", "[el infanto (setup 5)] Trend"),
    ("", "Trend"),
])
def test_image_station_title(station, expected):
    notifier = SlackNotifier("fake-token", "C-default", station_name=station)
    with patch.object(notifier, "_post_form", side_effect=[
        {"ok": True, "upload_url": "https://example.test/upload", "file_id": "F1"},
        {"ok": True},
    ]) as form, patch("core.slack_notifier.request.urlopen", return_value=io.BytesIO(b"ok")):
        assert notifier.send_image(b"png", "trend.png", title="Trend")
    fields = form.call_args.args[1]
    assert fields["channel_id"] == "C-default"
    assert json.loads(fields["files"]) == [{"id": "F1", "title": expected}]


def test_image_report_comment_has_station_header():
    notifier = SlackNotifier("fake-token", "C-default", station_name="*Station*")
    with patch.object(notifier, "_post_form", side_effect=[
        {"ok": True, "upload_url": "https://example.test/upload", "file_id": "F1"},
        {"ok": True},
    ]) as form, patch("core.slack_notifier.request.urlopen", return_value=io.BytesIO(b"ok")):
        assert notifier.send_image(b"png", "bug.png", comment="Bug description")
    fields = form.call_args.args[1]
    assert fields["initial_comment"] == "[*Station*]\nBug description"
    assert json.loads(fields["files"])[0]["title"] == "[Station] bug.png"


@pytest.mark.parametrize('reason,hint', [
    ('missing_scope', 'files:write'),
    ('not_in_channel', 'Invite'),
    ('channel_not_found', 'channel ID'),
])
def test_image_failure_explains_slack_rejection(reason, hint):
    notifier = SlackNotifier('fake-token', 'C-default')
    with patch.object(notifier, '_post_form', return_value={'ok': False, 'error': reason}):
        assert not notifier.send_image(b'png', 'bug.png')
    assert reason in notifier.last_image_error
    assert hint in notifier.last_image_error
    assert 'request upload URL' in notifier.last_image_error


def test_share_failure_is_not_reported_as_upload_success():
    notifier = SlackNotifier('fake-token', 'C-default')
    with patch.object(notifier, '_post_form', side_effect=[
        {'ok': True, 'upload_url': 'https://example.test/upload', 'file_id': 'F1'},
        {'ok': False, 'error': 'not_in_channel'},
    ]), patch('core.slack_notifier.request.urlopen', return_value=io.BytesIO(b'ok')):
        assert not notifier.send_image(b'png', 'bug.png')
    assert 'share image in channel' in notifier.last_image_error
