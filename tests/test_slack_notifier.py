import io
import json
from unittest.mock import patch

import pytest

from core.slack_notifier import SlackNotifier


def test_image_report_comment_and_title():
    notifier = SlackNotifier("fake-token", "C-default")
    with patch.object(notifier, "_post_form", side_effect=[
        {"ok": True, "upload_url": "https://example.test/upload", "file_id": "F1"},
        {"ok": True},
    ]) as form, patch("core.slack_notifier.request.urlopen", return_value=io.BytesIO(b"ok")):
        assert notifier.send_image(b"png", "bug.png", comment="Bug description")
    fields = form.call_args.args[1]
    assert fields["initial_comment"] == "Bug description"
    assert json.loads(fields["files"])[0]["title"] == "bug.png"


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
