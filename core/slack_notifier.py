"""
core/slack_notifier.py - Minimal Slack notifier (chat.postMessage).
"""

import json
from typing import Callable, Optional
from urllib import error, request
from urllib.parse import urlencode


class SlackNotifier:
    """Send plain-text Slack notifications with a bot token."""

    _POST_MESSAGE_URL = "https://slack.com/api/chat.postMessage"
    _GET_UPLOAD_URL = "https://slack.com/api/files.getUploadURLExternal"
    _COMPLETE_UPLOAD_URL = "https://slack.com/api/files.completeUploadExternal"

    def __init__(
        self,
        bot_token: str,
        default_target: str,
        log_callback: Callable[[str], None] = print,
        timeout_seconds: float = 8.0,
    ):
        self._token = (bot_token or "").strip()
        self._target = (default_target or "").strip()
        self._log = log_callback
        self._timeout = float(timeout_seconds)

    @property
    def enabled(self) -> bool:
        return bool(self._token and self._target)

    def send_message(self, text: str, target: Optional[str] = None) -> bool:
        """Post a message to Slack and return True on success."""
        channel = (target or self._target or "").strip()
        if not text or not self._token or not channel:
            return False

        payload = json.dumps({"channel": channel, "text": text}).encode("utf-8")
        req = request.Request(
            self._POST_MESSAGE_URL,
            data=payload,
            method="POST",
            headers={
                "Authorization": f"Bearer {self._token}",
                "Content-Type": "application/json; charset=utf-8",
            },
        )
        try:
            with request.urlopen(req, timeout=self._timeout) as resp:
                body = resp.read().decode("utf-8", errors="replace")
            data = json.loads(body)
            ok = bool(data.get("ok"))
            if not ok:
                self._log(f"Slack notify failed: {data.get('error', 'unknown_error')}")
            return ok
        except (error.URLError, TimeoutError) as exc:
            self._log(f"Slack notify failed: {exc}")
            return False
        except Exception as exc:
            self._log(f"Slack notify failed: {type(exc).__name__}: {exc}")
            return False

    def send_image(
        self,
        content: bytes,
        filename: str,
        target: Optional[str] = None,
        title: Optional[str] = None,
    ) -> bool:
        """Upload image bytes and share them in a Slack conversation."""
        channel = (target or self._target or "").strip()
        filename = (filename or "image.png").strip()
        if not content or not self._token or not channel:
            return False

        try:
            ticket = self._post_form(
                self._GET_UPLOAD_URL,
                {"filename": filename, "length": str(len(content))},
            )
            if not ticket.get("ok"):
                self._log(f"Slack image upload failed: {ticket.get('error', 'unknown_error')}")
                return False

            upload_url = str(ticket.get("upload_url") or "")
            file_id = str(ticket.get("file_id") or "")
            if not upload_url or not file_id:
                self._log("Slack image upload failed: upload ticket was incomplete")
                return False

            upload_request = request.Request(
                upload_url,
                data=content,
                method="POST",
                headers={"Content-Type": "application/octet-stream"},
            )
            with request.urlopen(upload_request, timeout=self._timeout) as resp:
                resp.read()

            completed = self._post_form(
                self._COMPLETE_UPLOAD_URL,
                {
                    "files": json.dumps([{"id": file_id, "title": title or filename}]),
                    "channel_id": channel,
                },
            )
            if not completed.get("ok"):
                self._log(f"Slack image upload failed: {completed.get('error', 'unknown_error')}")
                return False
            return True
        except (error.URLError, TimeoutError) as exc:
            self._log(f"Slack image upload failed: {exc}")
            return False
        except Exception as exc:
            self._log(f"Slack image upload failed: {type(exc).__name__}: {exc}")
            return False

    def _post_form(self, url: str, fields: dict) -> dict:
        payload = urlencode(fields).encode("utf-8")
        req = request.Request(
            url,
            data=payload,
            method="POST",
            headers={
                "Authorization": f"Bearer {self._token}",
                "Content-Type": "application/x-www-form-urlencoded",
            },
        )
        with request.urlopen(req, timeout=self._timeout) as resp:
            body = resp.read().decode("utf-8", errors="replace")
        return json.loads(body)
