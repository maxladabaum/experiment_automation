"""
core/slack_notifier.py - Minimal Slack notifier (chat.postMessage).
"""

import json
from typing import Callable, Optional
from urllib import error, request


class SlackNotifier:
    """Send plain-text Slack notifications with a bot token."""

    _POST_MESSAGE_URL = "https://slack.com/api/chat.postMessage"

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
