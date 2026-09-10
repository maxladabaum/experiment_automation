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
        station_name: str = "",
    ):
        self._token = (bot_token or "").strip()
        self._target = (default_target or "").strip()
        self._log = log_callback
        self._timeout = float(timeout_seconds)
        self._station_name = (station_name or "").strip()
        self.last_image_error = ""

    @property
    def enabled(self) -> bool:
        return bool(self._token and self._target)

    def send_message(self, text: str, target: Optional[str] = None) -> bool:
        """Post a message to Slack and return True on success."""
        channel = (target or self._target or "").strip()
        if not text or not self._token or not channel:
            return False

        if self._station_name:
            text = f"[{self._station_name}]\n{text}"
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
        comment: Optional[str] = None,
    ) -> bool:
        """Upload image bytes and share them in a Slack conversation."""
        self.last_image_error = ""
        channel = (target or self._target or "").strip()
        filename = (filename or "image.png").strip()
        if not content or not self._token or not channel:
            return self._image_upload_failed("Missing image, Slack token, or destination channel.")

        image_title = title or filename
        if self._station_name:
            # File titles are plain text, so omit Slack bold markers.
            plain_station = self._station_name.replace("*", "")
            image_title = f"[{plain_station}] {image_title}"

        stage = "request upload URL"
        try:
            ticket = self._post_form(
                self._GET_UPLOAD_URL,
                {"filename": filename, "length": str(len(content))},
            )
            if not ticket.get("ok"):
                return self._image_upload_failed(ticket.get("error", "unknown_error"), stage)

            upload_url = str(ticket.get("upload_url") or "")
            file_id = str(ticket.get("file_id") or "")
            if not upload_url or not file_id:
                return self._image_upload_failed("upload ticket was incomplete", stage)

            stage = "transfer image"
            upload_request = request.Request(
                upload_url,
                data=content,
                method="POST",
                headers={"Content-Type": "application/octet-stream"},
            )
            with request.urlopen(upload_request, timeout=self._timeout) as resp:
                resp.read()

            fields = {
                "files": json.dumps([{"id": file_id, "title": image_title}]),
                "channel_id": channel,
            }
            if comment:
                fields["initial_comment"] = (
                    f"[{self._station_name}]\n{comment}" if self._station_name else comment
                )
            stage = "share image in channel"
            completed = self._post_form(self._COMPLETE_UPLOAD_URL, fields)
            if not completed.get("ok"):
                return self._image_upload_failed(completed.get("error", "unknown_error"), stage)
            self._log(f"Slack image upload accepted: file {file_id}, channel {channel}")
            return True
        except (error.URLError, TimeoutError) as exc:
            return self._image_upload_failed(type(exc).__name__, stage)
        except Exception as exc:
            return self._image_upload_failed(type(exc).__name__, stage)

    def _image_upload_failed(self, reason: str, stage: str = "") -> bool:
        hints = {
            "missing_scope": "Add the files:write bot scope in Slack OAuth & Permissions, then reinstall the app to the workspace.",
            "not_in_channel": "Invite the Slack bot to the destination channel.",
            "no_permission": "Check that the Slack bot is a member of the destination channel and can upload files.",
            "channel_not_found": "Set EA_SLACK_TARGET to the channel ID, not its name, and check the bot has access.",
        }
        self.last_image_error = f"{stage}: {reason}" if stage else reason
        if reason in hints:
            self.last_image_error += "\n" + hints[reason]
        self._log(f"Slack image upload failed: {self.last_image_error}")
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
