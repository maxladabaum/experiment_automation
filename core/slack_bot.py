"""
core/slack_bot.py - Minimal Slack Events API listener for status pings.
"""

import hashlib
import hmac
import json
import threading
import time
from http.server import BaseHTTPRequestHandler, HTTPServer
from typing import Callable, Dict, Optional

from .slack_notifier import SlackNotifier


class SlackBotServer:
    """Listen for Slack Events API app_mention and reply with queue status."""

    def __init__(
        self,
        host: str,
        port: int,
        signing_secret: str,
        notifier: SlackNotifier,
        status_provider: Callable[[], Dict[str, Optional[str]]],
        log_callback: Callable[[str], None] = print,
    ):
        self._host = host
        self._port = int(port)
        self._signing_secret = (signing_secret or "").strip()
        self._notifier = notifier
        self._status_provider = status_provider
        self._log = log_callback
        self._server: Optional[HTTPServer] = None
        self._thread: Optional[threading.Thread] = None

    def start(self):
        if not self._signing_secret:
            self._log("Slack bot not started: missing signing secret.")
            return
        self._server = HTTPServer((self._host, self._port), self._make_handler())
        self._thread = threading.Thread(target=self._server.serve_forever, daemon=True)
        self._thread.start()
        self._log(f"Slack bot listening on http://{self._host}:{self._port}/slack/events")

    def stop(self):
        if self._server is not None:
            self._server.shutdown()
            self._server.server_close()
            self._server = None

    def _make_handler(self):
        parent = self

        class Handler(BaseHTTPRequestHandler):
            def do_POST(self):  # noqa: N802
                if self.path != "/slack/events":
                    self.send_response(404)
                    self.end_headers()
                    return

                length = int(self.headers.get("Content-Length", "0") or "0")
                body = self.rfile.read(length) if length > 0 else b""
                if not parent._verify_signature(self.headers, body):
                    self.send_response(403)
                    self.end_headers()
                    return

                try:
                    payload = json.loads(body.decode("utf-8"))
                except Exception:
                    self.send_response(400)
                    self.end_headers()
                    return

                if payload.get("type") == "url_verification":
                    challenge = payload.get("challenge", "")
                    self.send_response(200)
                    self.send_header("Content-Type", "text/plain")
                    self.end_headers()
                    self.wfile.write(challenge.encode("utf-8"))
                    return

                if payload.get("type") == "event_callback":
                    event = payload.get("event") or {}
                    if event.get("type") == "app_mention":
                        channel = event.get("channel")
                        if channel:
                            text = parent._format_status_text()
                            parent._notifier.send_message(text, target=channel)

                self.send_response(200)
                self.end_headers()

            def log_message(self, fmt, *args):  # noqa: N802
                return

        return Handler

    def _verify_signature(self, headers, body: bytes) -> bool:
        timestamp = headers.get("X-Slack-Request-Timestamp", "")
        signature = headers.get("X-Slack-Signature", "")
        if not timestamp or not signature:
            return False
        try:
            ts_int = int(timestamp)
        except ValueError:
            return False
        if abs(time.time() - ts_int) > 60 * 5:
            return False
        base = f"v0:{timestamp}:{body.decode('utf-8')}".encode("utf-8")
        digest = hmac.new(
            self._signing_secret.encode("utf-8"),
            base,
            hashlib.sha256,
        ).hexdigest()
        expected = f"v0={digest}"
        return hmac.compare_digest(expected, signature)

    def _format_status_text(self) -> str:
        status = self._status_provider() or {}
        state = (status.get("state") or "unknown").upper()
        idx = status.get("current_index") or "-"
        total = status.get("total") or "-"
        label = status.get("current_label") or "(none)"
        session_name = status.get("session_name") or "(none)"
        experiment_name = status.get("experiment_name") or "(none)"
        updated_at = status.get("updated_at") or "(unknown)"
        return (
            f"Queue status: {state} | step {idx}/{total} | {label}\n"
            f"Session: {session_name} | Experiment: {experiment_name}\n"
            f"Updated: {updated_at}"
        )
