from types import SimpleNamespace

from core.slack_bot import SlackBotServer
from gui.tab_bayesian_optimization import BayesianOptimizationTab


class _Notifier:
    def __init__(self):
        self.images = []

    def send_image(self, content, filename, target=None, title=None):
        self.images.append((content, filename, target, title))
        return True


def test_explicit_status_request_sends_provided_image():
    notifier = _Notifier()
    server = SlackBotServer(
        host="127.0.0.1",
        port=0,
        signing_secret="secret",
        notifier=notifier,
        status_provider=lambda: {},
        status_image_provider=lambda: (b"png", "trend.png", "Q trend"),
    )

    server._send_status_image({"text": "<@BOT> status"}, "C123")

    assert notifier.images == [(b"png", "trend.png", "C123", "Q trend")]


def test_non_status_request_does_not_send_image():
    notifier = _Notifier()
    server = SlackBotServer(
        host="127.0.0.1",
        port=0,
        signing_secret="secret",
        notifier=notifier,
        status_provider=lambda: {},
        status_image_provider=lambda: (b"png", "trend.png", "Q trend"),
    )

    server._send_status_image({"text": "<@BOT> eta"}, "C123")

    assert notifier.images == []


def test_active_bo_loop_renders_q_trend_png():
    tab = BayesianOptimizationTab.__new__(BayesianOptimizationTab)
    tab._auto_running = True
    tab._paired_queue_running = False
    tab._bo_session = SimpleNamespace(
        observations=[
            {"group_id": 1, "group_name": "Group 1", "iteration": 1, "Q_run": 0.3},
            {"group_id": 1, "group_name": "Group 1", "iteration": 2, "Q_run": 0.7},
        ]
    )

    content, filename, title = tab.get_slack_q_trend_image()

    assert content.startswith(b"\x89PNG\r\n\x1a\n")
    assert filename == "bo_q_score_trend.png"
    assert title == "BO Q-score trend"
