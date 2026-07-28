from __future__ import annotations

import threading
import sys
import time
import unittest
from types import SimpleNamespace
from unittest.mock import patch

from app.task_state import TaskPhase
from xiaoxin_assistant import (
    COLOR_ACCENT,
    COLOR_ACCENT_HOVER,
    COLOR_DANGER,
    COLOR_DANGER_HOVER,
    NetworkWorker,
    XiaoXinAssistant,
)


class _FakeButton:
    def __init__(self):
        self.options = {}

    def winfo_exists(self):
        return True

    def configure(self, **kwargs):
        self.options.update(kwargs)


class NetworkButtonStateTests(unittest.TestCase):
    def test_button_switches_between_start_and_stop_states(self):
        running = threading.Event()
        button = _FakeButton()
        app = SimpleNamespace(
            network=SimpleNamespace(running=running),
            network_button=button,
        )

        XiaoXinAssistant._refresh_network_button(app)
        self.assertEqual("▶  开始模拟", button.options["text"])
        self.assertEqual(COLOR_ACCENT, button.options["fg_color"])
        self.assertEqual(COLOR_ACCENT_HOVER, button.options["hover_color"])

        running.set()
        XiaoXinAssistant._refresh_network_button(app)
        self.assertEqual("■  停止模拟", button.options["text"])
        self.assertEqual(COLOR_DANGER, button.options["fg_color"])
        self.assertEqual(COLOR_DANGER_HOVER, button.options["hover_color"])

        running.clear()
        XiaoXinAssistant._refresh_network_button(app)
        self.assertEqual("▶  开始模拟", button.options["text"])

    def test_worker_reports_running_only_after_windivert_open_confirms(self):
        release_open = threading.Event()
        receive_block = threading.Event()
        states = []

        class FakeWinDivert:
            def __init__(self, _filter):
                pass

            def open(self):
                release_open.wait(timeout=1)

            def recv(self):
                receive_block.wait(timeout=1)
                raise RuntimeError("closed")

            def close(self):
                receive_block.set()

            def send(self, _packet):
                pass

        worker = NetworkWorker(lambda _text: None, lambda phase, message: states.append((phase, message)))
        fake_module = SimpleNamespace(WinDivert=FakeWinDivert)
        with patch("xiaoxin_assistant.is_admin", return_value=True), patch.dict(sys.modules, {"pydivert": fake_module}):
            worker.start(300, 2, "true")
            self.assertEqual(TaskPhase.STARTING, states[-1][0])
            self.assertFalse(any(phase == TaskPhase.RUNNING for phase, _ in states))

            release_open.set()
            deadline = time.monotonic() + 1
            while time.monotonic() < deadline and not any(phase == TaskPhase.RUNNING for phase, _ in states):
                time.sleep(0.01)
            self.assertTrue(any(phase == TaskPhase.RUNNING for phase, _ in states))
            worker.stop()


if __name__ == "__main__":
    unittest.main()
