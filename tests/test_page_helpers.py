from __future__ import annotations

import os
import tempfile
import threading
import unittest
from collections import deque
from types import SimpleNamespace
from unittest.mock import patch

import xiaoxin_assistant as xa
from xiaoxin_assistant import ADB_LIVE_PREVIEW_LINES, RECENT_PACKAGES_MAX, XiaoXinAssistant


class RecentPackagesTests(unittest.TestCase):
    def test_remember_moves_package_to_front_and_caps_list(self):
        with tempfile.TemporaryDirectory() as tmp:
            with patch.object(xa, "app_data_path", lambda rel: os.path.join(tmp, rel)):
                app = SimpleNamespace()
                app._load_recent_packages = lambda: XiaoXinAssistant._load_recent_packages(app)
                self.assertEqual([], app._load_recent_packages())
                for i in range(RECENT_PACKAGES_MAX + 3):
                    XiaoXinAssistant._remember_package(app, f"com.demo.app{i}")
                XiaoXinAssistant._remember_package(app, "com.demo.app5")
                XiaoXinAssistant._remember_package(app, "  ")
                items = app._load_recent_packages()
                self.assertEqual(RECENT_PACKAGES_MAX, len(items))
                self.assertEqual("com.demo.app5", items[0])
                self.assertEqual(1, items.count("com.demo.app5"))


class LivePreviewQueueTests(unittest.TestCase):
    def test_queue_keeps_newest_lines_and_schedules_one_flush(self):
        scheduled = []
        app = SimpleNamespace(
            _adb_live_lock=threading.Lock(),
            _adb_live_pending=deque(maxlen=ADB_LIVE_PREVIEW_LINES),
            _adb_live_flush_scheduled=False,
            _flush_adb_live_preview=lambda: None,
            after=lambda ms, fn: scheduled.append(ms),
        )
        for i in range(ADB_LIVE_PREVIEW_LINES + 50):
            XiaoXinAssistant._queue_adb_live_preview_line(app, f"line {i}\r\n")
        self.assertEqual(1, len(scheduled))
        self.assertEqual(ADB_LIVE_PREVIEW_LINES, len(app._adb_live_pending))
        self.assertEqual(f"line {ADB_LIVE_PREVIEW_LINES + 49}", app._adb_live_pending[-1])

    def test_flush_waits_while_page_is_hidden(self):
        app = SimpleNamespace(
            current_page="home",
            _adb_live_lock=threading.Lock(),
            _adb_live_pending=deque(["a", "b"], maxlen=ADB_LIVE_PREVIEW_LINES),
            _adb_live_flush_scheduled=True,
        )
        XiaoXinAssistant._flush_adb_live_preview(app)
        self.assertFalse(app._adb_live_flush_scheduled)
        self.assertEqual(["a", "b"], list(app._adb_live_pending))


if __name__ == "__main__":
    unittest.main()
