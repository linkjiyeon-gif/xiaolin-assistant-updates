from __future__ import annotations

import inspect
import unittest
from types import SimpleNamespace

from app.task_state import RuntimeTaskState, TaskPhase
from xiaoxin_assistant import NETWORK_USAGE_GUIDE, XiaoXinAssistant


class _FakeWidget:
    def __init__(self):
        self.options = {}

    def winfo_exists(self):
        return True

    def configure(self, **kwargs):
        self.options.update(kwargs)


class BasicUISmokeTests(unittest.TestCase):
    def _network_app(self, phase):
        state = RuntimeTaskState("network")
        state.set(phase)
        button = _FakeWidget()
        emergency_button = _FakeWidget()
        controls = [_FakeWidget(), _FakeWidget()]
        app = SimpleNamespace(
            task_states={"network": state},
            network=SimpleNamespace(running=SimpleNamespace(is_set=lambda: phase == TaskPhase.RUNNING)),
            network_button=button,
            network_emergency_button=emergency_button,
            network_parameter_widgets=controls,
        )
        app._task_snapshot = lambda name: app.task_states[name].snapshot()
        app._set_network_controls_locked = lambda locked: XiaoXinAssistant._set_network_controls_locked(app, locked)
        return app, button, controls

    def test_network_starting_and_running_states_lock_parameters(self):
        app, button, controls = self._network_app(TaskPhase.STARTING)
        XiaoXinAssistant._refresh_network_button(app)
        self.assertEqual("…  正在启动", button.options["text"])
        self.assertEqual("disabled", button.options["state"])
        self.assertEqual("normal", app.network_emergency_button.options["state"])
        self.assertTrue(all(control.options["state"] == "disabled" for control in controls))

        app, button, controls = self._network_app(TaskPhase.RUNNING)
        XiaoXinAssistant._refresh_network_button(app)
        self.assertEqual("■  停止模拟", button.options["text"])
        self.assertEqual("normal", button.options["state"])
        self.assertEqual("normal", app.network_emergency_button.options["state"])
        self.assertTrue(all(control.options["state"] == "disabled" for control in controls))

    def test_network_failure_unlocks_parameters(self):
        app, button, controls = self._network_app(TaskPhase.FAILED)
        XiaoXinAssistant._refresh_network_button(app)
        self.assertEqual("▶  开始模拟", button.options["text"])
        self.assertEqual("normal", button.options["state"])
        self.assertEqual("disabled", app.network_emergency_button.options["state"])
        self.assertTrue(all(control.options["state"] == "normal" for control in controls))

    def test_all_runtime_pages_have_a_refresh_target(self):
        source = XiaoXinAssistant._refresh_current_task_ui.__code__.co_consts
        flattened = " ".join(str(value) for value in source)
        for page in (
            "clicker",
            "network",
            "adb_log",
            "adb_tools",
            "ios_log",
            "compare",
            "localization",
            "config_validator",
        ):
            self.assertIn(page, flattened)

    def test_configuration_change_notification_is_not_exposed(self):
        sidebar_source = inspect.getsource(XiaoXinAssistant._build_sidebar)
        page_source = inspect.getsource(XiaoXinAssistant.show_page)
        startup_source = inspect.getsource(XiaoXinAssistant._finish_background_startup)
        for source in (sidebar_source, page_source, startup_source):
            self.assertNotIn("gitlab_monitor", source)
            self.assertNotIn("配置变更通知", source)

    def test_network_usage_guide_covers_the_complete_basic_workflow(self):
        for text in (
            "管理员重启",
            "目标 IP",
            "单向延迟",
            "抖动",
            "丢包",
            "开始模拟",
            "停止模拟",
            "非 Root 真机",
        ):
            self.assertIn(text, NETWORK_USAGE_GUIDE)


if __name__ == "__main__":
    unittest.main()
