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
    ensure_windivert_impostor_guard,
)


class _FakeButton:
    def __init__(self):
        self.options = {}

    def winfo_exists(self):
        return True

    def configure(self, **kwargs):
        self.options.update(kwargs)


class _Packet:
    def __init__(
        self,
        name,
        *,
        size=100,
        direction="outbound",
        src_addr="10.0.0.1",
        dst_addr="10.0.0.2",
        src_port=1000,
        dst_port=2000,
        protocol=6,
    ):
        self.name = name
        self.raw = b"x" * size
        self.direction = direction
        self.is_outbound = direction == "outbound"
        self.is_inbound = direction == "inbound"
        self.src_addr = src_addr
        self.dst_addr = dst_addr
        self.src_port = src_port
        self.dst_port = dst_port
        self.protocol = (protocol, 0)


class NetworkButtonStateTests(unittest.TestCase):
    @staticmethod
    def _wait_until(predicate, timeout=1.5):
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            if predicate():
                return True
            time.sleep(0.01)
        return bool(predicate())

    @staticmethod
    def _var(value):
        return SimpleNamespace(get=lambda: value)

    def _filter_app(self, protocol="全部", direction="全部", target_ip="", advanced=""):
        return SimpleNamespace(
            protocol_var=self._var(protocol),
            direction_var=self._var(direction),
            target_ip_var=self._var(target_ip),
            advanced_filter_var=self._var(advanced),
        )

    @staticmethod
    def _fake_pydivert(win_divert_type):
        win_divert_type.check_filter = staticmethod(lambda _filter: (True, 0, ""))
        return SimpleNamespace(WinDivert=win_divert_type)

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
        fake_module = self._fake_pydivert(FakeWinDivert)
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

    def test_default_and_advanced_filters_include_impostor_guard(self):
        default_filter = XiaoXinAssistant._build_filter(self._filter_app())
        self.assertEqual("!impostor", default_filter)

        advanced_filter = XiaoXinAssistant._build_filter(self._filter_app(advanced="tcp and outbound"))
        self.assertIn("!impostor", advanced_filter)
        self.assertIn("tcp and outbound", advanced_filter)
        self.assertEqual(1, advanced_filter.lower().count("!impostor"))
        self.assertEqual("!impostor and tcp", ensure_windivert_impostor_guard("!impostor and tcp"))

    def test_filter_builds_ipv4_ipv6_protocol_and_direction_combinations(self):
        ipv4_filter = XiaoXinAssistant._build_filter(
            self._filter_app(protocol="TCP", direction="出站", target_ip="10.0.2.35")
        )
        self.assertIn("tcp", ipv4_filter)
        self.assertIn("outbound", ipv4_filter)
        self.assertIn("ip.SrcAddr == 10.0.2.35", ipv4_filter)
        self.assertIn("!impostor", ipv4_filter)

        ipv6_filter = XiaoXinAssistant._build_filter(
            self._filter_app(protocol="UDP", direction="入站", target_ip="2001:db8::1")
        )
        self.assertIn("udp", ipv6_filter)
        self.assertIn("inbound", ipv6_filter)
        self.assertIn("ipv6.DstAddr == 2001:db8::1", ipv6_filter)

        icmp_filter = XiaoXinAssistant._build_filter(self._filter_app(protocol="ICMP"))
        self.assertIn("icmp", icmp_filter)

    def test_invalid_advanced_filter_is_rejected_before_start(self):
        class InvalidWinDivert:
            @staticmethod
            def check_filter(_filter):
                return False, 7, "unexpected token"

        with patch.dict(sys.modules, {"pydivert": SimpleNamespace(WinDivert=InvalidWinDivert)}):
            with self.assertRaisesRegex(ValueError, "位置 7"):
                NetworkWorker.validate_filter("tcp and (")

    def test_jitter_stays_in_range_and_never_becomes_negative(self):
        worker = NetworkWorker(lambda _text: None)
        worker.delay_ms = 100
        worker.jitter_ms = 80
        with patch("xiaoxin_assistant.random.uniform", side_effect=[-80, 80]):
            self.assertEqual(20, worker._effective_delay_ms())
            self.assertEqual(180, worker._effective_delay_ms())

        worker.delay_ms = 20
        worker.jitter_ms = 100
        with patch("xiaoxin_assistant.random.uniform", return_value=-100):
            self.assertEqual(0, worker._effective_delay_ms())

        worker.jitter_ms = 0
        self.assertEqual(20, worker._effective_delay_ms())

    def test_jitter_preserves_packet_order_within_the_same_flow(self):
        worker = NetworkWorker(lambda _text: None)
        worker._start_generation = 1
        worker.running.set()
        first = _Packet("first")
        second = _Packet("second")
        with patch.object(worker, "_effective_delay_ms", side_effect=[200, 0]):
            worker._schedule_packet(first, 1, now=100.0)
            worker._schedule_packet(second, 1, now=100.001)
        scheduled = [item[4].name for item in sorted(worker.waiting_heap)]
        self.assertEqual(["first", "second"], scheduled)

    def test_different_flows_are_scheduled_independently(self):
        worker = NetworkWorker(lambda _text: None)
        worker._start_generation = 1
        worker.running.set()
        slow_flow = _Packet("slow", dst_port=2000)
        fast_flow = _Packet("fast", dst_port=3000)
        with patch.object(worker, "_effective_delay_ms", side_effect=[200, 0]):
            worker._schedule_packet(slow_flow, 1, now=100.0)
            worker._schedule_packet(fast_flow, 1, now=100.001)
        scheduled = [item[4].name for item in sorted(worker.waiting_heap)]
        self.assertEqual(["fast", "slow"], scheduled)

    def test_explicit_reorder_can_reverse_queued_packets_and_counts_it(self):
        worker = NetworkWorker(lambda _text: None)
        worker._start_generation = 1
        worker.running.set()
        worker.delay_ms = 100
        worker.reorder_percent = 100
        first = _Packet("first")
        second = _Packet("second")
        worker._schedule_packet(first, 1, now=100.0)
        worker._schedule_packet(second, 1, now=100.001)
        scheduled = [item[4].name for item in sorted(worker.waiting_heap)]
        self.assertEqual(["second", "first"], scheduled)
        self.assertEqual(1, worker.stats_snapshot()["reordered"])

    def test_stale_flow_scheduling_state_is_cleaned(self):
        worker = NetworkWorker(lambda _text: None)
        worker._flow_last_send["stale"] = (1.0, 1.0)
        worker._flow_fair_finish["stale"] = (10.0, 1.0)
        worker._cleanup_flow_state_locked(1000.0, force=True)
        self.assertNotIn("stale", worker._flow_last_send)
        self.assertNotIn("stale", worker._flow_fair_finish)

    def test_upload_download_bandwidth_waits_are_independent(self):
        worker = NetworkWorker(lambda _text: None)
        worker.upload_bps = 8000
        worker.download_bps = 16000
        with worker.lock:
            worker._reset_bandwidth_state_locked(now=100.0)
            upload_wait = worker._bandwidth_wait_seconds_locked("outbound", 1000, 100.0)
            download_wait = worker._bandwidth_wait_seconds_locked("inbound", 1000, 100.0)
        self.assertAlmostEqual(1.0, upload_wait, places=3)
        self.assertAlmostEqual(0.5, download_wait, places=3)

    def test_zero_bandwidth_value_means_unlimited(self):
        worker = NetworkWorker(lambda _text: None)
        worker.upload_bps = 0
        worker.download_bps = 0
        with worker.lock:
            self.assertEqual(0, worker._bandwidth_wait_seconds_locked("outbound", 64000, 100.0))
            self.assertEqual(0, worker._bandwidth_wait_seconds_locked("inbound", 64000, 100.0))

    def test_fair_ready_queue_prevents_one_flow_from_monopolizing_direction(self):
        worker = NetworkWorker(lambda _text: None)
        worker._start_generation = 1
        worker.running.set()
        flow_a1 = _Packet("a1", dst_port=2000)
        flow_a2 = _Packet("a2", dst_port=2000)
        flow_b1 = _Packet("b1", dst_port=3000)
        for packet in (flow_a1, flow_a2, flow_b1):
            worker._enqueue_packet(packet, 100.0, 1, now=100.0)
        selected = []
        with worker.lock:
            worker._move_due_packets_locked(100.0)
            for _ in range(3):
                item, wait_seconds = worker._select_ready_packet_locked(100.0)
                self.assertEqual(0, wait_seconds)
                selected.append(item[0].name)
        self.assertEqual(["a1", "b1", "a2"], selected)

    def test_random_and_burst_loss_are_counted_separately(self):
        burst_worker = NetworkWorker(lambda _text: None)
        burst_worker.burst_trigger_percent = 100
        burst_worker.burst_length = 3
        self.assertTrue(burst_worker._should_drop_packet())
        self.assertTrue(burst_worker._should_drop_packet())
        self.assertTrue(burst_worker._should_drop_packet())
        burst_stats = burst_worker.stats_snapshot()
        self.assertEqual(3, burst_stats["burst_drop"])
        self.assertEqual(0, burst_stats["random_drop"])

        random_worker = NetworkWorker(lambda _text: None)
        random_worker.loss_percent = 100
        self.assertTrue(random_worker._should_drop_packet())
        random_stats = random_worker.stats_snapshot()
        self.assertEqual(1, random_stats["random_drop"])
        self.assertEqual(0, random_stats["burst_drop"])

    def test_delay_queue_enforces_packet_and_byte_capacity(self):
        class Packet:
            def __init__(self, size):
                self.raw = b"x" * size

        worker = NetworkWorker(lambda _text: None, max_queue_packets=2, max_queue_bytes=4096)
        worker._start_generation = 1
        worker.running.set()
        self.assertTrue(worker._enqueue_packet(Packet(100), time.monotonic() + 10, 1))
        self.assertTrue(worker._enqueue_packet(Packet(100), time.monotonic() + 10, 1))
        self.assertFalse(worker._enqueue_packet(Packet(100), time.monotonic() + 10, 1))
        self.assertEqual(1, worker.stats_snapshot()["queue_overflow"])

        byte_worker = NetworkWorker(lambda _text: None, max_queue_packets=10, max_queue_bytes=1024)
        byte_worker._start_generation = 1
        byte_worker.running.set()
        self.assertTrue(byte_worker._enqueue_packet(Packet(700), time.monotonic() + 10, 1))
        self.assertFalse(byte_worker._enqueue_packet(Packet(700), time.monotonic() + 10, 1))
        self.assertEqual(1, byte_worker.stats_snapshot()["queue_overflow"])

    def test_stop_clears_waiting_queue_and_counts_cancelled_packets(self):
        packet = SimpleNamespace(raw=b"x")
        worker = NetworkWorker(lambda _text: None, max_queue_packets=5)
        worker._start_generation = 1
        worker.running.set()
        worker._enqueue_packet(packet, time.monotonic() + 60, 1)
        worker._enqueue_packet(packet, time.monotonic() + 60, 1)
        worker.stop()
        stats = worker.stats_snapshot()
        self.assertEqual(0, stats["queue_count"])
        self.assertEqual(0, stats["queue_bytes"])
        self.assertEqual(2, stats["cancelled"])

    def test_one_hundred_percent_loss_drops_without_sending(self):
        closed = threading.Event()
        sent = []

        class FakeWinDivert:
            def __init__(self, _filter):
                self.first = True

            def open(self):
                pass

            def recv(self):
                if self.first:
                    self.first = False
                    return SimpleNamespace(raw=b"packet")
                closed.wait(timeout=1)
                raise RuntimeError("closed")

            def send(self, packet):
                sent.append(packet)

            def close(self):
                closed.set()

        worker = NetworkWorker(lambda _text: None)
        fake_module = self._fake_pydivert(FakeWinDivert)
        with patch("xiaoxin_assistant.is_admin", return_value=True), patch.dict(sys.modules, {"pydivert": fake_module}):
            worker.start(0, 100, "true")
            self.assertTrue(self._wait_until(lambda: worker.simulated_drop_count == 1))
            worker.stop()
        self.assertEqual([], sent)
        self.assertEqual(1, worker.stats_snapshot()["simulated_drop"])

    def test_send_failure_reports_runtime_failure(self):
        closed = threading.Event()
        states = []

        class FakeWinDivert:
            def __init__(self, _filter):
                self.first = True

            def open(self):
                pass

            def recv(self):
                if self.first:
                    self.first = False
                    return SimpleNamespace(raw=b"packet")
                closed.wait(timeout=1)
                raise RuntimeError("closed")

            def send(self, _packet):
                raise OSError("send failed")

            def close(self):
                closed.set()

        worker = NetworkWorker(lambda _text: None, lambda phase, message: states.append((phase, message)))
        fake_module = self._fake_pydivert(FakeWinDivert)
        with patch("xiaoxin_assistant.is_admin", return_value=True), patch.dict(sys.modules, {"pydivert": fake_module}):
            worker.start(0, 0, "true")
            self.assertTrue(self._wait_until(lambda: any(phase == TaskPhase.FAILED for phase, _ in states)))
        self.assertEqual(1, worker.stats_snapshot()["send_failure"])
        self.assertFalse(worker.running.is_set())
        self.assertTrue(
            self._wait_until(
                lambda: not worker.capture_thread.is_alive() and not worker.dispatch_thread.is_alive()
            )
        )
        self.assertIsNone(worker.handle)
        self.assertEqual(0, worker.stats_snapshot()["queue_count"])

    def test_runtime_receive_failure_reports_failure_not_normal_stop(self):
        states = []

        class FakeWinDivert:
            def __init__(self, _filter):
                pass

            def open(self):
                pass

            def recv(self):
                raise OSError("recv failed")

            def close(self):
                pass

        worker = NetworkWorker(lambda _text: None, lambda phase, message: states.append((phase, message)))
        fake_module = self._fake_pydivert(FakeWinDivert)
        with patch("xiaoxin_assistant.is_admin", return_value=True), patch.dict(sys.modules, {"pydivert": fake_module}):
            worker.start(0, 0, "true")
            self.assertTrue(self._wait_until(lambda: any(phase == TaskPhase.FAILED for phase, _ in states)))
        self.assertIn("接收数据包失败", states[-1][1])
        self.assertFalse(worker.running.is_set())
        self.assertTrue(
            self._wait_until(
                lambda: not worker.capture_thread.is_alive() and not worker.dispatch_thread.is_alive()
            )
        )
        self.assertIsNone(worker.handle)

    def test_startup_timeout_and_rapid_stop_release_worker(self):
        release_open = threading.Event()
        states = []

        class SlowWinDivert:
            def __init__(self, _filter):
                pass

            def open(self):
                release_open.wait(timeout=1)

            def recv(self):
                raise RuntimeError("closed")

            def close(self):
                release_open.set()

        worker = NetworkWorker(lambda _text: None, lambda phase, message: states.append((phase, message)))
        worker.startup_timeout_seconds = 0.05
        fake_module = self._fake_pydivert(SlowWinDivert)
        with patch("xiaoxin_assistant.is_admin", return_value=True), patch.dict(sys.modules, {"pydivert": fake_module}):
            worker.start(10, 0, "true")
            self.assertTrue(self._wait_until(lambda: any(phase == TaskPhase.FAILED for phase, _ in states)))
            self.assertFalse(worker.running.is_set())

            states.clear()
            release_open.clear()
            worker.startup_timeout_seconds = 1
            worker.start(10, 0, "true")
            worker.stop()
            self.assertFalse(worker.running.is_set())
            self.assertEqual(0, worker.stats_snapshot()["queue_count"])

    def test_auto_stop_restores_network_and_releases_threads(self):
        closed = threading.Event()
        states = []

        class FakeWinDivert:
            def __init__(self, _filter):
                pass

            def open(self):
                pass

            def recv(self):
                closed.wait(timeout=1)
                raise RuntimeError("closed")

            def close(self):
                closed.set()

        worker = NetworkWorker(lambda _text: None, lambda phase, message: states.append((phase, message)))
        fake_module = self._fake_pydivert(FakeWinDivert)
        with patch("xiaoxin_assistant.is_admin", return_value=True), patch.dict(sys.modules, {"pydivert": fake_module}):
            worker.start(0, 0, "true", auto_stop_seconds=0.05)
            self.assertTrue(
                self._wait_until(
                    lambda: any(phase == TaskPhase.IDLE and "自动停止" in message for phase, message in states)
                )
            )
        self.assertFalse(worker.running.is_set())
        self.assertEqual(0, worker.stats_snapshot()["queue_count"])

    def test_emergency_restore_stops_worker_immediately(self):
        running = threading.Event()
        running.set()
        stop_messages = []
        logs = []
        phases = []

        def stop(message):
            stop_messages.append(message)
            running.clear()

        app = SimpleNamespace(
            network=SimpleNamespace(running=running, stop=stop),
            _network_log=logs.append,
            _set_task_phase=lambda name, phase, message: phases.append((name, phase, message)),
            _refresh_network_button=lambda: None,
        )
        XiaoXinAssistant._emergency_restore_network(app)
        self.assertFalse(running.is_set())
        self.assertEqual(["已立即恢复正常网络"], stop_messages)
        self.assertIn("立即恢复", logs[0])
        self.assertEqual(TaskPhase.STOPPING, phases[0][1])

    def test_parse_network_config_converts_bandwidth_units_and_duration(self):
        values = {
            "delay_var": "100",
            "jitter_var": "20",
            "loss_var": "2",
            "reorder_var": "0",
            "upload_bandwidth_var": "512",
            "download_bandwidth_var": "1024",
            "bandwidth_unit_var": "Kbps",
            "burst_trigger_var": "5",
            "burst_length_var": "4",
            "network_duration_var": "1.5",
        }
        app = SimpleNamespace(**{name: self._var(value) for name, value in values.items()})
        app._build_filter = lambda: "!impostor"
        with patch.object(NetworkWorker, "validate_filter", return_value=True):
            config = XiaoXinAssistant._parse_network_config(app)
        self.assertEqual(512000, config["upload_bps"])
        self.assertEqual(1024000, config["download_bps"])
        self.assertEqual(90, config["duration_seconds"])

    def test_parse_network_config_rejects_half_configured_burst_loss(self):
        values = {
            "delay_var": "100",
            "jitter_var": "0",
            "loss_var": "0",
            "reorder_var": "0",
            "upload_bandwidth_var": "0",
            "download_bandwidth_var": "0",
            "bandwidth_unit_var": "Mbps",
            "burst_trigger_var": "5",
            "burst_length_var": "0",
            "network_duration_var": "0",
        }
        app = SimpleNamespace(**{name: self._var(value) for name, value in values.items()})
        app._build_filter = lambda: "!impostor"
        with self.assertRaisesRegex(ValueError, "都必须大于 0"):
            XiaoXinAssistant._parse_network_config(app)


if __name__ == "__main__":
    unittest.main()
