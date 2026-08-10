import ctypes
from collections import deque
from dataclasses import dataclass
import heapq
import ipaddress
import json
import itertools
import os
import random
import sys
import threading
import time
import tkinter as tk
import subprocess
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
from PIL import Image

try:
    import pystray
except Exception:
    pystray = None

from app.comparer import CompareOptions, ContentComparer, DiffItem
from app.constants import APP_NAME, APP_SUBTITLE, APP_VERSION, COMPARE_MODES, SUPPORTED_EXTENSIONS
from app.excel_exporter import export_diffs_to_excel, export_value_compare_diffs_to_excel
from app.adb_tools import ADBTools
from app.task_state import RuntimeTaskState, TaskPhase
from app.apk_info import analyze_apk, format_apk_info
from app.file_readers import FileReadError, get_file_info, read_file
from app.normalizer import NormalizeOptions, clean_ignore_fields
from app.translation_compare import DEFAULT_LANGUAGE_MAPPING_TEXT, TranslationCompareOptions, safe_int

from app.value_config_compare import (
    VALUE_COMPARE_MODE,
    ValueCompareOptions,
    build_value_compare_rule_preview,
    compare_value_reference_to_config,
    prepare_value_compare_options,
)
from app.admin_utils import is_windows_admin, relaunch_current_process_as_admin
from app.ios_log_tools import (
    DEFAULT_IOS_KEYWORDS,
    IOSLogFilter,
    IOSLogTools,
    format_devices,
    parse_keywords,
)
from app.localization_checker import (
    LocalizationCheckOptions,
    LocalizationColumn,
    check_localization_file,
    detect_dictionary_language_files,
    detect_language_columns,
    detect_multi_table_sheet_columns,
    parse_localization_table,
    auto_detect_source_language,
    localization_issues_to_tsv,
    export_localization_issues_to_excel,
    build_single_sheet_preview,
    recommend_target_languages_from_filename,
)
from app.localization_sheet_parser import (
    MULTI_FILE_DICTIONARY_STRUCTURE,
    MULTI_TABLE_SHEETS_STRUCTURE,
    MULTI_SHEET_STRUCTURE,
    SINGLE_SHEET_STRUCTURE,
    UNKNOWN_STRUCTURE,
    build_multi_sheet_preview,
    detect_localization_structure,
    normalize_localization_mode,
    parse_multi_sheet_localization,
)


from app.config_rule_templates import QUICK_TEMPLATES, TEMPLATE_HELP
from app.config_validator import (
    COMPOUND_FIELD_EXAMPLES,
    ConfigValidationOptions,
    detect_config_structure,
    find_reward_like_field,
    is_key_index_field,
    is_reward_like_field,
    parse_config_table,
    validate_config_table,
    validation_issues_to_tsv,
    export_validation_issues_to_excel,
    options_to_dict as config_options_to_dict,
    options_from_dict as config_options_from_dict,
)

from app.log_monitor import (
    DEFAULT_KEYWORDS,
    LogMonitor,
    export_keyword_hits_to_excel,
    export_keyword_hits_to_txt,
    format_keyword_hits,
    keyword_hits_to_tsv,
)
from app.crash_analyzer import (
    CrashAnalyzer,
    CrashStreamDetector,
    crash_records_to_tsv,
    export_crash_records_to_excel,
    format_crash_records,
)

from app.time_tools import (
    add_or_subtract_time_parts,
    boundary_times,
    countdown,
    convert_time_units,
    datetime_to_timestamps,
    format_datetime,
    readable_duration,
    time_difference,
    timestamp_pair,
    timestamp_to_datetime,
)
from app.update_manager import (
    DownloadCancelled,
    UpdateError,
    download_installer,
    fetch_update_info,
    is_newer_version,
    launch_update_helper,
)

Button = None
MouseController = None
keyboard = None
_pynput_load_attempted = False
_pynput_load_lock = threading.Lock()


def _ensure_pynput_loaded():
    """Import the global hotkey dependency only when the clicker is initialized."""
    global Button, MouseController, keyboard, _pynput_load_attempted
    if _pynput_load_attempted:
        return keyboard is not None and MouseController is not None
    with _pynput_load_lock:
        if _pynput_load_attempted:
            return keyboard is not None and MouseController is not None
        try:
            from pynput.mouse import Button as PynputButton, Controller as PynputMouseController
            from pynput import keyboard as pynput_keyboard

            Button = PynputButton
            MouseController = PynputMouseController
            keyboard = pynput_keyboard
        except Exception:
            Button = None
            MouseController = None
            keyboard = None
        _pynput_load_attempted = True
    return keyboard is not None and MouseController is not None


DEFAULT_PKG_NAME = "com.tinywarsurvivalexpress.android"
ADB_CMD_TIMEOUT = 10

NETWORK_USAGE_GUIDE = """使用说明
1. 权限准备：先点击左下角“管理员重启”，右上角显示“已就绪”后再使用。
2. 选择范围：目标 IP 建议填写要测试的服务器 IP；留空会影响本机全部匹配流量，启动前会再次确认。
3. 设置参数：协议和方向不确定时保持“全部”；延迟是单向延迟，抖动表示延迟上下浮动，丢包填 0～100。
4. 进阶模拟：带宽填 0 表示不限速；突发丢包用于模拟连续断流；乱序默认关闭，仅在专项测试时开启。
5. 快速开始：可先选择“弱网测试”或“移动网络”预设，再点击“开始模拟”。上下行同时生效时，额外 RTT 约为基础延迟的两倍。
6. 查看与结束：运行中观察速度、丢包、队列和失败统计；测试完成后点击“停止模拟”，紧急情况可点“立即恢复网络”。
示例：基础延迟 300ms、抖动 50ms，表示每个方向约延迟 250～350ms；上下行都生效时额外 RTT 通常约 600ms。
注意：当前仅影响这台电脑以及通过电脑联网的模拟器流量，不支持非 Root 真机直连弱网；高级规则仅建议熟悉 WinDivert 的用户填写。
"""

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


def resource_path(relative_path: str) -> str:
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def get_resource_path(*relative_paths: str) -> str:
    """Return the first existing bundled resource path, compatible with source and PyInstaller modes."""
    for relative_path in relative_paths:
        path = resource_path(relative_path)
        if os.path.exists(path):
            return path
    return resource_path(relative_paths[0]) if relative_paths else ""


def get_icon_path(kind: str = "ico") -> str:
    if kind.lower() == "png":
        return get_resource_path("assets/app.png")
    return get_resource_path("assets/app.ico")


def app_data_path(relative_path: str) -> str:
    if getattr(sys, "frozen", False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


def is_admin() -> bool:
    return is_windows_admin()


def relaunch_as_admin():
    result = relaunch_current_process_as_admin()
    if result.ok:
        try:
            messagebox.showinfo("管理员重启", "已请求管理员权限重启，当前窗口即将退出。")
        except Exception:
            pass
        sys.exit(0)
    if result.message == "当前已是管理员权限":
        messagebox.showinfo("管理员权限", result.message)
    else:
        messagebox.showwarning("管理员重启", result.message)


# ---------------- Mouse Auto Clicker ----------------
class ClickerWorker:
    def __init__(self):
        self.mouse = None
        self.running = threading.Event()
        self.thread = None
        self.hotkey_listener = None
        self.hotkey_getter = None
        self.toggle_callback = None
        self._last_hotkey_at = 0.0

    def start(self, interval: float, button_name: str):
        if self.running.is_set():
            return
        if self.mouse is None and _ensure_pynput_loaded():
            self.mouse = MouseController()
        if self.mouse is None:
            raise RuntimeError("pynput 未安装，无法使用连点功能")

        button = {
            "left": Button.left,
            "middle": Button.middle,
            "right": Button.right
        }.get(button_name, Button.left)

        interval = max(0.01, float(interval))
        self.running.set()
        self.thread = threading.Thread(target=self._loop, args=(interval, button), daemon=True)
        self.thread.start()

    def stop(self):
        self.running.clear()

    def _loop(self, interval, button):
        while self.running.is_set():
            try:
                self.mouse.click(button)
            except Exception:
                self.running.clear()
                break
            time.sleep(interval)

    def setup_hotkey(self, hotkey_getter, toggle_callback):
        if not _ensure_pynput_loaded():
            return
        if self.mouse is None:
            self.mouse = MouseController()
        self.hotkey_getter = hotkey_getter
        self.toggle_callback = toggle_callback

        def on_press(key):
            mapping = {
                "F6": keyboard.Key.f6, "F7": keyboard.Key.f7, "F8": keyboard.Key.f8,
                "F9": keyboard.Key.f9, "F10": keyboard.Key.f10, "F11": keyboard.Key.f11,
                "F12": keyboard.Key.f12,
            }
            target = mapping.get(self.hotkey_getter(), keyboard.Key.f8)
            now = time.monotonic()
            if key == target and self.toggle_callback and now - self._last_hotkey_at >= 0.35:
                self._last_hotkey_at = now
                self.toggle_callback()

        self.hotkey_listener = keyboard.Listener(on_press=on_press)
        self.hotkey_listener.daemon = True
        self.hotkey_listener.start()

    def close(self):
        self.stop()
        try:
            if self.hotkey_listener:
                self.hotkey_listener.stop()
        except Exception:
            pass


# ---------------- Network Delay / Loss ----------------
def ensure_windivert_impostor_guard(filter_text: str) -> str:
    """Protect WinDivert rules from capturing reinjected packets again."""
    cleaned = str(filter_text or "").strip() or "true"
    compact = "".join(cleaned.lower().split())
    if "!impostor" in compact or "notimpostor" in compact:
        return cleaned
    if cleaned.lower() == "true":
        return "!impostor"
    return f"(!impostor) and ({cleaned})"


class NetworkWorker:
    DEFAULT_MAX_QUEUE_PACKETS = 10000
    DEFAULT_MAX_QUEUE_BYTES = 64 * 1024 * 1024
    FLOW_STATE_TTL_SECONDS = 120.0
    FLOW_CLEANUP_INTERVAL_SECONDS = 30.0
    RATE_WINDOW_SECONDS = 1.0

    def __init__(
        self,
        log_callback,
        state_callback=None,
        max_queue_packets=None,
        max_queue_bytes=None,
    ):
        self.log_callback = log_callback
        self.state_callback = state_callback
        self.running = threading.Event()
        self.startup_event = threading.Event()
        self.startup_error = ""
        self.startup_timeout_seconds = 8
        self._start_generation = 0
        self.capture_thread = None
        self.dispatch_thread = None
        self.handle = None

        self.lock = threading.RLock()
        self.cond = threading.Condition(self.lock)
        self.waiting_heap = []
        self.ready_heaps = {"outbound": [], "inbound": [], "unknown": []}
        # Compatibility alias for older diagnostics that inspected the delay heap.
        self.heap = self.waiting_heap
        self.packet_seq = itertools.count()
        self.queue_bytes = 0
        self.max_queue_packets = max(1, int(max_queue_packets or self.DEFAULT_MAX_QUEUE_PACKETS))
        self.max_queue_bytes = max(1024, int(max_queue_bytes or self.DEFAULT_MAX_QUEUE_BYTES))

        self.delay_ms = 100
        self.jitter_ms = 0
        self.loss_percent = 0
        self.reorder_percent = 0
        self.burst_trigger_percent = 0
        self.burst_length = 0
        self.burst_remaining = 0
        self.upload_bps = 0.0
        self.download_bps = 0.0
        self.auto_stop_seconds = 0.0
        self.auto_stop_deadline = 0.0
        self.filter_text = "!impostor"

        self._flow_last_send = {}
        self._flow_fair_finish = {}
        self._direction_virtual_finish = {"outbound": 0.0, "inbound": 0.0, "unknown": 0.0}
        self._last_flow_cleanup_at = 0.0
        self._direction_cursor = 0
        self._tokens = {"outbound": 0.0, "inbound": 0.0, "unknown": 0.0}
        self._token_updated_at = {"outbound": 0.0, "inbound": 0.0, "unknown": 0.0}
        self._bandwidth_waiting_sequences = set()
        self._rate_events = {
            "outbound": deque(maxlen=20),
            "inbound": deque(maxlen=20),
            "unknown": deque(maxlen=20),
        }

        self.captured_count = 0
        self.dropped_count = 0
        self.simulated_drop_count = 0
        self.random_drop_count = 0
        self.burst_drop_count = 0
        self.reordered_count = 0
        self.bandwidth_wait_count = 0
        self.queue_overflow_count = 0
        self.send_failure_count = 0
        self.cancelled_count = 0
        self.sent_count = 0
        self.sent_bytes = {"outbound": 0, "inbound": 0, "unknown": 0}
        self.peak_queue_count = 0
        self.peak_queue_bytes = 0
        self.started_at = 0.0
        self.elapsed_at_stop = 0.0
        self._last_overflow_log_at = 0.0
        self.auto_stop_thread = None

    def _emit_state(self, phase, message=""):
        callback = self.state_callback
        if callback is None:
            return
        try:
            callback(phase, message)
        except Exception:
            pass

    @staticmethod
    def validate_filter(filter_text: str):
        try:
            import pydivert
        except Exception as exc:
            raise RuntimeError("未安装 pydivert / WinDivert，请先运行 run.bat 或 build_exe.bat 安装依赖") from exc
        try:
            valid, position, error = pydivert.WinDivert.check_filter(str(filter_text or "true"))
        except Exception as exc:
            raise ValueError(f"WinDivert 规则校验失败：{exc}") from exc
        if not valid:
            position_text = f"，位置 {position}" if position is not None else ""
            raise ValueError(f"高级规则语法无效{position_text}：{error or '无法解析过滤表达式'}")
        return True

    @staticmethod
    def _packet_size(packet) -> int:
        raw = getattr(packet, "raw", None)
        if raw is not None:
            try:
                return max(0, len(raw))
            except Exception:
                pass
        try:
            return max(0, len(packet))
        except Exception:
            return 0

    def _is_active(self, generation) -> bool:
        return generation == self._start_generation and self.running.is_set()

    def _queue_count_locked(self) -> int:
        return len(self.waiting_heap) + sum(len(heap) for heap in self.ready_heaps.values())

    def _clear_flow_state_locked(self):
        self._flow_last_send.clear()
        self._flow_fair_finish.clear()
        for direction in self._direction_virtual_finish:
            self._direction_virtual_finish[direction] = 0.0
        self._last_flow_cleanup_at = 0.0

    def _bandwidth_rate(self, direction: str) -> float:
        if direction == "outbound":
            return self.upload_bps
        if direction == "inbound":
            return self.download_bps
        configured = [rate for rate in (self.upload_bps, self.download_bps) if rate > 0]
        return min(configured) if configured else 0.0

    def _token_capacity(self, direction: str) -> float:
        rate_bytes = self._bandwidth_rate(direction) / 8.0
        if rate_bytes <= 0:
            return float("inf")
        return max(65535.0, rate_bytes * 0.25)

    def _reset_bandwidth_state_locked(self, now=None):
        now = time.monotonic() if now is None else now
        self._direction_cursor = 0
        self._bandwidth_waiting_sequences.clear()
        for direction in self._tokens:
            self._tokens[direction] = 0.0
            self._token_updated_at[direction] = now
            self._rate_events[direction].clear()

    def _refresh_tokens_locked(self, direction: str, now: float):
        rate_bytes = self._bandwidth_rate(direction) / 8.0
        if rate_bytes <= 0:
            self._tokens[direction] = float("inf")
            self._token_updated_at[direction] = now
            return
        previous = self._token_updated_at.get(direction, now) or now
        elapsed = max(0.0, now - previous)
        self._tokens[direction] = min(
            self._token_capacity(direction),
            self._tokens.get(direction, 0.0) + elapsed * rate_bytes,
        )
        self._token_updated_at[direction] = now

    def _bandwidth_wait_seconds_locked(self, direction: str, packet_size: int, now: float) -> float:
        rate_bytes = self._bandwidth_rate(direction) / 8.0
        if rate_bytes <= 0:
            return 0.0
        self._refresh_tokens_locked(direction, now)
        deficit = max(0.0, max(1, packet_size) - self._tokens[direction])
        return deficit / rate_bytes

    def _reset_stats_locked(self):
        self.waiting_heap.clear()
        for heap in self.ready_heaps.values():
            heap.clear()
        self.queue_bytes = 0
        self.captured_count = 0
        self.dropped_count = 0
        self.simulated_drop_count = 0
        self.random_drop_count = 0
        self.burst_drop_count = 0
        self.reordered_count = 0
        self.bandwidth_wait_count = 0
        self.queue_overflow_count = 0
        self.send_failure_count = 0
        self.cancelled_count = 0
        self.sent_count = 0
        for direction in self.sent_bytes:
            self.sent_bytes[direction] = 0
        self.peak_queue_count = 0
        self.peak_queue_bytes = 0
        self.started_at = 0.0
        self.elapsed_at_stop = 0.0
        self._last_overflow_log_at = 0.0
        self.packet_seq = itertools.count()
        self.burst_remaining = 0
        self.auto_stop_deadline = 0.0
        self._clear_flow_state_locked()
        self._reset_bandwidth_state_locked()

    def _prune_rate_events_locked(self, now: float):
        cutoff = now - self.RATE_WINDOW_SECONDS
        for events in self._rate_events.values():
            while events and events[0][0] < cutoff:
                events.popleft()

    def _current_rate_bps_locked(self, direction: str, now: float) -> float:
        self._prune_rate_events_locked(now)
        return sum(size for _, size in self._rate_events[direction]) * 8.0 / self.RATE_WINDOW_SECONDS

    def _record_rate_event_locked(self, direction: str, now: float, packet_size: int):
        bucket = int(now * 10) / 10.0
        events = self._rate_events[direction]
        if events and events[-1][0] == bucket:
            _, current_size = events.pop()
            events.append((bucket, current_size + packet_size))
        else:
            events.append((bucket, packet_size))

    def stats_snapshot(self):
        with self.lock:
            now = time.monotonic()
            elapsed = self.elapsed_at_stop
            if self.started_at:
                elapsed = max(0.0, now - self.started_at) if self.running.is_set() else self.elapsed_at_stop
            lost = self.random_drop_count + self.burst_drop_count + self.queue_overflow_count + self.send_failure_count
            loss_rate = (lost / self.captured_count * 100.0) if self.captured_count else 0.0
            remaining = 0.0
            if self.running.is_set() and self.auto_stop_deadline:
                remaining = max(0.0, self.auto_stop_deadline - now)
            return {
                "captured": self.captured_count,
                "sent": self.sent_count,
                "dropped": self.dropped_count,
                "simulated_drop": self.simulated_drop_count,
                "random_drop": self.random_drop_count,
                "burst_drop": self.burst_drop_count,
                "reordered": self.reordered_count,
                "bandwidth_wait": self.bandwidth_wait_count,
                "queue_overflow": self.queue_overflow_count,
                "send_failure": self.send_failure_count,
                "cancelled": self.cancelled_count,
                "queue_count": self._queue_count_locked(),
                "queue_bytes": self.queue_bytes,
                "peak_queue_count": self.peak_queue_count,
                "peak_queue_bytes": self.peak_queue_bytes,
                "loss_rate": loss_rate,
                "elapsed_seconds": elapsed,
                "remaining_seconds": remaining,
                "auto_stop_enabled": self.auto_stop_seconds > 0,
                "upload_bps": self._current_rate_bps_locked("outbound", now),
                "download_bps": self._current_rate_bps_locked("inbound", now),
            }

    def _effective_delay_ms(self) -> float:
        if self.jitter_ms <= 0:
            return float(self.delay_ms)
        return max(0.0, float(self.delay_ms) + random.uniform(-self.jitter_ms, self.jitter_ms))

    def _cancel_queued_locked(self):
        cancelled = self._queue_count_locked()
        if cancelled:
            self.cancelled_count += cancelled
            self.waiting_heap.clear()
            for heap in self.ready_heaps.values():
                heap.clear()
            self.queue_bytes = 0
        self._bandwidth_waiting_sequences.clear()
        return cancelled

    def _close_handle(self):
        handle = self.handle
        self.handle = None
        if handle is not None:
            try:
                handle.close()
            except Exception:
                pass

    def _fail_runtime(self, message: str, generation):
        with self.cond:
            if generation != self._start_generation:
                return
            if not self.running.is_set() and self.startup_error:
                return
            if self.started_at:
                self.elapsed_at_stop = max(0.0, time.monotonic() - self.started_at)
            self.startup_error = str(message or "弱网模拟运行异常")
            self.running.clear()
            self._cancel_queued_locked()
            self._clear_flow_state_locked()
            self.cond.notify_all()
        self.startup_event.set()
        self._close_handle()
        self._emit_state(TaskPhase.FAILED, self.startup_error)
        self.log_callback(self.startup_error)

    def start(
        self,
        delay_ms: int,
        loss_percent: int,
        filter_text: str,
        jitter_ms: int = 0,
        *,
        reorder_percent: int = 0,
        upload_bps: float = 0,
        download_bps: float = 0,
        burst_trigger_percent: int = 0,
        burst_length: int = 0,
        auto_stop_seconds: float = 0,
    ):
        if self.running.is_set():
            return

        if not is_admin():
            raise PermissionError("弱网功能需要管理员权限，请点击“管理员重启”后再使用")

        try:
            import pydivert  # noqa: F401
        except Exception as exc:
            raise RuntimeError("未安装 pydivert / WinDivert，请先运行 run.bat 或 build_exe.bat 安装依赖") from exc

        protected_filter = ensure_windivert_impostor_guard(filter_text)
        self.validate_filter(protected_filter)
        with self.cond:
            self.delay_ms = max(0, int(delay_ms))
            self.jitter_ms = max(0, int(jitter_ms))
            self.loss_percent = min(100, max(0, int(loss_percent)))
            self.reorder_percent = min(100, max(0, int(reorder_percent)))
            self.upload_bps = max(0.0, float(upload_bps))
            self.download_bps = max(0.0, float(download_bps))
            self.burst_trigger_percent = min(100, max(0, int(burst_trigger_percent)))
            self.burst_length = max(0, int(burst_length))
            self.auto_stop_seconds = max(0.0, float(auto_stop_seconds))
            self.filter_text = protected_filter
            self._reset_stats_locked()
            self.startup_error = ""
            self.startup_event.clear()
            self._start_generation += 1
            generation = self._start_generation
            self.running.set()

        self._emit_state(TaskPhase.STARTING, "正在初始化 WinDivert")
        self.capture_thread = threading.Thread(target=self._capture_loop, args=(generation,), daemon=True)
        self.dispatch_thread = threading.Thread(target=self._dispatch_loop, args=(generation,), daemon=True)
        self.capture_thread.start()
        self.dispatch_thread.start()
        if self.auto_stop_seconds > 0:
            self.auto_stop_thread = threading.Thread(target=self._auto_stop_loop, args=(generation,), daemon=True)
            self.auto_stop_thread.start()
        else:
            self.auto_stop_thread = None
        threading.Thread(target=self._watch_startup, args=(generation,), daemon=True).start()

    def _watch_startup(self, generation):
        if self.startup_event.wait(timeout=self.startup_timeout_seconds):
            return
        if not self._is_active(generation):
            return
        self._fail_runtime(f"WinDivert 启动确认超时（{self.startup_timeout_seconds} 秒）", generation)

    def _auto_stop_loop(self, generation):
        while self._is_active(generation):
            with self.cond:
                deadline = self.auto_stop_deadline
                if not deadline:
                    self.cond.wait(timeout=0.1)
                    continue
                remaining = deadline - time.monotonic()
                if remaining > 0:
                    self.cond.wait(timeout=min(remaining, 0.2))
                    continue
            if self._is_active(generation):
                self.log_callback("已达到设定的模拟时长，正在自动恢复正常网络")
                self.stop("已按计划自动停止并恢复正常网络")
            return

    def stop(self, status_message="弱网模拟已停止"):
        with self.cond:
            was_running = self.running.is_set()
            if was_running:
                self._emit_state(TaskPhase.STOPPING, "正在停止弱网模拟")
            if self.started_at:
                self.elapsed_at_stop = max(0.0, time.monotonic() - self.started_at)
            self.running.clear()
            self._start_generation += 1
            self._cancel_queued_locked()
            self._clear_flow_state_locked()
            self.cond.notify_all()
        self.startup_event.set()
        self._close_handle()

        deadline = time.monotonic() + 1.0
        current = threading.current_thread()
        for thread in (self.capture_thread, self.dispatch_thread, self.auto_stop_thread):
            if thread is not None and thread is not current and thread.is_alive():
                thread.join(timeout=max(0.0, deadline - time.monotonic()))
        if was_running:
            self.log_callback(status_message)
            self._emit_state(TaskPhase.IDLE, status_message)

    @staticmethod
    def _safe_packet_attr(packet, name, default=None):
        try:
            value = getattr(packet, name, default)
            return default if value is None else value
        except Exception:
            return default

    @classmethod
    def _packet_direction(cls, packet) -> str:
        if bool(cls._safe_packet_attr(packet, "is_outbound", False)):
            return "outbound"
        if bool(cls._safe_packet_attr(packet, "is_inbound", False)):
            return "inbound"
        direction = str(cls._safe_packet_attr(packet, "direction", "")).lower()
        if "out" in direction:
            return "outbound"
        if "in" in direction:
            return "inbound"
        return "unknown"

    @classmethod
    def _packet_flow_key(cls, packet, direction=None):
        direction = direction or cls._packet_direction(packet)
        protocol = cls._safe_packet_attr(packet, "protocol", "unknown")
        if isinstance(protocol, tuple):
            protocol = protocol[0]
        return (
            direction,
            str(protocol),
            str(cls._safe_packet_attr(packet, "src_addr", "")),
            str(cls._safe_packet_attr(packet, "dst_addr", "")),
            int(cls._safe_packet_attr(packet, "src_port", 0) or 0),
            int(cls._safe_packet_attr(packet, "dst_port", 0) or 0),
        )

    def _cleanup_flow_state_locked(self, now: float, force=False):
        if not force and now - self._last_flow_cleanup_at < self.FLOW_CLEANUP_INTERVAL_SECONDS:
            return
        cutoff = now - self.FLOW_STATE_TTL_SECONDS
        queued_flows = {item[5] for item in self.waiting_heap}
        for heap in self.ready_heaps.values():
            queued_flows.update(item[4] for item in heap)
        for state in (self._flow_last_send, self._flow_fair_finish):
            stale = [
                key
                for key, (_, last_seen) in state.items()
                if last_seen < cutoff and key not in queued_flows
            ]
            for key in stale:
                state.pop(key, None)
        self._last_flow_cleanup_at = now

    def _schedule_packet(self, packet, generation, now=None) -> bool:
        now = time.monotonic() if now is None else now
        direction = self._packet_direction(packet)
        flow_key = self._packet_flow_key(packet, direction)
        natural_send_at = now + self._effective_delay_ms() / 1000.0
        reordered = False
        with self.cond:
            if not self._is_active(generation):
                return False
            self._cleanup_flow_state_locked(now)
            previous_send_at, _ = self._flow_last_send.get(flow_key, (None, now))
            if (
                self.reorder_percent > 0
                and previous_send_at is not None
                and previous_send_at > now + 0.0001
                and random.random() < self.reorder_percent / 100.0
            ):
                lead = min(0.05, max(0.001, (previous_send_at - now) / 2.0))
                send_at = max(now, previous_send_at - lead)
                reordered = send_at < previous_send_at
            else:
                send_at = natural_send_at
                if previous_send_at is not None:
                    send_at = max(send_at, previous_send_at)
            self._flow_last_send[flow_key] = (send_at, now)
            return self._enqueue_packet(
                packet,
                send_at,
                generation,
                flow_key=flow_key,
                direction=direction,
                reordered=reordered,
                now=now,
            )

    def _enqueue_packet(
        self,
        packet,
        send_at: float,
        generation,
        *,
        flow_key=None,
        direction=None,
        reordered=False,
        now=None,
    ) -> bool:
        packet_size = self._packet_size(packet)
        should_log = False
        now = time.monotonic() if now is None else now
        direction = direction or self._packet_direction(packet)
        flow_key = flow_key or self._packet_flow_key(packet, direction)
        with self.cond:
            if not self._is_active(generation):
                return False
            if self._queue_count_locked() >= self.max_queue_packets or self.queue_bytes + packet_size > self.max_queue_bytes:
                self.queue_overflow_count += 1
                self.dropped_count += 1
                if now - self._last_overflow_log_at >= 1.0:
                    self._last_overflow_log_at = now
                    should_log = True
            else:
                previous_finish, _ = self._flow_fair_finish.get(flow_key, (0.0, now))
                baseline_finish = max(previous_finish, self._direction_virtual_finish[direction]) + max(1, packet_size)
                fair_rank = max(0.0, previous_finish - 0.5) if reordered else baseline_finish
                self._flow_fair_finish[flow_key] = (baseline_finish, now)
                sequence = next(self.packet_seq)
                heapq.heappush(
                    self.waiting_heap,
                    (send_at, sequence, fair_rank, packet_size, packet, flow_key, direction),
                )
                self.queue_bytes += packet_size
                if reordered:
                    self.reordered_count += 1
                self.peak_queue_count = max(self.peak_queue_count, self._queue_count_locked())
                self.peak_queue_bytes = max(self.peak_queue_bytes, self.queue_bytes)
                self.cond.notify()
                return True
        if should_log:
            self.log_callback(
                f"延迟队列已满，数据包被丢弃；上限 {self.max_queue_packets} 包 / {self.max_queue_bytes // (1024 * 1024)} MB"
            )
        return False

    def _move_due_packets_locked(self, now: float):
        while self.waiting_heap and self.waiting_heap[0][0] <= now:
            _, sequence, fair_rank, packet_size, packet, flow_key, direction = heapq.heappop(self.waiting_heap)
            heapq.heappush(
                self.ready_heaps[direction],
                (fair_rank, sequence, packet_size, packet, flow_key),
            )

    def _select_ready_packet_locked(self, now: float):
        directions = ("outbound", "inbound", "unknown")
        ordered = directions[self._direction_cursor :] + directions[: self._direction_cursor]
        candidates = []
        for order_index, direction in enumerate(ordered):
            heap = self.ready_heaps[direction]
            if not heap:
                continue
            _, sequence, packet_size, _, _ = heap[0]
            wait_seconds = self._bandwidth_wait_seconds_locked(direction, packet_size, now)
            candidates.append((wait_seconds, order_index, direction, sequence))
        if not candidates:
            return None, None
        wait_seconds, _, direction, sequence = min(candidates)
        if wait_seconds > 0:
            if sequence not in self._bandwidth_waiting_sequences:
                self._bandwidth_waiting_sequences.add(sequence)
                self.bandwidth_wait_count += 1
            return None, wait_seconds

        fair_rank, sequence, packet_size, packet, flow_key = heapq.heappop(self.ready_heaps[direction])
        rate = self._bandwidth_rate(direction)
        if rate > 0:
            self._tokens[direction] = max(0.0, self._tokens[direction] - max(1, packet_size))
        self._bandwidth_waiting_sequences.discard(sequence)
        self.queue_bytes = max(0, self.queue_bytes - packet_size)
        self._direction_virtual_finish[direction] = max(
            self._direction_virtual_finish[direction], fair_rank
        )
        self._direction_cursor = (directions.index(direction) + 1) % len(directions)
        return (packet, packet_size, direction, flow_key), 0.0

    def _should_drop_packet(self) -> bool:
        with self.lock:
            if self.burst_remaining > 0:
                self.burst_remaining -= 1
                self.burst_drop_count += 1
                self.dropped_count += 1
                return True
            if (
                self.burst_trigger_percent > 0
                and self.burst_length > 0
                and random.random() < self.burst_trigger_percent / 100.0
            ):
                self.burst_remaining = max(0, self.burst_length - 1)
                self.burst_drop_count += 1
                self.dropped_count += 1
                return True
            if self.loss_percent > 0 and random.random() < self.loss_percent / 100.0:
                self.random_drop_count += 1
                self.simulated_drop_count += 1
                self.dropped_count += 1
                return True
        return False

    def _capture_loop(self, generation):
        try:
            import pydivert

            self.log_callback(f"WinDivert 过滤规则：{self.filter_text}")
            handle = pydivert.WinDivert(self.filter_text)
            with self.lock:
                if not self._is_active(generation):
                    return
                self.handle = handle
            handle.open()
            if not self._is_active(generation):
                self._close_handle()
                return
            with self.cond:
                self.started_at = time.monotonic()
                if self.auto_stop_seconds > 0:
                    self.auto_stop_deadline = self.started_at + self.auto_stop_seconds
                self.cond.notify_all()
            self.startup_event.set()
            self._emit_state(TaskPhase.RUNNING, "弱网模拟运行中")
            self.log_callback("已开始拦截本机网络包")

            while self._is_active(generation):
                try:
                    packet = handle.recv()
                except Exception as exc:
                    if self._is_active(generation):
                        self._fail_runtime(f"运行中接收数据包失败：{exc}", generation)
                    return

                with self.lock:
                    self.captured_count += 1

                if self._should_drop_packet():
                    continue

                self._schedule_packet(packet, generation)

        except Exception as exc:
            if self._is_active(generation):
                self._fail_runtime(f"WinDivert 启动失败：{exc}", generation)
        finally:
            self.startup_event.set()

    def _dispatch_loop(self, generation):
        while self._is_active(generation):
            selected = None
            with self.cond:
                while self._is_active(generation) and selected is None:
                    now = time.monotonic()
                    self._move_due_packets_locked(now)
                    selected, bandwidth_wait = self._select_ready_packet_locked(now)
                    if selected is not None:
                        break
                    waits = [0.2]
                    if self.waiting_heap:
                        waits.append(max(0.0, self.waiting_heap[0][0] - now))
                    if bandwidth_wait is not None:
                        waits.append(max(0.0, bandwidth_wait))
                    positive_waits = [value for value in waits if value > 0]
                    self.cond.wait(timeout=min(positive_waits) if positive_waits else 0.01)

            if selected is None:
                continue
            packet, packet_size, direction, _ = selected
            if not self._is_active(generation):
                with self.lock:
                    self.cancelled_count += 1
                continue
            try:
                handle = self.handle
                if handle is None:
                    raise RuntimeError("WinDivert 句柄已不可用")
                handle.send(packet)
                with self.lock:
                    self.sent_count += 1
                    self.sent_bytes[direction] += packet_size
                    self._record_rate_event_locked(direction, time.monotonic(), packet_size)
            except Exception as exc:
                active = self._is_active(generation)
                with self.lock:
                    if active:
                        self.send_failure_count += 1
                        self.dropped_count += 1
                    else:
                        self.cancelled_count += 1
                if active:
                    self._fail_runtime(f"运行中发送数据包失败：{exc}", generation)
                break


# ---------------- ADB Log Capture ----------------
class ADBLogWorker:
    """Android logcat capture worker with per-device serial isolation."""

    def __init__(self, log_callback, state_callback=None, file_callback=None, line_callback=None, adb_tools=None):
        self.log_callback = log_callback
        self.state_callback = state_callback
        self.file_callback = file_callback
        self.line_callback = line_callback
        self.adb_tools = adb_tools or ADBTools(app_data_path(""))
        self.running = threading.Event()
        self.thread = None
        self.process = None  # compatibility: first running process
        self.pkg_name = DEFAULT_PKG_NAME
        self.log_dir = app_data_path("App_Logs")
        self.file_path = ""
        self.tasks = {}  # serial -> task dict
        self._lock = threading.RLock()

    def _adb_path(self) -> str:
        return self.adb_tools.adb_path()

    @staticmethod
    def _safe_name(text: str, fallback: str = "unknown") -> str:
        value = str(text or "").strip() or fallback
        return "".join(ch if ch.isalnum() or ch in "._-" else "_" for ch in value).strip("_") or fallback

    def start(self, pkg_name: str, serial: str = "", serials=None):
        if self.running.is_set():
            return
        self.pkg_name = (pkg_name or "").strip() or DEFAULT_PKG_NAME
        os.makedirs(self.log_dir, exist_ok=True)
        target_serials = list(serials or [])
        if not target_serials:
            target_serials = [self.adb_tools.resolve_serial(serial or None)]
        if not target_serials:
            raise RuntimeError("请先选择目标设备")
        self.running.set()
        self._emit_state("运行中")
        self._emit_log(f"目标包名：{self.pkg_name}")
        self._emit_log(f"日志抓取设备数：{len(target_serials)}")
        with self._lock:
            self.tasks.clear()
        for item_serial in target_serials:
            stop_event = threading.Event()
            task = {
                "serial": item_serial,
                "pkg_name": self.pkg_name,
                "status": "启动中",
                "process": None,
                "thread": None,
                "stop_event": stop_event,
                "file_path": "",
                "start_time": time.strftime("%Y-%m-%d %H:%M:%S"),
                "pid": "",
                "error": "",
                "file_size": 0,
                "updated_at": "",
            }
            thread = threading.Thread(target=self._loop_for_serial, args=(task,), daemon=True)
            task["thread"] = thread
            with self._lock:
                self.tasks[item_serial] = task
            thread.start()
        self.thread = threading.Thread(target=self._watch_tasks, daemon=True)
        self.thread.start()

    def start_all_ready_devices(self, pkg_name: str):
        data = self.adb_tools.list_devices()
        serials = [item.get("serial") for item in data.get("devices", []) if item.get("status") == "device"]
        if not serials:
            raise RuntimeError("未检测到可用 device 状态设备")
        self.start(pkg_name, serials=serials)

    def stop(self, serial: str = ""):
        with self._lock:
            targets = [self.tasks.get(serial)] if serial else list(self.tasks.values())
        for task in targets:
            if not task:
                continue
            try:
                task["stop_event"].set()
            except Exception:
                pass
            process = task.get("process")
            if process:
                try:
                    process.terminate()
                    process.wait(timeout=2)
                except Exception:
                    try:
                        process.kill()
                    except Exception:
                        pass
            task["status"] = "已停止"
        if not serial:
            self.running.clear()
            self.process = None
            self._emit_state("待命")

    def check_device_text(self) -> str:
        try:
            data = self.adb_tools.list_devices()
            return self.adb_tools.format_devices_text(data)
        except FileNotFoundError:
            return "未找到 adb.exe，请安装 Android platform-tools，或把 adb.exe 放到程序目录"
        except Exception as exc:
            return f"检测失败：{exc}"

    def open_log_dir(self):
        os.makedirs(self.log_dir, exist_ok=True)
        if os.name == "nt":
            os.startfile(self.log_dir)
        else:
            subprocess.Popen(["xdg-open", self.log_dir])

    def get_tasks_text(self) -> str:
        with self._lock:
            tasks = list(self.tasks.values())
        if not tasks:
            return "暂无日志抓取任务。"
        rows = ["serial\t状态\tPID\t文件大小\t日志文件\t错误"]
        for task in tasks:
            file_path = task.get("file_path", "")
            try:
                size = os.path.getsize(file_path) if file_path and os.path.exists(file_path) else int(task.get("file_size") or 0)
            except Exception:
                size = 0
            rows.append("\t".join([
                str(task.get("serial", "")),
                str(task.get("status", "")),
                str(task.get("pid", "")),
                str(size),
                str(file_path),
                str(task.get("error", "")),
            ]))
        return "\n".join(rows)

    def _emit_log(self, text: str):
        try:
            self.log_callback(text)
        except Exception:
            pass

    def _emit_state(self, text: str):
        if self.state_callback:
            try:
                self.state_callback(text)
            except Exception:
                pass

    def _emit_file(self, text: str):
        if self.file_callback:
            try:
                self.file_callback(text)
            except Exception:
                pass

    def _emit_line(self, line: str):
        if self.line_callback:
            try:
                self.line_callback(line)
            except Exception:
                pass

    def _run_cmd(self, cmd, timeout=ADB_CMD_TIMEOUT):
        # Backward-compatible helper for old call sites. Prefer adb_tools.run().
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="ignore",
                timeout=timeout,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
            )
        except subprocess.TimeoutExpired:
            return f"命令超时：{' '.join(cmd)}"
        output = (result.stdout or "").strip()
        error = (result.stderr or "").strip()
        if result.returncode != 0 and error:
            return error
        return output

    def _get_pid(self, serial: str):
        result = self.adb_tools.run(["shell", "pidof", "-s", self.pkg_name], timeout=5, serial=serial)
        pid = (result.stdout or "").strip()
        if pid.isdigit():
            return pid
        return ""

    def _start_logcat_process(self, task, cmd):
        old_process = task.get("process")
        if old_process:
            try:
                old_process.terminate()
                old_process.wait(timeout=2)
            except Exception:
                try:
                    old_process.kill()
                except Exception:
                    pass
        process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="ignore",
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        )
        task["process"] = process
        if not self.process:
            self.process = process
        return process

    def _build_log_file_path(self, serial: str, pkg_name: str) -> str:
        model = "device"
        try:
            model = self.adb_tools.get_prop("ro.product.model", serial=serial) or "device"
        except Exception:
            pass
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        safe_model = self._safe_name(model, "device")
        safe_serial = self._safe_name(serial, "serial")
        safe_pkg = self._safe_name(pkg_name, "pkg")
        return os.path.join(self.log_dir, f"logcat_{safe_model}_{safe_serial}_{safe_pkg}_{timestamp}.txt")

    def _loop_for_serial(self, task):
        serial = task.get("serial", "")
        pkg_name = task.get("pkg_name", self.pkg_name)
        try:
            task["status"] = "运行中"
            file_path = self._build_log_file_path(serial, pkg_name)
            task["file_path"] = file_path
            self.file_path = file_path if not self.file_path else self.file_path
            self._emit_file(self.get_tasks_text())
            self._emit_log(f"[{serial}] 正在清理旧 logcat")
            clear_result = self.adb_tools.run(["logcat", "-c"], timeout=8, serial=serial)
            if not clear_result.ok:
                self._emit_log(f"[{serial}] 清理 logcat 失败：{clear_result.output}")

            self._emit_log(f"[{serial}] 正在获取应用 PID")
            pid = self._get_pid(serial)
            task["pid"] = pid
            adb = self._adb_path()
            if pid:
                cmd = [adb, "-s", serial, "logcat", f"--pid={pid}", "-v", "threadtime"]
                filter_by_pkg = False
                self._emit_log(f"[{serial}] 发现进程 PID：{pid}")
            else:
                cmd = [adb, "-s", serial, "logcat", "-v", "threadtime"]
                filter_by_pkg = True
                self._emit_log(f"[{serial}] 应用当前未运行，改用包名关键词过滤模式")

            self._emit_log(f"[{serial}] 日志保存到：{file_path}")
            process = self._start_logcat_process(task, cmd)
            next_pid_check = time.monotonic() + 2
            with open(file_path, "w", encoding="utf-8", errors="ignore") as file:
                file.write(f"# device_serial={serial}\n# package={pkg_name}\n# start_time={task.get('start_time')}\n")
                while self.running.is_set() and not task["stop_event"].is_set() and process.poll() is None:
                    if filter_by_pkg and time.monotonic() >= next_pid_check:
                        next_pid_check = time.monotonic() + 2
                        pid = self._get_pid(serial)
                        if pid:
                            task["pid"] = pid
                            self._emit_log(f"[{serial}] 检测到应用已启动，切换为 PID 抓取：{pid}")
                            filter_by_pkg = False
                            process = self._start_logcat_process(task, [adb, "-s", serial, "logcat", f"--pid={pid}", "-v", "threadtime"])
                            continue
                    if not process.stdout:
                        time.sleep(0.05)
                        continue
                    line = process.stdout.readline()
                    if not line:
                        time.sleep(0.05)
                        continue
                    if filter_by_pkg and pkg_name and pkg_name not in line:
                        continue
                    file.write(line)
                    file.flush()
                    task["updated_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
                    try:
                        task["file_size"] = os.path.getsize(file_path)
                    except Exception:
                        pass
                    self._emit_line(line)
            if task["stop_event"].is_set() or not self.running.is_set():
                task["status"] = "已停止"
                self._emit_log(f"[{serial}] 日志抓取已停止")
            else:
                task["status"] = "异常"
                task["error"] = "logcat 进程已退出，可能是设备断开或 ADB 异常"
                self._emit_log(f"[{serial}] logcat 进程异常退出，请检查设备连接")
        except FileNotFoundError:
            task["status"] = "异常"
            task["error"] = "未找到 adb.exe"
            self._emit_log("未找到 adb.exe，请安装 Android platform-tools，或把 adb.exe 放到程序目录")
        except Exception as exc:
            task["status"] = "异常"
            task["error"] = str(exc)
            if "serial 已变化" in str(exc) or "offline" in str(exc).lower() or "device" in str(exc).lower():
                self._emit_log(f"[{serial}] 当前设备已断开或不可用，日志抓取已停止：{exc}")
            else:
                self._emit_log(f"[{serial}] 日志抓取失败：{exc}")
        finally:
            process = task.get("process")
            if process and process.poll() is None:
                try:
                    process.terminate()
                except Exception:
                    pass
            self._emit_file(self.get_tasks_text())

    def _watch_tasks(self):
        while self.running.is_set():
            with self._lock:
                threads = [task.get("thread") for task in self.tasks.values()]
            if threads and all((thread is None or not thread.is_alive()) for thread in threads):
                break
            time.sleep(0.3)
        self.running.clear()
        self.process = None
        self._emit_state("待命")
        self._emit_file(self.get_tasks_text())


# ---------------- Main App ----------------
FONT_FAMILY = "Microsoft YaHei UI"

COLOR_BG = ("#F4F7FB", "#0B1018")
COLOR_SIDEBAR = ("#FFFFFF", "#111820")
COLOR_SURFACE = ("#FFFFFF", "#151C26")
COLOR_SURFACE_2 = ("#F8FAFD", "#1B2430")
COLOR_HOVER = ("#EEF4FF", "#222D3B")
COLOR_SELECTED = ("#EAF2FF", "#263243")
COLOR_TEXT = ("#111827", "#E7ECF4")
COLOR_MUTED = ("#64748B", "#9AA6B8")
COLOR_BORDER = ("#DDE6F2", "#2B3544")
COLOR_ACCENT = ("#2F6FED", "#4F8CFF")
COLOR_ACCENT_HOVER = ("#245FD2", "#3A78EC")
COLOR_DANGER = ("#EF4444", "#F87171")
COLOR_DANGER_HOVER = ("#DC2626", "#EF4444")
COLOR_SUCCESS = ("#16A34A", "#22C55E")
COLOR_WARNING = ("#F59E0B", "#FBBF24")
UI_MAX_RENDER_ROWS = 3000
UI_PAGE_SIZE = 200


class XiaoXinAssistant(ctk.CTk):
    def __init__(self):
        try:
            if os.name == "nt":
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("XiaoLinAssistant.QAToolbox")
        except Exception:
            pass
        super().__init__()

        self.title(APP_NAME)
        self.default_window_width = 1400
        self.default_window_height = 850
        self.geometry(f"{self.default_window_width}x{self.default_window_height}")
        self.minsize(1280, 760)
        self._center_window(self.default_window_width, self.default_window_height)
        self.configure(fg_color=COLOR_BG)

        self._init_window_icon()
        self._page_cache = {}
        self._font_cache = {}
        self._configure_ttk_styles()

        self.current_page = "home"
        self.logo_image = None

        self.is_exiting = False
        self._close_prompt_active = False
        self._close_dialog = None
        self.tray_icon = None
        self.tray_thread = None
        self.tray_available = pystray is not None
        self.tray_error = ""
        self._tray_hide_tip_shown = False

        self.task_states = {
            name: RuntimeTaskState(name)
            for name in (
                "clicker",
                "network",
                "adb_log",
                "adb_tools",
                "ios_log",
                "compare",
                "localization",
                "config_validator",
            )
        }
        self.clicker = ClickerWorker()
        self.network = NetworkWorker(self._network_log, self._on_network_state)
        self.log_monitor = LogMonitor(DEFAULT_KEYWORDS, context_lines=20)
        self.crash_stream_detector = CrashStreamDetector(context_lines=50)
        self.crash_file_analyzer = CrashAnalyzer(context_lines=60)
        self.log_hit_records = []
        self.crash_records = []
        self._log_analysis_last_refresh = 0.0
        self.adb_tools = ADBTools(app_data_path(""))
        self.adb_log = ADBLogWorker(self._adb_log, self._set_adb_status, self._set_adb_file, self._on_live_log_line, adb_tools=self.adb_tools)

        # Clicker vars
        self.click_button_var = tk.StringVar(value="left")
        self.click_interval_var = tk.StringVar(value="极速模式（每秒100次）")
        self.hotkey_var = tk.StringVar(value="F8")
        self.clicker_status_var = tk.StringVar(value="运行状态：待命")
        self.clicker_custom_interval = 0.01

        # Network vars
        self.target_ip_var = tk.StringVar(value="")
        self.protocol_var = tk.StringVar(value="全部")
        self.direction_var = tk.StringVar(value="全部")
        self.delay_var = tk.StringVar(value="100")
        self.jitter_var = tk.StringVar(value="0")
        self.loss_var = tk.StringVar(value="0")
        self.reorder_var = tk.StringVar(value="0")
        self.upload_bandwidth_var = tk.StringVar(value="0")
        self.download_bandwidth_var = tk.StringVar(value="0")
        self.bandwidth_unit_var = tk.StringVar(value="Mbps")
        self.burst_trigger_var = tk.StringVar(value="0")
        self.burst_length_var = tk.StringVar(value="0")
        self.network_duration_var = tk.StringVar(value="0")
        self.advanced_filter_var = tk.StringVar(value="")
        self.network_guide_visible = True
        self.network_status_var = tk.StringVar(value="运行状态：待命")
        self.network_delay_hint_var = tk.StringVar(value="单向延迟 100ms；全部方向预计额外 RTT 约 200ms")
        self.network_stats_var = tk.StringVar(
            value="捕获 0 ｜ 放行 0 ｜ 随机丢包 0 ｜ 突发丢包 0 ｜ 乱序 0\n"
            "限速等待 0 ｜ 队列 0/峰值 0（0 B）｜ 溢出 0 ｜ 发送失败 0\n"
            "上行 0 bps ｜ 下行 0 bps ｜ 丢包率 0.00% ｜ 剩余 -- ｜ 运行 00:00"
        )
        for variable in (self.delay_var, self.jitter_var, self.direction_var):
            variable.trace_add("write", lambda *_args: self._update_network_delay_hint())

        # ADB log vars
        self.log_pkg_var = tk.StringVar(value=DEFAULT_PKG_NAME)
        self.adb_status_var = tk.StringVar(value="待命")
        self.adb_checking = threading.Event()
        self.adb_file_var = tk.StringVar(value="未生成")
        self.adb_selected_device_var = tk.StringVar(value="")
        self.adb_selected_device_label_var = tk.StringVar(value="当前设备：未选择")
        self.adb_device_status_badge_var = tk.StringVar(value="未选择")
        self.adb_device_hint_var = tk.StringVar(value="请点击刷新设备")
        self.adb_device_choice_values = []
        self.adb_device_status_labels = []
        self._adb_last_device_data = None
        # v2.4.28: 防止设备表格刷新时 selection_set 触发递归事件，造成 UI 未响应。
        self._adb_device_table_updating = False
        self.adb_log_selected_serials = set()

        # Log analysis vars
        self.log_analysis_enabled_var = tk.BooleanVar(value=True)
        self.log_context_lines_var = tk.StringVar(value="20")
        self.log_monitor_pkg_only_var = tk.BooleanVar(value=False)
        self.log_keyword_vars = {keyword: tk.BooleanVar(value=True) for keyword in DEFAULT_KEYWORDS}
        self.log_custom_keywords_textbox = None
        self.log_hits_box = None
        self.crash_results_box = None
        self.log_analysis_status_var = tk.StringVar(value="实时分析待命：启动日志抓取后自动监控")
        self.log_analysis_busy = threading.Event()

        # ADB tools vars
        self.adb_tools_devices_var = tk.StringVar(value="未检测")
        self.adb_tools_info_var = tk.StringVar(value="未读取设备信息")
        self.adb_tools_pkg_var = tk.StringVar(value="")
        self.adb_tools_foreground_var = tk.StringVar(value="未识别")
        self.adb_recording_status_var = tk.StringVar(value="未录屏")
        self.adb_tools_busy = threading.Event()
        self._adb_refresh_generation = 0
        self.adb_recording_session = None
        self.apk_path_var = tk.StringVar(value="")
        self.apk_status_var = tk.StringVar(value="未选择 APK")

        # File compare vars
        self.compare_actual_path_var = tk.StringVar(value="")
        self.compare_document_path_var = tk.StringVar(value="")
        self.compare_actual_status_var = tk.StringVar(value="未选择文件")
        self.compare_document_status_var = tk.StringVar(value="未选择文件")
        self.compare_mode_var = tk.StringVar(value="策划文档 vs 配置文件")
        self.compare_ignore_spaces_var = tk.BooleanVar(value=True)
        self.compare_ignore_newlines_var = tk.BooleanVar(value=False)
        self.compare_ignore_case_var = tk.BooleanVar(value=False)
        self.compare_ignore_fields_var = tk.StringVar(value="时间戳, timestamp, version, 版本号, id")
        self.compare_search_var = tk.StringVar(value="")
        self.compare_summary_var = tk.StringVar(value="请选择待比对文件和参考文档后开始比对")
        self.compare_all_diffs = []
        self.compare_visible_diffs = []
        self.compare_is_working = False
        self.compare_search_trace_bound = False
        self.compare_page_index = 1
        self.compare_page_jump_var = tk.StringVar(value="1")
        self.compare_page_status_var = tk.StringVar(value="第 1/1 页，每页 200 条")

        # Value reference vs config compare vars
        self.value_compare_ids_var = tk.StringVar(value="")
        self.value_compare_keywords_var = tk.StringVar(value="")
        self.value_compare_summary_var = tk.StringVar(value="通用结构化比对：选择 Sheet、主键、过滤条件和字段映射后，按主键输出确认差异")
        self.value_action_status_var = tk.StringVar(value="待执行")
        self.value_template_var = tk.StringVar(value="通用配置表")
        self.value_compare_rewards_var = tk.BooleanVar(value=True)
        self.value_compare_price_var = tk.BooleanVar(value=True)
        self.value_compare_pcid_var = tk.BooleanVar(value=True)
        self.value_compare_duration_var = tk.BooleanVar(value=True)
        self.value_compare_score_var = tk.BooleanVar(value=False)
        self.value_compare_limit_var = tk.BooleanVar(value=True)
        self.value_ref_sheet_var = tk.StringVar(value="配置参考表")
        self.value_cfg_sheet_var = tk.StringVar(value="")
        self.value_ref_header_row_var = tk.StringVar(value="auto")
        self.value_ref_data_start_row_var = tk.StringVar(value="auto")
        self.value_cfg_header_row_var = tk.StringVar(value="auto")
        self.value_cfg_data_start_row_var = tk.StringVar(value="auto")
        self.value_key_fields_var = tk.StringVar(value="auto")
        self.value_ref_filter_var = tk.StringVar(value="")
        self.value_cfg_filter_var = tk.StringVar(value="")
        self.value_ignore_fields_var = tk.StringVar(value="备注,说明,comment,desc,client_note,dev_note,程序不读,展示用字段,临时字段")
        self.value_zero_equal_fields_var = tk.StringVar(value="reward")
        self.value_percent_fields_var = tk.StringVar(value="rate,discount,percent")
        self.value_bool_fields_var = tk.StringVar(value="")
        self.value_case_insensitive_fields_var = tk.StringVar(value="")
        self.value_mapping_textbox = None
        self.value_compare_settings_card = None
        self.value_compare_button = None
        self.value_export_result_button = None
        self.value_open_result_button = None

        # Translation compare vars
        self.tr_source_header_row_var = tk.StringVar(value="1")
        self.tr_source_data_start_row_var = tk.StringVar(value="2")
        self.tr_source_key_column_var = tk.StringVar(value="")
        self.tr_source_sheet_var = tk.StringVar(value="")
        self.tr_config_header_row_var = tk.StringVar(value="1")
        self.tr_config_real_header_row_var = tk.StringVar(value="")
        self.tr_config_data_start_row_var = tk.StringVar(value="2")
        self.tr_config_key_column_var = tk.StringVar(value="")
        self.tr_config_sheet_var = tk.StringVar(value="")
        self.tr_check_extra_ids_var = tk.BooleanVar(value=True)
        self.tr_ignore_trim_var = tk.BooleanVar(value=True)
        self.tr_ignore_all_spaces_var = tk.BooleanVar(value=False)
        self.tr_ignore_newlines_var = tk.BooleanVar(value=False)
        self.tr_ignore_case_var = tk.BooleanVar(value=False)
        self.tr_ignore_width_var = tk.BooleanVar(value=False)
        self.tr_ignore_punctuation_var = tk.BooleanVar(value=False)
        self.tr_language_names = [
            "中文简体", "中文繁体", "英语", "日语", "韩语", "德语", "法语", "西班牙语",
            "葡萄牙语", "巴西葡语", "俄语", "意大利语", "泰语", "越南语", "印尼语",
            "土耳其语", "阿拉伯语", "乌克兰语"
        ]
        self.tr_language_vars = {name: tk.BooleanVar(value=True) for name in self.tr_language_names}
        self.tr_mapping_textbox = None
        self.translation_settings_card = None
        self.compare_result_card = None
        self.compare_summary_label = None
        self.compare_summary_card = None
        self.compare_local_scroll = None
        self.compare_last_export_path = ""
        self._last_value_compare_report = None
        self.tr_advanced_frame = None
        self.tr_advanced_visible = False
        self.tr_advanced_toggle_button = None
        self.tr_language_summary_var = tk.StringVar(value="")

        # Localization check vars
        self.loc_file_path_var = tk.StringVar(value="")
        self.loc_file_status_var = tk.StringVar(value="未选择文件")
        self.loc_header_row_var = tk.StringVar(value="1")
        self.loc_data_start_row_var = tk.StringVar(value="2")
        self.loc_key_column_var = tk.StringVar(value="")
        self.loc_sheet_var = tk.StringVar(value="")
        self.loc_structure_var = tk.StringVar(value="自动识别")
        self.loc_expected_languages_var = tk.StringVar(value="EN,DE,FR,RU,ES,JP,ID,KR,TH,IT,PT,TR,VN,ARB")
        self.loc_source_language_var = tk.StringVar(value="")
        self.loc_allowed_chinese_var = tk.StringVar(value="中文简体, 中文繁体, CN, cn, zh, zh_cn, TW, tw, zh_tw, 中文, 简体中文, 繁体中文")
        self.loc_ignore_ids_var = tk.StringVar(value="")
        self.loc_symbols_var = tk.StringVar(value="%, :, ：, +, -, /, \\, (), [], {}, \\n")
        self.loc_max_length_var = tk.StringVar(value="0")
        self.loc_length_ratio_var = tk.StringVar(value="0")
        self.loc_text_row_types_var = tk.StringVar(value="对白,说话,旁白,文本,标题,选项,按钮,提示,名字,名称,字幕,描述,转场文字")
        self.loc_non_text_row_types_var = tk.StringVar(value="转场动画,跳转,切换场景,模型出现,模型消失,玩法,镜头效果,关卡结束,分支节点")
        self.loc_unfinished_markers_var = tk.StringVar(value="TODO,TBD,FIXME,翻译未完成,翻譯未完成,翻译错误，修改中,翻譯錯誤，修改中,待翻译,待翻譯,未翻译,未翻譯,untranslated,need translation,needs translation")
        self.loc_search_var = tk.StringVar(value="")
        self.loc_summary_var = tk.StringVar(value="请选择翻译文件后开始检查")
        self.loc_progress_var = tk.StringVar(value="等待开始")
        self.loc_check_empty_var = tk.BooleanVar(value=True)
        self.loc_check_chinese_var = tk.BooleanVar(value=True)
        self.loc_check_placeholders_var = tk.BooleanVar(value=True)
        self.loc_check_tags_var = tk.BooleanVar(value=True)
        self.loc_check_symbols_var = tk.BooleanVar(value=True)
        self.loc_check_length_var = tk.BooleanVar(value=False)
        self.loc_check_duplicate_id_var = tk.BooleanVar(value=True)
        self.loc_check_invalid_id_var = tk.BooleanVar(value=True)
        self.loc_check_missing_id_var = tk.BooleanVar(value=True)
        self.loc_check_source_consistency_var = tk.BooleanVar(value=True)
        self.loc_check_numbers_var = tk.BooleanVar(value=True)
        self.loc_check_duplicate_translation_var = tk.BooleanVar(value=False)
        self.loc_language_vars = {}
        self.loc_language_frame = None
        self.loc_rule_detail_frame = None
        self.loc_progress_bar = None
        self.loc_rules_visible = False
        self.loc_rules_toggle_button = None
        self.loc_detected_headers = []
        self.loc_multi_table_candidates_cache = []
        self.loc_multi_table_cache_key = None
        self.loc_recommended_languages = []
        self.loc_recommended_languages = []
        self.loc_all_issues = []
        self.loc_visible_issues = []
        self.loc_is_working = False
        self.loc_stop_event = threading.Event()
        self.loc_search_trace_bound = False
        self.loc_page_index = 1
        self.loc_page_jump_var = tk.StringVar(value="1")
        self.loc_page_status_var = tk.StringVar(value="第 1/1 页，每页 200 条")

        # Config validator vars
        self.cfg_file_path_var = tk.StringVar(value="")
        self.cfg_file_status_var = tk.StringVar(value="未选择配置表")
        self.cfg_header_row_var = tk.StringVar(value="1")
        self.cfg_description_row_var = tk.StringVar(value="1")
        self.cfg_type_row_var = tk.StringVar(value="2")
        self.cfg_data_start_row_var = tk.StringVar(value="2")
        self.cfg_key_column_var = tk.StringVar(value="")
        self.cfg_sheet_var = tk.StringVar(value="")
        self.cfg_detect_summary_var = tk.StringVar(value="未识别表结构")
        self.cfg_current_headers = []
        self.cfg_quick_field_var = tk.StringVar(value="")
        self.cfg_quick_template_var = tk.StringVar(value="restrict 条件格式校验")
        self.cfg_quick_required_no_zero_var = tk.BooleanVar(value=False)
        self.cfg_quick_required_no_minus_one_var = tk.BooleanVar(value=False)
        self.cfg_quick_preview_var = tk.StringVar(value="请选择配置表并选择快捷规则模板")
        self.cfg_auto_type_enabled_var = tk.BooleanVar(value=True)
        self.cfg_compound_enabled_var = tk.BooleanVar(value=False)
        self.cfg_compound_field_var = tk.StringVar(value="")
        self.cfg_compound_allow_empty_var = tk.BooleanVar(value=True)
        self.cfg_compound_group_sep_var = tk.StringVar(value="|")
        self.cfg_compound_param_sep_var = tk.StringVar(value="*")
        self.cfg_compound_count_var = tk.StringVar(value="4")
        self.cfg_compound_type_var = tk.StringVar(value="整数")
        self.cfg_compound_allow_zero_var = tk.BooleanVar(value=True)
        self.cfg_compound_forbid_values_var = tk.StringVar(value="")
        self.cfg_compound_rules_var = tk.StringVar(value="")
        self.cfg_required_fields_var = tk.StringVar(value="")
        self.cfg_numeric_fields_var = tk.StringVar(value="")
        self.cfg_range_rules_var = tk.StringVar(value="")
        self.cfg_boolean_fields_var = tk.StringVar(value="")
        self.cfg_enum_rules_var = tk.StringVar(value="")
        self.cfg_time_fields_var = tk.StringVar(value="")
        self.cfg_time_range_rules_var = tk.StringVar(value="")
        self.cfg_reward_fields_var = tk.StringVar(value="")
        self.cfg_reference_rules_var = tk.StringVar(value="")
        self.cfg_interval_rules_var = tk.StringVar(value="")
        self.cfg_required_columns_var = tk.StringVar(value="")
        self.cfg_search_var = tk.StringVar(value="")
        self.cfg_summary_var = tk.StringVar(value="请选择配置表后开始校验")
        self.cfg_check_duplicate_var = tk.BooleanVar(value=True)
        self.cfg_check_empty_key_var = tk.BooleanVar(value=True)
        self.cfg_check_required_var = tk.BooleanVar(value=True)
        self.cfg_check_numeric_var = tk.BooleanVar(value=True)
        self.cfg_check_range_var = tk.BooleanVar(value=True)
        self.cfg_check_boolean_var = tk.BooleanVar(value=True)
        self.cfg_check_enum_var = tk.BooleanVar(value=True)
        self.cfg_check_time_var = tk.BooleanVar(value=True)
        self.cfg_check_time_range_var = tk.BooleanVar(value=True)
        self.cfg_check_reward_var = tk.BooleanVar(value=True)
        self.cfg_check_reference_var = tk.BooleanVar(value=True)
        self.cfg_check_interval_var = tk.BooleanVar(value=True)
        self.cfg_check_empty_rows_var = tk.BooleanVar(value=True)
        self.cfg_check_required_columns_var = tk.BooleanVar(value=True)
        self.cfg_rules_frame = None
        self.cfg_rules_visible = False
        self.cfg_rules_toggle_button = None
        self.cfg_all_issues = []
        self.cfg_visible_issues = []
        self.cfg_is_working = False
        self.cfg_stop_event = threading.Event()
        self.cfg_search_trace_bound = False
        self.cfg_page_index = 1
        self.cfg_page_jump_var = tk.StringVar(value="1")
        self.cfg_page_status_var = tk.StringVar(value="第 1/1 页，每页 200 条")

        # iOS log vars
        self.ios_tools = IOSLogTools(app_data_path(""))
        self.ios_tool_path_vars = {name: tk.StringVar(value=self.ios_tools.tool_paths.get(name, "")) for name in [
            "idevice_id.exe", "ideviceinfo.exe", "idevicesyslog.exe", "idevicecrashreport.exe"
        ]}
        self.ios_devices_var = tk.StringVar(value="未检测 iOS 设备")
        self.ios_selected_udid_var = tk.StringVar(value="")
        self.ios_filter_text_var = tk.StringVar(value="")
        self.ios_only_hits_var = tk.BooleanVar(value=False)
        self.ios_status_var = tk.StringVar(value="待命")
        self.ios_log_file_var = tk.StringVar(value="未生成")
        self.ios_tool_status_var = tk.StringVar(value="未检测工具路径")
        self.ios_busy = threading.Event()
        self.ios_log_running = threading.Event()
        self.ios_log_process = None
        self.ios_log_thread = None
        self.ios_log_full_path = ""
        self.ios_log_filtered_path = ""
        self.ios_display_lines = []
        self.ios_max_display_lines = 3000
        self.ios_keyword_vars = {keyword: tk.BooleanVar(value=True) for keyword in DEFAULT_IOS_KEYWORDS}
        self.ios_custom_keywords_textbox = None
        self.ios_devices_box = None
        self.ios_log_box = None

        # Time test tool vars
        self.time_result_boxes = {}
        self.time_default_tz_var = tk.StringVar(value="本地时区")
        self.time_custom_offset_var = tk.StringVar(value="+8")
        now_text = time.strftime("%Y-%m-%d %H:%M:%S")
        self.time_ts_input_var = tk.StringVar(value="")
        self.time_ts_unit_var = tk.StringVar(value="自动识别")
        self.time_dt_input_var = tk.StringVar(value=now_text)
        self.time_dt_tz_var = tk.StringVar(value="本地时区")
        self.time_unit_value_var = tk.StringVar(value="86400")
        self.time_unit_from_var = tk.StringVar(value="秒")
        self.time_diff_start_var = tk.StringVar(value=now_text)
        self.time_diff_end_var = tk.StringVar(value=now_text)
        self.time_add_base_var = tk.StringVar(value=now_text)
        self.time_add_direction_var = tk.StringVar(value="增加")
        self.time_add_days_var = tk.StringVar(value="7")
        self.time_add_hours_var = tk.StringVar(value="0")
        self.time_add_minutes_var = tk.StringVar(value="0")
        self.time_add_seconds_var = tk.StringVar(value="0")
        self.time_countdown_now_var = tk.StringVar(value="")
        self.time_countdown_end_var = tk.StringVar(value=now_text)
        self.time_boundary_base_var = tk.StringVar(value=now_text)

        # Online update vars
        self.update_status_var = tk.StringVar(value="尚未检查更新")
        self.update_busy = threading.Event()
        self.update_cancel_event = threading.Event()
        self.update_check_button = None
        self.update_progress_dialog = None
        self.update_progress_bar = None
        self.update_progress_label_var = tk.StringVar(value="")
        self.available_update_info = None

        # Theme vars
        self.theme_mode_var = tk.StringVar(value="深色模式")

        self._build_ui()
        self.after(50, self._finish_background_startup)

        self.protocol("WM_DELETE_WINDOW", self.on_close_window)

    def _finish_background_startup(self):
        """Start non-visual helpers after Tk has painted its first frame."""
        if self.is_exiting:
            return
        self._setup_tray()
        threading.Thread(
            target=self.clicker.setup_hotkey,
            args=(lambda: self.hotkey_var.get(), lambda: self.after(0, self.toggle_clicker)),
            daemon=True,
        ).start()
        self._tick_stats()
        if "--post-update" in sys.argv:
            self.after(
                1000,
                lambda: messagebox.showinfo(
                    "更新完成",
                    f"测试助手已成功更新到 v{APP_VERSION}。",
                    parent=self,
                ),
            )
        self.after(2200, lambda: self.check_for_updates(silent=True))

    def _center_window(self, width: int, height: int):
        """Center the default window on the current screen without changing resizable behavior."""
        try:
            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()
            x = max(0, int((screen_width - width) / 2))
            y = max(0, int((screen_height - height) / 2))
            self.geometry(f"{width}x{height}+{x}+{y}")
        except Exception:
            self.geometry(f"{width}x{height}")

    def _init_window_icon(self):
        try:
            if os.name == "nt":
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("XiaoLinAssistant.QAToolbox")
        except Exception:
            pass
        icon_path = get_icon_path("ico")
        if os.path.exists(icon_path):
            try:
                self.iconbitmap(icon_path)
            except Exception:
                pass
        png_path = get_icon_path("png")
        if os.path.exists(png_path):
            try:
                self._window_icon_image = tk.PhotoImage(file=png_path)
                self.iconphoto(True, self._window_icon_image)
            except Exception:
                pass

    def _set_dialog_icon(self, dialog):
        try:
            icon_path = get_icon_path("ico")
            if os.path.exists(icon_path):
                dialog.iconbitmap(icon_path)
        except Exception:
            pass

    def _ui_color(self, color_value):
        try:
            return self._apply_appearance_mode(color_value)
        except Exception:
            if isinstance(color_value, tuple):
                return color_value[1]
            return color_value

    def _font(self, size, weight="normal"):
        """Reuse CTkFont objects instead of creating one for every widget."""
        key = (int(size), str(weight))
        font = self._font_cache.get(key)
        if font is None:
            font = ctk.CTkFont(family=FONT_FAMILY, size=key[0], weight=key[1])
            self._font_cache[key] = font
        return font

    def _configure_ttk_styles(self):
        """Apply the active CTk palette to cached native ttk widgets."""
        try:
            style = ttk.Style(self)
            try:
                style.theme_use("clam")
            except Exception:
                pass
            bg = self._ui_color(COLOR_SURFACE_2)
            heading_bg = self._ui_color(COLOR_SURFACE)
            fg = self._ui_color(COLOR_TEXT)
            border = self._ui_color(COLOR_BORDER)
            selected = self._ui_color(COLOR_SELECTED)
            trough = self._ui_color(COLOR_BG)
            style.configure(
                "Xiao.Treeview",
                background=bg,
                fieldbackground=bg,
                foreground=fg,
                bordercolor=border,
                lightcolor=border,
                darkcolor=border,
                rowheight=28,
                borderwidth=0,
                relief="flat",
            )
            style.configure(
                "Xiao.Treeview.Heading",
                background=heading_bg,
                foreground=fg,
                bordercolor=border,
                relief="flat",
                font=(FONT_FAMILY, 10, "bold"),
            )
            style.map(
                "Xiao.Treeview",
                background=[("selected", selected)],
                foreground=[("selected", fg)],
            )
            style.configure(
                "TScrollbar",
                background=heading_bg,
                troughcolor=trough,
                bordercolor=border,
                lightcolor=border,
                darkcolor=border,
                arrowcolor=fg,
                relief="flat",
            )
            style.map(
                "TScrollbar",
                background=[("active", selected), ("pressed", selected)],
                arrowcolor=[("active", fg), ("pressed", fg)],
            )
        except Exception:
            pass

    def _refresh_ttk_theme_widgets(self):
        """Force already-cached ttk tables to repaint after an appearance switch."""
        pending = [self]
        while pending:
            parent = pending.pop()
            try:
                children = parent.winfo_children()
            except Exception:
                continue
            pending.extend(children)
            for child in children:
                try:
                    if isinstance(child, ttk.Treeview):
                        child.configure(style="Xiao.Treeview")
                        child.tag_configure("running", foreground=self._ui_color(COLOR_SUCCESS))
                        child.tag_configure("error", foreground=self._ui_color(COLOR_DANGER))
                        child.tag_configure("offline", foreground=self._ui_color(COLOR_WARNING))
                        child.tag_configure("stopped", foreground=self._ui_color(COLOR_MUTED))
                    elif isinstance(child, ttk.Scrollbar):
                        child.configure(style="TScrollbar")
                except Exception:
                    pass

    def _recreate_content_container(self, key: str):
        """Show a cached page container or create it on first use.

        Return True only when the caller needs to build the page widgets.
        """
        current = getattr(self, "content", None)
        try:
            if current is not None and current.winfo_exists():
                current.grid_remove()
        except Exception:
            pass

        cached = self._page_cache.get(key)
        try:
            if cached is not None and cached.winfo_exists():
                self.content = cached
                self.content.grid()
                return False
        except Exception:
            self._page_cache.pop(key, None)

        # A fixed frame avoids Windows Canvas trails from moving many CTk widgets.
        self.content = ctk.CTkFrame(self.main, fg_color=COLOR_BG, corner_radius=0)
        self.content.grid(row=1, column=0, sticky="nsew", padx=24, pady=(0, 24))
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(0, weight=1)
        self._page_cache[key] = self.content
        return True

    def _get_page_bounds(self, total_count: int, current_page: int):
        total_pages = max(1, (int(total_count) + UI_PAGE_SIZE - 1) // UI_PAGE_SIZE)
        current_page = max(1, min(int(current_page or 1), total_pages))
        start = (current_page - 1) * UI_PAGE_SIZE
        end = min(start + UI_PAGE_SIZE, int(total_count))
        return current_page, total_pages, start, end

    def _build_pager(self, parent, status_var, jump_var, prev_command, next_command, jump_command):
        pager = ctk.CTkFrame(parent, fg_color=COLOR_SURFACE)
        pager.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(pager, textvariable=status_var, text_color=COLOR_MUTED, font=self._font(12), anchor="w").grid(row=0, column=0, sticky="w", padx=(0, 12))
        ctk.CTkButton(pager, text="上一页", width=74, height=30, corner_radius=10, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(12), command=prev_command).grid(row=0, column=2, padx=(0, 8))
        ctk.CTkButton(pager, text="下一页", width=74, height=30, corner_radius=10, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(12), command=next_command).grid(row=0, column=3, padx=(0, 8))
        ctk.CTkEntry(pager, textvariable=jump_var, width=54, height=30, corner_radius=9, fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(12)).grid(row=0, column=4, padx=(0, 8))
        ctk.CTkButton(pager, text="跳转", width=58, height=30, corner_radius=10, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(12), command=jump_command).grid(row=0, column=5)
        return pager

    # ---------- UI base ----------
    def _build_ui(self):
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=0)
        self.grid_rowconfigure(0, weight=1)

        self.sidebar = ctk.CTkFrame(self, width=250, fg_color=COLOR_SIDEBAR, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)

        self.main = ctk.CTkFrame(self, fg_color=COLOR_BG, corner_radius=0)
        self.main.grid(row=0, column=1, sticky="nsew")
        self.main.grid_columnconfigure(0, weight=1)
        self.main.grid_rowconfigure(1, weight=1)

        self.right_panel = ctk.CTkFrame(self, width=310, fg_color=COLOR_SIDEBAR, corner_radius=0)
        self.right_panel.grid(row=0, column=2, sticky="nsew")
        self.right_panel.grid_propagate(False)

        self._build_sidebar()
        self._build_header()
        self._build_right_panel()

        self.content = None
        self.show_page("home")

    def _build_sidebar(self):
        self.sidebar.grid_rowconfigure(1, weight=1)
        self.sidebar.grid_columnconfigure(0, weight=1)

        brand = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        brand.grid(row=0, column=0, sticky="ew", padx=22, pady=(22, 18))

        icon_png = get_icon_path("png")
        if os.path.exists(icon_png):
            try:
                self.logo_image = ctk.CTkImage(Image.open(icon_png), size=(66, 66))
                ctk.CTkLabel(brand, image=self.logo_image, text="").pack(side="left", padx=(0, 12))
            except Exception:
                ctk.CTkLabel(brand, text="欣", width=66, height=66, corner_radius=22, fg_color=COLOR_ACCENT, text_color="white", font=self._font(28, "bold")).pack(side="left", padx=(0, 12))
        else:
            ctk.CTkLabel(brand, text="欣", width=66, height=66, corner_radius=22, fg_color=COLOR_ACCENT, text_color="white", font=self._font(28, "bold")).pack(side="left", padx=(0, 12))

        title_box = ctk.CTkFrame(brand, fg_color="transparent")
        title_box.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(title_box, text=APP_NAME, text_color=COLOR_TEXT, font=self._font(23, "bold")).pack(anchor="w")
        ctk.CTkLabel(title_box, text=APP_SUBTITLE, text_color=COLOR_MUTED, font=self._font(13)).pack(anchor="w", pady=(3, 0))

        self.nav_container = ctk.CTkFrame(self.sidebar, fg_color=COLOR_SIDEBAR, corner_radius=0)
        self.nav_container.grid(row=1, column=0, sticky="nsew", padx=0, pady=(0, 8))
        self.nav_container.grid_columnconfigure(0, weight=1)

        self.nav_buttons = {}
        # v2.4.15: 高频测试工具靠前，专项工具靠后；图标列和文本列使用固定布局保证左对齐。
        for key, icon, text in [
            ("home", "▦", "工具箱"),
            ("adb_tools", "▣", "ADB 工具"),
            ("adb_log", "▤", "日志抓取"),
            ("log_analysis", "◇", "日志分析"),
            ("ios_log", "◇", "iOS 日志"),
            ("compare", "≋", "文件比对"),
            ("config_validator", "表", "配置表校验"),
            ("time_tools", "时", "时间测试"),
            ("localization", "文", "多语言检查"),
            ("network", "◎", "网络弱网"),
            ("clicker", "⌁", "鼠标连点"),
        ]:
            self._nav_button(key, icon, text)

        bottom = ctk.CTkFrame(self.sidebar, fg_color=COLOR_SURFACE_2, corner_radius=14, border_width=1, border_color=COLOR_BORDER)
        bottom.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 22))
        bottom.grid_columnconfigure(0, weight=1)
        status = "管理员已开启" if is_admin() else "管理员未开启"
        dot_color = COLOR_SUCCESS if is_admin() else COLOR_DANGER

        row = ctk.CTkFrame(bottom, fg_color="transparent")
        row.grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 6))
        ctk.CTkLabel(row, text="", width=10, height=10, corner_radius=5, fg_color=dot_color).pack(side="left", padx=(0, 8))
        ctk.CTkLabel(row, text=status, text_color=COLOR_TEXT, font=self._font(13, "bold"), anchor="w").pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(bottom, text=f"版本 {APP_VERSION}", text_color=COLOR_MUTED, font=self._font(12), anchor="w").grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 10))

        ctk.CTkButton(
            bottom,
            text="管理员重启",
            height=38,
            corner_radius=12,
            fg_color=COLOR_SELECTED,
            hover_color=COLOR_HOVER,
            border_width=1,
            border_color=COLOR_BORDER,
            text_color=COLOR_TEXT,
            font=self._font(13, "bold"),
            command=relaunch_as_admin
        ).grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 14))

    def _nav_button(self, key, icon, text):
        parent = getattr(self, "nav_container", self.sidebar)
        item = ctk.CTkFrame(parent, height=38, corner_radius=10, fg_color="transparent")
        item.pack(fill="x", padx=16, pady=1)
        item.pack_propagate(False)
        item.grid_columnconfigure(0, minsize=24)
        item.grid_columnconfigure(1, weight=1)

        icon_label = ctk.CTkLabel(
            item,
            text=icon,
            width=24,
            anchor="center",
            text_color=COLOR_MUTED,
            font=self._font(14, "bold"),
        )
        icon_label.grid(row=0, column=0, sticky="nsw", padx=(12, 8), pady=0)
        text_label = ctk.CTkLabel(
            item,
            text=text,
            anchor="w",
            text_color=COLOR_MUTED,
            font=self._font(15, "bold"),
        )
        text_label.grid(row=0, column=1, sticky="nsew", padx=(0, 10), pady=0)

        def _go(_event=None, k=key):
            self.show_page(k)

        def _enter(_event=None):
            if self.current_page != key:
                item.configure(fg_color=COLOR_HOVER)

        def _leave(_event=None):
            if self.current_page != key:
                item.configure(fg_color="transparent")

        for widget in (item, icon_label, text_label):
            widget.bind("<Button-1>", _go)
            widget.bind("<Enter>", _enter)
            widget.bind("<Leave>", _leave)
        self.nav_buttons[key] = {"frame": item, "icon": icon_label, "label": text_label}

    def _set_nav_button_state(self, widgets, selected: bool):
        try:
            if isinstance(widgets, dict):
                widgets["frame"].configure(fg_color=COLOR_SELECTED if selected else "transparent")
                widgets["icon"].configure(text_color=COLOR_TEXT if selected else COLOR_MUTED)
                widgets["label"].configure(text_color=COLOR_TEXT if selected else COLOR_MUTED)
            else:
                widgets.configure(fg_color=COLOR_SELECTED if selected else "transparent", text_color=COLOR_TEXT if selected else COLOR_MUTED)
        except Exception:
            pass

    def _build_header(self):
        header = ctk.CTkFrame(self.main, fg_color="transparent", height=92)
        header.grid(row=0, column=0, sticky="ew", padx=24, pady=(24, 12))
        header.grid_columnconfigure(0, weight=1)

        title_wrap = ctk.CTkFrame(header, fg_color="transparent")
        title_wrap.grid(row=0, column=0, sticky="w")

        self.page_title = ctk.CTkLabel(title_wrap, text="", text_color=COLOR_TEXT, font=self._font(24, "bold"))
        self.page_title.pack(anchor="w")

        self.page_subtitle = ctk.CTkLabel(title_wrap, text="", text_color=COLOR_MUTED, font=self._font(14))
        self.page_subtitle.pack(anchor="w", pady=(5, 0))

        right = ctk.CTkFrame(header, fg_color="transparent")
        right.grid(row=0, column=1, sticky="e")
        ctk.CTkButton(right, text="反馈与建议", width=112, height=34, corner_radius=10, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_MUTED, font=self._font(13), command=self._show_feedback_dialog).pack(side="left", padx=(0, 12))
        self.page_badge = ctk.CTkLabel(
            right,
            text="●  已就绪",
            width=96,
            height=34,
            corner_radius=17,
            fg_color=COLOR_SURFACE_2,
            text_color=COLOR_SUCCESS,
            font=self._font(13, "bold")
        )
        self.page_badge.pack(side="left")
        self.page_badge_visible = True

    def _show_feedback_dialog(self):
        messagebox.showinfo(
            "反馈与建议",
            "想吐槽？请通过 IGGChat 召唤林凯，\n\n本人可能正在摸鱼但会假装认真倾听。",
            parent=self,
        )

    def _build_right_panel(self):
        wrap = ctk.CTkFrame(self.right_panel, fg_color=COLOR_SIDEBAR, corner_radius=0)
        wrap.pack(fill="both", expand=True, padx=18, pady=16)
        wrap.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(wrap, text="设置", text_color=COLOR_TEXT, font=self._font(18, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 10))

        self._panel_section_title(wrap, 1, "外观设置")
        theme_card = self._panel_card(wrap, 2)
        ctk.CTkLabel(theme_card, text="深色模式", text_color=COLOR_TEXT, font=self._font(14, "bold")).pack(anchor="w", padx=14, pady=(10, 2))
        ctk.CTkLabel(theme_card, text="启动后默认使用深色界面", text_color=COLOR_MUTED, font=self._font(12)).pack(anchor="w", padx=14, pady=(0, 6))
        self.theme_switch = ctk.CTkSwitch(theme_card, text="深色模式", variable=self.theme_mode_var, onvalue="深色模式", offvalue="浅色模式", command=self._on_theme_toggle, button_color=COLOR_ACCENT, progress_color=COLOR_ACCENT, text_color=COLOR_TEXT, font=self._font(13))
        self.theme_switch.select()
        self.theme_switch.pack(anchor="w", padx=14, pady=(0, 10))

        tray_card = self._panel_card(wrap, 3)
        ctk.CTkLabel(tray_card, text="系统托盘", text_color=COLOR_TEXT, font=self._font(14, "bold")).pack(anchor="w", padx=14, pady=(10, 2))
        tray_tip = "关闭可隐藏到托盘\n最小化保留任务栏"
        if not self.tray_available:
            tray_tip = "托盘依赖未安装\n运行 run.bat 安装"
        ctk.CTkLabel(
            tray_card,
            text=tray_tip,
            text_color=COLOR_MUTED,
            wraplength=170,
            justify="left",
            anchor="w",
            font=self._font(12),
        ).pack(anchor="w", fill="x", padx=14, pady=(0, 10))

        self._panel_section_title(wrap, 4, "快捷信息")
        hotkey_card = self._panel_card(wrap, 5)
        self._mini_info_row(hotkey_card, "连点热键", self.hotkey_var)
        self._mini_info_row(hotkey_card, "日志状态", self.adb_status_var)
        self._mini_info_row(hotkey_card, "录屏状态", self.adb_recording_status_var)
        self._mini_info_static(hotkey_card, "文件比对", "本地模式")
        self._mini_info_static(hotkey_card, "多语言检查", "本地模式")
        self._mini_info_static(hotkey_card, "配置表校验", "本地模式")
        self._panel_section_title(wrap, 6, "系统状态")
        state_card = self._panel_card(wrap, 7)
        self._mini_info_static(state_card, "管理员", "已开启" if is_admin() else "未开启")
        self._mini_info_static(state_card, "本地模式", "已启用")

        update_card = self._panel_card(wrap, 8)
        update_row = ctk.CTkFrame(update_card, fg_color="transparent")
        update_row.pack(fill="x", padx=12, pady=10)
        update_text = ctk.CTkFrame(update_row, fg_color="transparent")
        update_text.pack(side="left", fill="x", expand=True, padx=(2, 8))
        ctk.CTkLabel(
            update_text,
            text=f"在线更新  v{APP_VERSION}",
            text_color=COLOR_TEXT,
            anchor="w",
            font=self._font(12, "bold"),
        ).pack(fill="x")
        ctk.CTkLabel(
            update_text,
            textvariable=self.update_status_var,
            text_color=COLOR_MUTED,
            wraplength=116,
            justify="left",
            anchor="w",
            font=self._font(11),
        ).pack(fill="x", pady=(2, 0))
        self.update_check_button = ctk.CTkButton(
            update_row,
            text="检查更新",
            width=78,
            height=32,
            corner_radius=10,
            fg_color=COLOR_SURFACE_2,
            hover_color=COLOR_HOVER,
            border_width=1,
            border_color=COLOR_BORDER,
            text_color=COLOR_TEXT,
            font=self._font(12, "bold"),
            command=lambda: self.check_for_updates(silent=False),
        )
        self.update_check_button.pack(side="right")

    def _set_update_button_state(self, working: bool):
        button = getattr(self, "update_check_button", None)
        if button is None:
            return
        try:
            button.configure(
                state="disabled" if working else "normal",
                text="检查中…" if working else "检查更新",
            )
        except Exception:
            pass

    def check_for_updates(self, silent=False):
        if self.is_exiting or self.update_busy.is_set():
            return
        self.update_busy.set()
        self._set_update_button_state(True)
        self.update_status_var.set("正在检查更新…")

        def worker():
            try:
                info = fetch_update_info()
                self.after(0, lambda: self._on_update_check_success(info, silent))
            except Exception as exc:
                self.after(0, lambda message=str(exc): self._on_update_check_error(message, silent))

        threading.Thread(target=worker, daemon=True).start()

    def _finish_update_check(self):
        self.update_busy.clear()
        self._set_update_button_state(False)

    def _on_update_check_error(self, message, silent):
        self._finish_update_check()
        self.update_status_var.set("暂时无法检查更新")
        if not silent and not self.is_exiting:
            messagebox.showwarning("检查更新失败", message, parent=self)

    def _on_update_check_success(self, info, silent):
        self._finish_update_check()
        self.available_update_info = info
        if not is_newer_version(info.version, APP_VERSION):
            self.update_status_var.set("当前已是最新版本")
            if not silent:
                messagebox.showinfo("检查更新", f"当前版本 v{APP_VERSION} 已是最新版。", parent=self)
            return
        self.update_status_var.set(f"发现新版本 v{info.version}")
        if silent:
            try:
                if not self.winfo_viewable() or str(self.state()).lower() in {"iconic", "withdrawn"}:
                    return
            except Exception:
                return
        self._prompt_update(info)

    def _prompt_update(self, info):
        size_mb = info.size / 1024 / 1024
        published = f"\n发布时间：{info.published_at}" if info.published_at else ""
        message = (
            f"发现新版本 v{info.version}{published}\n"
            f"安装包大小：{size_mb:.1f} MB\n\n"
            f"{info.display_notes}\n\n"
            "是否现在下载？"
        )
        if messagebox.askyesno("测试助手更新", message, parent=self):
            self._start_update_download(info)

    def _start_update_download(self, info):
        if self.is_exiting or self.update_busy.is_set():
            return
        self.update_busy.set()
        self.update_cancel_event.clear()
        self._set_update_button_state(True)
        self.update_status_var.set(f"正在下载 v{info.version}…")
        self._open_update_progress_dialog(info)

        def progress(downloaded, total):
            try:
                self.after(0, lambda: self._update_download_progress(downloaded, total))
            except Exception:
                pass

        def worker():
            try:
                installer = download_installer(
                    info,
                    cancel_event=self.update_cancel_event,
                    progress_callback=progress,
                )
                self.after(0, lambda: self._on_update_download_success(info, installer))
            except DownloadCancelled:
                self.after(0, self._on_update_download_cancelled)
            except Exception as exc:
                self.after(0, lambda message=str(exc): self._on_update_download_error(message))

        threading.Thread(target=worker, daemon=True).start()

    def _open_update_progress_dialog(self, info):
        self._restore_window_for_modal()
        dialog = ctk.CTkToplevel(self)
        self.update_progress_dialog = dialog
        dialog.title(f"下载测试助手 v{info.version}")
        dialog.geometry("430x175")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.configure(fg_color=COLOR_BG)
        dialog.protocol("WM_DELETE_WINDOW", self._cancel_update_download)

        card = ctk.CTkFrame(dialog, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        card.pack(fill="both", expand=True, padx=14, pady=14)
        ctk.CTkLabel(card, text=f"正在下载 v{info.version}", text_color=COLOR_TEXT, font=self._font(15, "bold")).pack(anchor="w", padx=18, pady=(18, 8))
        self.update_progress_label_var.set("准备下载…")
        ctk.CTkLabel(card, textvariable=self.update_progress_label_var, text_color=COLOR_MUTED, font=self._font(12)).pack(anchor="w", padx=18)
        self.update_progress_bar = ctk.CTkProgressBar(card, height=12, progress_color=COLOR_ACCENT)
        self.update_progress_bar.pack(fill="x", padx=18, pady=(10, 12))
        self.update_progress_bar.set(0)
        ctk.CTkButton(
            card,
            text="取消下载",
            height=32,
            fg_color=COLOR_SURFACE_2,
            hover_color=COLOR_HOVER,
            border_width=1,
            border_color=COLOR_BORDER,
            text_color=COLOR_TEXT,
            command=self._cancel_update_download,
        ).pack(anchor="e", padx=18, pady=(0, 14))
        dialog.update_idletasks()
        dialog.grab_set()
        dialog.lift()
        dialog.focus_force()

    def _update_download_progress(self, downloaded, total):
        total = max(int(total or 0), 1)
        ratio = min(max(downloaded / total, 0.0), 1.0)
        if self.update_progress_bar is not None:
            self.update_progress_bar.set(ratio)
        self.update_progress_label_var.set(
            f"{downloaded / 1024 / 1024:.1f} MB / {total / 1024 / 1024:.1f} MB"
        )

    def _cancel_update_download(self):
        if self.update_busy.is_set():
            self.update_cancel_event.set()
            self.update_progress_label_var.set("正在取消下载…")

    def _close_update_progress_dialog(self):
        dialog = self.update_progress_dialog
        self.update_progress_dialog = None
        self.update_progress_bar = None
        if dialog is None:
            return
        try:
            dialog.grab_release()
        except Exception:
            pass
        try:
            dialog.destroy()
        except Exception:
            pass

    def _finish_update_download(self):
        self.update_busy.clear()
        self._set_update_button_state(False)
        self._close_update_progress_dialog()

    def _on_update_download_cancelled(self):
        self._finish_update_download()
        self.update_status_var.set("已取消更新下载")

    def _on_update_download_error(self, message):
        self._finish_update_download()
        self.update_status_var.set("更新下载失败")
        if not self.is_exiting:
            messagebox.showerror("下载更新失败", message, parent=self)

    def _on_update_download_success(self, info, installer_path):
        self._finish_update_download()
        self.update_status_var.set(f"v{info.version} 已下载并通过校验")
        if not messagebox.askyesno(
            "安装更新",
            "安装包已下载完成并通过 SHA-256 校验。\n\n"
            "现在安装将关闭测试助手，安装成功后会自动重新打开。是否继续？",
            parent=self,
        ):
            return
        try:
            launch_update_helper(installer_path, info.version)
        except UpdateError as exc:
            messagebox.showerror("无法安装更新", str(exc), parent=self)
            return
        self.update_status_var.set("正在退出并安装更新…")
        self.after(150, self.quit_app)

    def _panel_section_title(self, parent, row, text):
        ctk.CTkLabel(parent, text=text, text_color=COLOR_TEXT, font=self._font(14, "bold")).grid(row=row, column=0, sticky="w", pady=(4, 4))

    def _panel_card(self, parent, row):
        card = ctk.CTkFrame(parent, fg_color=COLOR_SURFACE, corner_radius=14, border_width=1, border_color=COLOR_BORDER)
        card.grid(row=row, column=0, sticky="ew", pady=(0, 8))
        return card

    def _mini_info_row(self, parent, label, variable):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=5)
        row.grid_columnconfigure(0, weight=1)
        row.grid_columnconfigure(1, weight=0, minsize=92)
        ctk.CTkLabel(row, text=label, text_color=COLOR_MUTED, anchor="w", font=self._font(13)).grid(row=0, column=0, sticky="w", padx=(0, 8))
        ctk.CTkLabel(row, textvariable=variable, text_color=COLOR_TEXT, anchor="e", justify="right", width=116, wraplength=116, font=self._font(13, "bold")).grid(row=0, column=1, sticky="e")

    def _mini_info_static(self, parent, label, value):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=5)
        row.grid_columnconfigure(0, weight=1)
        row.grid_columnconfigure(1, weight=0, minsize=92)
        ctk.CTkLabel(row, text=label, text_color=COLOR_MUTED, anchor="w", font=self._font(13)).grid(row=0, column=0, sticky="w", padx=(0, 8))
        ctk.CTkLabel(row, text=value, text_color=COLOR_TEXT, anchor="e", justify="right", width=116, wraplength=116, font=self._font(13, "bold")).grid(row=0, column=1, sticky="e")

    def _on_theme_toggle(self):
        mode = "dark" if self.theme_mode_var.get() == "深色模式" else "light"
        ctk.set_appearance_mode(mode)
        self.update_idletasks()
        self._configure_ttk_styles()
        self._refresh_ttk_theme_widgets()
        if hasattr(self, "theme_switch"):
            self.theme_switch.configure(text=self.theme_mode_var.get())
        self.configure(fg_color=COLOR_BG)
        self.show_page(self.current_page)

    def show_page(self, key):
        previous_page = getattr(self, "current_page", None)
        if previous_page == "ios_log" and key != "ios_log":
            # Cached pages keep their widget references; only stop active capture.
            try:
                self.ios_stop_log()
            except Exception:
                pass

        self.current_page = key

        for page, btn in self.nav_buttons.items():
            self._set_nav_button_state(btn, page == key)

        needs_build = self._recreate_content_container(key)

        if key == "home":
            self.page_title.configure(text="欢迎使用测试助手")
            self.page_subtitle.configure(text="高效的 QA 测试工具集合")
            self._set_badge(False)
            if needs_build:
                self._build_home_page()
        elif key == "clicker":
            self.page_title.configure(text="鼠标连点")
            self.page_subtitle.configure(text="自动点击工具，支持快捷键启动 / 停止")
            self._set_badge(self.clicker.running.is_set())
            if needs_build:
                self._build_clicker_page()
        elif key == "network":
            self.page_title.configure(text="网络弱网")
            self.page_subtitle.configure(text="为本机及模拟器流量模拟延迟、抖动和丢包")
            self._set_badge(self.network.running.is_set(), admin_required=True)
            if needs_build:
                self._build_network_page()
            self._refresh_network_button()
        elif key == "adb_log":
            self.page_title.configure(text="日志抓取")
            self.page_subtitle.configure(text="连接安卓设备后抓取指定包名的 ADB 日志")
            self._set_badge(self.adb_log.running.is_set())
            if needs_build:
                self._build_adb_log_page()
        elif key == "log_analysis":
            self.page_title.configure(text="日志分析")
            self.page_subtitle.configure(text="关键字监控、崩溃日志提取和异常报告导出")
            self._set_badge(self.adb_log.running.is_set())
            if needs_build:
                self._build_log_analysis_page()
        elif key == "adb_tools":
            self.page_title.configure(text="ADB 工具")
            self.page_subtitle.configure(text="设备检测、截图录屏、应用管理和 APK 信息查看")
            self._set_badge(self.adb_tools.is_recording())
            if needs_build:
                self._build_adb_tools_page()
        elif key == "ios_log":
            self.page_title.configure(text="iOS 日志")
            self.page_subtitle.configure(text="Windows 下通过 libimobiledevice 抓取 iPhone / iPad 实时日志")
            self._set_badge(self.ios_log_running.is_set())
            if needs_build:
                self._build_ios_log_page()
        elif key == "compare":
            self.page_title.configure(text="文件比对")
            self.page_subtitle.configure(text="比对参考文档（策划文档 / 配置说明）与待比对文件（实际配置文件 / 导出文件）")
            self._set_badge(self.compare_is_working)
            if needs_build:
                self._build_compare_page()
        elif key == "localization":
            self.page_title.configure(text="多语言检查")
            self.page_subtitle.configure(text="检查空翻译、残留中文、占位符、富文本标签、特殊符号和长度问题")
            self._set_badge(self.loc_is_working)
            if needs_build:
                self._build_localization_page()
        elif key == "config_validator":
            self.page_title.configure(text="配置表校验")
            self.page_subtitle.configure(text="检查主键、必填、数值、枚举、时间、奖励、引用和区间配置问题")
            self._set_badge(self.cfg_is_working)
            if needs_build:
                self._build_config_validator_page()
        elif key == "time_tools":
            self.page_title.configure(text="时间测试")
            self.page_subtitle.configure(text="时间戳转换、单位换算、活动时间、倒计时和跨天跨周跨月边界辅助")
            self._set_badge(False)
            if needs_build:
                self._build_time_tools_page()
        self._sync_legacy_task_states()
        self._refresh_current_task_ui()

    def _show_page_badge(self):
        if not hasattr(self, "page_badge"):
            return
        if not getattr(self, "page_badge_visible", True):
            self.page_badge.pack(side="left")
            self.page_badge_visible = True

    def _hide_page_badge(self):
        if not hasattr(self, "page_badge"):
            return
        if getattr(self, "page_badge_visible", True):
            self.page_badge.pack_forget()
            self.page_badge_visible = False

    def _set_badge(self, running=False, admin_required=False):
        if running:
            self._show_page_badge()
            self.page_badge.configure(text="●  运行中", fg_color=COLOR_SURFACE_2, text_color=COLOR_DANGER)
            return

        if admin_required:
            self._show_page_badge()
            if is_admin():
                self.page_badge.configure(text="●  已就绪", fg_color=COLOR_SURFACE_2, text_color=COLOR_SUCCESS)
            else:
                self.page_badge.configure(text="●  需管理员", fg_color=COLOR_SURFACE_2, text_color=COLOR_DANGER)
            return

        self._hide_page_badge()

    def _task_snapshot(self, name):
        state = getattr(self, "task_states", {}).get(name)
        return state.snapshot() if state is not None else None

    def _set_task_phase(self, name, phase, message=""):
        state = getattr(self, "task_states", {}).get(name)
        if state is None:
            return None
        snapshot = state.set(phase, message)
        try:
            self.after(0, self._refresh_current_task_ui)
        except Exception:
            pass
        return snapshot

    def _on_network_state(self, phase, message=""):
        self._set_task_phase("network", phase, message)
        labels = {
            TaskPhase.IDLE: "已停止",
            TaskPhase.STARTING: "正在启动…",
            TaskPhase.RUNNING: "运行中",
            TaskPhase.STOPPING: "正在停止…",
            TaskPhase.FAILED: f"运行异常：{message}" if message else "运行异常",
        }

        def update_status():
            try:
                self.network_status_var.set(labels.get(phase, message or str(phase)))
                self._refresh_network_button()
            except Exception:
                pass

        try:
            self.after(0, update_status)
        except Exception:
            pass

    def _sync_legacy_task_states(self):
        """Mirror legacy flags into one UI-facing task-state registry."""
        mappings = {
            "clicker": self.clicker.running.is_set(),
            "adb_log": self.adb_log.running.is_set(),
            "ios_log": self.ios_log_running.is_set(),
            "compare": bool(self.compare_is_working),
            "localization": bool(self.loc_is_working),
            "config_validator": bool(self.cfg_is_working),
        }
        for name, running in mappings.items():
            self.task_states[name].set(TaskPhase.RUNNING if running else TaskPhase.IDLE)

        adb_running = self.adb_tools_busy.is_set() or self.adb_tools.is_recording()
        adb_snapshot = self.task_states["adb_tools"].snapshot()
        if adb_snapshot.phase not in (TaskPhase.STARTING, TaskPhase.STOPPING, TaskPhase.FAILED):
            self.task_states["adb_tools"].set(TaskPhase.RUNNING if adb_running else TaskPhase.IDLE)

    def _set_task_badge(self, snapshot, admin_required=False):
        if snapshot is None:
            self._set_badge(False, admin_required=admin_required)
            return
        labels = {
            TaskPhase.STARTING: ("●  启动中", COLOR_ACCENT),
            TaskPhase.RUNNING: ("●  运行中", COLOR_DANGER),
            TaskPhase.STOPPING: ("●  停止中", COLOR_ACCENT),
            TaskPhase.FAILED: ("●  启动失败", COLOR_DANGER),
        }
        label = labels.get(snapshot.phase)
        if label:
            self._show_page_badge()
            self.page_badge.configure(text=label[0], fg_color=COLOR_SURFACE_2, text_color=label[1])
        else:
            self._set_badge(False, admin_required=admin_required)

    def _refresh_current_task_ui(self):
        page = getattr(self, "current_page", "home")
        page_tasks = {
            "clicker": ("clicker", self._refresh_clicker_button, False),
            "network": ("network", self._refresh_network_button, True),
            "adb_log": ("adb_log", self._refresh_adb_button, False),
            "log_analysis": ("adb_log", self._refresh_adb_button, False),
            "adb_tools": ("adb_tools", self._refresh_adb_tools_buttons, False),
            "ios_log": ("ios_log", self._refresh_ios_log_button, False),
            "compare": ("compare", self._refresh_compare_button, False),
            "localization": ("localization", self._refresh_localization_button, False),
            "config_validator": ("config_validator", self._refresh_config_validator_button, False),
        }
        item = page_tasks.get(page)
        if item is None:
            if page == "home":
                self._set_badge(False)
            return
        task_name, refresh, admin_required = item
        self._set_task_badge(self._task_snapshot(task_name), admin_required=admin_required)
        try:
            refresh()
        except Exception:
            pass

    def _card(self, parent):
        card = ctk.CTkFrame(parent, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        card.pack(fill="x", pady=(0, 16))
        card.grid_columnconfigure(1, weight=1)
        return card

    def _divider(self, parent, row):
        ctk.CTkFrame(parent, height=1, fg_color=COLOR_BORDER).grid(row=row, column=0, columnspan=3, sticky="ew", padx=24)

    def _entry_row(self, parent, row, label, variable, placeholder="", width=420):
        ctk.CTkLabel(parent, text=label, text_color=COLOR_TEXT, font=self._font(15, "bold")).grid(row=row, column=0, sticky="w", padx=(24, 22), pady=18)
        ctk.CTkEntry(parent, textvariable=variable, width=width, height=42, corner_radius=10, fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER, text_color=COLOR_TEXT, placeholder_text=placeholder, font=self._font(14)).grid(row=row, column=1, sticky="w", pady=18)

    def _option_row(self, parent, row, label, variable, values, width=420, command=None):
        ctk.CTkLabel(parent, text=label, text_color=COLOR_TEXT, font=self._font(15, "bold")).grid(row=row, column=0, sticky="w", padx=(24, 22), pady=18)
        ctk.CTkOptionMenu(parent, variable=variable, values=values, width=width, height=42, corner_radius=10, fg_color=COLOR_SURFACE_2, button_color=COLOR_SURFACE_2, button_hover_color=COLOR_HOVER, dropdown_fg_color=COLOR_SURFACE, dropdown_hover_color=COLOR_HOVER, dropdown_text_color=COLOR_TEXT, text_color=COLOR_TEXT, font=self._font(14), dropdown_font=self._font(14), command=command).grid(row=row, column=1, sticky="w", pady=18)


    # ---------- Time test tool page ----------
    def _time_now_string(self):
        return time.strftime("%Y-%m-%d %H:%M:%S")

    def _time_fmt_number(self, value):
        try:
            number = float(value)
        except Exception:
            return str(value)
        if abs(number - round(number)) < 1e-10:
            return str(int(round(number)))
        text = f"{number:.6f}".rstrip("0").rstrip(".")
        return text or "0"

    def _time_textbox(self, parent, row, key, height=260):
        box = ctk.CTkTextbox(
            parent,
            height=height,
            corner_radius=12,
            border_width=1,
            border_color=COLOR_BORDER,
            fg_color=COLOR_SURFACE_2,
            text_color=COLOR_TEXT,
            font=self._font(13),
        )
        box.grid(row=row, column=0, sticky="nsew", padx=20, pady=(0, 16))
        self.time_result_boxes[key] = box
        return box

    def _time_set_result(self, key, text):
        box = self.time_result_boxes.get(key)
        if box is None:
            return
        box.configure(state="normal")
        box.delete("1.0", "end")
        box.insert("end", text.strip() + "\n")
        box.configure(state="normal")

    def _time_copy_result(self, key):
        box = self.time_result_boxes.get(key)
        if box is None:
            return
        text = box.get("1.0", "end").strip()
        if not text:
            messagebox.showinfo("复制结果", "当前没有可复制的结果")
            return
        self.clipboard_clear()
        self.clipboard_append(text)
        messagebox.showinfo("复制结果", "已复制当前 Tab 结果")

    def _time_card(self, parent, title, row=0, column=0):
        card = ctk.CTkFrame(parent, fg_color=COLOR_SURFACE, corner_radius=14, border_width=1, border_color=COLOR_BORDER)
        card.grid(row=row, column=column, sticky="nsew", padx=(0, 8) if column == 0 else (8, 0), pady=(0, 12))
        card.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(card, text=title, text_color=COLOR_TEXT, font=self._font(16, "bold")).grid(row=0, column=0, columnspan=3, sticky="w", padx=18, pady=(14, 8))
        return card

    def _time_entry(self, parent, row, label, variable, placeholder="", width=220):
        ctk.CTkLabel(parent, text=label, text_color=COLOR_TEXT, font=self._font(13, "bold")).grid(row=row, column=0, sticky="w", padx=(18, 10), pady=7)
        entry = ctk.CTkEntry(parent, textvariable=variable, placeholder_text=placeholder, width=width, height=34, corner_radius=10, fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13))
        entry.grid(row=row, column=1, sticky="ew", padx=(0, 18), pady=7)
        return entry

    def _time_menu(self, parent, row, label, variable, values, width=220):
        ctk.CTkLabel(parent, text=label, text_color=COLOR_TEXT, font=self._font(13, "bold")).grid(row=row, column=0, sticky="w", padx=(18, 10), pady=7)
        ctk.CTkOptionMenu(parent, variable=variable, values=values, width=width, height=34, corner_radius=10, fg_color=COLOR_SURFACE_2, button_color=COLOR_SURFACE_2, button_hover_color=COLOR_HOVER, dropdown_fg_color=COLOR_SURFACE, dropdown_hover_color=COLOR_HOVER, dropdown_text_color=COLOR_TEXT, text_color=COLOR_TEXT, font=self._font(13), dropdown_font=self._font(13)).grid(row=row, column=1, sticky="ew", padx=(0, 18), pady=7)

    def _time_action_row(self, parent, row, buttons):
        action = ctk.CTkFrame(parent, fg_color="transparent")
        action.grid(row=row, column=0, columnspan=3, sticky="ew", padx=18, pady=(10, 16))
        for idx, (text, command, primary) in enumerate(buttons):
            ctk.CTkButton(
                action,
                text=text,
                width=110,
                height=32,
                corner_radius=10,
                fg_color=COLOR_ACCENT if primary else COLOR_SURFACE_2,
                hover_color=COLOR_ACCENT_HOVER if primary else COLOR_HOVER,
                border_width=0 if primary else 1,
                border_color=COLOR_BORDER,
                text_color="#FFFFFF" if primary else COLOR_TEXT,
                font=self._font(13, "bold"),
                command=command,
            ).pack(side="left", padx=(0, 10))

    def _build_time_tools_page(self):
        for child in self.content.winfo_children():
            try:
                child.destroy()
            except Exception:
                pass
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(0, weight=1)
        self.time_result_boxes = {}

        tabview = ctk.CTkTabview(self.content, fg_color=COLOR_SURFACE, segmented_button_fg_color=COLOR_SURFACE_2, segmented_button_selected_color=COLOR_ACCENT, segmented_button_selected_hover_color=COLOR_ACCENT_HOVER, segmented_button_unselected_color=COLOR_SURFACE_2, segmented_button_unselected_hover_color=COLOR_HOVER, text_color=COLOR_TEXT, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        tabview.grid(row=0, column=0, sticky="nsew")
        for name in ["时间戳转换", "单位换算", "时间差计算", "时间加减", "倒计时校验", "边界时间"]:
            tabview.add(name)

        self._build_timestamp_tab(tabview.tab("时间戳转换"))
        self._build_unit_tab(tabview.tab("单位换算"))
        self._build_diff_tab(tabview.tab("时间差计算"))
        self._build_add_tab(tabview.tab("时间加减"))
        self._build_countdown_tab(tabview.tab("倒计时校验"))
        self._build_boundary_tab(tabview.tab("边界时间"))

    def _build_timestamp_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)
        form = ctk.CTkFrame(tab, fg_color="transparent")
        form.grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 4))
        form.grid_columnconfigure((0, 1), weight=1, uniform="time_ts")

        left = self._time_card(form, "时间戳转日期时间", 0, 0)
        self._time_entry(left, 1, "时间戳", self.time_ts_input_var, "10 位秒级或 13 位毫秒级")
        self._time_menu(left, 2, "单位", self.time_ts_unit_var, ["自动识别", "秒", "毫秒"])
        self._time_action_row(left, 3, [("转换", self.time_convert_timestamp, True), ("复制结果", lambda: self._time_copy_result("timestamp"), False)])

        right = self._time_card(form, "日期时间转时间戳", 0, 1)
        self._time_entry(right, 1, "日期时间", self.time_dt_input_var, "2026-06-09 15:30:00")
        self._time_menu(right, 2, "时区", self.time_dt_tz_var, ["本地时区", "UTC", "UTC+8", "自定义 UTC 偏移"])
        self._time_entry(right, 3, "自定义偏移", self.time_custom_offset_var, "+8 / +08:00 / -05:30")
        self._time_action_row(right, 4, [("转换", self.time_convert_datetime, True), ("填入当前时间", lambda: self.time_dt_input_var.set(self._time_now_string()), False), ("复制结果", lambda: self._time_copy_result("timestamp"), False)])

        result_card = ctk.CTkFrame(tab, fg_color=COLOR_SURFACE, corner_radius=14, border_width=1, border_color=COLOR_BORDER)
        result_card.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 14))
        result_card.grid_columnconfigure(0, weight=1)
        result_card.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(result_card, text="转换结果", text_color=COLOR_TEXT, font=self._font(16, "bold")).grid(row=0, column=0, sticky="w", padx=20, pady=(14, 8))
        self._time_textbox(result_card, 1, "timestamp", height=250)
        self._time_set_result("timestamp", "提示：可输入 10 位秒级或 13 位毫秒级时间戳，也可输入日期时间转为秒级 / 毫秒级时间戳。")

    def _build_unit_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)
        form = ctk.CTkFrame(tab, fg_color=COLOR_SURFACE, corner_radius=14, border_width=1, border_color=COLOR_BORDER)
        form.grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 12))
        form.grid_columnconfigure(1, weight=1)
        self._time_entry(form, 0, "输入数值", self.time_unit_value_var, "支持小数")
        self._time_menu(form, 1, "输入单位", self.time_unit_from_var, ["秒", "分钟", "小时", "天", "周", "月", "年"])
        self._time_action_row(form, 2, [("换算", self.time_convert_unit_values, True), ("复制结果", lambda: self._time_copy_result("unit"), False)])
        result = ctk.CTkFrame(tab, fg_color=COLOR_SURFACE, corner_radius=14, border_width=1, border_color=COLOR_BORDER)
        result.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 14))
        result.grid_columnconfigure(0, weight=1)
        result.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(result, text="换算结果", text_color=COLOR_TEXT, font=self._font(16, "bold")).grid(row=0, column=0, sticky="w", padx=20, pady=(14, 8))
        self._time_textbox(result, 1, "unit", height=330)
        self.time_convert_unit_values()

    def _build_diff_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)
        form = ctk.CTkFrame(tab, fg_color=COLOR_SURFACE, corner_radius=14, border_width=1, border_color=COLOR_BORDER)
        form.grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 12))
        form.grid_columnconfigure(1, weight=1)
        self._time_entry(form, 0, "开始时间", self.time_diff_start_var, "2026-06-09 15:30:00")
        self._time_entry(form, 1, "结束时间", self.time_diff_end_var, "2026-06-12 19:42:05")
        self._time_menu(form, 2, "时区", self.time_default_tz_var, ["本地时区", "UTC", "UTC+8", "自定义 UTC 偏移"])
        self._time_entry(form, 3, "自定义偏移", self.time_custom_offset_var, "+8 / +08:00 / -05:30")
        self._time_action_row(form, 4, [("计算差值", self.time_calculate_difference, True), ("开始=当前", lambda: self.time_diff_start_var.set(self._time_now_string()), False), ("结束=当前", lambda: self.time_diff_end_var.set(self._time_now_string()), False), ("复制结果", lambda: self._time_copy_result("diff"), False)])
        result = ctk.CTkFrame(tab, fg_color=COLOR_SURFACE, corner_radius=14, border_width=1, border_color=COLOR_BORDER)
        result.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 14))
        result.grid_columnconfigure(0, weight=1)
        result.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(result, text="时间差结果", text_color=COLOR_TEXT, font=self._font(16, "bold")).grid(row=0, column=0, sticky="w", padx=20, pady=(14, 8))
        self._time_textbox(result, 1, "diff", height=300)
        self._time_set_result("diff", "提示：用于活动持续时间、礼包倒计时、任务刷新间隔、日志时间差排查。")

    def _build_add_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)
        form = ctk.CTkFrame(tab, fg_color=COLOR_SURFACE, corner_radius=14, border_width=1, border_color=COLOR_BORDER)
        form.grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 12))
        form.grid_columnconfigure(1, weight=1)
        self._time_entry(form, 0, "基准时间", self.time_add_base_var, "2026-06-09 15:30:00")
        self._time_menu(form, 1, "方向", self.time_add_direction_var, ["增加", "减少"])

        ctk.CTkLabel(form, text="时长", text_color=COLOR_TEXT, font=self._font(13, "bold")).grid(row=2, column=0, sticky="w", padx=(18, 10), pady=7)
        duration_frame = ctk.CTkFrame(form, fg_color="transparent")
        duration_frame.grid(row=2, column=1, sticky="ew", padx=(0, 18), pady=7)
        for idx in range(8):
            duration_frame.grid_columnconfigure(idx, weight=0)
        duration_frame.grid_columnconfigure(8, weight=1)
        self._time_duration_entry(duration_frame, 0, self.time_add_days_var, "7", "天")
        self._time_duration_entry(duration_frame, 2, self.time_add_hours_var, "0", "时")
        self._time_duration_entry(duration_frame, 4, self.time_add_minutes_var, "0", "分")
        self._time_duration_entry(duration_frame, 6, self.time_add_seconds_var, "0", "秒")

        self._time_action_row(form, 3, [("计算", self.time_add_or_subtract, True), ("基准=当前", lambda: self.time_add_base_var.set(self._time_now_string()), False), ("复制结果", lambda: self._time_copy_result("add"), False)])
        result = ctk.CTkFrame(tab, fg_color=COLOR_SURFACE, corner_radius=14, border_width=1, border_color=COLOR_BORDER)
        result.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 14))
        result.grid_columnconfigure(0, weight=1)
        result.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(result, text="时间加减结果", text_color=COLOR_TEXT, font=self._font(16, "bold")).grid(row=0, column=0, sticky="w", padx=20, pady=(14, 8))
        self._time_textbox(result, 1, "add", height=300)
        self._time_set_result("add", "提示：用于当前时间 + 7 天活动结束、当前时间 - 1 天跨天验证、服务器时间推进后状态预期计算；直接填写天 / 时 / 分 / 秒。")

    def _time_duration_entry(self, parent, column, variable, placeholder, unit_text):
        entry = ctk.CTkEntry(parent, textvariable=variable, placeholder_text=placeholder, width=96, height=34, corner_radius=10, fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13), justify="center")
        entry.grid(row=0, column=column, sticky="w", padx=(0, 6), pady=0)
        ctk.CTkLabel(parent, text=unit_text, text_color=COLOR_MUTED, font=self._font(13, "bold")).grid(row=0, column=column + 1, sticky="w", padx=(0, 14), pady=0)
        return entry

    def _build_countdown_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)
        form = ctk.CTkFrame(tab, fg_color=COLOR_SURFACE, corner_radius=14, border_width=1, border_color=COLOR_BORDER)
        form.grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 12))
        form.grid_columnconfigure(1, weight=1)
        self._time_entry(form, 0, "当前时间", self.time_countdown_now_var, "留空使用本机当前时间")
        self._time_entry(form, 1, "结束时间", self.time_countdown_end_var, "2026-06-16 15:30:00")
        self._time_menu(form, 2, "时区", self.time_default_tz_var, ["本地时区", "UTC", "UTC+8", "自定义 UTC 偏移"])
        self._time_entry(form, 3, "自定义偏移", self.time_custom_offset_var, "+8 / +08:00 / -05:30")
        self._time_action_row(form, 4, [("校验倒计时", self.time_check_countdown, True), ("当前=本机", lambda: self.time_countdown_now_var.set(self._time_now_string()), False), ("复制结果", lambda: self._time_copy_result("countdown"), False)])
        result = ctk.CTkFrame(tab, fg_color=COLOR_SURFACE, corner_radius=14, border_width=1, border_color=COLOR_BORDER)
        result.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 14))
        result.grid_columnconfigure(0, weight=1)
        result.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(result, text="倒计时结果", text_color=COLOR_TEXT, font=self._font(16, "bold")).grid(row=0, column=0, sticky="w", padx=20, pady=(14, 8))
        self._time_textbox(result, 1, "countdown", height=300)
        self._time_set_result("countdown", "提示：当前时间留空时使用本机当前时间，只读取时间，不会修改系统时间。")

    def _build_boundary_tab(self, tab):
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=1)
        form = ctk.CTkFrame(tab, fg_color=COLOR_SURFACE, corner_radius=14, border_width=1, border_color=COLOR_BORDER)
        form.grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 12))
        form.grid_columnconfigure(1, weight=1)
        self._time_entry(form, 0, "指定时间", self.time_boundary_base_var, "留空使用本机当前时间")
        self._time_menu(form, 1, "时区", self.time_default_tz_var, ["本地时区", "UTC", "UTC+8", "自定义 UTC 偏移"])
        self._time_entry(form, 2, "自定义偏移", self.time_custom_offset_var, "+8 / +08:00 / -05:30")
        self._time_action_row(form, 3, [("计算边界", self.time_calculate_boundaries, True), ("指定=当前", lambda: self.time_boundary_base_var.set(self._time_now_string()), False), ("复制结果", lambda: self._time_copy_result("boundary"), False)])
        result = ctk.CTkFrame(tab, fg_color=COLOR_SURFACE, corner_radius=14, border_width=1, border_color=COLOR_BORDER)
        result.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 14))
        result.grid_columnconfigure(0, weight=1)
        result.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(result, text="边界时间结果", text_color=COLOR_TEXT, font=self._font(16, "bold")).grid(row=0, column=0, sticky="w", padx=20, pady=(14, 8))
        self._time_textbox(result, 1, "boundary", height=330)
        self.time_calculate_boundaries()

    def time_convert_timestamp(self):
        try:
            data = timestamp_to_datetime(self.time_ts_input_var.get(), self.time_ts_unit_var.get())
            text = "\n".join([
                "【时间戳转日期时间】",
                f"输入时间戳：{self.time_ts_input_var.get().strip()}",
                f"识别单位：{data['unit']}",
                f"本地时间：{data['local']}",
                f"UTC 时间：{data['utc']}",
                f"秒级时间戳：{data['seconds']}",
                f"毫秒级时间戳：{data['milliseconds']}",
                "说明：10 位通常为秒级，13 位通常为毫秒级。",
            ])
        except Exception as exc:
            text = f"输入错误：{exc}"
        self._time_set_result("timestamp", text)

    def time_convert_datetime(self):
        try:
            data = datetime_to_timestamps(self.time_dt_input_var.get(), self.time_dt_tz_var.get(), self.time_custom_offset_var.get())
            text = "\n".join([
                "【日期时间转时间戳】",
                f"输入时间：{self.time_dt_input_var.get().strip()}",
                f"使用时区：{data['timezone']}",
                f"秒级时间戳：{data['seconds']}",
                f"毫秒级时间戳：{data['milliseconds']}",
                f"UTC 时间：{data['utc']}",
                f"本地时间：{data['local']}",
            ])
        except Exception as exc:
            text = f"输入错误：{exc}"
        self._time_set_result("timestamp", text)

    def time_convert_unit_values(self):
        try:
            values = convert_time_units(self.time_unit_value_var.get(), self.time_unit_from_var.get())
            lines = ["【时间单位换算】", f"输入：{self.time_unit_value_var.get().strip()} {self.time_unit_from_var.get()}"]
            for unit in ["秒", "分钟", "小时", "天", "周", "月", "年"]:
                suffix = "（按 30 天估算）" if unit == "月" else "（按 365 天估算）" if unit == "年" else ""
                lines.append(f"{unit}：{self._time_fmt_number(values[unit])}{suffix}")
            text = "\n".join(lines)
        except Exception as exc:
            text = f"输入错误：{exc}"
        self._time_set_result("unit", text)

    def time_calculate_difference(self):
        try:
            data = time_difference(self.time_diff_start_var.get(), self.time_diff_end_var.get(), self.time_default_tz_var.get(), self.time_custom_offset_var.get())
            earlier = "是" if data["is_negative"] else "否"
            signed_seconds = int(data["seconds"])
            lines = [
                "【时间差计算】",
                f"开始时间：{self.time_diff_start_var.get().strip()}",
                f"结束时间：{self.time_diff_end_var.get().strip()}",
                f"结束时间早于开始时间：{earlier}",
                f"相差秒数：{signed_seconds}",
                f"相差分钟：{self._time_fmt_number(data['minutes'])}",
                f"相差小时：{self._time_fmt_number(data['hours'])}",
                f"相差天数：{self._time_fmt_number(data['days'])}",
                f"可读格式：{data['readable']}",
            ]
            text = "\n".join(lines)
        except Exception as exc:
            text = f"输入错误：{exc}"
        self._time_set_result("diff", text)

    def time_add_or_subtract(self):
        try:
            data = add_or_subtract_time_parts(
                self.time_add_base_var.get(),
                self.time_add_direction_var.get(),
                self.time_add_days_var.get(),
                self.time_add_hours_var.get(),
                self.time_add_minutes_var.get(),
                self.time_add_seconds_var.get(),
                self.time_default_tz_var.get(),
                self.time_custom_offset_var.get(),
            )
            lines = [
                "【时间加减计算】",
                f"基准时间：{self.time_add_base_var.get().strip()}",
                f"计算方式：{self.time_add_direction_var.get()} {data['duration_display']}",
                f"计算后时间：{data['datetime']}",
                f"输入时长：{data['duration_display']}",
                f"换算秒数：{data['duration_seconds']} 秒",
                f"秒级时间戳：{data['seconds']}",
                f"毫秒级时间戳：{data['milliseconds']}",
            ]
            text = "\n".join(lines)
        except Exception as exc:
            text = f"输入错误：{exc}"
        self._time_set_result("add", text)

    def time_check_countdown(self):
        try:
            data = countdown(self.time_countdown_now_var.get(), self.time_countdown_end_var.get(), self.time_default_tz_var.get(), self.time_custom_offset_var.get())
            if data["ended"]:
                text = "\n".join([
                    "【倒计时校验】",
                    f"当前时间：{self.time_countdown_now_var.get().strip() or '本机当前时间'}",
                    f"结束时间：{self.time_countdown_end_var.get().strip()}",
                    "状态：已结束",
                    "剩余秒数：0",
                ])
            else:
                text = "\n".join([
                    "【倒计时校验】",
                    f"当前时间：{self.time_countdown_now_var.get().strip() or '本机当前时间'}",
                    f"结束时间：{self.time_countdown_end_var.get().strip()}",
                    "状态：未结束",
                    f"剩余秒数：{data['seconds']}",
                    f"剩余分钟：{self._time_fmt_number(data['minutes'])}",
                    f"剩余小时：{self._time_fmt_number(data['hours'])}",
                    f"剩余天数：{self._time_fmt_number(data['days'])}",
                    f"DD 天 HH:MM:SS：{data['dd_hhmmss']}",
                    f"HH:MM:SS：{data['hhmmss']}",
                ])
        except Exception as exc:
            text = f"输入错误：{exc}"
        self._time_set_result("countdown", text)

    def time_calculate_boundaries(self):
        try:
            data = boundary_times(self.time_boundary_base_var.get(), self.time_default_tz_var.get(), self.time_custom_offset_var.get())
            lines = ["【跨天 / 跨周 / 跨月边界】", f"指定时间：{self.time_boundary_base_var.get().strip() or '本机当前时间'}"]
            for label, dt in data.items():
                seconds, milliseconds = timestamp_pair(dt)
                lines.append("")
                lines.append(label)
                lines.append(f"日期时间：{format_datetime(dt)}")
                lines.append(f"秒级时间戳：{seconds}")
                lines.append(f"毫秒级时间戳：{milliseconds}")
            text = "\n".join(lines)
        except Exception as exc:
            text = f"输入错误：{exc}"
        self._time_set_result("boundary", text)

    # ---------- Home page ----------
    def _build_home_page(self):
        """Build a compact fixed home page.

        The previous home page used CTkScrollableFrame. On some Windows machines,
        scrolling many CTk cards in a canvas caused label text trails that looked like
        duplicated descriptions. This version avoids whole-page scrolling and rebuilds
        a clean fixed grid every time the page is opened.
        """
        for child in self.content.winfo_children():
            try:
                child.destroy()
            except Exception:
                pass

        wrap = ctk.CTkFrame(self.content, fg_color=COLOR_BG, corner_radius=0)
        wrap.grid(row=0, column=0, sticky="nsew")
        wrap.grid_columnconfigure((0, 1, 2), weight=1, uniform="home_cards")
        wrap.grid_rowconfigure(1, weight=1)

        hero = ctk.CTkFrame(wrap, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        hero.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        hero.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            hero,
            text="工具箱",
            text_color=COLOR_TEXT,
            fg_color=COLOR_SURFACE,
            font=self._font(22, "bold"),
        ).grid(row=0, column=0, sticky="w", padx=22, pady=(16, 2))
        ctk.CTkLabel(
            hero,
            text="选择需要的测试工具；高频 ADB、日志、iOS 与配置检查工具已优先展示。",
            text_color=COLOR_MUTED,
            fg_color=COLOR_SURFACE,
            wraplength=620,
            justify="left",
            font=self._font(14),
        ).grid(row=1, column=0, sticky="w", padx=22, pady=(0, 16))

        cards = ctk.CTkFrame(wrap, fg_color=COLOR_BG, corner_radius=0)
        cards.grid(row=1, column=0, columnspan=3, sticky="nsew")
        cards.grid_columnconfigure((0, 1, 2), weight=1, uniform="home_cards")
        for r in range(4):
            cards.grid_rowconfigure(r, weight=1, uniform="home_rows")

        tool_items = [
            (0, 0, "ADB 工具", "设备检测、截图录屏、应用管理", "adb_tools", "进入 ADB"),
            (0, 1, "日志抓取", "Android ADB logcat 日志抓取", "adb_log", "进入日志"),
            (0, 2, "日志分析", "关键字监控、崩溃提取、报告导出", "log_analysis", "进入分析"),
            (1, 0, "iOS 日志", "Windows 下抓取 iOS 实时日志", "ios_log", "进入 iOS"),
            (1, 1, "文件比对", "策划、配置和翻译表比对", "compare", "进入比对"),
            (1, 2, "配置表校验", "ID、必填、数值、枚举、时间、奖励", "config_validator", "进入校验"),
            (2, 0, "时间测试", "时间戳、倒计时、跨天周月边界", "time_tools", "进入时间"),
            (2, 1, "多语言检查", "空翻译、中文残留、占位符、标签", "localization", "进入检查"),
            (2, 2, "网络弱网", "指定 IP、协议、方向、延迟与丢包", "network", "进入弱网"),
            (3, 0, "鼠标连点", "左/中/右键连点，F 键热键", "clicker", "进入连点"),
        ]
        for item in tool_items:
            self._tool_card(cards, *item)

        recent = ctk.CTkFrame(wrap, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        recent.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(12, 0))
        recent.grid_columnconfigure((0, 1, 2, 3), weight=1, uniform="recent_status")
        ctk.CTkLabel(
            recent,
            text="最近状态",
            text_color=COLOR_TEXT,
            fg_color=COLOR_SURFACE,
            font=self._font(16, "bold"),
        ).grid(row=0, column=0, sticky="w", padx=18, pady=(12, 4), columnspan=4)
        recent_items = [
            ("连点", "运行中" if self.clicker.running.is_set() else "待命"),
            ("弱网", "运行中" if self.network.running.is_set() else "待命"),
            ("日志", "运行中" if self.adb_log.running.is_set() else "待命"),
            ("ADB 录屏", "运行中" if self.adb_tools.is_recording() else "待命"),
            ("iOS 日志", "运行中" if self.ios_log_running.is_set() else "待命"),
            ("文件比对", "运行中" if self.compare_is_working else "待命"),
            ("多语言", "运行中" if self.loc_is_working else "待命"),
            ("配置校验", "运行中" if self.cfg_is_working else "待命"),
        ]
        for idx, (label, value) in enumerate(recent_items):
            self._home_status_chip(recent, idx + 1, label, value)

    def _tool_card(self, parent, row, column, title, desc, page, button_text):
        # v2.4.20 首页新增“时间测试”后，4 行工具卡在 1400x850 / 1280x760 下会被压缩。
        # 按钮不再单独占用第三行，改为右侧固定按钮，避免底部被裁切。
        card = ctk.CTkFrame(parent, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER, height=86)
        card.grid(row=row, column=column, sticky="nsew", padx=(0, 8) if column == 0 else (8, 8) if column == 1 else (8, 0), pady=(0, 10))
        card.grid_propagate(False)
        card.grid_columnconfigure(0, weight=1, minsize=0)
        card.grid_columnconfigure(1, weight=0, minsize=92)
        card.grid_rowconfigure((0, 1), weight=1)
        ctk.CTkLabel(
            card,
            text=title,
            text_color=COLOR_TEXT,
            fg_color=COLOR_SURFACE,
            font=self._font(16, "bold"),
        ).grid(row=0, column=0, sticky="sw", padx=(14, 4), pady=(8, 0))
        ctk.CTkLabel(
            card,
            text=desc,
            text_color=COLOR_MUTED,
            fg_color=COLOR_SURFACE,
            wraplength=122,
            justify="left",
            font=self._font(11),
        ).grid(row=1, column=0, sticky="nw", padx=(14, 4), pady=(0, 8))
        ctk.CTkButton(
            card,
            text=button_text,
            width=82,
            height=30,
            corner_radius=10,
            fg_color=COLOR_ACCENT,
            hover_color=COLOR_ACCENT_HOVER,
            text_color="#FFFFFF",
            font=self._font(13, "bold"),
            command=lambda page=page: self.show_page(page),
        ).grid(row=0, column=1, rowspan=2, sticky="e", padx=(2, 10), pady=0)

    def _home_status_chip(self, parent, index, label, value):
        row = 1 + (index - 1) // 4
        column = (index - 1) % 4
        chip = ctk.CTkFrame(parent, fg_color=COLOR_SURFACE_2, corner_radius=10, border_width=1, border_color=COLOR_BORDER)
        chip.grid(row=row, column=column, sticky="ew", padx=(18 if column == 0 else 6, 18 if column == 3 else 6), pady=(4, 10 if row == 2 else 4))
        chip.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            chip,
            text=label,
            text_color=COLOR_MUTED,
            fg_color=COLOR_SURFACE_2,
            anchor="w",
            font=self._font(12),
        ).grid(row=0, column=0, sticky="w", padx=10, pady=(7, 0))
        ctk.CTkLabel(
            chip,
            text=value,
            text_color=COLOR_TEXT,
            fg_color=COLOR_SURFACE_2,
            anchor="w",
            font=self._font(13, "bold"),
        ).grid(row=1, column=0, sticky="w", padx=10, pady=(0, 7))

    # ---------- Clicker page ----------
    def _build_clicker_page(self):
        wrap = ctk.CTkFrame(self.content, fg_color=COLOR_BG)
        wrap.grid(row=0, column=0, sticky="nsew")
        wrap.grid_columnconfigure(0, weight=1)

        card = self._card(wrap)
        self._click_type_row(card)
        self._divider(card, 1)
        self._option_row(card, 2, "点击间隔", self.click_interval_var, ["自定义", "高效模式（每秒10次）", "极速模式（每秒100次）"], command=self._on_click_interval_changed)
        self._divider(card, 3)
        self._option_row(card, 4, "启动/停止热键", self.hotkey_var, ["F6", "F7", "F8", "F9", "F10", "F11", "F12"], command=lambda _: self._refresh_click_tip())

        action = ctk.CTkFrame(wrap, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        action.pack(fill="both", expand=True)
        action.grid_columnconfigure(0, weight=1)
        action.grid_rowconfigure(0, weight=1)

        center = ctk.CTkFrame(action, fg_color="transparent")
        center.grid(row=0, column=0, pady=50)

        self.click_start_button = ctk.CTkButton(center, text="▶  开始连点" if not self.clicker.running.is_set() else "■  停止连点", width=420, height=60, corner_radius=14, fg_color=COLOR_ACCENT if not self.clicker.running.is_set() else COLOR_DANGER, hover_color=COLOR_ACCENT_HOVER if not self.clicker.running.is_set() else COLOR_DANGER_HOVER, text_color="#FFFFFF", font=self._font(20, "bold"), command=self.toggle_clicker)
        self.click_start_button.pack(pady=(4, 18))

        self.click_tip_label = ctk.CTkLabel(center, text=self._click_tip_text(), text_color=COLOR_MUTED, font=self._font(16))
        self.click_tip_label.pack()

    def _click_type_row(self, parent):
        ctk.CTkLabel(parent, text="点击类型", text_color=COLOR_TEXT, font=self._font(15, "bold")).grid(row=0, column=0, sticky="w", padx=(24, 22), pady=18)

        box = ctk.CTkFrame(parent, fg_color="transparent")
        box.grid(row=0, column=1, sticky="w", pady=18)

        for text, val in [("鼠标左键", "left"), ("鼠标中键", "middle"), ("鼠标右键", "right")]:
            ctk.CTkRadioButton(box, text=text, value=val, variable=self.click_button_var, command=self._refresh_click_tip, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(14)).pack(side="left", padx=(0, 42))

    def _click_button_name(self):
        return {"left": "左键", "middle": "中键", "right": "右键"}.get(self.click_button_var.get(), "左键")

    def _click_tip_text(self):
        if self.clicker.running.is_set():
            return f"正在执行{self._click_button_name()}连点"
        return f"按  {self.hotkey_var.get()}  键开始{self._click_button_name()}连点"

    def _refresh_click_tip(self, *_):
        if hasattr(self, "click_tip_label"):
            self.click_tip_label.configure(text=self._click_tip_text())

    def _refresh_clicker_button(self):
        button = getattr(self, "click_start_button", None)
        if button is None:
            return
        try:
            if not button.winfo_exists():
                return
            running = self.clicker.running.is_set()
            button.configure(
                text="■  停止连点" if running else "▶  开始连点",
                fg_color=COLOR_DANGER if running else COLOR_ACCENT,
                hover_color=COLOR_DANGER_HOVER if running else COLOR_ACCENT_HOVER,
            )
            self._refresh_click_tip()
        except Exception:
            pass

    def _on_click_interval_changed(self, value):
        if value == "自定义":
            dialog = ctk.CTkInputDialog(title="自定义点击间隔", text="请输入点击间隔，单位毫秒，最小 10")
            raw = dialog.get_input()
            if raw is None:
                self.click_interval_var.set("极速模式（每秒100次）")
                return
            try:
                ms = int(raw.strip())
                if ms < 10 or ms > 10000:
                    raise ValueError
                self.clicker_custom_interval = ms / 1000
                self.click_interval_var.set(f"自定义（{ms}毫秒）")
            except Exception:
                messagebox.showerror("参数错误", "点击间隔必须是 10-10000 之间的整数")
                self.click_interval_var.set("极速模式（每秒100次）")

    def _click_interval_seconds(self):
        value = self.click_interval_var.get()
        if value.startswith("高效模式"):
            return 0.1
        if value.startswith("极速模式"):
            return 0.01
        return self.clicker_custom_interval

    def toggle_clicker(self):
        if self.clicker.running.is_set():
            self.clicker.stop()
        else:
            try:
                self.clicker.start(self._click_interval_seconds(), self.click_button_var.get())
            except Exception as exc:
                messagebox.showerror("启动失败", str(exc))
                return
        self._set_task_phase(
            "clicker",
            TaskPhase.RUNNING if self.clicker.running.is_set() else TaskPhase.IDLE,
        )
        self._refresh_clicker_button()

    # ---------- Network page ----------
    def _build_network_page(self):
        self.network_parameter_widgets = []
        wrap = ctk.CTkFrame(self.content, fg_color=COLOR_BG)
        wrap.grid(row=0, column=0, sticky="nsew")
        wrap.grid_columnconfigure(0, weight=1)
        wrap.grid_rowconfigure(2, weight=1)

        tabs = ctk.CTkTabview(
            wrap,
            height=280,
            corner_radius=14,
            border_width=1,
            border_color=COLOR_BORDER,
            fg_color=COLOR_SURFACE,
            segmented_button_fg_color=COLOR_SURFACE_2,
            segmented_button_selected_color=COLOR_ACCENT,
            segmented_button_selected_hover_color=COLOR_ACCENT_HOVER,
        )
        tabs.pack(fill="x", pady=(0, 10))
        basic_tab = tabs.add("基础设置")
        advanced_tab = tabs.add("进阶模拟")
        for tab in (basic_tab, advanced_tab):
            tab.grid_columnconfigure((0, 1), weight=1, uniform="network_settings")

        def inline_entry(parent, row, col, label, variable, placeholder=""):
            box = ctk.CTkFrame(parent, fg_color="transparent")
            box.grid(row=row, column=col, sticky="ew", padx=(0, 12), pady=4)
            box.grid_columnconfigure(1, weight=1)
            ctk.CTkLabel(box, text=label, width=150, text_color=COLOR_MUTED, font=self._font(12), anchor="w").grid(row=0, column=0, sticky="w", padx=(0, 8))
            entry = ctk.CTkEntry(box, textvariable=variable, placeholder_text=placeholder, height=34, corner_radius=9, fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(12))
            entry.grid(row=0, column=1, sticky="ew")
            return entry

        self.network_parameter_widgets.append(
            inline_entry(basic_tab, 0, 0, "目标 IP", self.target_ip_var, "留空 = 全部流量")
        )
        self.network_parameter_widgets.append(
            inline_entry(basic_tab, 0, 1, "高级规则", self.advanced_filter_var, "可选，自动加安全保护")
        )

        def compact_option(parent, row, col, label, variable, values):
            box = ctk.CTkFrame(parent, fg_color="transparent")
            box.grid(row=row, column=col, sticky="ew", padx=(0, 12), pady=4)
            box.grid_columnconfigure(1, weight=1)
            ctk.CTkLabel(box, text=label, width=150, text_color=COLOR_MUTED, font=self._font(12), anchor="w").grid(row=0, column=0, sticky="w", padx=(0, 8))
            menu = ctk.CTkOptionMenu(
                box,
                variable=variable,
                values=values,
                height=34,
                corner_radius=9,
                fg_color=COLOR_SURFACE_2,
                button_color=COLOR_SELECTED,
                button_hover_color=COLOR_HOVER,
                text_color=COLOR_TEXT,
                font=self._font(12),
            )
            menu.grid(row=0, column=1, sticky="ew")
            return menu

        self.network_parameter_widgets.append(compact_option(basic_tab, 1, 0, "协议类型", self.protocol_var, ["全部", "TCP", "UDP", "ICMP"]))
        self.network_parameter_widgets.append(compact_option(basic_tab, 1, 1, "方向", self.direction_var, ["全部", "出站", "入站"]))
        self.network_parameter_widgets.append(inline_entry(basic_tab, 2, 0, "基础单向延迟（ms）", self.delay_var, "例如 300"))
        self.network_parameter_widgets.append(inline_entry(basic_tab, 2, 1, "延迟抖动（±ms）", self.jitter_var, "0 = 固定延迟"))
        self.network_parameter_widgets.append(inline_entry(basic_tab, 3, 0, "随机丢包概率", self.loss_var, "0～100"))
        self.network_parameter_widgets.append(inline_entry(basic_tab, 3, 1, "数据包乱序概率", self.reorder_var, "0～100，默认关闭"))
        delay_hint = ctk.CTkLabel(
            basic_tab,
            textvariable=self.network_delay_hint_var,
            text_color=COLOR_MUTED,
            font=self._font(12),
            anchor="w",
            justify="left",
            wraplength=860,
        )
        delay_hint.grid(row=4, column=0, columnspan=2, sticky="ew", padx=(4, 12), pady=(4, 2))

        self.network_parameter_widgets.append(inline_entry(advanced_tab, 0, 0, "上行带宽", self.upload_bandwidth_var, "0 = 不限速"))
        self.network_parameter_widgets.append(inline_entry(advanced_tab, 0, 1, "下行带宽", self.download_bandwidth_var, "0 = 不限速"))
        self.network_parameter_widgets.append(compact_option(advanced_tab, 1, 0, "带宽单位", self.bandwidth_unit_var, ["Kbps", "Mbps"]))
        self.network_parameter_widgets.append(inline_entry(advanced_tab, 1, 1, "自动停止（分钟）", self.network_duration_var, "0 = 不自动停止"))
        self.network_parameter_widgets.append(inline_entry(advanced_tab, 2, 0, "突发丢包触发概率", self.burst_trigger_var, "0～100"))
        self.network_parameter_widgets.append(inline_entry(advanced_tab, 2, 1, "连续丢包数量", self.burst_length_var, "0 = 关闭"))

        preset = ctk.CTkFrame(wrap, fg_color=COLOR_BG)
        preset.pack(fill="x", pady=(0, 10))
        presets = [
            ("轻微延迟", dict(delay="100", jitter="0", loss="0")),
            ("弱网测试", dict(delay="300", jitter="50", loss="2")),
            ("移动网络", dict(delay="180", jitter="80", loss="3")),
            ("海外高延迟", dict(delay="350", jitter="100", loss="1")),
            ("低带宽", dict(delay="120", jitter="30", loss="1", upload="512", download="1024", unit="Kbps")),
            ("突发丢包", dict(delay="120", jitter="50", loss="0", burst_trigger="5", burst_length="4")),
            ("高丢包", dict(delay="100", jitter="30", loss="15")),
        ]
        preset.grid_columnconfigure(tuple(range(len(presets))), weight=1, uniform="network_presets")
        for index, (text, config) in enumerate(presets):
            preset_button = ctk.CTkButton(
                preset,
                text=text,
                height=34,
                corner_radius=12,
                fg_color=COLOR_SURFACE,
                hover_color=COLOR_HOVER,
                border_width=1,
                border_color=COLOR_BORDER,
                text_color=COLOR_TEXT,
                font=self._font(12, "bold"),
                command=lambda values=config: self._apply_network_preset(**values),
            )
            preset_button.grid(row=0, column=index, sticky="ew", padx=(0 if index == 0 else 4, 4))
            self.network_parameter_widgets.append(preset_button)

        action = ctk.CTkFrame(wrap, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        action.pack(fill="both", expand=True)
        action.grid_columnconfigure(0, weight=1)
        action.grid_rowconfigure(3, weight=1)

        actions = ctk.CTkFrame(action, fg_color="transparent")
        actions.grid(row=0, column=0, pady=(18, 8))
        self.network_button = ctk.CTkButton(actions, text="▶  开始模拟" if not self.network.running.is_set() else "■  停止模拟", width=320, height=48, corner_radius=14, fg_color=COLOR_ACCENT if not self.network.running.is_set() else COLOR_DANGER, hover_color=COLOR_ACCENT_HOVER if not self.network.running.is_set() else COLOR_DANGER_HOVER, text_color="#FFFFFF", font=self._font(18, "bold"), command=self.toggle_network)
        self.network_button.pack(side="left", padx=(0, 10))
        self.network_emergency_button = ctk.CTkButton(
            actions,
            text="立即恢复网络",
            width=150,
            height=48,
            corner_radius=14,
            fg_color="transparent",
            hover_color=COLOR_DANGER_HOVER,
            border_width=1,
            border_color=COLOR_DANGER,
            text_color=COLOR_DANGER,
            font=self._font(14, "bold"),
            command=self._emergency_restore_network,
        )
        self.network_emergency_button.pack(side="left")

        ctk.CTkLabel(action, textvariable=self.network_status_var, text_color=COLOR_TEXT, font=self._font(13, "bold")).grid(row=1, column=0, pady=(0, 6))
        ctk.CTkLabel(action, textvariable=self.network_stats_var, text_color=COLOR_MUTED, font=self._font(12), justify="center").grid(row=2, column=0, pady=(0, 8))

        self.network_bottom_frame = ctk.CTkFrame(action, fg_color="transparent")
        self.network_bottom_frame.grid(row=3, column=0, sticky="nsew", padx=20, pady=(0, 18))
        self.network_bottom_frame.grid_rowconfigure(0, weight=1)
        self.network_bottom_frame.grid_columnconfigure((0, 1), weight=1, uniform="network_bottom")

        guide_panel = ctk.CTkFrame(self.network_bottom_frame, fg_color=COLOR_SURFACE_2, corner_radius=12, border_width=1, border_color=COLOR_BORDER)
        guide_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        guide_panel.grid_columnconfigure(0, weight=1)
        guide_panel.grid_rowconfigure(1, weight=1)
        guide_header = ctk.CTkFrame(guide_panel, fg_color="transparent")
        guide_header.grid(row=0, column=0, sticky="ew", padx=12, pady=(8, 2))
        ctk.CTkLabel(guide_header, text="使用说明", text_color=COLOR_TEXT, font=self._font(13, "bold")).pack(side="left")
        self.network_guide_toggle_button = ctk.CTkButton(guide_header, text="收起", width=54, height=24, fg_color="transparent", hover_color=COLOR_HOVER, text_color=COLOR_MUTED, command=self._toggle_network_guide)
        self.network_guide_toggle_button.pack(side="right")
        self.network_guide_box = ctk.CTkTextbox(guide_panel, height=115, fg_color="transparent", text_color=COLOR_MUTED, font=self._font(12), wrap="word")
        self.network_guide_box.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
        self.network_guide_box.insert("end", NETWORK_USAGE_GUIDE)
        self.network_guide_box.configure(state="disabled")

        log_panel = ctk.CTkFrame(self.network_bottom_frame, fg_color=COLOR_SURFACE_2, corner_radius=12, border_width=1, border_color=COLOR_BORDER)
        log_panel.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        log_panel.grid_columnconfigure(0, weight=1)
        log_panel.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(log_panel, text="运行日志", text_color=COLOR_TEXT, font=self._font(13, "bold"), anchor="w").grid(row=0, column=0, sticky="ew", padx=12, pady=(8, 2))
        self.network_log_box = ctk.CTkTextbox(log_panel, height=115, fg_color="transparent", text_color=COLOR_TEXT, font=self._font(12), wrap="word")
        self.network_log_box.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
        self.network_log_box.insert("end", "等待启动。运行状态、生效规则和安全警告会显示在这里。\n")
        self.network_log_box.configure(state="disabled")
        self._refresh_network_button()

    def _apply_network_preset(
        self,
        delay="100",
        jitter="0",
        loss="0",
        reorder="0",
        upload="0",
        download="0",
        unit="Mbps",
        burst_trigger="0",
        burst_length="0",
        duration="0",
    ):
        if self.network.running.is_set():
            messagebox.showwarning("正在运行", "请先停止模拟，再切换预设")
            return
        self.delay_var.set(delay)
        self.jitter_var.set(jitter)
        self.loss_var.set(loss)
        self.reorder_var.set(reorder)
        self.upload_bandwidth_var.set(upload)
        self.download_bandwidth_var.set(download)
        self.bandwidth_unit_var.set(unit)
        self.burst_trigger_var.set(burst_trigger)
        self.burst_length_var.set(burst_length)
        self.network_duration_var.set(duration)

    def _toggle_network_guide(self):
        box = getattr(self, "network_guide_box", None)
        button = getattr(self, "network_guide_toggle_button", None)
        frame = getattr(self, "network_bottom_frame", None)
        if box is None or button is None or frame is None:
            return
        self.network_guide_visible = not self.network_guide_visible
        if self.network_guide_visible:
            box.grid()
            button.configure(text="收起")
            frame.grid_columnconfigure((0, 1), weight=1, uniform="network_bottom")
        else:
            box.grid_remove()
            button.configure(text="展开")
            frame.grid_columnconfigure(0, weight=0, minsize=150, uniform="")
            frame.grid_columnconfigure(1, weight=1, uniform="")

    def _emergency_restore_network(self):
        if not self.network.running.is_set():
            return
        self._network_log("用户触发立即恢复网络")
        self._set_task_phase("network", TaskPhase.STOPPING, "正在立即恢复正常网络")
        self.network.stop("已立即恢复正常网络")
        self._refresh_network_button()

    def _update_network_delay_hint(self):
        try:
            delay = max(0, int(self.delay_var.get().strip() or 0))
        except Exception:
            delay = 0
        try:
            jitter = max(0, int(self.jitter_var.get().strip() or 0))
        except Exception:
            jitter = 0
        if self.direction_var.get() == "全部":
            rtt = delay * 2
            text = f"单向延迟 {delay}±{jitter}ms；全部方向预计额外 RTT 约 {rtt}ms"
        else:
            text = f"单向延迟 {delay}±{jitter}ms；当前仅作用于{self.direction_var.get()}流量"
        try:
            self.network_delay_hint_var.set(text)
        except Exception:
            pass

    def _refresh_network_button(self):
        button = getattr(self, "network_button", None)
        if button is None:
            return
        try:
            if not button.winfo_exists():
                return
            snapshot = self._task_snapshot("network") if hasattr(self, "_task_snapshot") else None
            phase = snapshot.phase if snapshot is not None else (
                TaskPhase.RUNNING if self.network.running.is_set() else TaskPhase.IDLE
            )
            if phase == TaskPhase.STARTING:
                text, color, hover, state = "…  正在启动", COLOR_ACCENT, COLOR_ACCENT_HOVER, "disabled"
            elif phase == TaskPhase.STOPPING:
                text, color, hover, state = "…  正在停止", COLOR_DANGER, COLOR_DANGER_HOVER, "disabled"
            elif phase == TaskPhase.RUNNING:
                text, color, hover, state = "■  停止模拟", COLOR_DANGER, COLOR_DANGER_HOVER, "normal"
            else:
                text, color, hover, state = "▶  开始模拟", COLOR_ACCENT, COLOR_ACCENT_HOVER, "normal"
            button.configure(
                text=text,
                fg_color=color,
                hover_color=hover,
                state=state,
            )
            emergency = getattr(self, "network_emergency_button", None)
            if emergency is not None and emergency.winfo_exists():
                emergency.configure(
                    state="normal" if phase in (TaskPhase.STARTING, TaskPhase.RUNNING) else "disabled"
                )
            self._set_network_controls_locked(phase in (TaskPhase.STARTING, TaskPhase.RUNNING, TaskPhase.STOPPING))
        except Exception:
            pass

    def _set_network_controls_locked(self, locked):
        state = "disabled" if locked else "normal"
        for widget in getattr(self, "network_parameter_widgets", []):
            try:
                if widget is not None and widget.winfo_exists():
                    widget.configure(state=state)
            except Exception:
                pass

    def _build_filter(self):
        advanced = self.advanced_filter_var.get().strip()
        if advanced:
            return ensure_windivert_impostor_guard(advanced)

        parts = []
        protocol = self.protocol_var.get()
        if protocol == "TCP":
            parts.append("tcp")
        elif protocol == "UDP":
            parts.append("udp")
        elif protocol == "ICMP":
            parts.append("icmp")

        direction = self.direction_var.get()
        if direction == "出站":
            parts.append("outbound")
        elif direction == "入站":
            parts.append("inbound")

        target_ip = self.target_ip_var.get().strip()
        if target_ip:
            ip_obj = ipaddress.ip_address(target_ip)
            if ip_obj.version == 4:
                parts.append(f"(ip.SrcAddr == {target_ip} or ip.DstAddr == {target_ip})")
            else:
                parts.append(f"(ipv6.SrcAddr == {target_ip} or ipv6.DstAddr == {target_ip})")

        return ensure_windivert_impostor_guard(" and ".join(parts) if parts else "true")

    def _parse_network_config(self):
        try:
            delay = int(self.delay_var.get().strip())
            jitter = int(self.jitter_var.get().strip() or 0)
            loss = int(self.loss_var.get().strip())
            reorder = int(self.reorder_var.get().strip() or 0)
            upload = float(self.upload_bandwidth_var.get().strip() or 0)
            download = float(self.download_bandwidth_var.get().strip() or 0)
            burst_trigger = int(self.burst_trigger_var.get().strip() or 0)
            burst_length = int(self.burst_length_var.get().strip() or 0)
            duration_minutes = float(self.network_duration_var.get().strip() or 0)
        except ValueError as exc:
            raise ValueError("请检查弱网参数：概率和毫秒值需为整数，带宽和自动停止时间需为数字") from exc

        if delay < 0 or delay > 10000:
            raise ValueError("延迟时间必须在 0-10000 毫秒之间")
        if jitter < 0 or jitter > 10000:
            raise ValueError("延迟抖动必须在 0-10000 毫秒之间")
        if loss < 0 or loss > 100:
            raise ValueError("丢包概率必须在 0-100 之间")
        if reorder < 0 or reorder > 100:
            raise ValueError("数据包乱序概率必须在 0-100 之间")
        if upload < 0 or download < 0 or upload > 100000 or download > 100000:
            raise ValueError("上下行带宽必须在 0-100000 之间，0 表示不限速")
        if burst_trigger < 0 or burst_trigger > 100:
            raise ValueError("突发丢包触发概率必须在 0-100 之间")
        if burst_length < 0 or burst_length > 1000:
            raise ValueError("连续丢包数量必须在 0-1000 之间")
        if bool(burst_trigger) != bool(burst_length):
            raise ValueError("启用突发丢包时，触发概率和连续丢包数量都必须大于 0；关闭时请都填 0")
        if duration_minutes < 0 or duration_minutes > 1440:
            raise ValueError("自动停止时间必须在 0-1440 分钟之间，0 表示不自动停止")

        unit = self.bandwidth_unit_var.get().strip()
        if unit not in ("Kbps", "Mbps"):
            raise ValueError("带宽单位必须选择 Kbps 或 Mbps")
        multiplier = 1000.0 if unit == "Kbps" else 1000000.0
        filter_text = self._build_filter()
        NetworkWorker.validate_filter(filter_text)
        return {
            "delay": delay,
            "jitter": jitter,
            "loss": loss,
            "reorder": reorder,
            "upload_bps": upload * multiplier,
            "download_bps": download * multiplier,
            "upload_display": upload,
            "download_display": download,
            "bandwidth_unit": unit,
            "burst_trigger": burst_trigger,
            "burst_length": burst_length,
            "duration_seconds": duration_minutes * 60.0,
            "duration_minutes": duration_minutes,
            "filter_text": filter_text,
        }

    def _network_has_global_scope(self):
        return not self.target_ip_var.get().strip() and not self.advanced_filter_var.get().strip()

    def _network_risk_warnings(self, config):
        warnings = []
        if self._network_has_global_scope():
            warnings.append("目标 IP 和高级规则均为空，将影响本机全部匹配流量")
        if config["loss"] >= 50:
            warnings.append(f"随机丢包概率较高：{config['loss']}%")
        if config["reorder"] > 0:
            warnings.append(f"已主动启用数据包乱序：{config['reorder']}%")
        if config["burst_trigger"] >= 20 or config["burst_length"] >= 10:
            warnings.append(
                f"突发丢包强度较高：{config['burst_trigger']}% 概率连续丢 {config['burst_length']} 包"
            )
        limited_rates = [
            rate for rate in (config["upload_bps"], config["download_bps"]) if rate > 0
        ]
        if limited_rates and min(limited_rates) <= 128000:
            warnings.append("带宽限制低于或等于 128 Kbps，可能造成明显断流或超时")
        if config["delay"] >= 3000:
            warnings.append("单向基础延迟达到 3000ms 以上，可能触发大量请求超时")
        return warnings

    def toggle_network(self):
        snapshot = self._task_snapshot("network")
        phase = snapshot.phase if snapshot is not None else (
            TaskPhase.RUNNING if self.network.running.is_set() else TaskPhase.IDLE
        )
        if phase in (TaskPhase.STARTING, TaskPhase.STOPPING):
            return
        if phase == TaskPhase.RUNNING or self.network.running.is_set():
            self._set_task_phase("network", TaskPhase.STOPPING, "正在停止弱网模拟")
            self.network.stop()
        else:
            try:
                config = self._parse_network_config()
                warnings = self._network_risk_warnings(config)
                if warnings:
                    confirmed = messagebox.askyesno(
                        "确认弱网模拟范围与风险",
                        "启动前请确认：\n\n- "
                        + "\n- ".join(warnings)
                        + f"\n\n最终生效规则：\n{config['filter_text']}\n\n是否继续启动？",
                        icon="warning",
                    )
                    if not confirmed:
                        self._network_log("已取消启动：用户未确认模拟范围或高风险参数")
                        return
                self._set_task_phase("network", TaskPhase.STARTING, "正在初始化 WinDivert")
                self.network.start(
                    config["delay"],
                    config["loss"],
                    config["filter_text"],
                    jitter_ms=config["jitter"],
                    reorder_percent=config["reorder"],
                    upload_bps=config["upload_bps"],
                    download_bps=config["download_bps"],
                    burst_trigger_percent=config["burst_trigger"],
                    burst_length=config["burst_length"],
                    auto_stop_seconds=config["duration_seconds"],
                )
                self._network_log(
                    f"正在启动：单向延迟 {config['delay']}±{config['jitter']}ms，"
                    f"随机丢包 {config['loss']}%，乱序 {config['reorder']}%"
                )
                if self.direction_var.get() == "全部":
                    self._network_log(
                        f"上下行同时生效，预计额外 RTT 约 {config['delay'] * 2}ms（不含抖动）"
                    )
                if config["upload_bps"] or config["download_bps"]:
                    self._network_log(
                        f"带宽限制：上行 {config['upload_display']:g} {config['bandwidth_unit']}，"
                        f"下行 {config['download_display']:g} {config['bandwidth_unit']}（0 = 不限速）"
                    )
                if config["burst_trigger"]:
                    self._network_log(
                        f"突发丢包：{config['burst_trigger']}% 概率触发，连续丢弃 {config['burst_length']} 包"
                    )
                if config["duration_minutes"]:
                    self._network_log(f"将在 {config['duration_minutes']:g} 分钟后自动恢复正常网络")
                self._network_log(f"最终生效规则：{config['filter_text']}")
            except Exception as exc:
                self._set_task_phase("network", TaskPhase.FAILED, str(exc))
                messagebox.showerror("启动失败", str(exc))
        self._refresh_network_button()

    def _network_log(self, text):
        def write():
            if hasattr(self, "network_log_box") and self.network_log_box.winfo_exists():
                self.network_log_box.configure(state="normal")
                self.network_log_box.insert("end", f"[{time.strftime('%H:%M:%S')}] {text}\n")
                self.network_log_box.see("end")
                self.network_log_box.configure(state="disabled")
        try:
            self.after(0, write)
        except Exception:
            pass


    # ---------- ADB Log page ----------
    def _build_adb_log_page(self):
        """v2.4.30: Android 日志抓取页重构为状态摘要卡片 + 设备日志任务表。"""
        wrap = ctk.CTkFrame(self.content, fg_color=COLOR_BG)
        wrap.grid(row=0, column=0, sticky="nsew")
        wrap.grid_columnconfigure(0, weight=1)
        wrap.grid_rowconfigure(0, weight=1)

        page_tabs = ctk.CTkTabview(
            wrap,
            fg_color=COLOR_BG,
            segmented_button_fg_color=COLOR_SURFACE_2,
            segmented_button_selected_color=COLOR_ACCENT,
            segmented_button_selected_hover_color=COLOR_ACCENT_HOVER,
            segmented_button_unselected_color=COLOR_SURFACE_2,
            segmented_button_unselected_hover_color=COLOR_HOVER,
            text_color=COLOR_TEXT,
            corner_radius=12,
        )
        self.adb_log_page_tabs = page_tabs
        page_tabs.grid(row=0, column=0, sticky="nsew")
        control_tab = page_tabs.add("抓取控制")
        task_tab = page_tabs.add("设备日志任务")
        log_tab = page_tabs.add("运行日志")
        for tab in (control_tab, task_tab, log_tab):
            tab.configure(fg_color=COLOR_BG)
            tab.grid_columnconfigure(0, weight=1)
        control_tab.grid_rowconfigure(1, weight=1)
        task_tab.grid_rowconfigure(0, weight=1)
        log_tab.grid_rowconfigure(0, weight=1)

        config_card = self._card(control_tab)
        config_card.grid(row=0, column=0, sticky="ew", padx=0, pady=(0, 12))
        config_card.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            config_card,
            text="Android 日志抓取",
            text_color=COLOR_TEXT,
            font=self._font(18, "bold"),
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=24, pady=(18, 8))
        self._entry_row(config_card, 1, "应用包名", self.log_pkg_var, "例如 com.tinywarsurvivalexpress.android", width=520)

        # 第一行：设备操作栏。复用 ADB 多设备选择器，保留 serial 绑定逻辑。
        self._build_adb_device_selector(config_card, 2, show_actions=True)

        # 第二行：日志操作栏。所有运行状态统一进入摘要卡片和任务表，不再在按钮区堆叠文本。
        action_row = ctk.CTkFrame(config_card, fg_color="transparent")
        action_row.grid(row=3, column=0, columnspan=3, sticky="ew", padx=24, pady=(0, 18))
        # v2.4.31: narrow-window friendly action layout. Split buttons into
        # two rows and avoid placing six fixed-width buttons on one line.
        for col in range(5):
            action_row.grid_columnconfigure(col, weight=0)
        action_row.grid_columnconfigure(5, weight=1)

        ctk.CTkButton(action_row, text="清空已选 logcat", width=132, height=36, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self.clear_selected_adb_logcat).grid(row=0, column=0, padx=(0, 10), pady=4, sticky="w")
        ctk.CTkButton(action_row, text="开始抓取已选", width=132, height=36, corner_radius=12, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color="#FFFFFF", font=self._font(13, "bold"), command=self.start_selected_adb_log).grid(row=0, column=1, padx=(0, 10), pady=4, sticky="w")
        ctk.CTkButton(action_row, text="停止选中设备", width=132, height=36, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self.stop_selected_adb_log).grid(row=0, column=2, padx=(0, 10), pady=4, sticky="w")
        ctk.CTkButton(action_row, text="停止全部", width=108, height=36, corner_radius=12, fg_color=COLOR_DANGER, hover_color=COLOR_DANGER_HOVER, text_color="#FFFFFF", font=self._font(13, "bold"), command=self.stop_all_adb_log).grid(row=0, column=3, padx=(0, 10), pady=4, sticky="w")
        ctk.CTkButton(action_row, text="打开日志目录", width=132, height=36, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self.open_adb_log_dir).grid(row=1, column=0, padx=(0, 10), pady=(6, 4), sticky="w")
        ctk.CTkButton(action_row, text="复制日志目录", width=132, height=36, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self.copy_adb_log_dir).grid(row=1, column=1, padx=(0, 10), pady=(6, 4), sticky="w")

        # 保留“当前设备”兼容入口，但不再把文件路径和任务详情堆在按钮区。
        self.adb_log_button = ctk.CTkButton(action_row, text="▶  当前设备", width=132, height=36, corner_radius=12, fg_color=COLOR_ACCENT if not self.adb_log.running.is_set() else COLOR_DANGER, hover_color=COLOR_ACCENT_HOVER if not self.adb_log.running.is_set() else COLOR_DANGER_HOVER, text_color="#FFFFFF", font=self._font(13, "bold"), command=self.toggle_adb_log)
        self.adb_log_button.grid(row=1, column=2, padx=(0, 10), pady=(6, 4), sticky="w")
        ctk.CTkLabel(action_row, textvariable=self.adb_status_var, text_color=COLOR_MUTED, font=self._font(12)).grid(row=1, column=3, columnspan=3, sticky="w", padx=(0, 10), pady=(6, 4))

        self._build_adb_log_summary_card(control_tab)

        task_card = ctk.CTkFrame(task_tab, fg_color=COLOR_SURFACE, corner_radius=0)
        task_card.grid(row=0, column=0, sticky="nsew")
        task_card.grid_columnconfigure(0, weight=1)
        task_card.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(task_card, text="设备日志任务", text_color=COLOR_TEXT, font=self._font(16, "bold")).grid(row=0, column=0, sticky="w", padx=24, pady=(10, 6))
        self._build_adb_log_task_table(task_card)

        log_card = ctk.CTkFrame(log_tab, fg_color=COLOR_SURFACE, corner_radius=0)
        log_card.grid(row=0, column=0, sticky="nsew")
        log_card.grid_columnconfigure(0, weight=1)
        log_card.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(log_card, text="运行日志", text_color=COLOR_TEXT, font=self._font(16, "bold")).grid(row=0, column=0, sticky="w", padx=24, pady=(10, 6))
        self.adb_log_box = ctk.CTkTextbox(log_card, corner_radius=14, border_width=1, border_color=COLOR_BORDER, fg_color=COLOR_SURFACE_2, text_color=COLOR_TEXT, font=self._font(13))
        self.adb_log_box.grid(row=1, column=0, sticky="nsew", padx=24, pady=(0, 12))
        self.adb_log_box.insert("end", "提示：请先连接安卓设备，开启 USB 调试，并在手机上允许调试授权。多设备抓取时，每台设备会生成独立日志文件。\n")
        self.adb_log_box.configure(state="disabled")
        self._refresh_adb_button()
        self._refresh_adb_log_task_table()
        self._refresh_adb_log_summary()
        self._schedule_adb_log_task_refresh()

    def _build_adb_log_summary_card(self, parent):
        """Build compact two-column summary cards for narrow windows."""
        summary = ctk.CTkFrame(parent, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        summary.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        summary.grid_columnconfigure(0, weight=1)
        summary.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(summary, text="运行状态", text_color=COLOR_TEXT, font=self._font(16, "bold")).grid(row=0, column=0, sticky="w", padx=24, pady=(16, 8), columnspan=2)

        if not hasattr(self, "adb_log_running_count_var"):
            self.adb_log_running_count_var = tk.StringVar(value="0")
            self.adb_log_pkg_summary_var = tk.StringVar(value=DEFAULT_PKG_NAME)
            self.adb_log_dir_summary_var = tk.StringVar(value=self._short_path(self.adb_log.log_dir, 42))
            self.adb_log_updated_summary_var = tk.StringVar(value="-")

        self._summary_metric(summary, 1, 0, "运行中设备", self.adb_log_running_count_var)
        self._summary_metric(summary, 1, 1, "当前包名", self.adb_log_pkg_summary_var)
        self._summary_metric(summary, 2, 0, "日志目录", self.adb_log_dir_summary_var)
        self._summary_metric(summary, 2, 1, "最后更新时间", self.adb_log_updated_summary_var)

    def _summary_metric(self, parent, row, col, label, variable):
        box = ctk.CTkFrame(parent, fg_color=COLOR_SURFACE_2, corner_radius=12, border_width=1, border_color=COLOR_BORDER)
        box.grid(row=row, column=col, sticky="ew", padx=(24 if col == 0 else 8, 24 if col == 1 else 8), pady=(0, 10 if row == 1 else 16))
        box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(box, text=label, text_color=COLOR_MUTED, font=self._font(12), anchor="w").grid(row=0, column=0, sticky="ew", padx=14, pady=(8, 1))
        value_label = ctk.CTkLabel(box, textvariable=variable, text_color=COLOR_TEXT, font=self._font(13, "bold"), anchor="w")
        value_label.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 8))

    @staticmethod
    def _short_path(path, max_chars=58):
        value = str(path or "").strip()
        if not value:
            return "-"
        if len(value) <= max_chars:
            return value
        tail = os.path.basename(value.rstrip("/\\")) or value[-18:]
        head_len = max(12, max_chars - len(tail) - 5)
        return f"{value[:head_len]}...{tail}"

    @staticmethod
    def _format_file_size(size):
        try:
            value = int(size or 0)
        except Exception:
            value = 0
        units = ["B", "KB", "MB", "GB"]
        amount = float(value)
        idx = 0
        while amount >= 1024 and idx < len(units) - 1:
            amount /= 1024
            idx += 1
        if idx == 0:
            return f"{int(amount)} {units[idx]}"
        return f"{amount:.1f} {units[idx]}"

    def _adb_log_task_snapshot(self):
        with self.adb_log._lock:
            return dict(getattr(self.adb_log, "tasks", {}) or {})

    def _refresh_adb_log_summary(self):
        if not hasattr(self, "adb_log_running_count_var"):
            return
        tasks = self._adb_log_task_snapshot()
        running_count = sum(1 for task in tasks.values() if task.get("status") in ("启动中", "运行中"))
        latest = "-"
        for task in tasks.values():
            updated = task.get("updated_at") or task.get("start_time") or ""
            if updated and (latest == "-" or updated > latest):
                latest = updated
        self.adb_log_running_count_var.set(str(running_count))
        self.adb_log_pkg_summary_var.set(self._short_path(self.log_pkg_var.get() or DEFAULT_PKG_NAME, 38))
        self.adb_log_dir_summary_var.set(self._short_path(self.adb_log.log_dir, 42))
        self.adb_log_updated_summary_var.set(latest)

    def _schedule_adb_log_task_refresh(self):
        if getattr(self, "_adb_log_task_refresh_job", None):
            return
        def tick():
            self._adb_log_task_refresh_job = None
            try:
                if self.current_page == "adb_log":
                    self._refresh_adb_log_task_table()
                    self._refresh_adb_log_summary()
            finally:
                if self.current_page == "adb_log" or self.adb_log.running.is_set():
                    self._adb_log_task_refresh_job = self.after(1000, tick)
        self._adb_log_task_refresh_job = self.after(1000, tick)

    def copy_adb_log_dir(self):
        try:
            self.clipboard_clear()
            self.clipboard_append(self.adb_log.log_dir)
            self._adb_log("已复制日志目录路径")
        except Exception as exc:
            messagebox.showerror("复制失败", str(exc))

    def _open_path(self, path):
        if not path:
            raise FileNotFoundError("路径为空")
        target = path if os.path.exists(path) else os.path.dirname(path)
        if not target:
            raise FileNotFoundError(path)
        if os.name == "nt":
            os.startfile(target)
        else:
            subprocess.Popen(["xdg-open", target])

    def _open_adb_log_task_path(self, serial):
        task = self._adb_log_task_snapshot().get(serial) or {}
        path = task.get("file_path") or ""
        try:
            self._open_path(path)
        except Exception as exc:
            messagebox.showerror("打开失败", str(exc))

    def _copy_adb_log_task_path(self, serial):
        task = self._adb_log_task_snapshot().get(serial) or {}
        path = task.get("file_path") or ""
        if not path:
            messagebox.showinfo("复制路径", "该设备暂未生成日志文件。")
            return
        try:
            self.clipboard_clear()
            self.clipboard_append(path)
            self._adb_log(f"已复制日志路径：{serial}")
        except Exception as exc:
            messagebox.showerror("复制失败", str(exc))

    def _build_adb_log_task_table(self, parent):
        table_frame = ctk.CTkFrame(parent, fg_color=COLOR_SURFACE_2, corner_radius=12, border_width=1, border_color=COLOR_BORDER)
        table_frame.grid(row=1, column=0, sticky="nsew", padx=24, pady=(0, 16))
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)
        columns = ("checked", "serial", "model", "android", "status", "log_status", "file_size", "log_file", "action")
        tree = ttk.Treeview(table_frame, columns=columns, show="headings", style="Xiao.Treeview", height=8)
        headings = {
            "checked": "勾选",
            "serial": "serial",
            "model": "型号",
            "android": "Android",
            "status": "设备状态",
            "log_status": "日志状态",
            "file_size": "文件大小",
            "log_file": "日志文件",
            "action": "操作",
        }
        widths = {"checked": 64, "serial": 120, "model": 130, "android": 90, "status": 95, "log_status": 100, "file_size": 100, "log_file": 280, "action": 150}
        for col in columns:
            tree.heading(col, text=headings[col])
            anchor = "center" if col in ("checked", "log_status", "file_size", "action") else "w"
            tree.column(col, width=widths[col], minwidth=widths[col], anchor=anchor, stretch=col == "log_file")
        try:
            tree.tag_configure("running", foreground="#22C55E")
            tree.tag_configure("error", foreground="#F87171")
            tree.tag_configure("offline", foreground="#F59E0B")
            tree.tag_configure("stopped", foreground="#9AA6B8")
        except Exception:
            pass
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree.bind("<ButtonRelease-1>", self._on_adb_log_task_click)
        self.adb_log_task_tree = tree

    def _refresh_adb_log_task_table(self, data=None):
        tree = getattr(self, "adb_log_task_tree", None)
        if not tree or not tree.winfo_exists():
            return
        data = data or self._adb_last_device_data or {}
        devices = data.get("devices", []) or []
        tasks = self._adb_log_task_snapshot()
        try:
            tree.delete(*tree.get_children())
            if not devices and not tasks:
                self._refresh_adb_log_summary()
                return
            seen = set()
            for item in devices:
                serial = item.get("serial", "")
                seen.add(serial)
                task = tasks.get(serial) or {}
                checked = "☑" if serial in self.adb_log_selected_serials else "☐"
                file_path = task.get("file_path") or ""
                file_name = os.path.basename(file_path) if file_path else "-"
                try:
                    size = os.path.getsize(file_path) if file_path and os.path.exists(file_path) else int(task.get("file_size") or 0)
                except Exception:
                    size = int(task.get("file_size") or 0)
                status = item.get("status") or ""
                log_status = task.get("status") or "未启动"
                actions = ""
                if file_path:
                    actions = "打开 / 复制"
                if log_status in ("运行中", "启动中"):
                    actions = "停止 / " + (actions if actions else "打开 / 复制")
                values = (
                    checked,
                    serial,
                    item.get("model") or item.get("device") or item.get("product") or "未知型号",
                    item.get("android_version") or "",
                    status,
                    log_status,
                    self._format_file_size(size),
                    file_name,
                    actions,
                )
                tags = []
                if log_status in ("运行中", "启动中"):
                    tags.append("running")
                elif log_status in ("异常", "设备断开"):
                    tags.append("error")
                elif status in ("offline", "unauthorized", "disconnected"):
                    tags.append("offline")
                elif log_status in ("已停止", "未启动"):
                    tags.append("stopped")
                tree.insert("", "end", iid=serial, values=values, tags=tuple(tags))
            for serial, task in tasks.items():
                if serial in seen:
                    continue
                checked = "☑" if serial in self.adb_log_selected_serials else "☐"
                file_path = task.get("file_path") or ""
                file_name = os.path.basename(file_path) if file_path else "-"
                try:
                    size = os.path.getsize(file_path) if file_path and os.path.exists(file_path) else int(task.get("file_size") or 0)
                except Exception:
                    size = int(task.get("file_size") or 0)
                log_status = task.get("status") or "异常"
                action = "打开 / 复制" if file_path else ""
                if log_status in ("运行中", "启动中"):
                    action = "停止 / " + (action if action else "打开 / 复制")
                values = (checked, serial, "", "", "设备断开", log_status, self._format_file_size(size), file_name, action)
                tree.insert("", "end", iid=serial, values=values, tags=("error",))
            self._refresh_adb_log_summary()
        except Exception as exc:
            try:
                self._adb_log(f"刷新日志任务表失败：{exc}")
            except Exception:
                pass

    def _on_adb_log_task_click(self, event=None):
        tree = getattr(self, "adb_log_task_tree", None)
        if not tree:
            return
        row = tree.identify_row(event.y) if event else ""
        col = tree.identify_column(event.x) if event else ""
        if not row:
            return
        if col == "#1":
            if row in self.adb_log_selected_serials:
                self.adb_log_selected_serials.discard(row)
            else:
                self.adb_log_selected_serials.add(row)
            self._refresh_adb_log_task_table()
            return
        if col == "#9":
            bbox = tree.bbox(row, col)
            rel = (event.x - bbox[0]) if bbox else 0
            width = bbox[2] if bbox else 1
            task = self._adb_log_task_snapshot().get(row) or {}
            running = task.get("status") in ("运行中", "启动中")
            # 操作列按三段处理：停止 / 打开 / 复制。非运行状态下前半段打开、后半段复制。
            if running and rel < width / 3:
                self.adb_log.stop(row)
                self._adb_log(f"已请求停止设备日志抓取：{row}")
            elif (running and rel < width * 2 / 3) or (not running and rel < width / 2):
                self._open_adb_log_task_path(row)
            else:
                self._copy_adb_log_task_path(row)
            self._refresh_adb_log_task_table()
            return
        item = self._device_by_serial(row)
        if item and item.get("status") == "device":
            self.adb_tools.set_selected_serial(row)
            self._update_adb_device_choices(self._adb_last_device_data)

    def _selected_log_serials_or_current(self):
        serials = [s for s in sorted(self.adb_log_selected_serials) if s]
        if serials:
            return serials
        current = self.adb_tools.get_selected_serial()
        if current:
            return [current]
        return []

    def start_selected_adb_log(self):
        if self.adb_log.running.is_set():
            messagebox.showinfo("日志抓取", "日志抓取已在进行中，请先停止当前任务。")
            return
        try:
            data = self._update_adb_device_choices()
            serials = self._selected_log_serials_or_current()
            if not serials:
                ready = [item.get("serial") for item in (data or {}).get("devices", []) if item.get("status") == "device"]
                if len(ready) > 1:
                    messagebox.showwarning("请选择设备", "检测到多台 Android 设备，请勾选要抓取日志的设备。")
                    return
                serials = ready
            if not serials:
                messagebox.showwarning("无可用设备", "未检测到可用 device 状态设备。")
                return
            # 只允许 device 状态设备参与抓取
            valid = []
            for serial in serials:
                item = self._device_by_serial(serial, data)
                if item and item.get("status") == "device":
                    valid.append(serial)
            if not valid:
                messagebox.showwarning("无可用设备", "已选设备不是 device 状态，无法抓取日志。")
                return
            self.adb_log.start(self.log_pkg_var.get(), serials=valid)
            self._adb_log(f"已启动日志抓取：{len(valid)} 台")
        except Exception as exc:
            messagebox.showerror("启动失败", str(exc))
            return
        self._refresh_adb_button()
        self._refresh_adb_log_task_table()
        self._set_badge(self.adb_log.running.is_set())

    def stop_selected_adb_log(self):
        serials = self._selected_log_serials_or_current()
        if not serials:
            messagebox.showinfo("日志抓取", "请先勾选设备或选择当前设备。")
            return
        for serial in serials:
            self.adb_log.stop(serial)
            self._adb_log(f"已请求停止设备日志抓取：{serial}")
        self._refresh_adb_button()
        self._refresh_adb_log_task_table()

    def stop_all_adb_log(self):
        self.adb_log.stop()
        self._adb_log("已请求停止全部日志抓取")
        self._refresh_adb_button()
        self._refresh_adb_log_task_table()
        self._set_badge(False)

    def clear_selected_adb_logcat(self):
        serials = self._selected_log_serials_or_current()
        if not serials:
            messagebox.showinfo("清空 logcat", "请先勾选设备或选择当前设备。")
            return
        def worker():
            results = []
            for serial in serials:
                res = self.adb_tools.run(["logcat", "-c"], timeout=8, serial=serial)
                results.append((serial, res.ok, res.output))
            return results
        def success(results):
            for serial, ok, output in results:
                self._adb_log(f"[{serial}] 清空 logcat {'成功' if ok else '失败'}" + (f"：{output}" if output else ""))
        self._run_adb_tools_task("清空已选设备 logcat", worker, success)

    def toggle_adb_log(self):
        if self.adb_log.running.is_set():
            serial = self.adb_tools.get_selected_serial()
            if serial:
                self.adb_log.stop(serial)
                self._adb_log(f"已请求停止当前设备日志抓取：{serial}")
            else:
                self.adb_log.stop()
                self._adb_log("已请求停止日志抓取")
        else:
            try:
                self._update_adb_device_choices()
                serial = self.adb_tools.get_selected_serial()
                if not serial:
                    serial = self.adb_tools.resolve_serial()
                self.adb_log.start(self.log_pkg_var.get(), serial=serial)
            except Exception as exc:
                messagebox.showerror("启动失败", str(exc))
                return
        self._refresh_adb_button()
        self._refresh_adb_log_task_table()
        self._set_badge(self.adb_log.running.is_set())

    def start_all_adb_log(self):
        if self.adb_log.running.is_set():
            messagebox.showinfo("日志抓取", "日志抓取已在进行中，请先停止当前任务。")
            return
        try:
            data = self._update_adb_device_choices()
            ready = [item for item in (data or {}).get("devices", []) if item.get("status") == "device"]
            if not ready:
                messagebox.showwarning("无可用设备", "未检测到可用 device 状态设备。")
                return
            if len(ready) == 1:
                self.adb_tools.set_selected_serial(ready[0].get("serial", ""))
            self.adb_log.start_all_ready_devices(self.log_pkg_var.get())
            self._adb_log(f"已启动多设备日志抓取：{len(ready)} 台")
        except Exception as exc:
            messagebox.showerror("启动失败", str(exc))
            return
        self._refresh_adb_button()
        self._set_badge(self.adb_log.running.is_set())

    def check_adb_device(self):
        if self.adb_checking.is_set():
            self._adb_log("ADB 设备检测正在进行中")
            return

        self.adb_checking.set()
        self._adb_log("正在检测 ADB 设备")
        generation = self._begin_adb_refresh()

        def worker():
            try:
                data = self._scan_adb_devices_staged(generation, refresh_resolutions=True)
                text = self.adb_tools.format_devices_text(data)
                self._adb_log("ADB 设备检测结果：")
                for line in text.splitlines() or ["无输出"]:
                    self._adb_log(line)
            finally:
                self.adb_checking.clear()

        threading.Thread(target=worker, daemon=True).start()

    def _begin_adb_refresh(self):
        self._adb_refresh_generation += 1
        return self._adb_refresh_generation

    def _apply_adb_scan_snapshot(self, data, generation, hint=""):
        if generation != self._adb_refresh_generation:
            return
        self._update_adb_device_choices(data)
        text = self.adb_tools.format_devices_text(data)
        self.adb_tools_devices_var.set(text)
        box = getattr(self, "adb_tools_devices_box", None)
        if box is not None and box.winfo_exists():
            self._set_textbox_text(box, text)
        if hint:
            self.adb_device_hint_var.set(hint)

    def _apply_adb_detail_progress(self, item, completed, total, generation):
        if generation != self._adb_refresh_generation:
            return
        current = dict(self._adb_last_device_data or {})
        devices = [dict(device) for device in current.get("devices", []) or []]
        serial = item.get("serial")
        for index, device in enumerate(devices):
            if device.get("serial") == serial:
                devices[index] = dict(item)
                break
        current["devices"] = devices
        self._apply_adb_scan_snapshot(
            current,
            generation,
            f"已识别 {len(devices)} 台设备，正在读取详情 {completed}/{total}…",
        )

    def _scan_adb_devices_staged(self, generation, refresh_resolutions=True):
        quick = self.adb_tools.list_devices(detailed=False)
        ready_count = int(quick.get("ready_count") or 0)
        self.after(
            0,
            lambda data=quick, count=ready_count: self._apply_adb_scan_snapshot(
                data,
                generation,
                f"已完成快速扫描，正在读取在线设备详情 0/{count}…",
            ),
        )

        def progress(item, completed, total):
            self.after(
                0,
                lambda current=item, done=completed, count=total: self._apply_adb_detail_progress(
                    current,
                    done,
                    count,
                    generation,
                ),
            )

        detailed = self.adb_tools.enrich_device_data(
            quick,
            refresh_resolutions=refresh_resolutions,
            progress_callback=progress,
        )
        self.after(
            0,
            lambda data=detailed: self._apply_adb_scan_snapshot(
                data,
                generation,
                f"设备详情刷新完成，共 {len(data.get('devices', []) or [])} 台。",
            ),
        )
        return detailed

    def open_adb_log_dir(self):
        try:
            self.adb_log.open_log_dir()
        except Exception as exc:
            messagebox.showerror("打开失败", str(exc))

    def _refresh_adb_button(self):
        if hasattr(self, "adb_log_button") and self.adb_log_button.winfo_exists():
            running = self.adb_log.running.is_set()
            self.adb_log_button.configure(
                text="■  停止当前" if running else "▶  当前设备",
                fg_color=COLOR_DANGER if running else COLOR_ACCENT,
                hover_color=COLOR_DANGER_HOVER if running else COLOR_ACCENT_HOVER
            )

    def _set_adb_status(self, text):
        def update():
            self.adb_status_var.set(text)
            self._refresh_adb_button()
            self._refresh_adb_log_task_table()
            if self.current_page == "adb_log":
                self._set_badge(self.adb_log.running.is_set())
        try:
            self.after(0, update)
        except Exception:
            pass

    def _set_adb_file(self, text):
        def update():
            self.adb_file_var.set(text)
            self._refresh_adb_log_task_table()
        try:
            self.after(0, update)
        except Exception:
            pass

    def _adb_log(self, text):
        def write():
            if hasattr(self, "adb_log_box") and self.adb_log_box.winfo_exists():
                self.adb_log_box.configure(state="normal")
                self.adb_log_box.insert("end", f"[{time.strftime('%H:%M:%S')}] {text}\n")
                self.adb_log_box.see("end")
                self.adb_log_box.configure(state="disabled")
        try:
            self.after(0, write)
        except Exception:
            pass


    # ---------- Log Analysis page ----------
    def _build_log_analysis_page(self):
        """v2.4.19: 固定布局 + Tab 分区，避免整页滚动导致关键字区拖影。"""
        wrap = ctk.CTkFrame(self.content, fg_color=COLOR_BG)
        wrap.grid(row=0, column=0, sticky="nsew")
        wrap.grid_columnconfigure(0, weight=1)
        wrap.grid_rowconfigure(0, weight=1)

        tabview = ctk.CTkTabview(
            wrap,
            fg_color=COLOR_BG,
            segmented_button_fg_color=COLOR_SURFACE_2,
            segmented_button_selected_color=COLOR_ACCENT,
            segmented_button_selected_hover_color=COLOR_ACCENT_HOVER,
            segmented_button_unselected_color=COLOR_SURFACE_2,
            segmented_button_unselected_hover_color=COLOR_HOVER,
            text_color=COLOR_TEXT,
            corner_radius=12,
        )
        tabview.grid(row=0, column=0, sticky="nsew")

        config_tab = tabview.add("关键字配置")
        action_tab = tabview.add("操作与导入")
        hits_tab = tabview.add("命中记录")
        crash_tab = tabview.add("崩溃结果")
        for tab in (config_tab, action_tab, hits_tab, crash_tab):
            tab.configure(fg_color=COLOR_BG)
            tab.grid_columnconfigure(0, weight=1)
            tab.grid_rowconfigure(0, weight=1)

        # 关键字配置：控件只在进入页面时创建一次，后续刷新只更新变量值。
        config_card = ctk.CTkFrame(config_tab, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        config_card.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        config_card.grid_columnconfigure(0, weight=0)
        config_card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(config_card, text="关键字监控配置", text_color=COLOR_TEXT, font=self._font(18, "bold")).grid(row=0, column=0, sticky="w", padx=24, pady=(20, 10), columnspan=2)
        ctk.CTkCheckBox(config_card, text="启用实时关键字监控", variable=self.log_analysis_enabled_var, text_color=COLOR_TEXT, font=self._font(14)).grid(row=1, column=0, sticky="w", padx=24, pady=(0, 12), columnspan=2)

        keyword_frame = ctk.CTkFrame(config_card, fg_color=COLOR_SURFACE)
        keyword_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=24, pady=(0, 12))
        for col in range(3):
            keyword_frame.grid_columnconfigure(col, weight=1, uniform="log_keywords")
        for idx, keyword in enumerate(DEFAULT_KEYWORDS):
            ctk.CTkCheckBox(
                keyword_frame,
                text=keyword,
                variable=self.log_keyword_vars[keyword],
                width=190,
                text_color=COLOR_TEXT,
                font=self._font(13),
                command=self.apply_log_analysis_config,
            ).grid(row=idx // 3, column=idx % 3, sticky="w", padx=(0, 12), pady=5)

        ctk.CTkLabel(config_card, text="自定义关键字（每行一个）", text_color=COLOR_TEXT, font=self._font(15, "bold")).grid(row=3, column=0, sticky="nw", padx=(24, 22), pady=(12, 8))
        self.log_custom_keywords_textbox = ctk.CTkTextbox(config_card, height=120, corner_radius=12, border_width=1, border_color=COLOR_BORDER, fg_color=COLOR_SURFACE_2, text_color=COLOR_TEXT, font=self._font(13))
        self.log_custom_keywords_textbox.grid(row=3, column=1, sticky="ew", padx=(0, 24), pady=(12, 8))
        self.log_custom_keywords_textbox.insert("end", "")

        option_row = ctk.CTkFrame(config_card, fg_color=COLOR_SURFACE)
        option_row.grid(row=4, column=0, columnspan=2, sticky="ew", padx=24, pady=(0, 18))
        for col in range(6):
            option_row.grid_columnconfigure(col, weight=0)
        ctk.CTkLabel(option_row, text="上下文行数", text_color=COLOR_TEXT, font=self._font(14, "bold")).grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        ctk.CTkEntry(option_row, textvariable=self.log_context_lines_var, width=80, height=36, corner_radius=10, fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER, text_color=COLOR_TEXT).grid(row=0, column=1, sticky="w", padx=(0, 14), pady=4)
        ctk.CTkCheckBox(option_row, text="导入时只分析当前包名", variable=self.log_monitor_pkg_only_var, text_color=COLOR_TEXT, font=self._font(13)).grid(row=0, column=2, sticky="w", padx=(0, 14), pady=4)
        ctk.CTkButton(option_row, text="应用配置", width=104, height=34, corner_radius=12, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color="#FFFFFF", font=self._font(13, "bold"), command=self.apply_log_analysis_config).grid(row=0, column=3, sticky="w", pady=4)

        # 操作与导入：固定两行网格，不随窗口缩小裁切。
        action_card = ctk.CTkFrame(action_tab, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        action_card.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        action_card.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(action_card, text="操作", text_color=COLOR_TEXT, font=self._font(18, "bold")).grid(row=0, column=0, sticky="w", padx=24, pady=(20, 10))
        actions = ctk.CTkFrame(action_card, fg_color=COLOR_SURFACE)
        actions.grid(row=1, column=0, sticky="ew", padx=24, pady=(0, 14))
        for col in range(4):
            actions.grid_columnconfigure(col, weight=1, uniform="log_actions")
        for idx, (text, cmd, accent) in enumerate([
            ("导入日志文件分析", self.import_log_file_for_analysis, True),
            ("打开日志目录", self.open_adb_log_dir, False),
            ("复制命中记录", self.copy_keyword_hits, False),
            ("导出命中 TXT", self.export_keyword_hits_txt, False),
            ("导出命中 Excel", self.export_keyword_hits_excel, False),
            ("复制崩溃结果", self.copy_crash_records, False),
            ("导出崩溃 Excel", self.export_crash_records_excel, False),
            ("清空结果", self.clear_log_analysis_results, False),
        ]):
            ctk.CTkButton(
                actions,
                text=text,
                height=36,
                corner_radius=12,
                fg_color=COLOR_ACCENT if accent else COLOR_SURFACE_2,
                hover_color=COLOR_ACCENT_HOVER if accent else COLOR_HOVER,
                border_width=0 if accent else 1,
                border_color=COLOR_BORDER,
                text_color="#FFFFFF" if accent else COLOR_TEXT,
                font=self._font(13, "bold"),
                command=cmd,
            ).grid(row=idx // 4, column=idx % 4, sticky="ew", padx=(0 if idx % 4 == 0 else 8, 0), pady=5)
        ctk.CTkLabel(action_card, textvariable=self.log_analysis_status_var, text_color=COLOR_MUTED, wraplength=820, font=self._font(13), justify="left").grid(row=2, column=0, sticky="w", padx=24, pady=(0, 18))

        # 结果区：局部 Textbox 独立滚动，不再跟随整页滚动。
        hits_card = ctk.CTkFrame(hits_tab, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        hits_card.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        hits_card.grid_columnconfigure(0, weight=1)
        hits_card.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(hits_card, text="实时命中记录", text_color=COLOR_TEXT, font=self._font(18, "bold")).grid(row=0, column=0, sticky="w", padx=24, pady=(20, 10))
        self.log_hits_box = ctk.CTkTextbox(hits_card, corner_radius=12, border_width=1, border_color=COLOR_BORDER, fg_color=COLOR_SURFACE_2, text_color=COLOR_TEXT, font=self._font(13))
        self.log_hits_box.grid(row=1, column=0, sticky="nsew", padx=24, pady=(0, 20))

        crash_card = ctk.CTkFrame(crash_tab, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        crash_card.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        crash_card.grid_columnconfigure(0, weight=1)
        crash_card.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(crash_card, text="崩溃提取结果", text_color=COLOR_TEXT, font=self._font(18, "bold")).grid(row=0, column=0, sticky="w", padx=24, pady=(20, 10))
        self.crash_results_box = ctk.CTkTextbox(crash_card, corner_radius=12, border_width=1, border_color=COLOR_BORDER, fg_color=COLOR_SURFACE_2, text_color=COLOR_TEXT, font=self._font(13))
        self.crash_results_box.grid(row=1, column=0, sticky="nsew", padx=24, pady=(0, 20))

        self._refresh_log_analysis_views()

    def _get_log_analysis_keywords(self):
        keywords = [keyword for keyword in DEFAULT_KEYWORDS if self.log_keyword_vars.get(keyword) and self.log_keyword_vars[keyword].get()]
        if self.log_custom_keywords_textbox and self.log_custom_keywords_textbox.winfo_exists():
            custom_text = self.log_custom_keywords_textbox.get("1.0", "end").strip()
            keywords.extend([line.strip() for line in custom_text.splitlines() if line.strip()])
        return keywords

    def apply_log_analysis_config(self):
        try:
            context_lines = max(0, min(200, int(self.log_context_lines_var.get().strip() or "20")))
        except ValueError:
            context_lines = 20
            self.log_context_lines_var.set("20")
        self.log_monitor.configure(self._get_log_analysis_keywords(), context_lines=context_lines)
        self.log_analysis_status_var.set(f"已应用关键字配置：{len(self.log_monitor.keywords)} 个关键字，上下文前后各 {context_lines} 行")

    def _on_live_log_line(self, line: str):
        if not self.log_analysis_enabled_var.get():
            return
        try:
            new_hits = self.log_monitor.feed_line(line)
            new_crashes = self.crash_stream_detector.feed_line(line, fallback_package=self.log_pkg_var.get().strip())
            if new_hits or new_crashes:
                self.log_hit_records = self.log_monitor.get_hits()
                self.crash_records = self.crash_stream_detector.get_records()
                self._schedule_log_analysis_refresh()
        except Exception:
            pass

    def _schedule_log_analysis_refresh(self):
        now = time.monotonic()
        if now - self._log_analysis_last_refresh < 0.5:
            return
        self._log_analysis_last_refresh = now
        try:
            self.after(0, self._refresh_log_analysis_views)
        except Exception:
            pass

    def _refresh_log_analysis_views(self):
        if hasattr(self, "log_hits_box") and self.log_hits_box and self.log_hits_box.winfo_exists():
            self._set_textbox_text(self.log_hits_box, format_keyword_hits(self.log_hit_records, max_records=1000))
        if hasattr(self, "crash_results_box") and self.crash_results_box and self.crash_results_box.winfo_exists():
            self._set_textbox_text(self.crash_results_box, format_crash_records(self.crash_records, max_records=500))
        self.log_analysis_status_var.set(
            f"关键字命中 {len(self.log_hit_records)} 条｜崩溃结果 {len(self.crash_records)} 条｜日志显示仅渲染最近记录，导出包含完整结果"
        )

    def clear_log_analysis_results(self):
        self.log_monitor.clear()
        self.crash_stream_detector.clear()
        self.log_hit_records = []
        self.crash_records = []
        self._refresh_log_analysis_views()
        self.log_analysis_status_var.set("已清空日志分析结果")

    def import_log_file_for_analysis(self):
        if self.log_analysis_busy.is_set():
            messagebox.showinfo("提示", "日志分析正在进行中，请稍后")
            return
        path = filedialog.askopenfilename(title="选择日志文件", filetypes=[("日志文件", "*.txt *.log"), ("所有文件", "*.*")])
        if not path:
            return
        self.apply_log_analysis_config()
        self.log_analysis_busy.set()
        self.log_analysis_status_var.set("正在导入并分析日志文件...")

        def worker():
            try:
                with open(path, "r", encoding="utf-8", errors="ignore") as file:
                    lines = file.readlines()
                pkg = self.log_pkg_var.get().strip()
                if self.log_monitor_pkg_only_var.get() and pkg:
                    lines = [line for line in lines if pkg in line]
                hits = self.log_monitor.analyze_lines(lines)
                crashes = self.crash_file_analyzer.analyze_lines(lines, fallback_package=pkg)
                def ok():
                    self.log_hit_records = hits
                    self.crash_records = crashes
                    self.crash_stream_detector.clear()
                    for record in crashes:
                        self.crash_stream_detector.records.append(record)
                    self._refresh_log_analysis_views()
                    self.log_analysis_status_var.set(f"导入分析完成：关键字命中 {len(hits)} 条，崩溃结果 {len(crashes)} 条")
            except Exception as exc:
                def ok(exc=exc):
                    messagebox.showerror("分析失败", str(exc))
                    self.log_analysis_status_var.set("日志文件分析失败")
            finally:
                def done():
                    self.log_analysis_busy.clear()
                try:
                    self.after(0, ok)
                    self.after(0, done)
                except Exception:
                    self.log_analysis_busy.clear()
        threading.Thread(target=worker, daemon=True).start()

    def copy_keyword_hits(self):
        if not self.log_hit_records:
            messagebox.showinfo("提示", "暂无关键字命中记录")
            return
        self.clipboard_clear()
        self.clipboard_append(keyword_hits_to_tsv(self.log_hit_records))
        messagebox.showinfo("已复制", f"已复制 {len(self.log_hit_records)} 条关键字命中记录")

    def copy_crash_records(self):
        if not self.crash_records:
            messagebox.showinfo("提示", "暂无崩溃提取结果")
            return
        self.clipboard_clear()
        self.clipboard_append(crash_records_to_tsv(self.crash_records))
        messagebox.showinfo("已复制", f"已复制 {len(self.crash_records)} 条崩溃提取结果")

    def export_keyword_hits_txt(self):
        if not self.log_hit_records:
            messagebox.showinfo("提示", "暂无关键字命中记录")
            return
        output_path = filedialog.asksaveasfilename(title="导出关键字命中 TXT", defaultextension=".txt", filetypes=[("文本文件", "*.txt")], initialfile="日志关键字命中.txt")
        if not output_path:
            return
        try:
            export_keyword_hits_to_txt(self.log_hit_records, output_path)
            messagebox.showinfo("导出成功", f"已导出：{output_path}")
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))

    def export_keyword_hits_excel(self):
        if not self.log_hit_records:
            messagebox.showinfo("提示", "暂无关键字命中记录")
            return
        output_path = filedialog.asksaveasfilename(title="导出关键字命中 Excel", defaultextension=".xlsx", filetypes=[("Excel 文件", "*.xlsx")], initialfile="日志关键字命中.xlsx")
        if not output_path:
            return
        try:
            export_keyword_hits_to_excel(self.log_hit_records, output_path)
            messagebox.showinfo("导出成功", f"已导出：{output_path}")
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))

    def export_crash_records_excel(self):
        if not self.crash_records:
            messagebox.showinfo("提示", "暂无崩溃提取结果")
            return
        output_path = filedialog.asksaveasfilename(title="导出崩溃日志 Excel", defaultextension=".xlsx", filetypes=[("Excel 文件", "*.xlsx")], initialfile="崩溃日志提取报告.xlsx")
        if not output_path:
            return
        try:
            export_crash_records_to_excel(self.crash_records, output_path)
            messagebox.showinfo("导出成功", f"已导出：{output_path}")
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))

    # ---------- ADB Tools page ----------
    def _build_adb_tools_page(self):
        """v2.4.19: ADB 工具页改为 Tab 固定布局 + 局部 Textbox 滚动。"""
        wrap = ctk.CTkFrame(self.content, fg_color=COLOR_BG)
        wrap.grid(row=0, column=0, sticky="nsew")
        wrap.grid_columnconfigure(0, weight=1)
        wrap.grid_rowconfigure(0, weight=1)

        tabview = ctk.CTkTabview(
            wrap,
            fg_color=COLOR_BG,
            segmented_button_fg_color=COLOR_SURFACE_2,
            segmented_button_selected_color=COLOR_ACCENT,
            segmented_button_selected_hover_color=COLOR_ACCENT_HOVER,
            segmented_button_unselected_color=COLOR_SURFACE_2,
            segmented_button_unselected_hover_color=COLOR_HOVER,
            text_color=COLOR_TEXT,
            corner_radius=12,
        )
        tabview.grid(row=0, column=0, sticky="nsew")

        device_tab = tabview.add("设备状态")
        media_tab = tabview.add("截图 / 录屏")
        app_tab = tabview.add("应用管理")
        apk_tab = tabview.add("APK 信息")
        log_tab = tabview.add("操作日志")
        for tab in (device_tab, media_tab, app_tab, apk_tab, log_tab):
            tab.configure(fg_color=COLOR_BG)
            tab.grid_columnconfigure(0, weight=1)
            tab.grid_rowconfigure(0, weight=1)

        # 设备状态：v2.4.27 多设备表格 + 当前设备详情
        device_card = ctk.CTkFrame(device_tab, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        device_card.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        device_card.grid_columnconfigure(0, weight=1)
        device_card.grid_rowconfigure(3, weight=1)
        device_card.grid_rowconfigure(4, weight=1)
        ctk.CTkLabel(device_card, text="设备状态", text_color=COLOR_TEXT, font=self._font(18, "bold")).grid(row=0, column=0, sticky="w", padx=24, pady=(20, 8))
        self._build_adb_device_selector(device_card, 1, show_actions=True)

        device_actions = ctk.CTkFrame(device_card, fg_color="transparent")
        device_actions.grid(row=2, column=0, sticky="ew", padx=24, pady=(0, 12))
        ctk.CTkButton(device_actions, text="读取当前设备详情", width=156, height=36, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self.adb_tools_read_device_info).grid(row=0, column=0, padx=(0, 10), pady=4)
        ctk.CTkButton(device_actions, text="识别前台应用", width=132, height=36, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self.adb_tools_get_foreground).grid(row=0, column=1, padx=(0, 10), pady=4)
        ctk.CTkLabel(device_actions, textvariable=self.adb_device_hint_var, text_color=COLOR_MUTED, wraplength=720, justify="left", font=self._font(12)).grid(row=0, column=2, sticky="w", padx=(8, 0), pady=4)

        self._build_adb_device_table(device_card, row=3)
        self._build_adb_device_detail_card(device_card, row=4)

        # 截图 / 录屏
        media_card = ctk.CTkFrame(media_tab, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        media_card.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        media_card.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(media_card, text="截图 / 录屏", text_color=COLOR_TEXT, font=self._font(18, "bold")).grid(row=0, column=0, sticky="w", padx=24, pady=(20, 10))
        media_buttons = ctk.CTkFrame(media_card, fg_color=COLOR_SURFACE)
        media_buttons.grid(row=1, column=0, sticky="w", padx=24, pady=(0, 14))
        ctk.CTkButton(media_buttons, text="一键截图", width=112, height=36, corner_radius=12, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color="#FFFFFF", font=self._font(13, "bold"), command=self.adb_tools_screenshot).grid(row=0, column=0, padx=(0, 10), pady=4)
        ctk.CTkButton(media_buttons, text="打开截图目录", width=132, height=36, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=lambda: self._open_local_dir(self.adb_tools.screenshot_dir)).grid(row=0, column=1, padx=(0, 10), pady=4)
        self.adb_record_button = ctk.CTkButton(media_buttons, text="开始录屏", width=112, height=36, corner_radius=12, fg_color=COLOR_DANGER if self.adb_tools.is_recording() else COLOR_ACCENT, hover_color=COLOR_DANGER_HOVER if self.adb_tools.is_recording() else COLOR_ACCENT_HOVER, text_color="#FFFFFF", font=self._font(13, "bold"), command=self.adb_tools_toggle_recording)
        self.adb_record_button.grid(row=0, column=2, padx=(0, 10), pady=4)
        ctk.CTkButton(media_buttons, text="打开录屏目录", width=132, height=36, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=lambda: self._open_local_dir(self.adb_tools.recording_dir)).grid(row=0, column=3, pady=4)
        ctk.CTkLabel(media_card, textvariable=self.adb_recording_status_var, text_color=COLOR_MUTED, wraplength=760, font=self._font(13), justify="left").grid(row=2, column=0, sticky="w", padx=24, pady=(0, 18))

        # 应用管理
        app_card = ctk.CTkFrame(app_tab, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        app_card.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        app_card.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(app_card, text="应用管理", text_color=COLOR_TEXT, font=self._font(18, "bold")).grid(row=0, column=0, sticky="w", padx=24, pady=(20, 4), columnspan=3)
        self._entry_row(app_card, 1, "应用包名", self.adb_tools_pkg_var, "例如 com.example.game", width=520)
        app_buttons = ctk.CTkFrame(app_card, fg_color=COLOR_SURFACE)
        app_buttons.grid(row=2, column=0, columnspan=3, sticky="ew", padx=24, pady=(0, 18))
        for col in range(3):
            app_buttons.grid_columnconfigure(col, weight=1, uniform="adb_app_btns")
        for idx, (text, cmd, danger) in enumerate([
            ("启动应用", self.adb_tools_start_app, False),
            ("停止应用", self.adb_tools_stop_app, False),
            ("重启应用", self.adb_tools_restart_app, False),
            ("清除数据", self.adb_tools_clear_app_data, True),
            ("卸载应用", self.adb_tools_uninstall_app, True),
        ]):
            ctk.CTkButton(
                app_buttons,
                text=text,
                height=38,
                corner_radius=12,
                fg_color=COLOR_DANGER if danger else COLOR_SURFACE_2,
                hover_color=COLOR_DANGER_HOVER if danger else COLOR_HOVER,
                border_width=0 if danger else 1,
                border_color=COLOR_BORDER,
                text_color="#FFFFFF" if danger else COLOR_TEXT,
                font=self._font(13, "bold"),
                command=cmd,
            ).grid(row=idx // 3, column=idx % 3, sticky="ew", padx=(0 if idx % 3 == 0 else 8, 0), pady=5)

        # APK 信息
        apk_card = ctk.CTkFrame(apk_tab, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        apk_card.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        apk_card.grid_columnconfigure(1, weight=1)
        apk_card.grid_rowconfigure(4, weight=1)
        ctk.CTkLabel(apk_card, text="APK 信息", text_color=COLOR_TEXT, font=self._font(18, "bold")).grid(row=0, column=0, sticky="w", padx=24, pady=(20, 4), columnspan=3)
        self._entry_row(apk_card, 1, "APK 文件", self.apk_path_var, "请选择本地 APK", width=520)
        apk_buttons = ctk.CTkFrame(apk_card, fg_color=COLOR_SURFACE)
        apk_buttons.grid(row=2, column=0, columnspan=3, sticky="w", padx=24, pady=(0, 14))
        ctk.CTkButton(apk_buttons, text="选择 APK", width=112, height=36, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self.adb_tools_choose_apk).grid(row=0, column=0, padx=(0, 10))
        ctk.CTkButton(apk_buttons, text="解析 APK 信息", width=132, height=36, corner_radius=12, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color="#FFFFFF", font=self._font(13, "bold"), command=self.adb_tools_analyze_apk).grid(row=0, column=1, padx=(0, 10))
        ctk.CTkButton(apk_buttons, text="安装 APK", width=112, height=36, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self.adb_tools_install_apk).grid(row=0, column=2)
        ctk.CTkLabel(apk_card, textvariable=self.apk_status_var, text_color=COLOR_MUTED, wraplength=760, font=self._font(13)).grid(row=3, column=0, columnspan=3, sticky="w", padx=24, pady=(0, 8))
        self.apk_info_box = ctk.CTkTextbox(apk_card, corner_radius=12, border_width=1, border_color=COLOR_BORDER, fg_color=COLOR_SURFACE_2, text_color=COLOR_TEXT, font=self._font(13))
        self.apk_info_box.grid(row=4, column=0, columnspan=3, sticky="nsew", padx=24, pady=(0, 20))
        self._set_textbox_text(self.apk_info_box, "请选择 APK 后点击解析。")

        # 操作日志
        log_card = ctk.CTkFrame(log_tab, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        log_card.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        log_card.grid_columnconfigure(0, weight=1)
        log_card.grid_rowconfigure(2, weight=1)
        ctk.CTkLabel(log_card, text="操作日志", text_color=COLOR_TEXT, font=self._font(18, "bold")).grid(row=0, column=0, sticky="w", padx=24, pady=(20, 10))
        log_buttons = ctk.CTkFrame(log_card, fg_color=COLOR_SURFACE)
        log_buttons.grid(row=1, column=0, sticky="w", padx=24, pady=(0, 10))
        ctk.CTkButton(log_buttons, text="复制日志", width=112, height=34, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self.adb_tools_copy_log).grid(row=0, column=0, padx=(0, 10))
        ctk.CTkButton(log_buttons, text="清空日志", width=112, height=34, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self.adb_tools_clear_log).grid(row=0, column=1)
        self.adb_tools_log_box = ctk.CTkTextbox(log_card, corner_radius=12, border_width=1, border_color=COLOR_BORDER, fg_color=COLOR_SURFACE_2, text_color=COLOR_TEXT, font=self._font(13))
        self.adb_tools_log_box.grid(row=2, column=0, sticky="nsew", padx=24, pady=(0, 20))
        self.adb_tools_log_box.insert("end", "ADB 工具已就绪。所有操作均在本地执行，不上传文件。\n")
        self.adb_tools_log_box.configure(state="disabled")
        self._refresh_adb_tools_buttons()

    def _build_adb_device_selector(self, parent, row, show_actions=False):
        """Build a reusable responsive device toolbar for ADB pages.

        v2.4.31: split action buttons and device selector into two lines when
        actions are shown. This avoids clipping in non-fullscreen windows while
        keeping the serial-bound device selection logic unchanged.
        """
        try:
            parent.grid_columnconfigure(0, weight=1)
        except Exception:
            pass
        toolbar = ctk.CTkFrame(parent, fg_color="transparent")
        toolbar.grid(row=row, column=0, columnspan=3, sticky="ew", padx=24, pady=(4, 12))
        toolbar.grid_columnconfigure(0, weight=0)
        toolbar.grid_columnconfigure(1, weight=1)
        toolbar.grid_columnconfigure(2, weight=0)

        if show_actions:
            action_line = ctk.CTkFrame(toolbar, fg_color="transparent")
            action_line.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 4))
            action_line.grid_columnconfigure(0, weight=0)
            action_line.grid_columnconfigure(1, weight=0)
            action_line.grid_columnconfigure(2, weight=1)
            ctk.CTkButton(action_line, text="刷新设备", width=96, height=36, corner_radius=12, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color="#FFFFFF", font=self._font(13, "bold"), command=self.adb_tools_check_devices).grid(row=0, column=0, padx=(0, 8), pady=4, sticky="w")
            ctk.CTkButton(action_line, text="重启 ADB", width=96, height=36, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self.restart_adb_server).grid(row=0, column=1, padx=(0, 8), pady=4, sticky="w")
            selector_row = 1
        else:
            selector_row = 0

        ctk.CTkLabel(
            toolbar,
            text="当前设备",
            text_color=COLOR_TEXT,
            font=self._font(14, "bold"),
        ).grid(row=selector_row, column=0, sticky="w", padx=(0, 10), pady=4)
        values = self.adb_device_choice_values or ["未选择设备"]
        menu = ctk.CTkOptionMenu(
            toolbar,
            variable=self.adb_selected_device_var,
            values=values,
            width=360,
            height=36,
            corner_radius=10,
            fg_color=COLOR_SURFACE_2,
            button_color=COLOR_SURFACE_2,
            button_hover_color=COLOR_HOVER,
            dropdown_fg_color=COLOR_SURFACE,
            dropdown_hover_color=COLOR_HOVER,
            dropdown_text_color=COLOR_TEXT,
            text_color=COLOR_TEXT,
            font=self._font(13),
            dropdown_font=self._font(13),
            command=self._on_adb_device_selected,
        )
        menu.grid(row=selector_row, column=1, sticky="ew", padx=(0, 10), pady=4)
        if not hasattr(self, "adb_device_menus"):
            self.adb_device_menus = []
        self.adb_device_menus.append(menu)

        status_label = ctk.CTkLabel(
            toolbar,
            textvariable=self.adb_device_status_badge_var,
            width=96,
            height=28,
            corner_radius=14,
            fg_color=COLOR_SURFACE_2,
            text_color=COLOR_TEXT,
            font=self._font(12, "bold"),
        )
        status_label.grid(row=selector_row, column=2, sticky="e", pady=4)
        self.adb_device_status_labels.append(status_label)

    def _adb_device_label_from_item(self, item):
        serial = str(item.get("serial", ""))
        model = str(item.get("model") or item.get("device") or item.get("product") or "未知型号")
        version = str(item.get("android_version") or "?")
        status = str(item.get("status") or "?")
        connection = str(item.get("connection") or "")
        resolution = str(item.get("resolution") or "未知")
        return f"{serial} | {model} | Android {version} | {resolution} | {status} | {connection}"

    def _adb_status_colors(self, status):
        value = str(status or "").lower()
        if value == "device":
            return COLOR_SUCCESS, "#FFFFFF"
        if value == "unauthorized":
            return COLOR_WARNING, "#111111"
        if value == "offline":
            return COLOR_DANGER, "#FFFFFF"
        if value in ("未选择", "未检测"):
            return COLOR_SURFACE_2, COLOR_MUTED
        return COLOR_SURFACE_2, COLOR_TEXT

    def _update_adb_status_badge(self, status):
        text = status or "未选择"
        self.adb_device_status_badge_var.set(text)
        fg, tc = self._adb_status_colors(text)
        for label in getattr(self, "adb_device_status_labels", []):
            try:
                if label.winfo_exists():
                    label.configure(fg_color=fg, text_color=tc)
            except Exception:
                pass

    def _build_adb_device_table(self, parent, row):
        table_frame = ctk.CTkFrame(parent, fg_color=COLOR_SURFACE_2, corner_radius=12, border_width=1, border_color=COLOR_BORDER)
        table_frame.grid(row=row, column=0, sticky="nsew", padx=24, pady=(0, 12))
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)
        columns = ("current", "serial", "model", "brand", "android", "sdk", "resolution", "status", "connection", "foreground")
        tree = ttk.Treeview(table_frame, columns=columns, show="headings", style="Xiao.Treeview", height=7)
        headings = {
            "current": "当前",
            "serial": "serial",
            "model": "型号",
            "brand": "品牌",
            "android": "Android",
            "sdk": "API",
            "resolution": "分辨率",
            "status": "状态",
            "connection": "连接方式",
            "foreground": "前台包名",
        }
        widths = {"current": 52, "serial": 120, "model": 130, "brand": 90, "android": 90, "sdk": 70, "resolution": 118, "status": 98, "connection": 90, "foreground": 220}
        for col in columns:
            tree.heading(col, text=headings[col])
            tree.column(col, width=widths[col], minwidth=widths[col], anchor="w", stretch=col == "foreground")
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        # v2.4.28: 使用鼠标释放事件处理用户点击；避免程序化 selection_set 触发递归刷新。
        tree.bind("<ButtonRelease-1>", self._on_adb_device_table_select)
        self.adb_device_table = tree

    def _build_adb_device_detail_card(self, parent, row):
        detail = ctk.CTkFrame(parent, fg_color=COLOR_SURFACE_2, corner_radius=12, border_width=1, border_color=COLOR_BORDER)
        detail.grid(row=row, column=0, sticky="nsew", padx=24, pady=(0, 20))
        detail.grid_columnconfigure(0, weight=1)
        detail.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(detail, text="当前设备详情", text_color=COLOR_TEXT, font=self._font(15, "bold")).grid(row=0, column=0, sticky="w", padx=16, pady=(12, 6))
        self.adb_device_detail_box = ctk.CTkTextbox(detail, height=150, corner_radius=10, border_width=0, fg_color=COLOR_SURFACE_2, text_color=COLOR_TEXT, font=self._font(13))
        self.adb_device_detail_box.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 12))
        self._set_textbox_text(self.adb_device_detail_box, "未检测到 Android 设备，请连接设备并开启 USB 调试，然后点击刷新设备。")

    def _device_by_serial(self, serial, data=None):
        data = data or self._adb_last_device_data or {}
        for item in data.get("devices", []) or []:
            if item.get("serial") == serial:
                return item
        return None

    def _format_device_detail(self, item=None):
        if not item:
            return "未检测到 Android 设备，请连接设备并开启 USB 调试，然后点击刷新设备。"
        status = item.get("status", "")
        tips = []
        if status == "unauthorized":
            tips.append("提示：设备未授权，请在手机上允许 USB 调试授权。")
        elif status == "offline":
            tips.append("提示：设备处于 offline 状态，请重新插拔设备或重启 ADB 服务。")
        lines = [
            f"serial：{item.get('serial', '') or '未获取'}",
            f"型号：{item.get('model') or item.get('device') or item.get('product') or '未获取'}",
            f"品牌：{item.get('brand') or '未获取'}",
            f"Android 版本：{item.get('android_version') or '未获取'}",
            f"API Level：{item.get('sdk') or '未获取'}",
            f"分辨率：{item.get('resolution') or '未知'}",
            f"连接方式：{item.get('connection') or '未获取'}",
            f"当前前台包名：{item.get('foreground_package') or '未识别'}",
            f"状态：{status or '未获取'}",
            f"最后刷新时间：{item.get('last_refresh') or '未获取'}",
        ]
        if tips:
            lines.append("")
            lines.extend(tips)
        return "\n".join(lines)

    def _refresh_adb_device_table(self, data=None):
        data = data or self._adb_last_device_data or {}
        tree = getattr(self, "adb_device_table", None)
        if not tree or not tree.winfo_exists():
            return
        self._adb_device_table_updating = True
        try:
            tree.delete(*tree.get_children())
            selected = data.get("selected_serial") or self.adb_tools.get_selected_serial()
            devices = data.get("devices", []) or []
            if not devices:
                self.adb_device_hint_var.set("未检测到 Android 设备，请连接设备并开启 USB 调试，然后点击刷新设备。")
                self._set_textbox_text(self.adb_device_detail_box, self._format_device_detail(None))
                return
            for item in devices:
                serial = item.get("serial", "")
                if not serial:
                    continue
                current = "✓" if serial == selected else ""
                values = (
                    current,
                    serial,
                    item.get("model") or item.get("device") or item.get("product") or "未知型号",
                    item.get("brand") or "",
                    item.get("android_version") or "",
                    item.get("sdk") or "",
                    item.get("resolution") or "未知",
                    item.get("status") or "",
                    item.get("connection") or "",
                    item.get("foreground_package") or "",
                )
                tree.insert("", "end", iid=serial, values=values)
                if serial == selected:
                    try:
                        tree.selection_set(serial)
                        tree.see(serial)
                    except Exception:
                        pass
            ready = [item for item in devices if item.get("status") == "device"]
            if len(ready) > 1 and not selected:
                self.adb_device_hint_var.set("检测到多台 Android 设备，请先选择目标设备。")
            elif len(ready) == 1 and selected:
                self.adb_device_hint_var.set("已自动选择唯一可用设备。")
            else:
                self.adb_device_hint_var.set(f"已识别设备 {len(devices)} 台，可用 device 状态 {len(ready)} 台。")
        except Exception as exc:
            try:
                self.adb_device_hint_var.set(f"刷新设备表格失败：{exc}")
            except Exception:
                pass
        finally:
            self._adb_device_table_updating = False
            self._refresh_adb_device_detail(data)

    def _refresh_adb_device_detail(self, data=None):
        data = data or self._adb_last_device_data or {}
        serial = data.get("selected_serial") or self.adb_tools.get_selected_serial()
        item = self._device_by_serial(serial, data) if serial else None
        if hasattr(self, "adb_device_detail_box") and self.adb_device_detail_box.winfo_exists():
            self._set_textbox_text(self.adb_device_detail_box, self._format_device_detail(item))

    def _on_adb_device_table_select(self, event=None):
        # v2.4.28: 只响应用户点击，不响应刷新表格过程中的程序化选中。
        if getattr(self, "_adb_device_table_updating", False):
            return
        tree = getattr(self, "adb_device_table", None)
        if not tree:
            return
        serial = ""
        if event is not None:
            try:
                serial = tree.identify_row(event.y)
            except Exception:
                serial = ""
        if not serial:
            selected = tree.selection()
            if selected:
                serial = selected[0]
        if not serial:
            return
        item = self._device_by_serial(serial)
        if not item:
            return
        if item.get("status") != "device":
            self._refresh_adb_device_detail(self._adb_last_device_data)
            status = item.get("status") or "不可用"
            if status == "unauthorized":
                self.adb_device_hint_var.set("设备未授权，请在手机上允许 USB 调试授权。")
            elif status == "offline":
                self.adb_device_hint_var.set("设备处于 offline 状态，请重新插拔设备或重启 ADB 服务。")
            else:
                self.adb_device_hint_var.set(f"目标设备不可用：{serial}={status}")
            return

        # 选中设备本身不再触发 ADB 查询；只更新内存状态和 UI，避免点击行时卡住主线程。
        self.adb_tools.set_selected_serial(serial)
        label = self._adb_device_label_from_item(item)
        self.adb_selected_device_var.set(label)
        self.adb_selected_device_label_var.set(f"当前设备：{label}")
        self._update_adb_status_badge("device")
        if isinstance(self._adb_last_device_data, dict):
            self._adb_last_device_data["selected_serial"] = serial
        self._refresh_adb_device_table(self._adb_last_device_data)
        self._refresh_adb_log_task_table(self._adb_last_device_data)
        try:
            self._adb_tools_log(f"已选择目标设备：{serial}")
            self._adb_log(f"已选择目标设备：{serial}")
        except Exception:
            pass

    def _update_adb_device_choices(self, data=None):
        try:
            data = data or self.adb_tools.list_devices()
            self._adb_last_device_data = data
            devices = data.get("devices", []) or []
            choices = []
            for item in devices:
                choices.append(self._adb_device_label_from_item(item))
            if not choices:
                choices = ["未选择设备"]
            self.adb_device_choice_values = choices
            selected_serial = data.get("selected_serial") or self.adb_tools.get_selected_serial()
            selected_label = ""
            selected_status = "未选择"
            for item in devices:
                if item.get("serial") == selected_serial:
                    selected_label = self._adb_device_label_from_item(item)
                    selected_status = item.get("status") or "未选择"
                    break
            if selected_label:
                self.adb_selected_device_var.set(selected_label)
                self.adb_selected_device_label_var.set(f"当前设备：{selected_label}")
            else:
                self.adb_selected_device_var.set("未选择设备")
                self.adb_selected_device_label_var.set("当前设备：未选择")
            self._update_adb_status_badge(selected_status)
            for menu in getattr(self, "adb_device_menus", []):
                try:
                    if menu.winfo_exists():
                        menu.configure(values=choices)
                except Exception:
                    pass
            self._refresh_adb_device_table(data)
            self._refresh_adb_log_task_table(data)
            return data
        except Exception as exc:
            self.adb_selected_device_label_var.set(f"当前设备：识别失败｜{exc}")
            self._update_adb_status_badge("异常")
            return None

    def _on_adb_device_selected(self, choice):
        text = str(choice or "").strip()
        if not text or text.startswith("未选择"):
            self.adb_tools.set_selected_serial("")
            self.adb_selected_device_label_var.set("当前设备：未选择")
            self._update_adb_status_badge("未选择")
            return
        serial = text.split("|", 1)[0].strip()
        item = self._device_by_serial(serial)
        if item and item.get("status") != "device":
            self.adb_tools.set_selected_serial("")
            self.adb_selected_device_label_var.set("当前设备：未选择")
            self._update_adb_status_badge(item.get("status") or "不可用")
            if item.get("status") == "unauthorized":
                messagebox.showwarning("设备未授权", "设备未授权，请在手机上允许 USB 调试授权。")
            elif item.get("status") == "offline":
                messagebox.showwarning("设备 offline", "设备处于 offline 状态，请重新插拔设备或重启 ADB 服务。")
            return
        self.adb_tools.set_selected_serial(serial)
        self.adb_selected_device_label_var.set(f"当前设备：{text}")
        self._update_adb_status_badge(item.get("status") if item else "device")
        self._refresh_adb_device_table(self._adb_last_device_data)
        try:
            self._adb_tools_log(f"已选择目标设备：{serial}")
            self._adb_log(f"已选择目标设备：{serial}")
        except Exception:
            pass

    def _current_adb_serial_or_warn(self):
        try:
            serial = self.adb_tools.resolve_serial()
            self.adb_tools.set_selected_serial(serial)
            return serial
        except Exception as exc:
            messagebox.showwarning("请选择设备", str(exc))
            return ""

    def _set_textbox_text(self, textbox, text):
        try:
            textbox.configure(state="normal")
            textbox.delete("1.0", "end")
            textbox.insert("end", text or "")
            textbox.configure(state="disabled")
        except Exception:
            pass

    def _adb_tools_log(self, text):
        line = f"[{time.strftime('%H:%M:%S')}] {text}\n"
        def write():
            if hasattr(self, "adb_tools_log_box") and self.adb_tools_log_box.winfo_exists():
                self.adb_tools_log_box.configure(state="normal")
                self.adb_tools_log_box.insert("end", line)
                self.adb_tools_log_box.see("end")
                self.adb_tools_log_box.configure(state="disabled")
        try:
            self.after(0, write)
        except Exception:
            pass

    def _run_adb_tools_task(self, title, worker, on_success=None):
        if self.adb_tools_busy.is_set():
            self._adb_tools_log("已有 ADB 操作正在执行，请稍后再试")
            return
        self.adb_tools_busy.set()
        self._set_task_phase("adb_tools", TaskPhase.STARTING, title)
        self._adb_tools_log(f"开始：{title}")
        self._refresh_adb_tools_buttons()

        def run():
            try:
                self._set_task_phase("adb_tools", TaskPhase.RUNNING, title)
                result = worker()
                def ok():
                    if on_success:
                        on_success(result)
                    self._set_task_phase("adb_tools", TaskPhase.SUCCEEDED, f"{title}已完成")
                    self._adb_tools_log(f"完成：{title}")
            except Exception as exc:
                def ok(exc=exc):
                    self._set_task_phase("adb_tools", TaskPhase.FAILED, str(exc))
                    self._adb_tools_log(f"失败：{title}｜{exc}")
                    messagebox.showerror("操作失败", str(exc))
            finally:
                def done():
                    self.adb_tools_busy.clear()
                    if self.adb_tools.is_recording():
                        self._set_task_phase("adb_tools", TaskPhase.RUNNING, "录屏中")
                    elif self._task_snapshot("adb_tools").phase != TaskPhase.FAILED:
                        self._set_task_phase("adb_tools", TaskPhase.IDLE)
                    self._refresh_adb_tools_buttons()
                try:
                    self.after(0, ok)
                    self.after(0, done)
                except Exception:
                    self.adb_tools_busy.clear()
        threading.Thread(target=run, daemon=True).start()

    def _refresh_adb_tools_buttons(self):
        if hasattr(self, "adb_record_button") and self.adb_record_button.winfo_exists():
            recording = self.adb_tools.is_recording()
            self.adb_record_button.configure(
                text="停止录屏" if recording else "开始录屏",
                fg_color=COLOR_DANGER if recording else COLOR_ACCENT,
                hover_color=COLOR_DANGER_HOVER if recording else COLOR_ACCENT_HOVER,
            )
        if self.adb_tools.is_recording():
            self.adb_recording_status_var.set("录屏中")

    def restart_adb_server(self):
        def worker():
            kill = self.adb_tools.run(["kill-server"], timeout=8)
            start = self.adb_tools.run(["start-server"], timeout=8)
            generation = self._begin_adb_refresh()
            data = self._scan_adb_devices_staged(generation, refresh_resolutions=True)
            return kill, start, data
        def success(result):
            kill, start, data = result
            self._adb_tools_log("ADB 服务已重启" if start.ok else f"ADB 服务重启可能失败：{start.output}")
            self._update_adb_device_choices(data)
            text = self.adb_tools.format_devices_text(data)
            self.adb_tools_devices_var.set(text)
            if hasattr(self, "adb_tools_devices_box") and self.adb_tools_devices_box.winfo_exists():
                self._set_textbox_text(self.adb_tools_devices_box, text)
        self._run_adb_tools_task("重启 ADB 服务", worker, success)

    def adb_tools_check_devices(self):
        def worker():
            generation = self._begin_adb_refresh()
            return self._scan_adb_devices_staged(generation, refresh_resolutions=True)
        def success(data):
            self._update_adb_device_choices(data)
            text = self.adb_tools.format_devices_text(data)
            self.adb_tools_devices_var.set(text)
            if hasattr(self, "adb_tools_devices_box") and self.adb_tools_devices_box.winfo_exists():
                self._set_textbox_text(self.adb_tools_devices_box, text)
        self._run_adb_tools_task("检测设备", worker, success)

    def adb_tools_read_device_info(self):
        def worker():
            return self.adb_tools.get_device_info()
        def success(info):
            text = "\n".join(f"{k}：{v or '未获取'}" for k, v in info.items())
            self.adb_tools_info_var.set(text)
            if info.get("前台包名"):
                self.adb_tools_pkg_var.set(info.get("前台包名", ""))
            if hasattr(self, "adb_tools_info_box") and self.adb_tools_info_box.winfo_exists():
                self._set_textbox_text(self.adb_tools_info_box, text)
        self._run_adb_tools_task("读取设备信息", worker, success)

    def adb_tools_screenshot(self):
        def worker():
            return self.adb_tools.take_screenshot()
        def success(path):
            self._adb_tools_log(f"截图已保存：{path}")
            messagebox.showinfo("截图完成", f"截图已保存：\n{path}")
        self._run_adb_tools_task("一键截图", worker, success)

    def adb_tools_toggle_recording(self):
        if self.adb_tools.is_recording():
            def worker():
                return self.adb_tools.stop_recording()
            def success(path):
                self.adb_recording_status_var.set("录屏已保存")
                self._adb_tools_log(f"录屏已保存：{path}")
                messagebox.showinfo("录屏完成", f"录屏已保存：\n{path}")
            self._run_adb_tools_task("停止录屏", worker, success)
        else:
            def worker():
                return self.adb_tools.start_recording()
            def success(session):
                self.adb_recording_session = session
                self.adb_recording_status_var.set("录屏中")
                self._adb_tools_log(f"录屏已开始：{session.device_path}")
                self._refresh_adb_tools_buttons()
            self._run_adb_tools_task("开始录屏", worker, success)

    def adb_tools_choose_apk(self):
        path = filedialog.askopenfilename(title="选择 APK 文件", filetypes=[("APK 文件", "*.apk"), ("所有文件", "*.*")])
        if path:
            self.apk_path_var.set(path)
            self.apk_status_var.set("已选择 APK")

    def adb_tools_install_apk(self):
        path = self.apk_path_var.get().strip()
        if not path:
            path = filedialog.askopenfilename(title="选择 APK 文件", filetypes=[("APK 文件", "*.apk"), ("所有文件", "*.*")])
            if path:
                self.apk_path_var.set(path)
        if not path:
            return
        def worker():
            return self.adb_tools.install_apk(path)
        def success(result):
            text = result.output or ("安装成功" if result.ok else "安装失败")
            self._adb_tools_log(f"APK 安装结果：{text}")
            if result.ok:
                messagebox.showinfo("安装完成", text)
            else:
                messagebox.showerror("安装失败", text)
        self._run_adb_tools_task("安装 APK", worker, success)

    def adb_tools_analyze_apk(self):
        path = self.apk_path_var.get().strip()
        if not path:
            self.adb_tools_choose_apk()
            path = self.apk_path_var.get().strip()
        if not path:
            return
        def worker():
            return format_apk_info(analyze_apk(path))
        def success(text):
            self.apk_status_var.set("解析完成")
            if hasattr(self, "apk_info_box") and self.apk_info_box.winfo_exists():
                self._set_textbox_text(self.apk_info_box, text)
        self._run_adb_tools_task("解析 APK 信息", worker, success)

    def adb_tools_get_foreground(self):
        def worker():
            return self.adb_tools.get_foreground_app()
        def success(fg):
            pkg = fg.get("package", "")
            activity = fg.get("activity", "")
            full_activity = fg.get("full_activity", "")
            source = fg.get("source", "")
            command = fg.get("command", "")
            raw = fg.get("raw", "")
            if pkg:
                self.adb_tools_pkg_var.set(pkg)
            text = (
                f"前台包名：{pkg or '未识别'}\n"
                f"Activity：{activity or '未识别'}\n"
                f"完整 Activity：{full_activity or '未识别'}\n"
                f"识别来源：{source or '未识别'}" + (f"｜{command}" if command else "") +
                f"\n原始信息：{raw or '无'}"
            )
            self.adb_tools_foreground_var.set(text)
            self._adb_tools_log(text)
        self._run_adb_tools_task("识别前台应用", worker, success)

    def _get_adb_tools_pkg(self):
        pkg = self.adb_tools_pkg_var.get().strip()
        if not pkg:
            messagebox.showwarning("缺少包名", "请先输入应用包名，或点击“识别前台应用”。")
            return ""
        return pkg

    def adb_tools_start_app(self):
        pkg = self._get_adb_tools_pkg()
        if not pkg:
            return
        self._run_adb_tools_task("启动应用", lambda: self.adb_tools.start_app(pkg), lambda r: self._adb_tools_log(r.output or ("启动命令已执行" if r.ok else "启动失败")))

    def adb_tools_stop_app(self):
        pkg = self._get_adb_tools_pkg()
        if not pkg:
            return
        self._run_adb_tools_task("停止应用", lambda: self.adb_tools.stop_app(pkg), lambda r: self._adb_tools_log(r.output or ("停止命令已执行" if r.ok else "停止失败")))

    def adb_tools_restart_app(self):
        pkg = self._get_adb_tools_pkg()
        if not pkg:
            return
        self._run_adb_tools_task("重启应用", lambda: self.adb_tools.restart_app(pkg), lambda r: self._adb_tools_log(r.output or ("重启命令已执行" if r.ok else "重启失败")))

    def adb_tools_clear_app_data(self):
        pkg = self._get_adb_tools_pkg()
        if not pkg:
            return
        if not messagebox.askyesno("确认清除数据", f"确定清除应用数据？\n{pkg}"):
            return
        self._run_adb_tools_task("清除应用数据", lambda: self.adb_tools.clear_app_data(pkg), lambda r: self._adb_tools_log(r.output or ("清除完成" if r.ok else "清除失败")))

    def adb_tools_uninstall_app(self):
        pkg = self._get_adb_tools_pkg()
        if not pkg:
            return
        if not messagebox.askyesno("确认卸载应用", f"确定卸载应用？\n{pkg}"):
            return
        self._run_adb_tools_task("卸载应用", lambda: self.adb_tools.uninstall_app(pkg), lambda r: self._adb_tools_log(r.output or ("卸载完成" if r.ok else "卸载失败")))

    def adb_tools_copy_log(self):
        if hasattr(self, "adb_tools_log_box") and self.adb_tools_log_box.winfo_exists():
            text = self.adb_tools_log_box.get("1.0", "end").strip()
            self.clipboard_clear()
            self.clipboard_append(text)
            messagebox.showinfo("已复制", "ADB 操作日志已复制到剪贴板")

    def adb_tools_clear_log(self):
        if hasattr(self, "adb_tools_log_box") and self.adb_tools_log_box.winfo_exists():
            self.adb_tools_log_box.configure(state="normal")
            self.adb_tools_log_box.delete("1.0", "end")
            self.adb_tools_log_box.configure(state="disabled")

    def _open_local_dir(self, path):
        try:
            self.adb_tools.open_dir(path)
        except Exception as exc:
            messagebox.showerror("打开失败", str(exc))



    # ---------- iOS Log page ----------
    def _clear_ios_page_refs(self):
        """Release iOS page widget refs after page destroy to prevent stale updates."""
        for name in (
            "ios_custom_keywords_textbox",
            "ios_devices_box",
            "ios_tool_result_box",
            "ios_log_box",
            "ios_log_button",
        ):
            try:
                setattr(self, name, None)
            except Exception:
                pass

    def _build_ios_log_page(self):
        """Build iOS log page with fixed Tab layout to avoid CTk scroll ghosting.

        v2.4.17: The previous page-level scroll layout rendered many CheckBox widgets
        inside a CTkScrollableFrame. On Windows this could show duplicate keyword text
        while scrolling. This version creates keyword controls once inside a fixed tab,
        and only the device/tool/log Textbox widgets scroll locally.
        """
        self._clear_ios_page_refs()
        self.content.grid_rowconfigure(0, weight=1)
        self.content.grid_columnconfigure(0, weight=1)

        tabview = ctk.CTkTabview(
            self.content,
            fg_color=COLOR_BG,
            segmented_button_fg_color=COLOR_SURFACE_2,
            segmented_button_selected_color=COLOR_ACCENT,
            segmented_button_selected_hover_color=COLOR_ACCENT_HOVER,
            segmented_button_unselected_color=COLOR_SURFACE_2,
            segmented_button_unselected_hover_color=COLOR_HOVER,
            text_color=COLOR_TEXT,
            corner_radius=14,
        )
        tabview.grid(row=0, column=0, sticky="nsew")
        tab_setup = tabview.add("工具与设备")
        tab_log = tabview.add("日志过滤与抓取")
        for tab in (tab_setup, tab_log):
            tab.configure(fg_color=COLOR_BG)
            tab.grid_columnconfigure(0, weight=1)
            tab.grid_rowconfigure(0, weight=1)

        # ---- Tab 1: tools and devices ----
        setup_wrap = ctk.CTkFrame(tab_setup, fg_color=COLOR_BG)
        setup_wrap.grid(row=0, column=0, sticky="nsew", padx=2, pady=8)
        setup_wrap.grid_columnconfigure(0, weight=1)
        setup_wrap.grid_columnconfigure(1, weight=1)
        setup_wrap.grid_rowconfigure(0, weight=0)
        setup_wrap.grid_rowconfigure(1, weight=1)

        tool_card = ctk.CTkFrame(setup_wrap, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        tool_card.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 14))
        tool_card.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(tool_card, text="iOS 工具路径配置", text_color=COLOR_TEXT, font=self._font(18, "bold")).grid(row=0, column=0, sticky="w", padx=24, pady=(18, 8), columnspan=3)
        ctk.CTkLabel(
            tool_card,
            text="依赖 libimobiledevice。本工具不内置第三方 exe，可放到 tools/ios/ 或手动选择路径。",
            text_color=COLOR_MUTED,
            font=self._font(13),
            wraplength=760,
            anchor="w",
        ).grid(row=1, column=0, columnspan=3, sticky="ew", padx=24, pady=(0, 8))

        row = 2
        for tool_name, label in [
            ("idevice_id.exe", "idevice_id.exe"),
            ("ideviceinfo.exe", "ideviceinfo.exe"),
            ("idevicesyslog.exe", "idevicesyslog.exe"),
            ("idevicecrashreport.exe", "idevicecrashreport.exe（可选）"),
        ]:
            ctk.CTkLabel(tool_card, text=label, text_color=COLOR_TEXT, font=self._font(14, "bold")).grid(row=row, column=0, sticky="w", padx=(24, 16), pady=5)
            ctk.CTkEntry(tool_card, textvariable=self.ios_tool_path_vars[tool_name], height=34, corner_radius=10, fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER, text_color=COLOR_TEXT).grid(row=row, column=1, sticky="ew", pady=5)
            ctk.CTkButton(tool_card, text="选择", width=76, height=32, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=lambda n=tool_name: self.ios_choose_tool_path(n)).grid(row=row, column=2, padx=(10, 24), pady=5)
            row += 1

        tool_buttons = ctk.CTkFrame(tool_card, fg_color=COLOR_SURFACE)
        tool_buttons.grid(row=row, column=0, columnspan=3, sticky="ew", padx=24, pady=(8, 14))
        tool_buttons.grid_columnconfigure((0, 1, 2), weight=0)
        ctk.CTkButton(tool_buttons, text="检测工具", width=112, height=34, corner_radius=12, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color="#FFFFFF", font=self._font(13, "bold"), command=self.ios_detect_tools).grid(row=0, column=0, padx=(0, 10), pady=0)
        ctk.CTkButton(tool_buttons, text="保存路径配置", width=128, height=34, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self.ios_save_tool_paths).grid(row=0, column=1, padx=(0, 10), pady=0)
        ctk.CTkButton(tool_buttons, text="打开 iOS 工具目录", width=152, height=34, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self.ios_open_tools_dir).grid(row=0, column=2, pady=0)

        result_card = ctk.CTkFrame(setup_wrap, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        result_card.grid(row=1, column=0, sticky="nsew", padx=(0, 8), pady=(0, 0))
        result_card.grid_columnconfigure(0, weight=1)
        result_card.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(result_card, text="iOS 工具检测结果", text_color=COLOR_TEXT, font=self._font(17, "bold")).grid(row=0, column=0, sticky="w", padx=20, pady=(18, 8))
        self.ios_tool_result_box = ctk.CTkTextbox(result_card, height=210, corner_radius=12, border_width=1, border_color=COLOR_BORDER, fg_color=COLOR_SURFACE_2, text_color=COLOR_TEXT, font=self._font(13))
        self.ios_tool_result_box.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 18))
        self._set_textbox_text(self.ios_tool_result_box, self.ios_tool_status_var.get())

        device_card = ctk.CTkFrame(setup_wrap, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        device_card.grid(row=1, column=1, sticky="nsew", padx=(8, 0), pady=(0, 0))
        device_card.grid_columnconfigure(0, weight=1)
        device_card.grid_rowconfigure(2, weight=1)
        ctk.CTkLabel(device_card, text="设备检测", text_color=COLOR_TEXT, font=self._font(17, "bold")).grid(row=0, column=0, sticky="w", padx=20, pady=(18, 8))
        device_row = ctk.CTkFrame(device_card, fg_color=COLOR_SURFACE)
        device_row.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 10))
        device_row.grid_columnconfigure(2, weight=1)
        ctk.CTkButton(device_row, text="检测 iOS 设备", width=132, height=34, corner_radius=12, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color="#FFFFFF", font=self._font(13, "bold"), command=self.ios_detect_devices).grid(row=0, column=0, padx=(0, 12))
        ctk.CTkLabel(device_row, text="UDID", text_color=COLOR_TEXT, font=self._font(13, "bold")).grid(row=0, column=1, padx=(0, 8))
        ctk.CTkEntry(device_row, textvariable=self.ios_selected_udid_var, height=34, corner_radius=10, fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER, text_color=COLOR_TEXT, placeholder_text="留空则使用默认连接设备").grid(row=0, column=2, sticky="ew")
        self.ios_devices_box = ctk.CTkTextbox(device_card, height=210, corner_radius=12, border_width=1, border_color=COLOR_BORDER, fg_color=COLOR_SURFACE_2, text_color=COLOR_TEXT, font=self._font(13))
        self.ios_devices_box.grid(row=2, column=0, sticky="nsew", padx=20, pady=(0, 18))
        self._set_textbox_text(self.ios_devices_box, self.ios_devices_var.get())

        # ---- Tab 2: filter and log ----
        log_wrap = ctk.CTkFrame(tab_log, fg_color=COLOR_BG)
        log_wrap.grid(row=0, column=0, sticky="nsew", padx=2, pady=8)
        log_wrap.grid_columnconfigure(0, weight=1)
        log_wrap.grid_rowconfigure(1, weight=1)

        filter_card = ctk.CTkFrame(log_wrap, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        filter_card.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        filter_card.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(filter_card, text="日志过滤", text_color=COLOR_TEXT, font=self._font(18, "bold")).grid(row=0, column=0, sticky="w", padx=24, pady=(18, 8), columnspan=3)
        keyword_frame = ctk.CTkFrame(filter_card, fg_color=COLOR_SURFACE)
        keyword_frame.grid(row=1, column=0, columnspan=3, sticky="ew", padx=24, pady=(0, 8))
        for col in range(4):
            keyword_frame.grid_columnconfigure(col, minsize=150)
        for idx, keyword in enumerate(DEFAULT_IOS_KEYWORDS):
            ctk.CTkCheckBox(keyword_frame, text=keyword, variable=self.ios_keyword_vars[keyword], width=150, text_color=COLOR_TEXT, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, font=self._font(13)).grid(row=idx // 4, column=idx % 4, sticky="w", padx=(0, 10), pady=3)

        ctk.CTkLabel(filter_card, text="自定义关键字（每行一个）", text_color=COLOR_TEXT, font=self._font(14, "bold")).grid(row=2, column=0, sticky="nw", padx=(24, 18), pady=(6, 8))
        self.ios_custom_keywords_textbox = ctk.CTkTextbox(filter_card, height=60, corner_radius=12, border_width=1, border_color=COLOR_BORDER, fg_color=COLOR_SURFACE_2, text_color=COLOR_TEXT, font=self._font(13))
        self.ios_custom_keywords_textbox.grid(row=2, column=1, columnspan=2, sticky="ew", padx=(0, 24), pady=(6, 8))
        ctk.CTkLabel(filter_card, text="App / Bundle ID / 进程名过滤", text_color=COLOR_TEXT, font=self._font(14, "bold")).grid(row=3, column=0, sticky="w", padx=(24, 18), pady=(0, 14))
        ctk.CTkEntry(filter_card, textvariable=self.ios_filter_text_var, height=34, corner_radius=10, fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER, text_color=COLOR_TEXT, placeholder_text="例如 com.company.game / Unity / 进程名").grid(row=3, column=1, sticky="ew", pady=(0, 14))
        ctk.CTkCheckBox(filter_card, text="只显示命中日志", variable=self.ios_only_hits_var, text_color=COLOR_TEXT, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, font=self._font(13)).grid(row=3, column=2, sticky="w", padx=(12, 24), pady=(0, 14))

        log_card = ctk.CTkFrame(log_wrap, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        log_card.grid(row=1, column=0, sticky="nsew")
        log_card.grid_columnconfigure(0, weight=1)
        log_card.grid_rowconfigure(4, weight=1)
        ctk.CTkLabel(log_card, text="实时日志", text_color=COLOR_TEXT, font=self._font(18, "bold")).grid(row=0, column=0, sticky="w", padx=24, pady=(18, 8), columnspan=3)
        action_row = ctk.CTkFrame(log_card, fg_color=COLOR_SURFACE)
        action_row.grid(row=1, column=0, columnspan=3, sticky="ew", padx=24, pady=(0, 8))
        for col in range(4):
            action_row.grid_columnconfigure(col, weight=0)
        self.ios_log_button = ctk.CTkButton(action_row, text="▶  开始抓取 iOS 日志", width=180, height=36, corner_radius=12, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color="#FFFFFF", font=self._font(13, "bold"), command=self.ios_toggle_log)
        self.ios_log_button.grid(row=0, column=0, padx=(0, 10), pady=0)
        ctk.CTkButton(action_row, text="复制当前日志", width=120, height=36, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self.ios_copy_current_log).grid(row=0, column=1, padx=(0, 10), pady=0)
        ctk.CTkButton(action_row, text="清空显示", width=104, height=36, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self.ios_clear_display).grid(row=0, column=2, padx=(0, 10), pady=0)
        ctk.CTkButton(action_row, text="导入到日志分析", width=140, height=36, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self.ios_import_saved_log_to_analysis).grid(row=0, column=3, pady=0)
        ctk.CTkLabel(log_card, textvariable=self.ios_status_var, text_color=COLOR_MUTED, font=self._font(13)).grid(row=2, column=0, sticky="w", padx=24, pady=(0, 4))
        ctk.CTkLabel(log_card, textvariable=self.ios_log_file_var, text_color=COLOR_MUTED, wraplength=760, font=self._font(13)).grid(row=3, column=0, sticky="w", padx=24, pady=(0, 8))
        self.ios_log_box = ctk.CTkTextbox(log_card, height=260, corner_radius=12, border_width=1, border_color=COLOR_BORDER, fg_color=COLOR_SURFACE_2, text_color=COLOR_TEXT, font=self._font(13))
        self.ios_log_box.grid(row=4, column=0, columnspan=3, sticky="nsew", padx=24, pady=(0, 18))
        self.ios_log_box.insert("end", "提示：Windows 抓取 iOS 日志需要 libimobiledevice 工具和 Apple Mobile Device Support / iTunes 驱动。\n")
        self.ios_log_box.configure(state="disabled")
        self._refresh_ios_log_button()

    def _sync_ios_tool_paths_from_ui(self):
        for name, var in self.ios_tool_path_vars.items():
            self.ios_tools.set_tool_path(name, var.get())

    def ios_choose_tool_path(self, tool_name: str):
        path = filedialog.askopenfilename(title=f"选择 {tool_name}", filetypes=[("EXE 文件", "*.exe"), ("所有文件", "*.*")])
        if path:
            self.ios_tool_path_vars[tool_name].set(path)
            self.ios_tools.set_tool_path(tool_name, path)

    def ios_save_tool_paths(self):
        try:
            self._sync_ios_tool_paths_from_ui()
            self.ios_tools.save_config()
            self.ios_tool_status_var.set(f"路径配置已保存：{self.ios_tools.config_path}")
            messagebox.showinfo("保存成功", f"iOS 工具路径配置已保存：\n{self.ios_tools.config_path}")
        except Exception as exc:
            messagebox.showerror("保存失败", str(exc))

    def ios_detect_tools(self):
        self._sync_ios_tool_paths_from_ui()
        data = self.ios_tools.detect_tools()
        lines = ["iOS 工具检测结果："]
        for name, item in data.items():
            status = "已找到" if item.get("exists") else "未找到"
            source = item.get("source") or "未找到"
            path = item.get("path") or ""
            note = item.get("note") or ""
            lines.append(f"- {name}：{status}｜来源：{source}｜路径：{path}")
            if note:
                lines.append(f"  提示：{note}")
        missing = self.ios_tools.missing_required_tools()
        if missing:
            lines.append("\n缺少必要工具：" + ", ".join(missing))
            lines.append("请将 libimobiledevice 工具放入当前程序目录的 tools/ios，或手动选择 exe 路径。")
            lines.append("\n已检查位置：")
            lines.append(self.ios_tools.checked_locations_summary())
        else:
            lines.append("\n必要工具已就绪。")
        text = "\n".join(lines)
        self.ios_tool_status_var.set(text)
        if hasattr(self, "ios_tool_result_box") and self.ios_tool_result_box and self.ios_tool_result_box.winfo_exists():
            self._set_textbox_text(self.ios_tool_result_box, text)
    def _run_ios_task(self, title, worker, on_success=None):
        if self.ios_busy.is_set():
            self.ios_append_system_log("已有 iOS 操作正在执行，请稍后再试")
            return
        self.ios_busy.set()
        self.ios_status_var.set(f"正在执行：{title}")
        self.ios_append_system_log(f"开始：{title}")

        def run():
            try:
                result = worker()
                def ok():
                    if on_success:
                        on_success(result)
                    self.ios_append_system_log(f"完成：{title}")
                    if not self.ios_log_running.is_set():
                        self.ios_status_var.set("待命")
            except Exception as exc:
                def ok(exc=exc):
                    self.ios_append_system_log(f"失败：{title}｜{exc}")
                    self.ios_status_var.set("操作失败")
                    messagebox.showerror("操作失败", str(exc))
            finally:
                def done():
                    self.ios_busy.clear()
                    self._refresh_ios_log_button()
                try:
                    self.after(0, ok)
                    self.after(0, done)
                except Exception:
                    self.ios_busy.clear()
        threading.Thread(target=run, daemon=True).start()

    def ios_detect_devices(self):
        def worker():
            self._sync_ios_tool_paths_from_ui()
            return self.ios_tools.list_devices()
        def success(data):
            devices, result = data
            text = format_devices(devices, result)
            self.ios_devices_var.set(text)
            if devices and not self.ios_selected_udid_var.get().strip():
                self.ios_selected_udid_var.set(devices[0].udid)
            if hasattr(self, "ios_devices_box") and self.ios_devices_box.winfo_exists():
                self._set_textbox_text(self.ios_devices_box, text)
        self._run_ios_task("检测 iOS 设备", worker, success)

    def _get_ios_keywords(self):
        keywords = [kw for kw, var in self.ios_keyword_vars.items() if var.get()]
        if self.ios_custom_keywords_textbox and self.ios_custom_keywords_textbox.winfo_exists():
            keywords.extend(parse_keywords(self.ios_custom_keywords_textbox.get("1.0", "end")))
        return parse_keywords("\n".join(keywords))

    def ios_toggle_log(self):
        if self.ios_log_running.is_set():
            self.ios_stop_log()
        else:
            self.ios_start_log()

    def ios_start_log(self):
        if self.ios_log_running.is_set():
            return
        self._sync_ios_tool_paths_from_ui()
        missing = self.ios_tools.missing_required_tools()
        if missing:
            messagebox.showwarning("缺少 iOS 工具", "缺少必要工具：" + ", ".join(missing) + "\n请先配置 libimobiledevice 工具路径。")
            self.ios_detect_tools()
            return
        udid = self.ios_selected_udid_var.get().strip()
        keywords = self._get_ios_keywords()
        log_filter = IOSLogFilter(keywords=keywords, text_filter=self.ios_filter_text_var.get(), only_hits=self.ios_only_hits_var.get())
        try:
            full_path, filtered_path = self.ios_tools.make_log_paths(udid)
            self.ios_log_full_path = str(full_path)
            self.ios_log_filtered_path = str(filtered_path)
            self.ios_log_process = self.ios_tools.start_syslog_process(udid)
        except Exception as exc:
            messagebox.showerror("启动失败", str(exc))
            return

        self.ios_display_lines.clear()
        self.ios_log_running.set()
        self.ios_status_var.set("正在抓取 iOS 日志")
        self.ios_log_file_var.set(f"完整日志：{self.ios_log_full_path}\n过滤日志：{self.ios_log_filtered_path}")
        self.ios_append_system_log("iOS 日志抓取已启动")
        self.ios_log_thread = threading.Thread(target=self._ios_log_worker, args=(log_filter,), daemon=True)
        self.ios_log_thread.start()
        self._refresh_ios_log_button()
        self._set_badge(True)

    def ios_stop_log(self):
        if not self.ios_log_running.is_set() and self.ios_log_process is None:
            return
        self.ios_log_running.clear()
        proc = self.ios_log_process
        self.ios_log_process = None
        if proc:
            try:
                proc.terminate()
                proc.wait(timeout=2)
            except Exception:
                try:
                    proc.kill()
                except Exception:
                    pass
        self.ios_status_var.set("已停止 iOS 日志抓取")
        self.ios_append_system_log(f"iOS 日志抓取已停止。完整日志：{self.ios_log_full_path}")
        self._refresh_ios_log_button()
        if self.current_page == "ios_log":
            self._set_badge(False)

    def _ios_log_worker(self, log_filter: IOSLogFilter):
        try:
            with open(self.ios_log_full_path, "a", encoding="utf-8", errors="ignore") as full_file, open(self.ios_log_filtered_path, "a", encoding="utf-8", errors="ignore") as filtered_file:
                proc = self.ios_log_process
                if not proc or not proc.stdout:
                    raise RuntimeError("idevicesyslog 进程未正常启动")
                for line in proc.stdout:
                    if not self.ios_log_running.is_set():
                        break
                    text = line.rstrip("\n")
                    full_file.write(text + "\n")
                    full_file.flush()
                    self._on_ios_live_log_line(text)
                    display = log_filter.should_display(text)
                    if display:
                        filtered_file.write(text + "\n")
                        filtered_file.flush()
                        is_hit = log_filter.is_hit(text)
                        self._append_ios_log_line(text, hit=is_hit)
        except Exception as exc:
            self.after(0, lambda exc=exc: self.ios_append_system_log(f"iOS 日志抓取异常：{exc}"))
        finally:
            self.ios_log_running.clear()
            try:
                self.after(0, self._refresh_ios_log_button)
            except Exception:
                pass

    def _append_ios_log_line(self, line: str, hit: bool = False):
        prefix = "[命中] " if hit else ""
        text = prefix + line
        self.ios_display_lines.append(text)
        if len(self.ios_display_lines) > self.ios_max_display_lines:
            self.ios_display_lines = self.ios_display_lines[-self.ios_max_display_lines:]
        def write():
            if hasattr(self, "ios_log_box") and self.ios_log_box and self.ios_log_box.winfo_exists():
                self.ios_log_box.configure(state="normal")
                self.ios_log_box.insert("end", text + "\n")
                # Trim UI occasionally to avoid very large widget content.
                if len(self.ios_display_lines) == self.ios_max_display_lines:
                    self.ios_log_box.delete("1.0", "200.0")
                self.ios_log_box.see("end")
                self.ios_log_box.configure(state="disabled")
        try:
            self.after(0, write)
        except Exception:
            pass

    def ios_append_system_log(self, text: str):
        self._append_ios_log_line(f"[{time.strftime('%H:%M:%S')}] {text}", hit=False)

    def _on_ios_live_log_line(self, line: str):
        if not self.log_analysis_enabled_var.get():
            return
        try:
            new_hits = self.log_monitor.feed_line(line)
            fallback = self.ios_filter_text_var.get().strip()
            new_crashes = self.crash_stream_detector.feed_line(line, fallback_package=fallback)
            if new_hits or new_crashes:
                self.log_hit_records = self.log_monitor.get_hits()
                self.crash_records = self.crash_stream_detector.get_records()
                self._schedule_log_analysis_refresh()
        except Exception:
            pass

    def _refresh_ios_log_button(self):
        if hasattr(self, "ios_log_button") and self.ios_log_button and self.ios_log_button.winfo_exists():
            running = self.ios_log_running.is_set()
            self.ios_log_button.configure(
                text="■  停止 iOS 日志" if running else "▶  开始抓取 iOS 日志",
                fg_color=COLOR_DANGER if running else COLOR_ACCENT,
                hover_color=COLOR_DANGER_HOVER if running else COLOR_ACCENT_HOVER,
            )

    def ios_copy_current_log(self):
        text = "\n".join(self.ios_display_lines)
        if not text:
            messagebox.showinfo("提示", "当前没有可复制的 iOS 日志")
            return
        self.clipboard_clear()
        self.clipboard_append(text)
        messagebox.showinfo("已复制", f"已复制 {len(self.ios_display_lines)} 行 iOS 日志")

    def ios_clear_display(self):
        self.ios_display_lines.clear()
        if hasattr(self, "ios_log_box") and self.ios_log_box and self.ios_log_box.winfo_exists():
            self._set_textbox_text(self.ios_log_box, "")

    def ios_open_log_dir(self):
        try:
            self.ios_tools.open_log_dir()
        except Exception as exc:
            messagebox.showerror("打开失败", str(exc))

    def ios_open_tools_dir(self):
        try:
            self.ios_tools.open_tools_dir()
        except Exception as exc:
            messagebox.showerror("打开失败", str(exc))

    def ios_import_saved_log_to_analysis(self):
        path = self.ios_log_full_path if self.ios_log_full_path and os.path.exists(self.ios_log_full_path) else ""
        if not path:
            path = filedialog.askopenfilename(title="选择 iOS 日志文件", filetypes=[("日志文件", "*.txt *.log"), ("所有文件", "*.*")])
        if not path:
            return
        if self.log_analysis_busy.is_set():
            messagebox.showinfo("提示", "日志分析正在进行中，请稍后")
            return
        self.apply_log_analysis_config()
        self.log_analysis_busy.set()
        self.log_analysis_status_var.set("正在导入 iOS 日志并分析...")

        def worker():
            try:
                with open(path, "r", encoding="utf-8", errors="ignore") as file:
                    lines = file.readlines()
                hits = self.log_monitor.analyze_lines(lines)
                crashes = self.crash_file_analyzer.analyze_lines(lines, fallback_package=self.ios_filter_text_var.get().strip())
                def ok():
                    self.log_hit_records = hits
                    self.crash_records = crashes
                    self.crash_stream_detector.clear()
                    for record in crashes:
                        self.crash_stream_detector.records.append(record)
                    self._refresh_log_analysis_views()
                    self.log_analysis_status_var.set(f"iOS 日志导入分析完成：关键字命中 {len(hits)} 条，崩溃结果 {len(crashes)} 条")
                    messagebox.showinfo("导入完成", "iOS 日志已导入日志分析页。")
            except Exception as exc:
                def ok(exc=exc):
                    messagebox.showerror("分析失败", str(exc))
                    self.log_analysis_status_var.set("iOS 日志分析失败")
            finally:
                def done():
                    self.log_analysis_busy.clear()
                try:
                    self.after(0, ok)
                    self.after(0, done)
                except Exception:
                    self.log_analysis_busy.clear()
        threading.Thread(target=worker, daemon=True).start()

    # ---------- File Compare page ----------
    def _build_compare_page(self):
        """Build file compare page with fixed layout.

        v2.4.20 compare_ui_tab_layout_fix:
        The compare page previously created a local CTkScrollableFrame and placed the
        whole file-compare UI into one long Canvas-backed area. On Windows, Canvas
        scrolling with many embedded native widgets can leave redraw trails. The page
        now uses a fixed top area, a compact tabbed value-config settings card and a
        fixed result area; long content uses Treeview/Textbox native scrollbars.
        """
        self.compare_local_scroll = None
        self.value_tabview = None
        self.value_preview_textbox = None
        self.value_result_textbox = None

        wrap = ctk.CTkFrame(self.content, fg_color=COLOR_BG, corner_radius=0)
        wrap.grid(row=0, column=0, sticky="nsew")
        wrap.grid_columnconfigure(0, weight=1)
        wrap.grid_rowconfigure(0, weight=1)

        self.compare_file_tab_name = "文件与模式"
        self.compare_special_tab_name = "专项设置"
        self.compare_result_tab_name = "比对结果"
        page_tabs = ctk.CTkTabview(
            wrap,
            fg_color=COLOR_BG,
            segmented_button_fg_color=COLOR_SURFACE_2,
            segmented_button_selected_color=COLOR_ACCENT,
            segmented_button_selected_hover_color=COLOR_ACCENT_HOVER,
            segmented_button_unselected_color=COLOR_SURFACE_2,
            segmented_button_unselected_hover_color=COLOR_HOVER,
            text_color=COLOR_TEXT,
            corner_radius=12,
        )
        self.compare_page_tabs = page_tabs
        page_tabs.grid(row=0, column=0, sticky="nsew")
        file_tab = page_tabs.add(self.compare_file_tab_name)
        special_tab = page_tabs.add(self.compare_special_tab_name)
        result_tab = page_tabs.add(self.compare_result_tab_name)
        for tab in (file_tab, special_tab, result_tab):
            tab.configure(fg_color=COLOR_BG)
            tab.grid_columnconfigure(0, weight=1)
            tab.grid_rowconfigure(0, weight=1)

        file_page = ctk.CTkFrame(file_tab, fg_color=COLOR_BG, corner_radius=0)
        file_page.grid(row=0, column=0, sticky="nsew")
        file_page.grid_columnconfigure(0, weight=1)
        special_page = ctk.CTkFrame(special_tab, fg_color=COLOR_BG, corner_radius=0)
        special_page.grid(row=0, column=0, sticky="nsew")
        special_page.grid_columnconfigure(0, weight=1)
        special_page.grid_rowconfigure(0, weight=1)
        self.compare_page_wrap = special_page

        file_wrap = ctk.CTkFrame(file_page, fg_color=COLOR_BG, corner_radius=0)
        self.compare_file_wrap = file_wrap
        file_wrap.pack(fill="x", pady=(0, 12))
        file_wrap.grid_columnconfigure((0, 1), weight=1, uniform="compare_files")

        self._compare_file_panel(
            parent=file_wrap,
            column=0,
            title="待比对文件（实际配置文件 / 导出文件）",
            path_var=self.compare_actual_path_var,
            status_var=self.compare_actual_status_var,
            button_text="选择待比对文件",
            command=lambda: self.select_compare_file("actual"),
        )
        self._compare_file_panel(
            parent=file_wrap,
            column=1,
            title="参考文档（策划文档 / 配置说明）",
            path_var=self.compare_document_path_var,
            status_var=self.compare_document_status_var,
            button_text="选择参考文档",
            command=lambda: self.select_compare_file("document"),
        )

        option_card = self._card(file_page)
        self.compare_option_card = option_card
        self._option_row(option_card, 0, "比对模式", self.compare_mode_var, COMPARE_MODES, width=360, command=lambda _=None: self._on_compare_mode_changed())
        self._divider(option_card, 1)

        option_line = ctk.CTkFrame(option_card, fg_color=COLOR_SURFACE, corner_radius=0)
        option_line.grid(row=2, column=0, columnspan=3, sticky="ew", padx=24, pady=10)
        ctk.CTkCheckBox(option_line, text="忽略空格", variable=self.compare_ignore_spaces_var, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(14)).pack(side="left", padx=(0, 24))
        ctk.CTkCheckBox(option_line, text="忽略换行", variable=self.compare_ignore_newlines_var, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(14)).pack(side="left", padx=(0, 24))
        ctk.CTkCheckBox(option_line, text="忽略大小写", variable=self.compare_ignore_case_var, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(14)).pack(side="left")

        self._divider(option_card, 3)
        self._entry_row(option_card, 4, "忽略字段", self.compare_ignore_fields_var, "例如 时间戳, version, id", width=640)

        # Heavy mode-specific editors are created only when that mode is used.
        self.compare_special_hint = ctk.CTkLabel(
            special_page,
            text="当前模式不需要专项设置。\n普通文件比对请在“文件与模式”配置后，到“比对结果”开始执行。",
            text_color=COLOR_MUTED,
            font=self._font(16),
            justify="center",
        )
        self.compare_special_hint.pack(fill="both", expand=True)
        if self._is_value_config_mode():
            self._build_value_compare_settings(special_page)
        elif self._is_translation_mode():
            self._build_translation_compare_settings(special_page)

        result_card = ctk.CTkFrame(result_tab, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        self.compare_result_card = result_card
        result_card.pack(fill="both", expand=True)
        result_card.grid_columnconfigure(0, weight=1)
        result_card.grid_rowconfigure(3, weight=1)

        action_row = ctk.CTkFrame(result_card, fg_color=COLOR_SURFACE, corner_radius=0)
        action_row.grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 8))
        action_row.grid_columnconfigure(1, weight=1)

        self.compare_button = ctk.CTkButton(action_row, text="开始比对", width=140, height=40, corner_radius=12, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color="#FFFFFF", font=self._font(15, "bold"), command=self.start_compare)
        self.compare_button.grid(row=0, column=0, padx=(0, 12), sticky="w")

        search_entry = ctk.CTkEntry(action_row, textvariable=self.compare_search_var, placeholder_text="搜索差异结果", height=40, corner_radius=10, fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(14))
        search_entry.grid(row=0, column=1, sticky="ew", padx=(0, 12))
        ctk.CTkButton(action_row, text="复制结果", width=104, height=40, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(14, "bold"), command=self.copy_compare_visible_diffs).grid(row=0, column=2, padx=(0, 10))
        ctk.CTkButton(action_row, text="导出 Excel", width=112, height=40, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(14, "bold"), command=self.export_compare_excel).grid(row=0, column=3)

        self.compare_summary_card = ctk.CTkFrame(result_card, fg_color=COLOR_SURFACE_2, corner_radius=12, border_width=1, border_color=COLOR_BORDER)
        self.compare_summary_card.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 8))
        self.compare_summary_card.grid_columnconfigure(0, weight=1)
        self.compare_summary_label = ctk.CTkLabel(
            self.compare_summary_card,
            textvariable=self.compare_summary_var,
            text_color=COLOR_TEXT,
            font=self._font(14, "bold"),
            anchor="w",
            justify="left",
            wraplength=860,
        )
        self.compare_summary_label.grid(row=0, column=0, sticky="ew", padx=14, pady=10)
        self.compare_summary_card.bind("<Configure>", lambda event: self._update_compare_summary_wrap(event.width))
        ctk.CTkLabel(result_card, text="提示：商业化/礼包/活动奖励建议用“数值参考表 vs 游戏配置表”；普通临时差异用“普通文本 / 文件比对”；多语言表用翻译专项模式。", text_color=COLOR_MUTED, font=self._font(13), anchor="w", justify="left", wraplength=920).grid(row=2, column=0, sticky="ew", padx=18, pady=(0, 8))

        table_frame = ctk.CTkFrame(result_card, fg_color=COLOR_SURFACE_2, corner_radius=14, border_width=1, border_color=COLOR_BORDER)
        table_frame.grid(row=3, column=0, sticky="nsew", padx=18, pady=(0, 18))
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)

        columns = ("index", "diff_type", "document_content", "file_content", "location", "remark")
        self.compare_tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=13, style="Xiao.Treeview")
        self.compare_tree.heading("index", text="序号")
        self.compare_tree.heading("diff_type", text="差异类型")
        self.compare_tree.heading("document_content", text="文档内容")
        self.compare_tree.heading("file_content", text="文件内容")
        self.compare_tree.heading("location", text="所在位置")
        self.compare_tree.heading("remark", text="备注")
        self.compare_tree.column("index", width=58, anchor="center", stretch=False)
        self.compare_tree.column("diff_type", width=96, anchor="center", stretch=False)
        self.compare_tree.column("document_content", width=300, anchor="w")
        self.compare_tree.column("file_content", width=300, anchor="w")
        self.compare_tree.column("location", width=170, anchor="w")
        self.compare_tree.column("remark", width=230, anchor="w")

        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.compare_tree.yview)
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.compare_tree.xview)
        self.compare_tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.compare_tree.grid(row=0, column=0, sticky="nsew")
        self.compare_tree.bind("<Double-1>", self.show_compare_diff_detail)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        pager = self._build_pager(
            result_card,
            self.compare_page_status_var,
            self.compare_page_jump_var,
            self.compare_prev_page,
            self.compare_next_page,
            self.compare_jump_page,
        )
        pager.grid(row=4, column=0, sticky="ew", padx=18, pady=(0, 14))

        if not self.compare_search_trace_bound:
            self.compare_search_var.trace_add("write", lambda *_: self.apply_compare_search_filter())
            self.compare_search_trace_bound = True
        self.refresh_compare_table(self.compare_visible_diffs)
        self._refresh_compare_button()
        self._update_translation_settings_visibility()
        self._refresh_compare_scroll_region()


    def _update_compare_summary_wrap(self, width=None):
        try:
            if self.compare_summary_label is not None and self.compare_summary_label.winfo_exists():
                card_width = int(width or self.compare_summary_card.winfo_width() or 900)
                self.compare_summary_label.configure(wraplength=max(420, card_width - 34))
        except Exception:
            pass

    def _update_localization_result_hint_wrap(self, width=None):
        try:
            label = getattr(self, "loc_result_hint_label", None)
            if label is not None and label.winfo_exists():
                card_width = int(width or getattr(self, "loc_result_card", label.master).winfo_width() or 760)
                label.configure(wraplength=max(320, card_width - 36))
        except Exception:
            pass

    def _refresh_compare_scroll_region(self):
        """Compatibility hook after removing compare-page full Canvas scrolling."""
        def _do_refresh():
            try:
                self.update_idletasks()
            except Exception:
                pass
            try:
                if getattr(self, "value_tabview", None) is not None and self.value_tabview.winfo_exists():
                    self.value_tabview.update_idletasks()
            except Exception:
                pass
        try:
            self.after_idle(_do_refresh)
        except Exception:
            _do_refresh()

    def _build_value_compare_result_summary_text(self, diffs, report, export_path: str = "") -> str:
        counts = self._count_compare_by_type(diffs)
        confirmed_field = counts.get("字段差异", 0) + counts.get("值不一致", 0)
        confirmed_missing = counts.get("漏配", 0)
        confirmed_extra = counts.get("多配", 0)
        confirmed_count = confirmed_field + confirmed_missing + confirmed_extra
        lines = ["结构化比对完成"]
        try:
            lines.append(f"参考表记录：{getattr(report, 'reference_records', 0)}")
            lines.append(f"配置表记录：{getattr(report, 'config_records', 0)}")
            lines.append(f"高可信匹配：{getattr(report, 'high_confidence_matches', 0)}")
            lines.append(f"中可信匹配：{getattr(report, 'medium_confidence_matches', 0)}")
            lines.append(f"低可信匹配：{getattr(report, 'low_confidence_matches', 0)}")
        except Exception:
            pass
        lines.extend([
            f"确认字段差异：{confirmed_field}",
            f"确认漏配：{confirmed_missing}",
            f"确认多配：{confirmed_extra}",
            f"待确认匹配项：{counts.get('待确认匹配项', 0) + counts.get('待确认', 0)}",
            f"过滤风险：{counts.get('过滤风险', 0)}",
        ])
        if confirmed_count == 0:
            lines.append("确认差异：0")
        else:
            lines.append(f"确认差异：{confirmed_count}")
        if counts.get('过滤风险', 0):
            lines.append("提示：当前差异可能包含非目标范围配置，请确认配置表过滤条件。")
        if export_path:
            lines.append(f"结果文件：{export_path}")
        else:
            lines.append("结果文件：尚未导出，请点击“导出 Excel”")
        return "\n".join(lines)


    def _value_tab_label_entry(self, parent, row, label_text, variable, placeholder="", col=0):
        """Small helper for value-compare tab forms."""
        base_col = col * 2
        ctk.CTkLabel(
            parent,
            text=label_text,
            text_color=COLOR_MUTED,
            font=self._font(12),
            anchor="w",
        ).grid(row=row, column=base_col, sticky="w", padx=(0, 8), pady=6)
        ctk.CTkEntry(
            parent,
            textvariable=variable,
            placeholder_text=placeholder,
            height=32,
            corner_radius=9,
            fg_color=COLOR_SURFACE_2,
            border_color=COLOR_BORDER,
            text_color=COLOR_TEXT,
            font=self._font(12),
        ).grid(row=row, column=base_col + 1, sticky="ew", padx=(0, 14) if col == 0 else (0, 0), pady=6)

    def _set_textbox_text(self, textbox, text: str, disabled: bool = True):
        if textbox is None:
            return
        try:
            if disabled:
                textbox.configure(state="normal")
            textbox.delete("1.0", "end")
            textbox.insert("1.0", text or "")
            if disabled:
                textbox.configure(state="disabled")
        except Exception:
            pass

    def _copy_ref_filter_to_config(self):
        self.value_cfg_filter_var.set(self.value_ref_filter_var.get().strip())
        self.value_compare_summary_var.set("已将参考表过滤条件同步到配置表过滤条件。")
        self._refresh_value_preview_tab(silent=True)

    def _refresh_value_preview_tab(self, silent: bool = False):
        actual_path = self.compare_actual_path_var.get().strip()
        document_path = self.compare_document_path_var.get().strip()
        if not actual_path or not document_path:
            text = "请先选择参考表和配置表，再刷新执行预览。"
            self._set_textbox_text(getattr(self, "value_preview_textbox", None), text)
            if not silent:
                messagebox.showinfo("执行预览", text)
            return
        try:
            options = prepare_value_compare_options(self.get_value_compare_options(), actual_path, document_path)
            preview = build_value_compare_rule_preview(actual_path, document_path, options)
            self._set_textbox_text(getattr(self, "value_preview_textbox", None), preview)
        except Exception as exc:
            self._set_textbox_text(getattr(self, "value_preview_textbox", None), f"预览失败：{exc}")
            if not silent:
                messagebox.showwarning("执行预览失败", str(exc))

    def _set_value_result_summary_text(self, text: str):
        self._set_textbox_text(getattr(self, "value_result_textbox", None), text or "暂无比对结果")

    def _build_value_compare_compact_file_row(self, parent, row=1):
        """Compact file selector shown only inside the structured compare workspace.

        The generic file cards are intentionally hidden in value-config mode to avoid
        forcing the Notebook below the fold on normal 1366x768/1440x900 windows.
        """
        compact = ctk.CTkFrame(parent, fg_color=COLOR_SURFACE_2, corner_radius=12, border_width=1, border_color=COLOR_BORDER)
        self.value_compact_file_row = compact
        compact.grid(row=row, column=0, sticky="ew", padx=18, pady=(0, 8))
        compact.grid_columnconfigure(1, weight=1)
        compact.grid_columnconfigure(4, weight=1)

        ctk.CTkLabel(compact, text="配置表", text_color=COLOR_TEXT, font=self._font(13, "bold"), anchor="w").grid(row=0, column=0, sticky="w", padx=(12, 8), pady=(10, 4))
        ctk.CTkEntry(compact, textvariable=self.compare_actual_path_var, state="readonly", height=32, corner_radius=9, fg_color=COLOR_SURFACE, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(12)).grid(row=0, column=1, sticky="ew", pady=(10, 4))
        ctk.CTkButton(compact, text="选择配置表", width=96, height=32, corner_radius=9, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color="#FFFFFF", font=self._font(12, "bold"), command=lambda: self.select_compare_file("actual")).grid(row=0, column=2, sticky="e", padx=(8, 18), pady=(10, 4))

        ctk.CTkLabel(compact, text="参考表", text_color=COLOR_TEXT, font=self._font(13, "bold"), anchor="w").grid(row=0, column=3, sticky="w", padx=(0, 8), pady=(10, 4))
        ctk.CTkEntry(compact, textvariable=self.compare_document_path_var, state="readonly", height=32, corner_radius=9, fg_color=COLOR_SURFACE, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(12)).grid(row=0, column=4, sticky="ew", pady=(10, 4))
        ctk.CTkButton(compact, text="选择参考表", width=96, height=32, corner_radius=9, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color="#FFFFFF", font=self._font(12, "bold"), command=lambda: self.select_compare_file("document")).grid(row=0, column=5, sticky="e", padx=(8, 12), pady=(10, 4))

        ctk.CTkLabel(compact, textvariable=self.compare_actual_status_var, text_color=COLOR_MUTED, font=self._font(11), anchor="w").grid(row=1, column=1, sticky="ew", pady=(0, 8))
        ctk.CTkLabel(compact, textvariable=self.compare_document_status_var, text_color=COLOR_MUTED, font=self._font(11), anchor="w").grid(row=1, column=4, sticky="ew", pady=(0, 8))

    def _build_value_compare_settings(self, parent):
        """Build value-config settings with tabbed layout instead of long Canvas scrolling."""
        card = ctk.CTkFrame(parent, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        self.value_compare_settings_card = card
        card.pack(fill="x", pady=(0, 12))
        card.grid_columnconfigure(0, weight=1)
        # v2.4.20 compare_structured_workspace_fix:
        # Structured value compare now owns the main working area in value mode.
        # The large generic file cards are hidden and this compact file row + Notebook
        # get the usable vertical space, so users do not need to maximize the window.
        card.grid_rowconfigure(0, weight=0)
        card.grid_rowconfigure(1, weight=0)
        card.grid_rowconfigure(2, weight=0)
        card.grid_rowconfigure(3, weight=1, minsize=300)
        card.grid_rowconfigure(4, weight=0, minsize=52)

        header = ctk.CTkFrame(card, fg_color=COLOR_SURFACE, corner_radius=0)
        header.grid(row=0, column=0, sticky="ew", padx=18, pady=(14, 6))
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text="数值参考表 vs 游戏配置表（结构化比对）", text_color=COLOR_TEXT, font=self._font(16, "bold"), anchor="w").grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(header, text="结构化比对专用工作区：紧凑文件区 + 分区 Tab + 底部固定操作栏。", text_color=COLOR_MUTED, font=self._font(12), anchor="e", wraplength=560).grid(row=0, column=1, sticky="e", padx=(12, 0))

        self._build_value_compare_compact_file_row(card, row=1)

        info_box = ctk.CTkFrame(card, fg_color=COLOR_SURFACE_2, corner_radius=12, border_width=1, border_color=COLOR_BORDER)
        info_box.grid(row=2, column=0, sticky="ew", padx=18, pady=(0, 8))
        info_box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(info_box, textvariable=self.value_compare_summary_var, text_color=COLOR_MUTED, font=self._font(13), anchor="w", wraplength=920, justify="left").grid(row=0, column=0, sticky="ew", padx=12, pady=8)

        tabview = ctk.CTkTabview(
            card,
            height=330,
            fg_color=COLOR_SURFACE,
            segmented_button_fg_color=COLOR_SURFACE_2,
            segmented_button_selected_color=COLOR_ACCENT,
            segmented_button_selected_hover_color=COLOR_ACCENT_HOVER,
            segmented_button_unselected_color=COLOR_SURFACE_2,
            segmented_button_unselected_hover_color=COLOR_HOVER,
            text_color=COLOR_TEXT,
            corner_radius=12,
            border_width=1,
            border_color=COLOR_BORDER,
        )
        self.value_tabview = tabview
        tabview.grid(row=3, column=0, sticky="nsew", padx=18, pady=(0, 8))
        for name in ["基础设置", "主键与过滤", "字段映射", "标准化规则", "执行预览", "结果摘要"]:
            tabview.add(name)
            tab = tabview.tab(name)
            tab.configure(fg_color=COLOR_SURFACE)
            tab.grid_columnconfigure(0, weight=1)
            for c in (1, 3):
                tab.grid_columnconfigure(c, weight=1)
            # Allow taller widgets inside preview/result tabs to use available space without
            # forcing the outer action bar out of view.
            tab.grid_rowconfigure(10, weight=1)

        # Tab 1: Basic settings
        basic = tabview.tab("基础设置")
        self._value_tab_label_entry(basic, 0, "规则模板", self.value_template_var, "通用配置表", 0)
        template_menu = ctk.CTkOptionMenu(
            basic,
            variable=self.value_template_var,
            values=["通用配置表", "商店商品配置", "活动配置", "奖励配置", "商业化礼包配置"],
            height=32,
            corner_radius=9,
            fg_color=COLOR_SURFACE_2,
            button_color=COLOR_ACCENT,
            button_hover_color=COLOR_ACCENT_HOVER,
            text_color=COLOR_TEXT,
            font=self._font(12),
            command=lambda *_: self._on_value_template_changed(),
        )
        # Replace the entry created by helper with a real OptionMenu in the same cell.
        try:
            for child in basic.grid_slaves(row=0, column=1):
                child.destroy()
        except Exception:
            pass
        template_menu.grid(row=0, column=1, sticky="ew", padx=(0, 14), pady=6)
        ctk.CTkLabel(basic, text="商店商品配置模板会推荐双侧 type=3；活动模板仅提示填写 activity_id / event_id / group。", text_color=COLOR_MUTED, font=self._font(12), anchor="w", justify="left", wraplength=420).grid(row=0, column=2, columnspan=2, sticky="ew", pady=6)

        self._value_tab_label_entry(basic, 1, "参考表 Sheet", self.value_ref_sheet_var, "默认 配置参考表；留空自动选择", 0)
        self._value_tab_label_entry(basic, 1, "配置表 Sheet", self.value_cfg_sheet_var, "CSV 可空；Excel 留空自动选择", 1)
        self._value_tab_label_entry(basic, 2, "参考字段名行", self.value_ref_header_row_var, "auto 或 3", 0)
        self._value_tab_label_entry(basic, 2, "配置字段名行", self.value_cfg_header_row_var, "auto 或 3", 1)
        self._value_tab_label_entry(basic, 3, "参考数据起始", self.value_ref_data_start_row_var, "auto 或 4", 0)
        self._value_tab_label_entry(basic, 3, "配置数据起始", self.value_cfg_data_start_row_var, "auto 或 4", 1)

        # Tab 2: Keys and filters
        filters = tabview.tab("主键与过滤")
        self._value_tab_label_entry(filters, 0, "主键字段", self.value_key_fields_var, "auto / #id / type+item_id", 0)
        self._value_tab_label_entry(filters, 1, "参考表过滤", self.value_ref_filter_var, "例如 type=3；type=3;open=1", 0)
        self._value_tab_label_entry(filters, 1, "配置表过滤", self.value_cfg_filter_var, "例如 type=3；group in 1,2,3", 1)
        button_row = ctk.CTkFrame(filters, fg_color=COLOR_SURFACE, corner_radius=0)
        button_row.grid(row=2, column=0, columnspan=4, sticky="ew", pady=(8, 4))
        ctk.CTkButton(button_row, text="同步参考过滤到配置表", height=34, corner_radius=10, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self._copy_ref_filter_to_config).pack(side="left", padx=(0, 10))
        ctk.CTkButton(button_row, text="刷新过滤 / 执行预览", height=34, corner_radius=10, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color="#FFFFFF", font=self._font(13, "bold"), command=lambda: self._refresh_value_preview_tab()).pack(side="left")
        ctk.CTkLabel(filters, text="说明：过滤条件必须同时进入预览、执行和 Excel 摘要；字段不存在或语法错误会拦截执行。", text_color=COLOR_MUTED, font=self._font(12), anchor="w", justify="left", wraplength=850).grid(row=3, column=0, columnspan=4, sticky="ew", pady=(10, 0))

        # Tab 3: Field mapping
        mapping = tabview.tab("字段映射")
        mapping.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(mapping, text="字段映射（可空，同名字段自动匹配；手动格式：参考字段=>配置字段，每行一组）", text_color=COLOR_MUTED, font=self._font(12), anchor="w").grid(row=0, column=0, sticky="ew", pady=(0, 4))
        self.value_mapping_textbox = ctk.CTkTextbox(mapping, height=86, corner_radius=10, fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(12))
        self.value_mapping_textbox.grid(row=1, column=0, sticky="ew")
        self.value_mapping_textbox.insert("1.0", "# 可选：item=>reward_id\n# price_practical=>price")
        self._value_tab_label_entry(mapping, 2, "忽略字段", self.value_ignore_fields_var, "备注,说明,comment", 0)

        # Tab 4: Normalization rules
        rules = tabview.tab("标准化规则")
        self._value_tab_label_entry(rules, 0, "空值=0 字段", self.value_zero_equal_fields_var, "例如 reward", 0)
        self._value_tab_label_entry(rules, 0, "百分比字段", self.value_percent_fields_var, "rate,discount,percent", 1)
        self._value_tab_label_entry(rules, 1, "布尔字段", self.value_bool_fields_var, "可空，例如 enable,open", 0)
        self._value_tab_label_entry(rules, 1, "忽略大小写字段", self.value_case_insensitive_fields_var, "可空，例如 name", 1)
        checks = ctk.CTkFrame(rules, fg_color=COLOR_SURFACE, corner_radius=0)
        checks.grid(row=2, column=0, columnspan=4, sticky="ew", pady=(10, 0))
        for idx, (text, var) in enumerate([
            ("比对 reward/奖励字段", self.value_compare_rewards_var),
            ("比对价格字段", self.value_compare_price_var),
            ("比对 pcid", self.value_compare_pcid_var),
            ("比对持续时间", self.value_compare_duration_var),
            ("比对积分/分数", self.value_compare_score_var),
            ("比对限购次数", self.value_compare_limit_var),
        ]):
            ctk.CTkCheckBox(checks, text=text, variable=var, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13)).grid(row=idx // 3, column=idx % 3, sticky="w", padx=(0, 18), pady=4)

        # Tab 5: Preview
        preview_tab = tabview.tab("执行预览")
        preview_tab.grid_columnconfigure(0, weight=1)
        preview_tab.grid_rowconfigure(1, weight=1)
        preview_buttons = ctk.CTkFrame(preview_tab, fg_color=COLOR_SURFACE, corner_radius=0)
        preview_buttons.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ctk.CTkButton(preview_buttons, text="刷新预览", width=120, height=34, corner_radius=10, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=lambda: self._refresh_value_preview_tab()).pack(side="left", padx=(0, 10))
        self.value_preview_textbox = ctk.CTkTextbox(preview_tab, height=180, corner_radius=10, fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(12))
        self.value_preview_textbox.grid(row=1, column=0, sticky="nsew")
        self._set_textbox_text(self.value_preview_textbox, "点击“刷新预览”查看过滤条件、主键、数据量和风险提示。")

        # Tab 6: Result summary
        result_tab = tabview.tab("结果摘要")
        result_tab.grid_columnconfigure(0, weight=1)
        result_tab.grid_rowconfigure(0, weight=1)
        self.value_result_textbox = ctk.CTkTextbox(result_tab, height=210, corner_radius=10, fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13))
        self.value_result_textbox.grid(row=0, column=0, sticky="nsew")
        self._set_value_result_summary_text("暂无比对结果。完成后会显示参考表过滤后数量、配置表过滤后数量、漏配、多配、字段差异、待确认、过滤风险和导出路径。")

        # Bottom fixed action bar. It is outside all tabs so the core actions are always visible
        # even when a tab contains more rows than the available height.
        action_bar = ctk.CTkFrame(card, fg_color=COLOR_SURFACE_2, corner_radius=12, border_width=1, border_color=COLOR_BORDER)
        self.value_action_bar = action_bar
        action_bar.grid(row=4, column=0, sticky="ew", padx=18, pady=(0, 12))
        action_bar.grid_columnconfigure(3, weight=1)

        self.value_compare_button = ctk.CTkButton(
            action_bar,
            text="开始比对",
            width=118,
            height=34,
            corner_radius=10,
            fg_color=COLOR_ACCENT,
            hover_color=COLOR_ACCENT_HOVER,
            text_color="#FFFFFF",
            font=self._font(13, "bold"),
            command=self.start_compare,
        )
        self.value_compare_button.grid(row=0, column=0, sticky="w", padx=(12, 8), pady=9)

        self.value_export_result_button = ctk.CTkButton(
            action_bar,
            text="导出 Excel",
            width=104,
            height=34,
            corner_radius=10,
            fg_color=COLOR_SURFACE,
            hover_color=COLOR_HOVER,
            border_width=1,
            border_color=COLOR_BORDER,
            text_color=COLOR_TEXT,
            font=self._font(13, "bold"),
            command=self.export_compare_excel,
        )
        self.value_export_result_button.grid(row=0, column=1, sticky="w", padx=(0, 8), pady=9)

        self.value_open_result_button = ctk.CTkButton(
            action_bar,
            text="打开结果",
            width=104,
            height=34,
            corner_radius=10,
            fg_color=COLOR_SURFACE,
            hover_color=COLOR_HOVER,
            border_width=1,
            border_color=COLOR_BORDER,
            text_color=COLOR_TEXT,
            font=self._font(13, "bold"),
            command=self.open_compare_last_export,
        )
        self.value_open_result_button.grid(row=0, column=2, sticky="w", padx=(0, 12), pady=9)

        ctk.CTkLabel(
            action_bar,
            textvariable=self.value_action_status_var,
            text_color=COLOR_MUTED,
            font=self._font(12),
            anchor="w",
            justify="left",
            wraplength=620,
        ).grid(row=0, column=3, sticky="ew", padx=(0, 12), pady=9)

    def _clear_value_mapping_textbox(self):
        try:
            if self.value_mapping_textbox is not None and self.value_mapping_textbox.winfo_exists():
                self.value_mapping_textbox.delete("1.0", "end")
                self.value_mapping_textbox.insert("1.0", "# 可选：item=>reward_id\n# price_practical=>price")
        except Exception:
            pass

    def _reset_value_compare_runtime_state(self, reason: str = ""):
        self.compare_all_diffs = []
        self.compare_visible_diffs = []
        self.compare_last_export_path = ""
        self._last_value_compare_report = None
        try:
            self.clear_compare_table()
        except Exception:
            pass
        try:
            self.value_action_status_var.set("待执行 / 可开始比对")
            self.compare_summary_var.set(reason or "已重置结构化比对状态，请刷新预览后开始比对。")
            self._set_value_result_summary_text("暂无比对结果。切换文件后已清空旧结果，请重新开始比对。")
            self._set_textbox_text(getattr(self, "value_preview_textbox", None), reason or "已清空旧预览，请点击“刷新预览”。")
        except Exception:
            pass
        self._refresh_value_export_buttons()

    def _reset_value_reference_state(self):
        old_sheet = self.value_ref_sheet_var.get().strip()
        self.value_ref_sheet_var.set("auto")
        self.value_ref_header_row_var.set("auto")
        self.value_ref_data_start_row_var.set("auto")
        self.value_ref_filter_var.set("")
        self._clear_value_mapping_textbox()
        msg = "已选择新的参考表，旧参考 Sheet、字段行、过滤条件、字段映射和预览结果已清空；将使用 auto 重新扫描当前文件。"
        if old_sheet and old_sheet.lower() not in {"auto", "自动"}:
            msg += f"\n旧 Sheet「{old_sheet}」不会继续沿用。"
        self.value_compare_summary_var.set(msg)
        self._reset_value_compare_runtime_state(msg)

    def _reset_value_config_state(self):
        self.value_cfg_sheet_var.set("")
        self.value_cfg_header_row_var.set("auto")
        self.value_cfg_data_start_row_var.set("auto")
        self.value_cfg_filter_var.set("")
        self._clear_value_mapping_textbox()
        msg = "已选择新的配置表，旧字段行、数据起始行、过滤条件、字段缓存和比对结果已清空；将重新扫描当前配置表。"
        self.value_compare_summary_var.set(msg)
        self._reset_value_compare_runtime_state(msg)

    def _refresh_value_export_buttons(self):
        has_result = bool(getattr(self, "compare_all_diffs", None)) and self._is_value_config_mode()
        has_export = bool(getattr(self, "compare_last_export_path", "")) and os.path.exists(getattr(self, "compare_last_export_path", ""))
        try:
            if getattr(self, "value_export_result_button", None) is not None and self.value_export_result_button.winfo_exists():
                self.value_export_result_button.configure(state="normal" if has_result and not self.compare_is_working else "disabled")
        except Exception:
            pass
        try:
            if getattr(self, "value_open_result_button", None) is not None and self.value_open_result_button.winfo_exists():
                self.value_open_result_button.configure(state="normal" if has_export else "disabled")
        except Exception:
            pass

    def _on_value_template_changed(self):
        template = self.value_template_var.get().strip()
        compact = template.lower()
        if any(k in compact for k in ["商业化", "礼包", "mall"]):
            for variable in (self.value_ref_filter_var, self.value_cfg_filter_var):
                if variable.get().strip() in {"type=3", "type in 2000,13"}:
                    variable.set("")
            self.value_compare_summary_var.set("商业化礼包专项：使用已验证的礼包/档位解析与奖励比对，不强制 name 主键，也不预设单一 type 过滤。")
        elif any(k in compact for k in ["商店", "shop"]):
            if not self.value_ref_filter_var.get().strip():
                self.value_ref_filter_var.set("type=3")
            if not self.value_cfg_filter_var.get().strip():
                self.value_cfg_filter_var.set("type=3")
            self.value_compare_summary_var.set("商店商品配置模板：已推荐参考表过滤 type=3、配置表过滤 type=3；可按实际业务手动修改。")
        elif "活动" in template:
            self.value_compare_summary_var.set("活动配置模板：请填写 activity_id / event_id / group 等过滤条件，避免全表参与比对。")
        else:
            self.value_compare_summary_var.set("通用结构化比对：如只想比对部分数据，请同时填写参考表和配置表过滤条件。")
        self._refresh_value_preview_tab(silent=True)


    def _build_translation_compare_settings(self, parent):
        card = ctk.CTkFrame(parent, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        self.translation_settings_card = card
        card.pack(fill="x", pady=(0, 16))
        card.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(card, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=18, pady=(16, 6))
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            header,
            text="翻译源文件 vs 游戏翻译配置表设置",
            text_color=COLOR_TEXT,
            font=self._font(16, "bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(
            header,
            text="仅在该模式下显示，减少普通文件比对页面的滚动卡顿。",
            text_color=COLOR_MUTED,
            font=self._font(12),
            anchor="e",
        ).grid(row=0, column=1, sticky="e", padx=(12, 0))

        ctk.CTkLabel(
            card,
            text="约定：左侧选择游戏配置表，右侧选择翻译源文件。支持 xlsx/csv；主键和语言列可自动识别，也可手动填写。",
            text_color=COLOR_MUTED,
            font=self._font(13),
            anchor="w",
            wraplength=900,
        ).grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 10))

        grid = ctk.CTkFrame(card, fg_color="transparent")
        grid.grid(row=2, column=0, sticky="ew", padx=18, pady=(0, 8))
        for col in range(5):
            grid.grid_columnconfigure(col, weight=1)

        self._small_entry(grid, 0, 0, "源表头行", self.tr_source_header_row_var, "1")
        self._small_entry(grid, 0, 1, "源数据行", self.tr_source_data_start_row_var, "2")
        self._small_entry(grid, 0, 2, "源主键列", self.tr_source_key_column_var, "自动")
        self._small_entry(grid, 0, 3, "源 Sheet", self.tr_source_sheet_var, "默认第一个")
        self._small_entry(grid, 0, 4, "配置表头行", self.tr_config_header_row_var, "1")
        self._small_entry(grid, 1, 0, "配置真实字段行", self.tr_config_real_header_row_var, "可空")
        self._small_entry(grid, 1, 1, "配置数据行", self.tr_config_data_start_row_var, "2")
        self._small_entry(grid, 1, 2, "配置主键列", self.tr_config_key_column_var, "自动")
        self._small_entry(grid, 1, 3, "配置 Sheet", self.tr_config_sheet_var, "默认第一个")

        checks = ctk.CTkFrame(card, fg_color="transparent")
        checks.grid(row=3, column=0, sticky="ew", padx=18, pady=(4, 8))
        for idx, (text, var) in enumerate([
            ("检查配置表多余 ID", self.tr_check_extra_ids_var),
            ("忽略首尾空格", self.tr_ignore_trim_var),
            ("忽略所有空格", self.tr_ignore_all_spaces_var),
            ("忽略换行", self.tr_ignore_newlines_var),
            ("忽略大小写", self.tr_ignore_case_var),
            ("忽略全角半角", self.tr_ignore_width_var),
            ("忽略中英标点", self.tr_ignore_punctuation_var),
        ]):
            ctk.CTkCheckBox(checks, text=text, variable=var, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13)).grid(row=idx // 4, column=idx % 4, sticky="w", padx=(0, 14), pady=4)

        compact = ctk.CTkFrame(card, fg_color=COLOR_SURFACE_2, corner_radius=12, border_width=1, border_color=COLOR_BORDER)
        compact.grid(row=4, column=0, sticky="ew", padx=18, pady=(0, 10))
        compact.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(compact, text="参与语言", text_color=COLOR_TEXT, font=self._font(14, "bold"), anchor="w").grid(row=0, column=0, sticky="w", padx=12, pady=10)
        ctk.CTkLabel(compact, textvariable=self.tr_language_summary_var, text_color=COLOR_MUTED, font=self._font(12), anchor="w", wraplength=390).grid(row=0, column=1, sticky="ew", padx=(0, 12), pady=10)
        ctk.CTkButton(compact, text="全选", width=58, height=30, corner_radius=9, fg_color=COLOR_SURFACE, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(12), command=lambda: self._set_translation_languages("all")).grid(row=0, column=2, padx=(0, 8), pady=10)
        ctk.CTkButton(compact, text="常用", width=58, height=30, corner_radius=9, fg_color=COLOR_SURFACE, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(12), command=lambda: self._set_translation_languages("common")).grid(row=0, column=3, padx=(0, 8), pady=10)
        ctk.CTkButton(compact, text="清空", width=58, height=30, corner_radius=9, fg_color=COLOR_SURFACE, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(12), command=lambda: self._set_translation_languages("none")).grid(row=0, column=4, padx=(0, 8), pady=10)
        self.tr_advanced_toggle_button = ctk.CTkButton(compact, text="展开语言映射 / 参与语言", width=170, height=30, corner_radius=9, fg_color=COLOR_SELECTED, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(12, "bold"), command=self._toggle_translation_advanced)
        self.tr_advanced_toggle_button.grid(row=0, column=5, padx=(0, 12), pady=10)

        self.tr_advanced_frame = ctk.CTkFrame(card, fg_color="transparent")
        self.tr_advanced_frame.grid(row=5, column=0, sticky="ew", padx=18, pady=(0, 16))
        self.tr_advanced_frame.grid_columnconfigure(0, weight=1)

        lang_box = ctk.CTkFrame(self.tr_advanced_frame, fg_color=COLOR_SURFACE_2, corner_radius=12, border_width=1, border_color=COLOR_BORDER)
        lang_box.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        lang_box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(lang_box, text="语言勾选（默认折叠，避免滚动时渲染过重）", text_color=COLOR_TEXT, font=self._font(14, "bold"), anchor="w").grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 4))
        lang_grid = ctk.CTkFrame(lang_box, fg_color="transparent")
        lang_grid.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 10))
        for col in range(4):
            lang_grid.grid_columnconfigure(col, weight=1)
        for index, name in enumerate(self.tr_language_names):
            row = index // 4
            col = index % 4
            ctk.CTkCheckBox(
                lang_grid,
                text=name,
                variable=self.tr_language_vars[name],
                fg_color=COLOR_ACCENT,
                hover_color=COLOR_ACCENT_HOVER,
                border_color=COLOR_BORDER,
                text_color=COLOR_TEXT,
                font=self._font(12),
                width=140,
                command=self._refresh_translation_language_summary,
            ).grid(row=row, column=col, sticky="w", padx=(0, 10), pady=4)

        mapping_box = ctk.CTkFrame(self.tr_advanced_frame, fg_color="transparent")
        mapping_box.grid(row=1, column=0, sticky="ew")
        mapping_box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(mapping_box, text="语言列映射（格式：语言: 源列候选 => 配置列候选；用 / 分隔候选名）", text_color=COLOR_MUTED, font=self._font(12), anchor="w").grid(row=0, column=0, sticky="ew", pady=(0, 4))
        self.tr_mapping_textbox = ctk.CTkTextbox(mapping_box, height=86, corner_radius=10, fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(12))
        self.tr_mapping_textbox.grid(row=1, column=0, sticky="ew")
        self.tr_mapping_textbox.insert("1.0", DEFAULT_LANGUAGE_MAPPING_TEXT)

        self.tr_advanced_frame.grid_remove()
        self.tr_advanced_visible = False
        if self.tr_advanced_toggle_button is not None:
            self.tr_advanced_toggle_button.configure(text="展开语言映射 / 参与语言")
        self._refresh_translation_language_summary()

    def _small_entry(self, parent, row, col, label, variable, placeholder=""):
        box = ctk.CTkFrame(parent, fg_color="transparent")
        box.grid(row=row, column=col, sticky="ew", padx=(0, 12), pady=5)
        box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(box, text=label, text_color=COLOR_MUTED, font=self._font(12), anchor="w").grid(row=0, column=0, sticky="ew")
        entry = ctk.CTkEntry(box, textvariable=variable, placeholder_text=placeholder, height=34, corner_radius=9, fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(12))
        entry.grid(row=1, column=0, sticky="ew", pady=(3, 0))
        return entry

    def _toggle_translation_advanced(self):
        if self.tr_advanced_frame is None:
            return
        self.tr_advanced_visible = not self.tr_advanced_visible
        card = getattr(self, "translation_settings_card", None)
        if card is not None:
            for child in card.winfo_children():
                info = child.grid_info()
                if not info:
                    continue
                row = int(info.get("row", -1))
                if row in (1, 2, 3):
                    if self.tr_advanced_visible:
                        child.grid_remove()
                    else:
                        child.grid()
        if self.tr_advanced_visible:
            self.tr_advanced_frame.grid()
            if self.tr_advanced_toggle_button is not None:
                self.tr_advanced_toggle_button.configure(text="返回基础设置")
        else:
            self.tr_advanced_frame.grid_remove()
            if self.tr_advanced_toggle_button is not None:
                self.tr_advanced_toggle_button.configure(text="展开语言映射 / 参与语言")
        # 避免在折叠/展开时强制同步重绘，降低滚动区域残影和闪烁概率
        try:
            self.after_idle(lambda: None)
        except Exception:
            pass

    def _set_translation_languages(self, mode):
        common = {"中文简体", "中文繁体", "英语", "日语", "韩语"}
        for name, var in self.tr_language_vars.items():
            if mode == "all":
                var.set(True)
            elif mode == "none":
                var.set(False)
            elif mode == "common":
                var.set(name in common)
        self._refresh_translation_language_summary()

    def _refresh_translation_language_summary(self):
        enabled = [name for name, var in self.tr_language_vars.items() if var.get()]
        if not enabled:
            text = "未选择语言"
        elif len(enabled) <= 6:
            text = "、".join(enabled)
        else:
            text = "、".join(enabled[:6]) + f" 等 {len(enabled)} 种"
        try:
            self.tr_language_summary_var.set(text)
        except Exception:
            pass

    def _set_compare_common_options_visible(self, visible: bool):
        """Fold plain-text compare options when structured value compare is active.

        The value-config editor needs vertical room. Keeping the generic ignore-space /
        ignore-line / ignore-field rows visible above the tabbed editor was squeezing the
        Notebook down to an unusable strip on 1366x768 and similar Windows screens.
        """
        card = getattr(self, "compare_option_card", None)
        if card is None:
            return
        try:
            for child in card.winfo_children():
                info = child.grid_info()
                if not info:
                    continue
                row = int(info.get("row", -1))
                if row in (2, 3, 4):
                    if visible:
                        child.grid()
                    else:
                        child.grid_remove()
        except Exception:
            pass

    def _update_compare_mode_settings_visibility(self):
        wrap = getattr(self, "compare_page_wrap", None)
        if wrap is not None:
            if self._is_value_config_mode() and self.value_compare_settings_card is None:
                self._build_value_compare_settings(wrap)
            elif self._is_translation_mode() and self.translation_settings_card is None:
                self._build_translation_compare_settings(wrap)

        value_card = getattr(self, "value_compare_settings_card", None)
        translation_card = getattr(self, "translation_settings_card", None)
        hint = getattr(self, "compare_special_hint", None)
        value_mode = self._is_value_config_mode()
        translation_mode = self._is_translation_mode()

        def hide(card):
            if card is None:
                return
            try:
                card.pack_forget()
            except Exception:
                pass

        def show(card):
            if card is None:
                return
            try:
                card.pack_forget()
                card.pack(fill="both", expand=True)
            except Exception:
                pass

        self._set_compare_common_options_visible(not value_mode)

        if value_mode:
            show(value_card)
            hide(translation_card)
            hide(hint)
        elif translation_mode:
            hide(value_card)
            show(translation_card)
            hide(hint)
        else:
            hide(value_card)
            hide(translation_card)
            show(hint)
        self._refresh_compare_scroll_region()

    def _update_translation_settings_visibility(self):
        self._update_compare_mode_settings_visibility()

    def _on_compare_mode_changed(self):
        self._update_compare_mode_settings_visibility()
        tabs = getattr(self, "compare_page_tabs", None)
        if tabs is not None:
            target = self.compare_special_tab_name if (self._is_value_config_mode() or self._is_translation_mode()) else self.compare_file_tab_name
            try:
                tabs.set(target)
            except Exception:
                pass
        self.refresh_compare_table(self.compare_visible_diffs)
        self._refresh_compare_scroll_region()

    def _is_translation_mode(self):
        return self.compare_mode_var.get() == "翻译源文件 vs 游戏翻译配置表"

    def _is_value_config_mode(self):
        return self.compare_mode_var.get() == VALUE_COMPARE_MODE

    def _compare_file_panel(self, parent, column, title, path_var, status_var, button_text, command):
        panel = ctk.CTkFrame(parent, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        panel.grid(row=0, column=column, padx=(0, 8) if column == 0 else (8, 0), sticky="nsew")
        panel.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(panel, text=title, text_color=COLOR_TEXT, font=self._font(16, "bold"), anchor="w").grid(row=0, column=0, padx=18, pady=(18, 6), sticky="ew")
        ctk.CTkEntry(panel, textvariable=path_var, state="readonly", height=38, corner_radius=10, fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13)).grid(row=1, column=0, padx=18, pady=6, sticky="ew")
        ctk.CTkLabel(panel, textvariable=status_var, text_color=COLOR_MUTED, anchor="w", font=self._font(13)).grid(row=2, column=0, padx=18, pady=(0, 10), sticky="ew")
        ctk.CTkButton(panel, text=button_text, height=38, corner_radius=12, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color="#FFFFFF", font=self._font(14, "bold"), command=command).grid(row=3, column=0, padx=18, pady=(0, 18), sticky="ew")

    def select_compare_file(self, role: str):
        filetypes = [
            ("支持的文件", " ".join(f"*{ext}" for ext in sorted(SUPPORTED_EXTENSIONS))),
            ("所有文件", "*.*"),
        ]
        if role == "actual":
            paths = filedialog.askopenfilenames(title="选择一个或多个游戏配置表", filetypes=filetypes)
            if not paths:
                return
            path = ";".join(paths)
            self.compare_actual_path_var.set(path)
            if len(paths) == 1:
                self.compare_actual_status_var.set(f"已选择：{get_file_info(paths[0])}")
            else:
                names = "、".join(os.path.basename(p) for p in paths[:3])
                more = f" 等 {len(paths)} 个文件" if len(paths) > 3 else ""
                self.compare_actual_status_var.set(f"已选择 {len(paths)} 个配置表：{names}{more}")
            self._reset_value_config_state()
        else:
            path = filedialog.askopenfilename(title="选择参考表", filetypes=filetypes)
            if not path:
                return
            self.compare_document_path_var.set(path)
            self.compare_document_status_var.set(f"已选择：{get_file_info(path)}")
            self._reset_value_reference_state()
        self._auto_recommend_compare_mode()
        if self._is_value_config_mode():
            self._refresh_value_preview_tab(silent=True)


    def _auto_recommend_compare_mode(self):
        actual_names = " ".join(os.path.basename(p.strip()) for p in self.compare_actual_path_var.get().strip().split(";") if p.strip())
        names = " ".join([
            actual_names,
            os.path.basename(self.compare_document_path_var.get().strip()),
        ]).lower()
        if any(keyword in names for keyword in ["月卡", "周卡", "礼包", "累充", "双子星", "无限", "无尽", "商城", "mall", "神秘商店", "shop_basic"]):
            self.compare_mode_var.set(VALUE_COMPARE_MODE)
            if "神秘商店" in names or "shop_basic" in names:
                try:
                    self.value_template_var.set("商店商品配置")
                    self._on_value_template_changed()
                except Exception:
                    pass
        elif any(keyword in names for keyword in ["翻译", "多语言", "language", "localization", "text"]):
            self.compare_mode_var.set("翻译源文件 vs 游戏翻译配置表")
        elif self.compare_actual_path_var.get().strip() and self.compare_document_path_var.get().strip():
            actual_ext = os.path.splitext(self.compare_actual_path_var.get().strip())[1].lower()
            doc_ext = os.path.splitext(self.compare_document_path_var.get().strip())[1].lower()
            if actual_ext in {".csv", ".xlsx", ".json"} and doc_ext in {".docx", ".xlsx", ".pdf", ".txt"}:
                self.compare_mode_var.set("策划文档 vs 配置文件")
        try:
            self._on_compare_mode_changed()
        except Exception:
            pass

    def start_compare(self):
        if self.compare_is_working:
            return
        actual_path = self.compare_actual_path_var.get().strip()
        document_path = self.compare_document_path_var.get().strip()
        if not actual_path or not document_path:
            messagebox.showwarning("提示", "请先选择待比对文件和参考文档")
            return

        self.compare_is_working = True
        self.compare_last_export_path = ""
        self._refresh_compare_button()
        self._refresh_value_export_buttons()
        self._set_badge(True)
        self.compare_summary_var.set("正在读取文件并比对...")
        try:
            self.value_action_status_var.set("正在比对...")
        except Exception:
            pass
        self.clear_compare_table()

        if self._is_value_config_mode():
            options = prepare_value_compare_options(self.get_value_compare_options(), actual_path, document_path)
            # Keep UI, preview, execution and export using the same canonical filter parameters.
            try:
                self.value_template_var.set(getattr(options, "template_name", "") or "通用配置表")
                self.value_ref_filter_var.set(getattr(options, "reference_filter", "") or "")
                self.value_cfg_filter_var.set(getattr(options, "config_filter", "") or "")
            except Exception:
                pass
            try:
                preview = build_value_compare_rule_preview(actual_path, document_path, options)
            except Exception as exc:
                self.compare_is_working = False
                self._refresh_compare_button()
                self._set_badge(False)
                self.compare_summary_var.set("结构化比对已拦截：过滤条件或表结构未通过校验")
                try:
                    self.value_action_status_var.set("规则校验失败")
                except Exception:
                    pass
                messagebox.showwarning("结构化比对规则校验失败", str(exc))
                return
            if not messagebox.askyesno("结构化比对规则预览", preview + "\n\n是否开始比对？"):
                self.compare_is_working = False
                self._refresh_compare_button()
                self._set_badge(False)
                self.compare_summary_var.set("已取消结构化比对")
                try:
                    self.value_action_status_var.set("已取消")
                except Exception:
                    pass
                return
            threading.Thread(target=self._value_compare_worker, args=(actual_path, document_path, options), daemon=True).start()
            return

        options = self.get_compare_options()
        threading.Thread(target=self._compare_worker, args=(actual_path, document_path, options), daemon=True).start()


    def get_value_compare_options(self):
        mapping_text = ""
        if self.value_mapping_textbox is not None:
            try:
                mapping_text = self.value_mapping_textbox.get("1.0", "end").strip()
            except Exception:
                mapping_text = ""
        return ValueCompareOptions(
            manual_ids=self.value_compare_ids_var.get().strip(),
            keywords=self.value_compare_keywords_var.get().strip(),
            compare_rewards=self.value_compare_rewards_var.get(),
            compare_price=self.value_compare_price_var.get(),
            compare_pcid=self.value_compare_pcid_var.get(),
            compare_duration=self.value_compare_duration_var.get(),
            compare_score=self.value_compare_score_var.get(),
            compare_limit_time=self.value_compare_limit_var.get(),
            reference_sheet=self.value_ref_sheet_var.get().strip(),
            config_sheet=self.value_cfg_sheet_var.get().strip(),
            reference_header_row=self.value_ref_header_row_var.get().strip(),
            reference_data_start_row=self.value_ref_data_start_row_var.get().strip(),
            config_header_row=self.value_cfg_header_row_var.get().strip(),
            config_data_start_row=self.value_cfg_data_start_row_var.get().strip(),
            key_fields=self.value_key_fields_var.get().strip(),
            reference_filter=self.value_ref_filter_var.get().strip(),
            config_filter=self.value_cfg_filter_var.get().strip(),
            template_name=self.value_template_var.get().strip(),
            field_mappings=mapping_text,
            ignore_fields=self.value_ignore_fields_var.get().strip(),
            zero_equal_fields=self.value_zero_equal_fields_var.get().strip(),
            percent_fields=self.value_percent_fields_var.get().strip(),
            bool_fields=self.value_bool_fields_var.get().strip(),
            case_insensitive_fields=self.value_case_insensitive_fields_var.get().strip(),
            enable_general_compare=True,
        )

    def _value_compare_worker(self, actual_path: str, document_path: str, options):
        try:
            diffs, report = compare_value_reference_to_config(actual_path, document_path, options)
            self.after(0, lambda: self._on_value_compare_success(actual_path, document_path, diffs, report))
        except Exception as exc:
            message = f"专项比对失败：{exc}"
            self.after(0, lambda msg=message: self._on_compare_error(msg))

    def _on_value_compare_success(self, actual_path: str, document_path: str, diffs, report):
        self.compare_is_working = False
        self._refresh_compare_button()
        self._set_badge(False)
        self.compare_all_diffs = diffs
        self.compare_visible_diffs = diffs
        self._last_value_compare_report = report
        self.compare_actual_status_var.set(f"配置表读取完成：{os.path.basename(actual_path)}")
        self.compare_document_status_var.set(f"参考表读取完成：{os.path.basename(document_path)}")
        try:
            self.value_compare_summary_var.set(report.summary())
        except Exception:
            pass
        self.apply_compare_search_filter()
        summary_text = self._build_value_compare_result_summary_text(diffs, report, self.compare_last_export_path)
        self.compare_summary_var.set(summary_text)
        self._set_value_result_summary_text(summary_text)
        try:
            self.value_action_status_var.set("比对完成，等待导出")
        except Exception:
            pass
        self._refresh_value_export_buttons()
        try:
            if getattr(self, "value_tabview", None) is not None and self.value_tabview.winfo_exists():
                self.value_tabview.set("结果摘要")
        except Exception:
            pass
        self._update_compare_summary_wrap()
        self._refresh_compare_scroll_region()

    def get_compare_options(self):
        normalize_options = NormalizeOptions(
            ignore_spaces=self.compare_ignore_spaces_var.get(),
            ignore_newlines=self.compare_ignore_newlines_var.get(),
            ignore_case=self.compare_ignore_case_var.get(),
        )
        ignore_fields = clean_ignore_fields(self.compare_ignore_fields_var.get(), normalize_options)
        translation_options = None
        if self._is_translation_mode():
            mapping_text = ""
            if self.tr_mapping_textbox is not None:
                try:
                    mapping_text = self.tr_mapping_textbox.get("1.0", "end").strip()
                except Exception:
                    mapping_text = DEFAULT_LANGUAGE_MAPPING_TEXT
            enabled_languages = [name for name, var in self.tr_language_vars.items() if var.get()]
            translation_options = TranslationCompareOptions(
                source_header_row=safe_int(self.tr_source_header_row_var.get(), 1),
                source_data_start_row=safe_int(self.tr_source_data_start_row_var.get(), 2),
                source_key_column=self.tr_source_key_column_var.get().strip(),
                source_sheet=self.tr_source_sheet_var.get().strip(),
                config_header_row=safe_int(self.tr_config_header_row_var.get(), 1),
                config_real_header_row=safe_int(self.tr_config_real_header_row_var.get(), 0) if self.tr_config_real_header_row_var.get().strip() else 0,
                config_data_start_row=safe_int(self.tr_config_data_start_row_var.get(), 2),
                config_key_column=self.tr_config_key_column_var.get().strip(),
                config_sheet=self.tr_config_sheet_var.get().strip(),
                language_mapping_text=mapping_text,
                enabled_languages=enabled_languages,
                ignore_trim=self.tr_ignore_trim_var.get(),
                ignore_all_spaces=self.tr_ignore_all_spaces_var.get(),
                ignore_newlines=self.tr_ignore_newlines_var.get(),
                ignore_case=self.tr_ignore_case_var.get(),
                ignore_full_half_width=self.tr_ignore_width_var.get(),
                ignore_punctuation=self.tr_ignore_punctuation_var.get(),
                check_extra_ids=self.tr_check_extra_ids_var.get(),
            )
        return CompareOptions(mode=self.compare_mode_var.get(), normalize=normalize_options, ignore_fields=ignore_fields, translation_options=translation_options)

    def _compare_worker(self, actual_path: str, document_path: str, options):
        try:
            actual = read_file(actual_path)
            document = read_file(document_path)
            diffs = ContentComparer().compare(actual=actual, document=document, options=options)
            self.after(0, lambda: self._on_compare_success(actual, document, diffs))
        except FileReadError as exc:
            message = str(exc)
            self.after(0, lambda msg=message: self._on_compare_error(msg))
        except Exception as exc:
            message = f"比对失败：{exc}"
            self.after(0, lambda msg=message: self._on_compare_error(msg))

    def _on_compare_success(self, actual, document, diffs):
        self.compare_is_working = False
        self._refresh_compare_button()
        self._set_badge(False)
        self.compare_all_diffs = diffs
        self.compare_visible_diffs = diffs
        self.compare_actual_status_var.set(f"读取成功：{actual.ext or '无扩展名'} / 行数 {len(actual.lines)} / 字段 {len(actual.fields)}")
        self.compare_document_status_var.set(f"读取成功：{document.ext or '无扩展名'} / 行数 {len(document.lines)} / 字段 {len(document.fields)}")
        self.apply_compare_search_filter()
        if not diffs:
            self.compare_summary_var.set("比对完成：未发现差异")
        else:
            counts = self._count_compare_by_type(diffs)
            top_counts = " | ".join(f"{key} {value}" for key, value in sorted(counts.items())[:6])
            self.compare_summary_var.set(f"比对完成：共 {len(diffs)} 条差异 | {top_counts}")
        try:
            self.compare_page_tabs.set(self.compare_result_tab_name)
        except Exception:
            pass

    def _on_compare_error(self, message: str):
        self.compare_is_working = False
        self._refresh_compare_button()
        self._set_badge(False)
        self.compare_summary_var.set("比对失败")
        messagebox.showerror("错误", message)

    @staticmethod
    def _count_compare_by_type(diffs):
        result = {}
        for item in diffs:
            result[item.diff_type] = result.get(item.diff_type, 0) + 1
        return result

    def apply_compare_search_filter(self):
        keyword = self.compare_search_var.get().strip().lower()
        if not keyword:
            self.compare_visible_diffs = self.compare_all_diffs
        else:
            self.compare_visible_diffs = [item for item in self.compare_all_diffs if keyword in self._compare_diff_to_text(item).lower()]
        self.compare_page_index = 1
        self.refresh_compare_table(self.compare_visible_diffs)

    def refresh_compare_table(self, diffs):
        if not hasattr(self, "compare_tree") or not self.compare_tree.winfo_exists():
            return
        self._configure_compare_tree_columns(self._is_translation_mode() or self._has_translation_diffs(diffs), self._is_value_config_mode())
        self.clear_compare_table()
        self.compare_page_index, total_pages, start, end = self._get_page_bounds(len(diffs), self.compare_page_index)
        self.compare_page_jump_var.set(str(self.compare_page_index))
        self.compare_page_status_var.set(f"第 {self.compare_page_index}/{total_pages} 页，每页 {UI_PAGE_SIZE} 条，当前筛选 {len(diffs)} 条")
        for item in list(diffs)[start:end]:
            if self._has_translation_diffs([item]):
                values = (
                    item.index,
                    item.diff_type,
                    item.item_id,
                    item.language,
                    item.source_row,
                    item.config_row,
                    self._shorten_compare_text(item.document_content),
                    self._shorten_compare_text(item.file_content),
                    self._shorten_compare_text(item.remark, 180),
                )
            else:
                values = (
                    item.index,
                    item.diff_type,
                    self._shorten_compare_text(item.document_content),
                    self._shorten_compare_text(item.file_content),
                    self._shorten_compare_text(item.location, 120),
                    self._shorten_compare_text(item.remark, 160),
                )
            self.compare_tree.insert("", "end", values=values)
        self._refresh_compare_scroll_region()

    def compare_prev_page(self):
        self.compare_page_index = max(1, self.compare_page_index - 1)
        self.refresh_compare_table(self.compare_visible_diffs)

    def compare_next_page(self):
        self.compare_page_index += 1
        self.refresh_compare_table(self.compare_visible_diffs)

    def compare_jump_page(self):
        try:
            self.compare_page_index = int(self.compare_page_jump_var.get().strip() or "1")
        except Exception:
            self.compare_page_index = 1
        self.refresh_compare_table(self.compare_visible_diffs)

    @staticmethod
    def _has_translation_diffs(diffs):
        return any(getattr(item, "item_id", "") or getattr(item, "language", "") or getattr(item, "source_row", "") or getattr(item, "config_row", "") for item in diffs)

    def _configure_compare_tree_columns(self, translation_columns=False, value_columns=False):
        if translation_columns:
            columns = ("index", "diff_type", "item_id", "language", "source_row", "config_row", "document_content", "file_content", "remark")
            headings = {
                "index": "序号", "diff_type": "差异类型", "item_id": "ID", "language": "语言",
                "source_row": "源文件行号", "config_row": "配置表行号", "document_content": "源文件文本",
                "file_content": "配置表文本", "remark": "备注"
            }
            widths = {"index": 58, "diff_type": 110, "item_id": 120, "language": 90, "source_row": 86, "config_row": 86, "document_content": 260, "file_content": 260, "remark": 220}
        elif value_columns:
            columns = ("index", "diff_type", "document_content", "file_content", "location", "remark")
            headings = {"index": "序号", "diff_type": "差异类型", "document_content": "参考表内容", "file_content": "配置表内容", "location": "商品/档位", "remark": "问题说明 / 建议处理"}
            widths = {"index": 58, "diff_type": 130, "document_content": 360, "file_content": 360, "location": 260, "remark": 360}
        else:
            columns = ("index", "diff_type", "document_content", "file_content", "location", "remark")
            headings = {"index": "序号", "diff_type": "差异类型", "document_content": "文档内容", "file_content": "文件内容", "location": "所在位置", "remark": "备注"}
            widths = {"index": 58, "diff_type": 96, "document_content": 300, "file_content": 300, "location": 170, "remark": 230}
        current = tuple(self.compare_tree["columns"])
        if current != columns:
            self.compare_tree.configure(columns=columns)
        for col in columns:
            self.compare_tree.heading(col, text=headings[col])
            self.compare_tree.column(col, width=widths[col], anchor="center" if col in {"index", "diff_type", "language", "source_row", "config_row"} else "w", stretch=col not in {"index", "diff_type", "language", "source_row", "config_row"})

    def clear_compare_table(self):
        if not hasattr(self, "compare_tree") or not self.compare_tree.winfo_exists():
            return
        for item in self.compare_tree.get_children():
            self.compare_tree.delete(item)

    @staticmethod
    def _shorten_compare_text(text: str, max_len: int = 220) -> str:
        text = "" if text is None else str(text).replace("\n", " ⏎ ")
        return text if len(text) <= max_len else text[:max_len] + "..."

    @staticmethod
    def _compare_diff_to_text(item) -> str:
        if getattr(item, "item_id", "") or getattr(item, "language", "") or getattr(item, "source_row", "") or getattr(item, "config_row", ""):
            return "\t".join([str(item.index), item.diff_type, item.item_id, item.language, item.source_row, item.config_row, item.document_content, item.file_content, item.remark])
        return "\t".join([str(item.index), item.diff_type, item.document_content, item.file_content, item.location, item.remark])


    def show_compare_diff_detail(self, event=None):
        if not hasattr(self, "compare_tree") or not self.compare_tree.winfo_exists():
            return
        selected = self.compare_tree.selection()
        if not selected:
            return
        item_values = self.compare_tree.item(selected[0], "values")
        if not item_values:
            return
        try:
            index = int(item_values[0])
        except Exception:
            return
        target = None
        for diff in self.compare_visible_diffs:
            if getattr(diff, "index", None) == index:
                target = diff
                break
        if target is None:
            return
        top = ctk.CTkToplevel(self)
        top.title("差异详情")
        top.geometry("760x520")
        top.transient(self)
        top.grab_set()
        top.grid_columnconfigure(0, weight=1)
        top.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(top, text="差异详情", text_color=COLOR_TEXT, font=self._font(18, "bold"), anchor="w").grid(row=0, column=0, sticky="ew", padx=18, pady=(16, 8))
        detail = (
            f"序号：{target.index}\n"
            f"差异类型：{target.diff_type}\n"
            f"所在位置：{target.location}\n\n"
            f"参考表 / 文档内容：\n{target.document_content}\n\n"
            f"配置表 / 文件内容：\n{target.file_content}\n\n"
            f"备注 / 建议：\n{target.remark}"
        )
        box = ctk.CTkTextbox(top, corner_radius=12, fg_color=COLOR_SURFACE_2, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13), wrap="word")
        box.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 12))
        box.insert("1.0", detail)
        box.configure(state="disabled")
        buttons = ctk.CTkFrame(top, fg_color="transparent")
        buttons.grid(row=2, column=0, sticky="e", padx=18, pady=(0, 16))
        def copy_detail():
            self.clipboard_clear()
            self.clipboard_append(detail)
            messagebox.showinfo("已复制", "已复制当前差异详情")
        ctk.CTkButton(buttons, text="复制详情", width=110, height=34, corner_radius=12, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color="#FFFFFF", font=self._font(13, "bold"), command=copy_detail).grid(row=0, column=0, padx=(0, 10))
        ctk.CTkButton(buttons, text="关闭", width=88, height=34, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=top.destroy).grid(row=0, column=1)

    def copy_compare_visible_diffs(self):
        if not self.compare_visible_diffs:
            messagebox.showinfo("提示", "当前没有可复制的差异结果")
            return
        if self._has_translation_diffs(self.compare_visible_diffs):
            headers = ["序号", "差异类型", "ID", "语言", "源文件行号", "配置表行号", "源文件文本", "配置表文本", "备注"]
        elif self._is_value_config_mode():
            headers = ["序号", "差异类型", "参考表内容", "配置表内容", "商品/档位", "问题说明 / 建议处理"]
        else:
            headers = ["序号", "差异类型", "文档内容", "文件内容", "所在位置", "备注"]
        rows = ["\t".join(headers)]
        for item in self.compare_visible_diffs:
            rows.append(self._compare_diff_to_text(item))
        text = "\n".join(rows)
        self.clipboard_clear()
        self.clipboard_append(text)
        messagebox.showinfo("提示", f"已复制 {len(self.compare_visible_diffs)} 条差异结果")

    def open_compare_last_export(self):
        """Open the last exported compare report, or ask the user to export first."""
        path = getattr(self, "compare_last_export_path", "") or ""
        if not path or not os.path.exists(path):
            messagebox.showinfo("提示", "暂无可打开的导出结果，请先点击“导出 Excel”保存报告。")
            return
        try:
            os.startfile(path)
        except Exception as exc:
            messagebox.showerror("打开失败", str(exc))

    def export_compare_excel(self):
        if not self.compare_visible_diffs:
            if self._is_value_config_mode():
                messagebox.showinfo("提示", "请先开始结构化比对，完成后再点击“导出 Excel”。")
            else:
                messagebox.showinfo("提示", "当前没有可导出的差异结果")
            return
        output_path = filedialog.asksaveasfilename(
            title="导出差异报告",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile="差异比对结果.xlsx",
        )
        if not output_path:
            return
        try:
            if self._is_value_config_mode():
                export_value_compare_diffs_to_excel(self.compare_visible_diffs, output_path)
                self.compare_last_export_path = output_path
                summary_text = self._build_value_compare_result_summary_text(self.compare_all_diffs, getattr(self, "_last_value_compare_report", None) or type("_Report", (), {"reference_records": 0, "config_records": 0, "matched_records": 0})(), output_path)
                self.compare_summary_var.set(summary_text)
                self._set_value_result_summary_text(summary_text)
                try:
                    self.value_action_status_var.set("比对完成，结果已导出")
                except Exception:
                    pass
                try:
                    if getattr(self, "value_tabview", None) is not None and self.value_tabview.winfo_exists():
                        self.value_tabview.set("结果摘要")
                except Exception:
                    pass
                self._refresh_value_export_buttons()
                self._update_compare_summary_wrap()
                self._refresh_compare_scroll_region()
            else:
                export_diffs_to_excel(self.compare_visible_diffs, output_path)
            messagebox.showinfo("导出成功", f"已导出：{output_path}")
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))

    def _refresh_compare_button(self):
        state = "disabled" if self.compare_is_working else "normal"
        text = "比对中..." if self.compare_is_working else "开始比对"
        fg = COLOR_DANGER if self.compare_is_working else COLOR_ACCENT
        hover = COLOR_DANGER_HOVER if self.compare_is_working else COLOR_ACCENT_HOVER
        for button_name in ("compare_button", "value_compare_button"):
            button = getattr(self, button_name, None)
            try:
                if button is not None and button.winfo_exists():
                    button.configure(state=state, text=text, fg_color=fg, hover_color=hover)
            except Exception:
                pass
        try:
            if getattr(self, "value_action_status_var", None) is not None and self.compare_is_working:
                self.value_action_status_var.set("比对中...")
        except Exception:
            pass
        try:
            self._refresh_value_export_buttons()
        except Exception:
            pass


    # ---------- Localization check page ----------
    def _make_localization_card(self, parent, row=0, column=0, padx=0, pady=(0, 14)):
        card = ctk.CTkFrame(parent, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        card.grid(row=row, column=column, sticky="nsew", padx=padx, pady=pady)
        card.grid_columnconfigure(0, weight=1)
        return card

    def _sync_localization_status_text(self, *_):
        if hasattr(self, "loc_status_box") and self.loc_status_box is not None:
            try:
                if self.loc_status_box.winfo_exists():
                    self._set_textbox_text(self.loc_status_box, self.loc_summary_var.get())
                    self.loc_status_box.update_idletasks()
            except Exception:
                pass

    def _build_localization_page(self):
        # 多语言检查页改为固定 Tab 分区，避免在一个长滚动页面里堆叠大量控件造成拖影。
        for child in self.content.winfo_children():
            try:
                child.destroy()
            except Exception:
                pass

        wrap = ctk.CTkFrame(self.content, fg_color=COLOR_BG)
        wrap.grid(row=0, column=0, sticky="nsew")
        wrap.grid_columnconfigure(0, weight=1)
        wrap.grid_rowconfigure(0, weight=1)

        tabview = ctk.CTkTabview(
            wrap,
            fg_color=COLOR_SURFACE,
            segmented_button_fg_color=COLOR_SURFACE_2,
            segmented_button_selected_color=COLOR_ACCENT,
            segmented_button_selected_hover_color=COLOR_ACCENT_HOVER,
            segmented_button_unselected_color=COLOR_SURFACE_2,
            segmented_button_unselected_hover_color=COLOR_HOVER,
            text_color=COLOR_TEXT,
            corner_radius=16,
            border_width=1,
            border_color=COLOR_BORDER,
        )
        tabview.grid(row=0, column=0, sticky="nsew")

        basic_tab = tabview.add("基础配置")
        rule_tab = tabview.add("检查规则")
        result_tab = tabview.add("检查结果")
        for tab in (basic_tab, rule_tab, result_tab):
            tab.configure(fg_color=COLOR_BG)
            tab.grid_columnconfigure(0, weight=1)
            tab.grid_rowconfigure(0, weight=0)

        # 基础配置
        basic_tab.grid_rowconfigure(2, weight=1)
        file_card = self._make_localization_card(basic_tab, row=0)
        file_card.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(file_card, text="翻译文件", text_color=COLOR_TEXT, font=self._font(15, "bold")).grid(row=0, column=0, sticky="w", padx=(24, 22), pady=18)
        ctk.CTkEntry(file_card, textvariable=self.loc_file_path_var, state="readonly", height=38, corner_radius=10, fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13)).grid(row=0, column=1, sticky="ew", pady=18)
        ctk.CTkButton(file_card, text="选择文件", width=96, height=38, corner_radius=12, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color="#FFFFFF", font=self._font(14, "bold"), command=self.select_localization_file).grid(row=0, column=2, padx=(12, 8), pady=18)
        ctk.CTkButton(file_card, text="选择文件夹", width=108, height=38, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(14, "bold"), command=self.select_localization_folder).grid(row=0, column=3, padx=(0, 24), pady=18)
        self._divider(file_card, 1)
        ctk.CTkLabel(file_card, textvariable=self.loc_file_status_var, text_color=COLOR_MUTED, font=self._font(13), anchor="w", wraplength=860).grid(row=2, column=0, columnspan=4, sticky="ew", padx=24, pady=(0, 14))

        parse_card = self._make_localization_card(basic_tab, row=1)
        parse_card.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)
        self._small_entry(parse_card, 0, 0, "表头所在行", self.loc_header_row_var, "默认 1，多 Sheet 可自动识别")
        self._small_entry(parse_card, 0, 1, "数据起始行", self.loc_data_start_row_var, "默认 2")
        self._small_entry(parse_card, 0, 2, "主键列名", self.loc_key_column_var, "自动识别 ID / #tid / Text ID")
        self._small_entry(parse_card, 0, 3, "Sheet 名称", self.loc_sheet_var, "单 Sheet 使用；批量模式留空")
        structure_box = ctk.CTkFrame(parse_card, fg_color="transparent")
        structure_box.grid(row=0, column=4, sticky="ew", padx=(0, 24), pady=6)
        structure_box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(structure_box, text="表格结构", text_color=COLOR_MUTED, font=self._font(12), anchor="w").grid(row=0, column=0, sticky="ew")
        self.loc_structure_menu = ctk.CTkOptionMenu(structure_box, values=["自动识别", "单 Sheet 多语言列", "多 Sheet 翻译表", "多 Sheet 语言页", "多文件字典模式"], variable=self.loc_structure_var, width=170, height=34, fg_color=COLOR_SURFACE_2, button_color=COLOR_SELECTED, button_hover_color=COLOR_HOVER, text_color=COLOR_TEXT, font=self._font(12))
        self.loc_structure_menu.grid(row=1, column=0, sticky="ew", pady=(3, 0))
        action = ctk.CTkFrame(parse_card, fg_color="transparent")
        action.grid(row=1, column=0, columnspan=5, sticky="ew", padx=24, pady=(4, 8))
        ctk.CTkButton(action, text="识别结构/语言", width=135, height=36, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self.detect_localization_columns).pack(side="left", padx=(0, 10))
        ctk.CTkButton(action, text="全选语言", width=100, height=36, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=lambda: self._set_localization_languages("all")).pack(side="left", padx=(0, 10))
        ctk.CTkButton(action, text="清空语言", width=100, height=36, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=lambda: self._set_localization_languages("none")).pack(side="left", padx=(0, 10))
        ctk.CTkButton(action, text="多 Sheet 说明", width=110, height=36, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self.show_localization_multisheet_help).pack(side="left")
        self.loc_structure_preview_box = ctk.CTkTextbox(parse_card, height=96, wrap="word", fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER, border_width=1, text_color=COLOR_TEXT, font=self._font(12), corner_radius=12)
        self.loc_structure_preview_box.grid(row=2, column=0, columnspan=5, sticky="ew", padx=24, pady=(0, 14))
        self._set_textbox_text(self.loc_structure_preview_box, "结构预览：选择翻译文件或字典文件夹后点击“识别结构/语言”。支持单 Sheet、多版本 Sheet、多 Sheet 语言页，以及 dictionary_*.xlsx 多文件字典模式。")

        language_card = self._make_localization_card(basic_tab, row=2, pady=(0, 0))
        language_card.grid_rowconfigure(2, weight=1)
        header = ctk.CTkFrame(language_card, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=24, pady=(16, 8))
        header.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(header, text="源语言列", text_color=COLOR_TEXT, font=self._font(14, "bold")).grid(row=0, column=0, sticky="w", padx=(0, 12))
        self.loc_source_menu = ctk.CTkOptionMenu(header, values=self.loc_detected_headers or ["未识别"], variable=self.loc_source_language_var, width=220, fg_color=COLOR_SURFACE_2, button_color=COLOR_SELECTED, button_hover_color=COLOR_HOVER, text_color=COLOR_TEXT, font=self._font(13))
        self.loc_source_menu.grid(row=0, column=1, sticky="w")
        ctk.CTkLabel(language_card, text="参与检查语言列（点击“识别语言列”后自动生成，可手动勾选）", text_color=COLOR_MUTED, font=self._font(12), anchor="w").grid(row=1, column=0, sticky="ew", padx=24, pady=(0, 8))
        self.loc_language_frame = ctk.CTkFrame(language_card, fg_color=COLOR_SURFACE_2, corner_radius=12, height=170)
        self.loc_language_frame.grid(row=2, column=0, sticky="nsew", padx=24, pady=(0, 16))
        self.loc_language_frame.grid_columnconfigure(tuple(range(6)), weight=1)
        self._refresh_localization_language_checks([])

        # 检查规则
        rule_tab.grid_rowconfigure(1, weight=1)
        rule_card = self._make_localization_card(rule_tab, row=0)
        rule_card.grid_columnconfigure(0, weight=1)
        rule_top = ctk.CTkFrame(rule_card, fg_color="transparent")
        rule_top.grid(row=0, column=0, sticky="ew", padx=24, pady=(16, 8))
        for idx, (text, var) in enumerate([
            ("空翻译", self.loc_check_empty_var), ("残留中文", self.loc_check_chinese_var),
            ("占位符", self.loc_check_placeholders_var), ("标签", self.loc_check_tags_var),
            ("特殊符号", self.loc_check_symbols_var), ("长度", self.loc_check_length_var),
            ("重复 ID", self.loc_check_duplicate_id_var), ("无效 ID", self.loc_check_invalid_id_var),
            ("ID 缺失/多出", self.loc_check_missing_id_var), ("源文本一致", self.loc_check_source_consistency_var),
            ("数值一致", self.loc_check_numbers_var),
            ("译文复用分析", self.loc_check_duplicate_translation_var),
        ]):
            ctk.CTkCheckBox(rule_top, text=text, variable=var, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13)).grid(row=idx // 4, column=idx % 4, sticky="w", padx=(0, 18), pady=5)
        self.loc_rules_toggle_button = ctk.CTkButton(rule_card, text="展开高级规则", width=130, height=32, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(12, "bold"), command=self._toggle_localization_rules)
        self.loc_rules_toggle_button.grid(row=1, column=0, sticky="w", padx=24, pady=(0, 8))
        self.loc_rule_detail_frame = ctk.CTkFrame(rule_card, fg_color="transparent")
        self.loc_rule_detail_frame.grid(row=2, column=0, sticky="ew", padx=24, pady=(0, 16))
        self.loc_rule_detail_frame.grid_columnconfigure((0, 1), weight=1)
        self._small_entry(self.loc_rule_detail_frame, 0, 0, "允许包含中文的语言", self.loc_allowed_chinese_var, "逗号分隔")
        self._small_entry(self.loc_rule_detail_frame, 0, 1, "忽略 ID", self.loc_ignore_ids_var, "逗号或换行分隔")
        self._small_entry(self.loc_rule_detail_frame, 1, 0, "检查符号", self.loc_symbols_var, "例如 %, { }, \\n")
        self._small_entry(self.loc_rule_detail_frame, 1, 1, "最大字符数", self.loc_max_length_var, "0 = 不限制")
        self._small_entry(self.loc_rule_detail_frame, 2, 0, "源文长度倍率", self.loc_length_ratio_var, "0 = 不限制，例如 2")
        self._small_entry(self.loc_rule_detail_frame, 2, 1, "预期语言 Sheet", self.loc_expected_languages_var, "EN,DE,FR...；多 Sheet 模式用于检查缺失")
        self._small_entry(self.loc_rule_detail_frame, 3, 0, "文本类行类型", self.loc_text_row_types_var, "逗号分隔；用于空源文判断")
        self._small_entry(self.loc_rule_detail_frame, 3, 1, "非文本控制行类型", self.loc_non_text_row_types_var, "逗号分隔；源文和译文均空时忽略")
        self._small_entry(self.loc_rule_detail_frame, 4, 0, "待翻译完整标记", self.loc_unfinished_markers_var, "TODO/TBD/FIXME 使用完整边界")
        self.loc_rule_detail_frame.grid_remove()
        self.loc_rules_visible = False

        help_card = self._make_localization_card(rule_tab, row=1, pady=(0, 0))
        help_card.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(help_card, text="检查规则说明", text_color=COLOR_TEXT, font=self._font(15, "bold"), anchor="w").grid(row=0, column=0, sticky="ew", padx=18, pady=(14, 8))
        self.loc_rule_help_box = ctk.CTkTextbox(help_card, height=210, wrap="word", fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER, border_width=1, text_color=COLOR_MUTED, font=self._font(13), corner_radius=12)
        self.loc_rule_help_box.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 16))
        self._set_textbox_text(self.loc_rule_help_box, """空翻译：检查已勾选语言列是否为空。\n残留中文：检查非中文语言中是否残留中文字符，可通过“允许包含中文的语言”排除中文列。\n占位符：检查 %s、%d、{0}、{name}、\\n 等占位符是否缺失、多余或顺序异常。\n标签：检查 <color>、</color>、<b>、</b>、[b]、[/b] 等富文本标签是否成对。\n特殊符号：检查源文中的关键符号是否在译文中缺失。\n长度：按最大字符数或源文长度倍率检查文本过长风险。\n重复 ID / 无效 ID：检查主键重复或主键为空。
多 Sheet 语言页：支持 EN、DE、FR、JP 等语言分 Sheet 的表格，结果中会显示语言、Sheet 和行号。
ID 缺失/多出：以 EN 优先作为基准，检查其他语言页是否缺少或多出 ID。
源文本一致：检查同一 ID 在不同语言 Sheet 中的源文本是否一致。
数值一致：检查源文本与译文中的业务数值是否一致，会忽略 {0}、%1$s、[value1]、富文本标签和颜色值中的数字。
译文复用分析：默认关闭；开启后仅作为低置信度辅助信息，不计入错误总数。""")

        # 检查结果
        result_tab.grid_rowconfigure(1, weight=1)
        status_card = self._make_localization_card(result_tab, row=0)
        status_card.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(status_card, text="处理状态", text_color=COLOR_TEXT, font=self._font(15, "bold"), anchor="w").grid(row=0, column=0, sticky="ew", padx=18, pady=(14, 8))
        self.loc_status_box = ctk.CTkTextbox(status_card, height=74, wrap="word", fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER, border_width=1, text_color=COLOR_TEXT, font=self._font(13), corner_radius=12)
        self.loc_status_box.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 8))
        ctk.CTkLabel(status_card, textvariable=self.loc_progress_var, text_color=COLOR_MUTED, font=self._font(12), anchor="w").grid(row=2, column=0, sticky="ew", padx=18, pady=(0, 6))
        self.loc_progress_bar = ctk.CTkProgressBar(status_card, height=10, progress_color=COLOR_ACCENT, fg_color=COLOR_SURFACE_2)
        self.loc_progress_bar.grid(row=3, column=0, sticky="ew", padx=18, pady=(0, 14))
        self.loc_progress_bar.set(0)
        self._sync_localization_status_text()
        if not getattr(self, "loc_status_trace_bound", False):
            self.loc_summary_var.trace_add("write", self._sync_localization_status_text)
            self.loc_status_trace_bound = True

        result_card = ctk.CTkFrame(result_tab, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        result_card.grid(row=1, column=0, sticky="nsew", pady=(0, 0))
        self.loc_result_card = result_card
        result_card.grid_columnconfigure(0, weight=1)
        result_card.grid_rowconfigure(3, weight=1)

        result_action = ctk.CTkFrame(result_card, fg_color="transparent")
        result_action.grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 8))
        result_action.grid_columnconfigure(2, weight=1)
        self.loc_check_button = ctk.CTkButton(result_action, text="开始检查", width=130, height=40, corner_radius=12, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color="#FFFFFF", font=self._font(15, "bold"), command=self.start_localization_check)
        self.loc_check_button.grid(row=0, column=0, padx=(0, 10), sticky="w")
        self.loc_stop_button = ctk.CTkButton(result_action, text="停止", width=78, height=40, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(14, "bold"), command=self.stop_localization_check)
        self.loc_stop_button.grid(row=0, column=1, padx=(0, 12), sticky="w")
        search_entry = ctk.CTkEntry(result_action, textvariable=self.loc_search_var, placeholder_text="搜索 ID / 语言 / 问题类型 / 文本", height=40, corner_radius=10, fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(14))
        search_entry.grid(row=0, column=2, sticky="ew", padx=(0, 12))
        ctk.CTkButton(result_action, text="复制结果", width=104, height=40, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(14, "bold"), command=self.copy_localization_results).grid(row=0, column=3, padx=(0, 10))
        ctk.CTkButton(result_action, text="导出 Excel", width=112, height=40, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(14, "bold"), command=self.export_localization_excel).grid(row=0, column=4)

        self.loc_result_hint_label = ctk.CTkLabel(
            result_card,
            text="提示：占位符、特殊符号和长度倍率会基于源语言列进行检查；结果区按每页 200 条分页渲染，导出仍使用完整结果。双击任意结果行可查看完整源文本、当前文本、问题说明和建议处理。",
            text_color=COLOR_MUTED,
            font=self._font(13),
            anchor="w",
            justify="left",
            wraplength=760,
        )
        self.loc_result_hint_label.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 8))
        result_card.bind("<Configure>", lambda event: self._update_localization_result_hint_wrap(event.width))

        table_frame = ctk.CTkFrame(result_card, fg_color=COLOR_SURFACE_2, corner_radius=14, border_width=1, border_color=COLOR_BORDER)
        table_frame.grid(row=3, column=0, sticky="nsew", padx=18, pady=(0, 18))
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)
        columns = ("index", "level", "issue_type", "language", "sheet", "row", "item_id", "source", "current", "remark", "suggestion")
        self.loc_tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=16, style="Xiao.Treeview")
        headings = {"index": "序号", "level": "问题等级", "issue_type": "问题类型", "language": "语言", "sheet": "Sheet", "row": "行号", "item_id": "ID", "source": "源文本", "current": "当前文本", "remark": "问题说明", "suggestion": "建议处理"}
        widths = {"index": 58, "level": 100, "issue_type": 120, "language": 90, "sheet": 110, "row": 70, "item_id": 130, "source": 420, "current": 420, "remark": 460, "suggestion": 520}
        for col in columns:
            self.loc_tree.heading(col, text=headings[col])
            self.loc_tree.column(
                col,
                width=widths[col],
                minwidth=widths[col],
                anchor="center" if col in {"index", "level", "issue_type", "language", "sheet", "row"} else "w",
                stretch=False,
            )
        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.loc_tree.yview)
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.loc_tree.xview)
        self.loc_tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.loc_tree_issue_map = {}
        self.loc_tree.bind("<Double-1>", self.show_localization_issue_detail)
        self.loc_tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        pager = self._build_pager(
            result_card,
            self.loc_page_status_var,
            self.loc_page_jump_var,
            self.loc_prev_page,
            self.loc_next_page,
            self.loc_jump_page,
        )
        pager.grid(row=4, column=0, sticky="ew", padx=18, pady=(0, 14))

        if not self.loc_search_trace_bound:
            self.loc_search_var.trace_add("write", lambda *_: self.apply_localization_search_filter())
            self.loc_search_trace_bound = True
        self._refresh_localization_button()
        self.refresh_localization_table(self.loc_visible_issues)

    def _reset_localization_detection_state_for_new_file(self):
        """选择新的多语言文件后清理上一份文件的结构识别缓存，避免模式残留。"""
        try:
            self.loc_structure_var.set("自动识别")
        except Exception:
            pass
        self.loc_detected_headers = []
        self.loc_multi_table_candidates_cache = []
        self.loc_multi_table_cache_key = None
        self.loc_language_vars = {}
        try:
            self.loc_source_language_var.set("")
        except Exception:
            pass
        try:
            self.loc_search_var.set("")
        except Exception:
            pass
        self.loc_all_issues = []
        self.loc_visible_issues = []
        self.loc_page_index = 1
        self.loc_progress_var.set("等待开始")
        if self.loc_progress_bar is not None:
            try:
                self.loc_progress_bar.set(0)
            except Exception:
                pass
        if hasattr(self, "loc_source_menu") and self.loc_source_menu is not None:
            try:
                self.loc_source_menu.configure(values=["未识别"])
            except Exception:
                pass
        if hasattr(self, "loc_structure_preview_box") and self.loc_structure_preview_box is not None:
            try:
                self._set_textbox_text(
                    self.loc_structure_preview_box,
                    "结构预览：已切换新文件，表格结构已重置为“自动识别”。请点击“识别结构/语言”重新识别当前文件。"
                )
            except Exception:
                pass
        if hasattr(self, "loc_language_frame") and self.loc_language_frame is not None:
            try:
                self._refresh_localization_language_checks([])
            except Exception:
                pass
        if hasattr(self, "loc_tree") and self.loc_tree is not None:
            try:
                if self.loc_tree.winfo_exists():
                    self.clear_localization_table()
            except Exception:
                pass

    def _localization_multi_table_cache_key(self, path=None):
        return (
            os.path.abspath(path or self.loc_file_path_var.get().strip()),
            str(safe_int(self.loc_header_row_var.get(), 1)),
            str(safe_int(self.loc_data_start_row_var.get(), 2)),
            self.loc_key_column_var.get().strip(),
        )

    def select_localization_file(self):
        path = filedialog.askopenfilename(title="选择翻译文件", filetypes=[("翻译文件", "*.xlsx *.csv"), ("所有文件", "*.*")])
        if not path:
            return
        self.loc_file_path_var.set(path)
        self._reset_localization_detection_state_for_new_file()
        self.loc_file_status_var.set(f"已选择：{get_file_info(path)}")
        self.loc_summary_var.set("已选择新文件，表格结构已重置为自动识别，请点击“识别结构/语言”重新识别")

    def select_localization_folder(self):
        path = filedialog.askdirectory(title="选择多文件字典文件夹")
        if not path:
            return
        self.loc_file_path_var.set(path)
        self._reset_localization_detection_state_for_new_file()
        try:
            count = len([name for name in os.listdir(path) if name.lower().startswith("dictionary_") and name.lower().endswith(".xlsx")])
        except Exception:
            count = 0
        self.loc_structure_var.set("多文件字典模式" if count else "自动识别")
        self.loc_file_status_var.set(f"已选择文件夹：{path} / dictionary 文件 {count} 个")
        self.loc_summary_var.set("已选择字典文件夹，请点击“识别结构/语言”识别源语言和目标语言")

    def show_localization_multisheet_help(self):
        messagebox.showinfo(
            "多 Sheet 模式说明",
            "多 Sheet 翻译表：适用于一个 Excel 内有多个版本 Sheet，每个 Sheet 都是 Text ID + CN/EN/DE/FR... 多语言列。工具会批量检查所有符合结构的 Sheet，并在结果中保留 Sheet 名称。\n\n"
            "多 Sheet 语言页：适用于每个语言一个 Sheet 的翻译表，例如 EN / DE / FR / JP。每个语言 Sheet 需要包含 ID、源文本和翻译列。\n\n"
            "多文件字典模式：适用于 dictionary_ChineseSimplified.xlsx / dictionary_English.xlsx 这类一个语言一个文件的字典。工具会按 ID 对齐 Contents，忽略“所属模块”列检查，但会把模块写入报告。"
        )

    def detect_localization_columns(self):
        path = self.loc_file_path_var.get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择翻译文件")
            return
        try:
            mode = normalize_localization_mode(self.loc_structure_var.get().strip() or "自动识别")
            precomputed_multi_table = None
            if mode == MULTI_FILE_DICTIONARY_STRUCTURE:
                structure = MULTI_FILE_DICTIONARY_STRUCTURE
            elif mode == MULTI_TABLE_SHEETS_STRUCTURE:
                structure = MULTI_TABLE_SHEETS_STRUCTURE
            elif mode == MULTI_SHEET_STRUCTURE:
                structure = MULTI_SHEET_STRUCTURE
            elif mode == SINGLE_SHEET_STRUCTURE:
                structure = SINGLE_SHEET_STRUCTURE
            else:
                dictionary_files, _dictionary_columns = detect_dictionary_language_files(path)
                if len(dictionary_files) >= 2:
                    structure = MULTI_FILE_DICTIONARY_STRUCTURE
                else:
                    structure = detect_localization_structure(path)
                if structure == SINGLE_SHEET_STRUCTURE and not self.loc_sheet_var.get().strip():
                    auto_options = LocalizationCheckOptions(
                        header_row=safe_int(self.loc_header_row_var.get(), 1),
                        data_start_row=safe_int(self.loc_data_start_row_var.get(), 2),
                        key_column=self.loc_key_column_var.get().strip(),
                        sheet_name="",
                        table_structure="多 Sheet 翻译表",
                    )
                    precomputed_multi_table = detect_multi_table_sheet_columns(path, auto_options)
                    if len(precomputed_multi_table[0]) >= 2:
                        structure = MULTI_TABLE_SHEETS_STRUCTURE
                if structure == MULTI_FILE_DICTIONARY_STRUCTURE:
                    self.loc_structure_var.set("多文件字典模式")
                elif structure == MULTI_TABLE_SHEETS_STRUCTURE:
                    self.loc_structure_var.set("多 Sheet 翻译表")
                elif structure == MULTI_SHEET_STRUCTURE:
                    self.loc_structure_var.set("多 Sheet 语言页")
                elif structure == SINGLE_SHEET_STRUCTURE:
                    self.loc_structure_var.set("单 Sheet 多语言列")
                else:
                    raise FileReadError("未能自动识别多语言表结构，请手动选择「单 Sheet 多语言列」「多 Sheet 翻译表」「多 Sheet 语言页」或「多文件字典模式」")

            if structure == MULTI_FILE_DICTIONARY_STRUCTURE:
                dictionary_files, columns = detect_dictionary_language_files(path)
                if len(dictionary_files) < 2:
                    raise FileReadError("未识别到多文件字典：请选择包含 dictionary_*.xlsx 的文件夹")
                self.loc_detected_headers = [col.header for col in columns]
                self._refresh_localization_language_checks(columns)
                source = self.loc_source_language_var.get().strip()
                if source not in self.loc_detected_headers:
                    source = "CN" if "CN" in self.loc_detected_headers else ("EN" if "EN" in self.loc_detected_headers else self.loc_detected_headers[0])
                self.loc_source_language_var.set(source)
                if hasattr(self, "loc_source_menu"):
                    self.loc_source_menu.configure(values=self.loc_detected_headers or ["未识别"])
                preview_lines = [
                    f"结构：多文件字典模式（识别到 {len(dictionary_files)} 份 dictionary_*.xlsx）",
                    f"当前源语言：{source}（可在“源语言列”下拉框切换，例如 CN 或 EN）",
                    "字段规则：B列 ID；C列 所属模块（仅报告定位，不参与文本检查）；D列 Contents；第5行开始为数据。",
                    "",
                    "语言文件：",
                ]
                for info in dictionary_files:
                    marker = " ← 当前源语言" if info.code == source else ""
                    preview_lines.append(f"- {info.code} / {info.name}：{info.file_name}{marker}")
                if hasattr(self, "loc_structure_preview_box"):
                    self._set_textbox_text(self.loc_structure_preview_box, "\n".join(preview_lines[:50]))
                self.loc_file_status_var.set(f"识别成功：多文件字典模式 / 文件 {len(dictionary_files)} 个 / 语言 {', '.join(self.loc_detected_headers)}")
                self.loc_summary_var.set("多文件字典识别完成，可在源语言下拉框选择 CN 或 EN 后开始检查")
                return

            if structure == MULTI_TABLE_SHEETS_STRUCTURE:
                base_options = LocalizationCheckOptions(
                    header_row=safe_int(self.loc_header_row_var.get(), 1),
                    data_start_row=safe_int(self.loc_data_start_row_var.get(), 2),
                    key_column=self.loc_key_column_var.get().strip(),
                    sheet_name="",
                    table_structure="多 Sheet 翻译表",
                )
                if precomputed_multi_table is not None:
                    candidates, columns = precomputed_multi_table
                else:
                    candidates, columns = detect_multi_table_sheet_columns(path, base_options)
                if not candidates:
                    raise FileReadError("未识别到可批量检查的 Sheet：请确认每个版本 Sheet 都包含 Text ID、CN 原文和至少 2 个语言翻译列")
                self.loc_multi_table_candidates_cache = list(candidates)
                self.loc_multi_table_cache_key = self._localization_multi_table_cache_key(path)
                preview_lines = [f"结构：多 Sheet 翻译表（将批量检查 {len(candidates)} 个 Sheet）", ""]
                for sheet_name, score, reason in candidates:
                    preview_lines.append(f"- {sheet_name}：评分 {score}；{reason}")
                self.loc_detected_headers = [col.header for col in columns]
                self._refresh_localization_language_checks(columns)
                source = "CN" if "CN" in self.loc_detected_headers else auto_detect_source_language([col.header for col in columns], self.loc_detected_headers)
                self.loc_source_language_var.set(source or "未识别")
                if hasattr(self, "loc_source_menu"):
                    values = self.loc_detected_headers or ["未识别"]
                    self.loc_source_menu.configure(values=values)
                preview_lines.insert(1, f"语言并集：{', '.join(self.loc_detected_headers) or '未识别'}")
                preview_lines.insert(2, "提示：预览阶段只读取每个 Sheet 前 80 行；完整记录数会在开始检查后统计。")
                if hasattr(self, "loc_structure_preview_box"):
                    self._set_textbox_text(self.loc_structure_preview_box, "\n".join(preview_lines[:40]))
                self.loc_file_status_var.set(f"识别成功：多 Sheet 翻译表 / Sheet {len(candidates)} 个 / 语言并集 {len(columns)} 个")
                recommendation = f"；已按文件名推荐 {', '.join(self.loc_recommended_languages)}，最终以勾选为准" if self.loc_recommended_languages else ""
                self.loc_summary_var.set(f"多 Sheet 翻译表识别完成，可开始批量检查{recommendation}")
                return

            if structure == MULTI_SHEET_STRUCTURE:
                enabled = []
                dataset = parse_multi_sheet_localization(path, safe_int(self.loc_header_row_var.get(), 0), enabled)
                columns = [LocalizationColumn(name=info.language, header=info.sheet_name, index=idx) for idx, info in enumerate(dataset.sheet_infos) if not info.error]
                self.loc_detected_headers = [col.header for col in columns]
                self._refresh_localization_language_checks(columns)
                self.loc_source_language_var.set("每个 Sheet 的源文本列")
                if hasattr(self, "loc_source_menu"):
                    self.loc_source_menu.configure(values=["每个 Sheet 的源文本列"])
                preview = build_multi_sheet_preview(dataset)
                if hasattr(self, "loc_structure_preview_box"):
                    self._set_textbox_text(self.loc_structure_preview_box, preview)
                self.loc_file_status_var.set(f"识别成功：多 Sheet 语言页 / 语言页 {len(columns)} 个 / 文本 ID {len(dataset.text_ids)} 个 / 翻译条目 {len(dataset.entries)} 条")
                self.loc_summary_var.set("多 Sheet 语言页识别完成，可开始检查")
                return

            options = LocalizationCheckOptions(
                header_row=safe_int(self.loc_header_row_var.get(), 1),
                data_start_row=safe_int(self.loc_data_start_row_var.get(), 2),
                key_column=self.loc_key_column_var.get().strip(),
                sheet_name=self.loc_sheet_var.get().strip(),
                table_structure="单 Sheet 多语言列",
            )
            table = parse_localization_table(options, path)
            columns = detect_language_columns(table.headers, table.key_col_index, table.display_headers, table.type_headers)
            self.loc_detected_headers = [col.header for col in columns]
            self._refresh_localization_language_checks(columns)
            source = auto_detect_source_language(table.headers, self.loc_detected_headers)
            self.loc_source_language_var.set(source or "未识别")
            if hasattr(self, "loc_source_menu"):
                values = self.loc_detected_headers or ["未识别"]
                self.loc_source_menu.configure(values=values)
            if hasattr(self, "loc_structure_preview_box"):
                self._set_textbox_text(self.loc_structure_preview_box, build_single_sheet_preview(table, columns))
            self.loc_file_status_var.set(f"识别成功：Sheet={table.sheet_name} / {table.structure_type} / 表头第 {table.header_row} 行 / 数据第 {table.data_start_row} 行 / 记录 {len(table.records)} 条 / 语言列 {len(columns)} 个")
            recommendation = f"；已按文件名推荐 {', '.join(self.loc_recommended_languages)}，最终以勾选为准" if self.loc_recommended_languages else ""
            self.loc_summary_var.set(f"语言列识别完成，可开始检查{recommendation}")
        except Exception as exc:
            messagebox.showerror("识别失败", str(exc))

    def _refresh_localization_language_checks(self, columns):
        if self.loc_language_frame is None:
            return
        for child in self.loc_language_frame.winfo_children():
            child.destroy()
        self.loc_language_vars = {}
        if not columns:
            ctk.CTkLabel(self.loc_language_frame, text="尚未识别语言列。选择文件后点击“识别语言列”。", text_color=COLOR_MUTED, font=self._font(13)).grid(row=0, column=0, sticky="w", pady=4)
            return
        recommended = recommend_target_languages_from_filename(self.loc_file_path_var.get().strip(), columns)
        self.loc_recommended_languages = list(recommended)
        recommended_set = set(recommended)
        for idx, col in enumerate(columns):
            code = str(getattr(col, "code", "") or getattr(col, "header", "") or "").upper()
            var = tk.BooleanVar(value=(not recommended_set or code in recommended_set))
            self.loc_language_vars[col.header] = var
            text = col.name if col.name == col.header else f"{col.name}（{col.header}）"
            ctk.CTkCheckBox(self.loc_language_frame, text=text, variable=var, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(12)).grid(row=idx // 6, column=idx % 6, sticky="w", padx=(0, 8), pady=4)
        if recommended:
            self.loc_summary_var.set(f"已根据文件名推荐本次检查语言：{', '.join(recommended)}；最终范围以当前勾选为准")

    def _set_localization_languages(self, mode):
        for var in self.loc_language_vars.values():
            var.set(mode == "all")

    def _toggle_localization_rules(self):
        if self.loc_rule_detail_frame is None:
            return
        self.loc_rules_visible = not self.loc_rules_visible
        if self.loc_rules_visible:
            self.loc_rule_detail_frame.grid()
            if self.loc_rules_toggle_button is not None:
                self.loc_rules_toggle_button.configure(text="收起高级规则")
        else:
            self.loc_rule_detail_frame.grid_remove()
            if self.loc_rules_toggle_button is not None:
                self.loc_rules_toggle_button.configure(text="展开高级规则")

    def get_localization_options(self):
        enabled = [header for header, var in self.loc_language_vars.items() if var.get()]
        try:
            ratio = float(self.loc_length_ratio_var.get().strip() or 0)
        except Exception:
            ratio = 0.0
        options = LocalizationCheckOptions(
            header_row=safe_int(self.loc_header_row_var.get(), 1),
            data_start_row=safe_int(self.loc_data_start_row_var.get(), 2),
            key_column=self.loc_key_column_var.get().strip(),
            sheet_name=self.loc_sheet_var.get().strip(),
            source_language_column=self.loc_source_language_var.get().strip() if self.loc_source_language_var.get().strip() != "未识别" else "",
            enabled_language_columns=enabled,
            ignore_ids_text=self.loc_ignore_ids_var.get(),
            allowed_chinese_languages_text=self.loc_allowed_chinese_var.get(),
            symbols_text=self.loc_symbols_var.get(),
            max_length=safe_int(self.loc_max_length_var.get(), 0) if self.loc_max_length_var.get().strip() else 0,
            length_ratio=ratio,
            check_empty=self.loc_check_empty_var.get(),
            check_chinese=self.loc_check_chinese_var.get(),
            check_placeholders=self.loc_check_placeholders_var.get(),
            check_tags=self.loc_check_tags_var.get(),
            check_symbols=self.loc_check_symbols_var.get(),
            check_length=self.loc_check_length_var.get(),
            check_duplicate_id=self.loc_check_duplicate_id_var.get(),
            check_invalid_id=self.loc_check_invalid_id_var.get(),
            ui_limit=UI_MAX_RENDER_ROWS,
            table_structure=self.loc_structure_var.get(),
            expected_languages_text=self.loc_expected_languages_var.get(),
            check_missing_id=self.loc_check_missing_id_var.get(),
            check_source_consistency=self.loc_check_source_consistency_var.get(),
            check_numbers=self.loc_check_numbers_var.get(),
            check_duplicate_translation=self.loc_check_duplicate_translation_var.get(),
            text_row_types_text=self.loc_text_row_types_var.get(),
            non_text_row_types_text=self.loc_non_text_row_types_var.get(),
            unfinished_markers_text=self.loc_unfinished_markers_var.get(),
        )
        if (
            normalize_localization_mode(options.table_structure or "自动识别") == MULTI_TABLE_SHEETS_STRUCTURE
            and self.loc_multi_table_candidates_cache
            and self.loc_multi_table_cache_key == self._localization_multi_table_cache_key(self.loc_file_path_var.get().strip())
        ):
            options.precomputed_multi_table_sheets = list(self.loc_multi_table_candidates_cache)
        return options

    def start_localization_check(self):
        if self.loc_is_working:
            return
        path = self.loc_file_path_var.get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择翻译文件")
            return
        self.loc_is_working = True
        self.loc_stop_event.clear()
        self._refresh_localization_button()
        self._set_badge(True)
        self._update_localization_progress({"message": "准备开始多语言检查…", "ratio": 0})
        self.loc_summary_var.set("正在检查多语言文本，请稍候...\n当前正在读取翻译文件并执行空翻译、残留中文、未完成占位文本、占位符、标签、特殊符号、数值一致、拼写疑似、术语一致性和译文重复等规则检查；数值顺序变化、JP 正常汉字、UI 方括号术语会按疑似/兼容规则降级。")
        self.after(50, self.update_idletasks)
        self.clear_localization_table()
        options = self.get_localization_options()
        threading.Thread(target=self._localization_worker, args=(path, options), daemon=True).start()

    def stop_localization_check(self):
        if self.loc_is_working:
            self.loc_stop_event.set()
            self.loc_summary_var.set("正在请求停止检查，请稍候...")
            self._update_localization_progress({"message": "正在请求停止检查，请稍候…"})
            self.after(50, self.update_idletasks)

    def _update_localization_progress(self, payload):
        try:
            message = str(payload.get("message") or "")
            if not message:
                current = int(payload.get("current") or 0)
                total = int(payload.get("total") or 0)
                message = f"处理中：{current}/{total}" if total else "处理中…"
            ratio = payload.get("ratio", None)
            if ratio is not None and self.loc_progress_bar is not None and self.loc_progress_bar.winfo_exists():
                self.loc_progress_bar.set(min(max(float(ratio), 0.0), 1.0))
            self.loc_progress_var.set(message)
        except Exception:
            pass

    def _localization_worker(self, path, options):
        try:
            def progress(payload):
                self.after(0, lambda data=dict(payload): self._update_localization_progress(data))

            issues, table, columns = check_localization_file(path, options, self.loc_stop_event, progress)
            self.after(0, lambda: self._on_localization_success(issues, table, columns))
        except FileReadError as exc:
            self.after(0, lambda msg=str(exc): self._on_localization_error(msg))
        except Exception as exc:
            self.after(0, lambda msg=f"多语言检查失败：{exc}": self._on_localization_error(msg))

    def _on_localization_success(self, issues, table, columns):
        self.loc_is_working = False
        self._refresh_localization_button()
        self._set_badge(False)
        self._update_localization_progress({"message": "检查完成，结果已刷新", "ratio": 1})
        self.loc_all_issues = issues
        self.loc_visible_issues = issues
        if columns and not self.loc_language_vars:
            self._refresh_localization_language_checks(columns)
        extra = ""
        try:
            extra = f" / {table.structure_type} / 表头第 {table.header_row} 行 / 数据第 {table.data_start_row} 行"
        except Exception:
            extra = ""
        self.loc_file_status_var.set(f"读取成功：Sheet={table.sheet_name}{extra} / 记录 {len(table.records)} 条 / 语言列 {len(columns)} 个")
        self.apply_localization_search_filter()
        counts = {}
        level_counts = {}
        category_counts = {}
        for issue in issues:
            counts[issue.issue_type] = counts.get(issue.issue_type, 0) + 1
            level = getattr(issue, "issue_level", "普通问题")
            level_counts[level] = level_counts.get(level, 0) + 1
            category = getattr(issue, "result_category", "明确问题")
            category_counts[category] = category_counts.get(category, 0) + 1
        if not issues:
            self.loc_summary_var.set("检查完成：未发现问题")
        else:
            level_top = " | ".join(f"{k} {v}" for k, v in category_counts.items())
            top = " | ".join(f"{k} {v}" for k, v in sorted(counts.items())[:8])
            self.loc_summary_var.set(f"检查完成：共 {len(issues)} 条问题 | {level_top} | {top}")
        self.after(50, self.update_idletasks)

    def _on_localization_error(self, message):
        self.loc_is_working = False
        self._refresh_localization_button()
        self._set_badge(False)
        self._update_localization_progress({"message": "检查失败，请查看错误提示", "ratio": 0})
        self.loc_summary_var.set(f"检查失败：{message}")
        self.after(50, self.update_idletasks)
        messagebox.showerror("检查失败", message)

    def apply_localization_search_filter(self):
        keyword = self.loc_search_var.get().strip().lower()
        if not keyword:
            self.loc_visible_issues = self.loc_all_issues
        else:
            self.loc_visible_issues = [item for item in self.loc_all_issues if keyword in self._localization_issue_to_text(item).lower()]
        self.loc_page_index = 1
        self.refresh_localization_table(self.loc_visible_issues)

    def refresh_localization_table(self, issues):
        if not hasattr(self, "loc_tree") or not self.loc_tree.winfo_exists():
            return
        self.clear_localization_table()
        self.loc_page_index, total_pages, start, end = self._get_page_bounds(len(issues), self.loc_page_index)
        self.loc_page_jump_var.set(str(self.loc_page_index))
        self.loc_page_status_var.set(f"第 {self.loc_page_index}/{total_pages} 页，每页 {UI_PAGE_SIZE} 条，当前筛选 {len(issues)} 条")
        self.loc_tree_issue_map = {}
        for item in list(issues)[start:end]:
            tree_item = self.loc_tree.insert("", "end", values=(
                item.index,
                getattr(item, "issue_level", "普通问题"),
                item.issue_type,
                item.language,
                item.sheet_name,
                item.row_number,
                item.item_id,
                self._shorten_compare_text(item.source_text, 500),
                self._shorten_compare_text(item.current_text, 500),
                self._shorten_compare_text(item.remark, 500),
                self._shorten_compare_text(item.suggestion, 500),
            ))
            self.loc_tree_issue_map[tree_item] = item

    def loc_prev_page(self):
        self.loc_page_index = max(1, self.loc_page_index - 1)
        self.refresh_localization_table(self.loc_visible_issues)

    def loc_next_page(self):
        self.loc_page_index += 1
        self.refresh_localization_table(self.loc_visible_issues)

    def loc_jump_page(self):
        try:
            self.loc_page_index = int(self.loc_page_jump_var.get().strip() or "1")
        except Exception:
            self.loc_page_index = 1
        self.refresh_localization_table(self.loc_visible_issues)

    def clear_localization_table(self):
        if not hasattr(self, "loc_tree") or not self.loc_tree.winfo_exists():
            return
        for item in self.loc_tree.get_children():
            self.loc_tree.delete(item)
        self.loc_tree_issue_map = {}

    def show_localization_issue_detail(self, event=None):
        if not hasattr(self, "loc_tree") or not self.loc_tree.winfo_exists():
            return
        tree_item = self.loc_tree.focus()
        issue = getattr(self, "loc_tree_issue_map", {}).get(tree_item)
        if issue is None:
            return
        detail = (
            f"序号：{issue.index}\n"
            f"问题等级：{getattr(issue, 'issue_level', '普通问题')}\n"
            f"问题类型：{issue.issue_type}\n"
            f"ID：{issue.item_id}\n"
            f"语言：{issue.language}\n"
            f"行号：{issue.row_number}\n\n"
            f"【源文本】\n{issue.source_text or ''}\n\n"
            f"【当前文本】\n{issue.current_text or ''}\n\n"
            f"【命中规则】\n{getattr(issue, 'rule_name', issue.issue_type) or ''}\n\n"
            f"【问题说明】\n{issue.remark or ''}\n\n"
            f"【建议处理】\n{issue.suggestion or ''}\n\n"
            f"【是否疑似误报】\n{'是' if getattr(issue, 'is_false_positive', False) else '否'}"
        )
        top = ctk.CTkToplevel(self)
        top.title("多语言检查结果详情")
        top.geometry("920x620")
        top.minsize(760, 480)
        top.configure(fg_color=COLOR_BG)
        try:
            top.transient(self)
            top.grab_set()
        except Exception:
            pass
        top.grid_columnconfigure(0, weight=1)
        top.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(top, text="多语言检查结果详情", text_color=COLOR_TEXT, font=self._font(18, "bold"), anchor="w").grid(row=0, column=0, sticky="ew", padx=18, pady=(16, 8))
        box = ctk.CTkTextbox(top, wrap="word", fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER, border_width=1, text_color=COLOR_TEXT, font=self._font(13), corner_radius=12)
        box.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 12))
        box.insert("end", detail)
        box.configure(state="disabled")
        buttons = ctk.CTkFrame(top, fg_color="transparent")
        buttons.grid(row=2, column=0, sticky="ew", padx=18, pady=(0, 16))
        buttons.grid_columnconfigure(0, weight=1)
        def copy_detail():
            self.clipboard_clear()
            self.clipboard_append(detail)
            messagebox.showinfo("已复制", "已复制当前问题详情")
        ctk.CTkButton(buttons, text="复制详情", width=110, height=34, corner_radius=12, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color="#FFFFFF", font=self._font(13, "bold"), command=copy_detail).grid(row=0, column=1, padx=(0, 10))
        ctk.CTkButton(buttons, text="关闭", width=90, height=34, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=top.destroy).grid(row=0, column=2)

    @staticmethod
    def _localization_issue_to_text(item):
        return "\t".join([str(item.index), getattr(item, "issue_level", "普通问题"), item.issue_type, item.item_id, item.language, item.row_number, item.source_text, item.current_text, item.remark, item.suggestion])

    def copy_localization_results(self):
        if not self.loc_visible_issues:
            messagebox.showinfo("提示", "当前没有可复制的检查结果")
            return
        text = localization_issues_to_tsv(self.loc_visible_issues)
        self.clipboard_clear()
        self.clipboard_append(text)
        messagebox.showinfo("已复制", f"已复制 {len(self.loc_visible_issues)} 条检查结果")

    def export_localization_excel(self):
        if not self.loc_visible_issues:
            messagebox.showinfo("提示", "当前没有可导出的检查结果")
            return
        output_path = filedialog.asksaveasfilename(
            title="导出多语言检查报告",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile="多语言文本质量检查报告.xlsx",
        )
        if not output_path:
            return
        try:
            export_localization_issues_to_excel(self.loc_visible_issues, output_path)
            messagebox.showinfo("导出成功", f"已导出：{output_path}")
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))

    def _refresh_localization_button(self):
        if hasattr(self, "loc_check_button") and self.loc_check_button.winfo_exists():
            self.loc_check_button.configure(
                state="disabled" if self.loc_is_working else "normal",
                text="检查中..." if self.loc_is_working else "开始检查",
                fg_color=COLOR_DANGER if self.loc_is_working else COLOR_ACCENT,
                hover_color=COLOR_DANGER_HOVER if self.loc_is_working else COLOR_ACCENT_HOVER,
            )
        if hasattr(self, "loc_stop_button") and self.loc_stop_button.winfo_exists():
            self.loc_stop_button.configure(state="normal" if self.loc_is_working else "disabled")



    # ---------- Config validator page ----------
    def _build_config_validator_page(self):
        # 配置表校验改为“场景化字段校验”入口，减少抽象规则配置和整页滚动。
        for child in self.content.winfo_children():
            try:
                child.destroy()
            except Exception:
                pass
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(0, weight=1)

        wrap = ctk.CTkFrame(self.content, fg_color=COLOR_BG)
        wrap.grid(row=0, column=0, sticky="nsew")
        wrap.grid_columnconfigure(0, weight=1)
        wrap.grid_rowconfigure(0, weight=1)

        self.cfg_tabview = ctk.CTkTabview(
            wrap,
            fg_color=COLOR_BG,
            segmented_button_fg_color=COLOR_SURFACE_2,
            segmented_button_selected_color=COLOR_ACCENT,
            segmented_button_selected_hover_color=COLOR_ACCENT_HOVER,
            segmented_button_unselected_color=COLOR_SURFACE_2,
            segmented_button_unselected_hover_color=COLOR_HOVER,
            text_color=COLOR_TEXT,
            corner_radius=12,
        )
        self.cfg_tabview.grid(row=0, column=0, sticky="nsew")

        basic_tab = self.cfg_tabview.add("基础配置")
        quick_tab = self.cfg_tabview.add("快捷校验")
        type_tab = self.cfg_tabview.add("类型行自动校验")
        compound_tab = self.cfg_tabview.add("复合字段校验")
        advanced_tab = self.cfg_tabview.add("高级规则")
        switch_tab = self.cfg_tabview.add("检查开关")
        result_tab = self.cfg_tabview.add("校验结果")
        for tab in (basic_tab, quick_tab, type_tab, compound_tab, advanced_tab, switch_tab, result_tab):
            tab.configure(fg_color=COLOR_BG)
            tab.grid_columnconfigure(0, weight=1)

        # 基础配置
        file_card = self._card(basic_tab)
        file_card.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(file_card, text="配置表文件", text_color=COLOR_TEXT, font=self._font(15, "bold")).grid(row=0, column=0, sticky="w", padx=(24, 22), pady=18)
        ctk.CTkEntry(file_card, textvariable=self.cfg_file_path_var, state="readonly", height=38, corner_radius=10, fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13)).grid(row=0, column=1, sticky="ew", pady=18)
        ctk.CTkButton(file_card, text="选择配置表", width=120, height=38, corner_radius=12, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color="#FFFFFF", font=self._font(14, "bold"), command=self.select_config_file).grid(row=0, column=2, padx=(12, 24), pady=18)
        ctk.CTkLabel(file_card, textvariable=self.cfg_file_status_var, text_color=COLOR_MUTED, font=self._font(13), anchor="w", wraplength=980, justify="left").grid(row=1, column=0, columnspan=3, sticky="ew", padx=24, pady=(0, 14))

        parse_card = self._card(basic_tab)
        parse_card.grid_columnconfigure((0, 1, 2, 3, 4), weight=1, uniform="cfg_basic")
        self._small_entry(parse_card, 0, 0, "说明行", self.cfg_description_row_var, "通常 1")
        self._small_entry(parse_card, 0, 1, "类型行", self.cfg_type_row_var, "通常 2")
        self._small_entry(parse_card, 0, 2, "字段名行", self.cfg_header_row_var, "通常 3")
        self._small_entry(parse_card, 0, 3, "数据起始行", self.cfg_data_start_row_var, "通常 4")
        self._small_entry(parse_card, 0, 4, "主键列名", self.cfg_key_column_var, "自动识别 #id")
        self._small_entry(parse_card, 1, 0, "Sheet 名称", self.cfg_sheet_var, "默认第一个 Sheet")
        action = ctk.CTkFrame(parse_card, fg_color="transparent")
        action.grid(row=1, column=1, columnspan=4, sticky="ew", padx=0, pady=(24, 16))
        ctk.CTkButton(action, text="自动识别表结构", width=130, height=36, corner_radius=12, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color="#FFFFFF", font=self._font(13, "bold"), command=self.detect_config_structure_for_current_file).pack(side="left", padx=(0, 10))
        ctk.CTkButton(action, text="预览字段", width=100, height=36, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self.preview_config_fields).pack(side="left", padx=(0, 10))
        ctk.CTkButton(action, text="保存规则", width=100, height=36, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self.save_config_rules).pack(side="left", padx=(0, 10))
        ctk.CTkButton(action, text="加载规则", width=100, height=36, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self.load_config_rules).pack(side="left", padx=(0, 10))
        ctk.CTkButton(action, text="清空规则", width=100, height=36, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self.clear_config_rules).pack(side="left")
        detect_box = ctk.CTkTextbox(parse_card, height=118, corner_radius=12, border_width=1, border_color=COLOR_BORDER, fg_color=COLOR_SURFACE_2, text_color=COLOR_TEXT, font=self._font(12), wrap="word")
        detect_box.grid(row=2, column=0, columnspan=5, sticky="ew", padx=24, pady=(0, 16))
        self.cfg_detect_box = detect_box
        self._set_textbox_text(self.cfg_detect_box, self.cfg_detect_summary_var.get())

        # 快捷校验
        quick_card = self._card(quick_tab)
        quick_card.grid_columnconfigure((0, 1, 2), weight=1, uniform="quick_cfg")
        ctk.CTkLabel(quick_card, text="快捷字段校验", text_color=COLOR_TEXT, font=self._font(17, "bold"), anchor="w").grid(row=0, column=0, columnspan=3, sticky="ew", padx=24, pady=(18, 8))
        ctk.CTkLabel(quick_card, text="先选字段，再选模板；工具会自动转换为底层规则执行。适合 restrict、reward、必填字段等高频检查。", text_color=COLOR_MUTED, font=self._font(12), anchor="w").grid(row=1, column=0, columnspan=3, sticky="ew", padx=24, pady=(0, 10))
        field_box = ctk.CTkFrame(quick_card, fg_color="transparent")
        field_box.grid(row=2, column=0, sticky="ew", padx=(24, 12), pady=6)
        field_box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(field_box, text="字段名", text_color=COLOR_MUTED, font=self._font(12), anchor="w").grid(row=0, column=0, sticky="ew")
        self.cfg_quick_field_menu = ctk.CTkOptionMenu(field_box, variable=self.cfg_quick_field_var, values=[""], height=34, fg_color=COLOR_SURFACE_2, button_color=COLOR_SURFACE_2, button_hover_color=COLOR_HOVER, text_color=COLOR_TEXT, font=self._font(12), command=lambda *_: self.update_config_quick_preview())
        self.cfg_quick_field_menu.grid(row=1, column=0, sticky="ew", pady=(3, 0))
        tpl_box = ctk.CTkFrame(quick_card, fg_color="transparent")
        tpl_box.grid(row=2, column=1, sticky="ew", padx=(0, 12), pady=6)
        tpl_box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(tpl_box, text="规则模板", text_color=COLOR_MUTED, font=self._font(12), anchor="w").grid(row=0, column=0, sticky="ew")
        self.cfg_quick_template_menu = ctk.CTkOptionMenu(tpl_box, variable=self.cfg_quick_template_var, values=QUICK_TEMPLATES, height=34, fg_color=COLOR_SURFACE_2, button_color=COLOR_SURFACE_2, button_hover_color=COLOR_HOVER, text_color=COLOR_TEXT, font=self._font(12), command=lambda *_: self.update_config_quick_preview())
        self.cfg_quick_template_menu.grid(row=1, column=0, sticky="ew", pady=(3, 0))
        opt_box = ctk.CTkFrame(quick_card, fg_color="transparent")
        opt_box.grid(row=2, column=2, sticky="ew", padx=(0, 24), pady=6)
        ctk.CTkCheckBox(opt_box, text="0 视为无效", variable=self.cfg_quick_required_no_zero_var, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(12), command=self.update_config_quick_preview).grid(row=0, column=0, sticky="w", pady=(0, 4))
        ctk.CTkCheckBox(opt_box, text="-1 视为无效", variable=self.cfg_quick_required_no_minus_one_var, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(12), command=self.update_config_quick_preview).grid(row=1, column=0, sticky="w")
        self.cfg_quick_preview_box = ctk.CTkTextbox(quick_card, height=150, corner_radius=12, border_width=1, border_color=COLOR_BORDER, fg_color=COLOR_SURFACE_2, text_color=COLOR_TEXT, font=self._font(12), wrap="word")
        self.cfg_quick_preview_box.grid(row=3, column=0, columnspan=3, sticky="ew", padx=24, pady=(10, 18))
        self._set_textbox_text(self.cfg_quick_preview_box, self.cfg_quick_preview_var.get())

        # 类型行自动校验
        type_card = self._card(type_tab)
        type_card.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(type_card, text="类型行自动校验", text_color=COLOR_TEXT, font=self._font(17, "bold"), anchor="w").grid(row=0, column=0, sticky="ew", padx=24, pady=(18, 8))
        ctk.CTkCheckBox(type_card, text="启用类型行自动校验：uint/int/float/bool 自动检查空值和格式，uint[]/int[] 自动检查数组元素", variable=self.cfg_auto_type_enabled_var, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13)).grid(row=1, column=0, sticky="w", padx=24, pady=(0, 10))
        type_help = ctk.CTkTextbox(type_card, height=180, corner_radius=12, border_width=1, border_color=COLOR_BORDER, fg_color=COLOR_SURFACE_2, text_color=COLOR_MUTED, font=self._font(12), wrap="word")
        type_help.grid(row=2, column=0, sticky="ew", padx=24, pady=(0, 18))
        type_help.insert("1.0", "适用场景：配置表第 2 行为类型行，第 3 行为程序字段名行。\n\n示例：chessboard_stage 表中 attack_attribute / defence_attribute / health_attribute 类型为 uint，如果数据行为空，会输出：uint 类型字段不能为空。\n\n说明：string 默认允许为空；uint/int/float/bool 默认不能为空；uint[]/int[] 为空不报错，不为空时检查数组元素类型。")
        type_help.configure(state="disabled")

        # 复合字段校验
        comp_card = self._card(compound_tab)
        comp_card.grid_columnconfigure((0, 1, 2, 3), weight=1, uniform="compound_cfg")
        ctk.CTkLabel(comp_card, text="复合字段参数个数校验", text_color=COLOR_TEXT, font=self._font(17, "bold"), anchor="w").grid(row=0, column=0, columnspan=4, sticky="ew", padx=24, pady=(18, 8))
        ctk.CTkCheckBox(comp_card, text="启用复合字段校验", variable=self.cfg_compound_enabled_var, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13)).grid(row=1, column=0, sticky="w", padx=24, pady=(0, 8))
        comp_field_box = ctk.CTkFrame(comp_card, fg_color="transparent")
        comp_field_box.grid(row=2, column=0, sticky="ew", padx=(24, 12), pady=5)
        comp_field_box.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(comp_field_box, text="字段名", text_color=COLOR_MUTED, font=self._font(12), anchor="w").grid(row=0, column=0, sticky="ew")
        self.cfg_compound_field_menu = ctk.CTkOptionMenu(comp_field_box, variable=self.cfg_compound_field_var, values=[""], height=34, fg_color=COLOR_SURFACE_2, button_color=COLOR_SURFACE_2, button_hover_color=COLOR_HOVER, text_color=COLOR_TEXT, font=self._font(12))
        self.cfg_compound_field_menu.grid(row=1, column=0, sticky="ew", pady=(3, 0))
        self._small_entry(comp_card, 2, 1, "一级分隔符", self.cfg_compound_group_sep_var, "多组用 |，无则留空")
        self._small_entry(comp_card, 2, 2, "二级分隔符", self.cfg_compound_param_sep_var, "每组用 *")
        self._small_entry(comp_card, 2, 3, "每组参数数量", self.cfg_compound_count_var, "例如 4")
        self._small_entry(comp_card, 3, 0, "元素类型", self.cfg_compound_type_var, "整数")
        self._small_entry(comp_card, 3, 1, "禁止值", self.cfg_compound_forbid_values_var, "例如 0，多个用逗号")
        ctk.CTkCheckBox(comp_card, text="允许字段为空", variable=self.cfg_compound_allow_empty_var, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(12)).grid(row=3, column=2, sticky="w", padx=(0, 12), pady=(25, 4))
        ctk.CTkCheckBox(comp_card, text="允许参数为 0", variable=self.cfg_compound_allow_zero_var, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(12)).grid(row=3, column=3, sticky="w", padx=(0, 24), pady=(25, 4))
        comp_btns = ctk.CTkFrame(comp_card, fg_color="transparent")
        comp_btns.grid(row=4, column=0, columnspan=4, sticky="ew", padx=24, pady=(8, 8))
        ctk.CTkButton(comp_btns, text="添加 / 更新复合规则", width=150, height=34, corner_radius=12, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color="#FFFFFF", font=self._font(13, "bold"), command=self.add_config_compound_rule).pack(side="left", padx=(0, 10))
        ctk.CTkButton(comp_btns, text="套用 reward 四参数", width=140, height=34, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=lambda: self.apply_compound_template("reward4")).pack(side="left", padx=(0, 10))
        ctk.CTkButton(comp_btns, text="套用 restrict x|y", width=140, height=34, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=lambda: self.apply_compound_template("restrict")).pack(side="left", padx=(0, 10))
        ctk.CTkButton(comp_btns, text="清空复合规则", width=120, height=34, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self.clear_config_compound_rules).pack(side="left")
        self.cfg_compound_rules_box = ctk.CTkTextbox(comp_card, height=150, corner_radius=12, border_width=1, border_color=COLOR_BORDER, fg_color=COLOR_SURFACE_2, text_color=COLOR_TEXT, font=self._font(12), wrap="word")
        self.cfg_compound_rules_box.grid(row=5, column=0, columnspan=4, sticky="ew", padx=24, pady=(0, 18))
        self._set_textbox_text(self.cfg_compound_rules_box, self.cfg_compound_rules_var.get() or "规则格式：字段=reward;允许空=是;一级=|;二级=*;数量=4;类型=整数;允许0=是;禁止值=\n")
        try:
            self.cfg_compound_rules_box.configure(state="normal")
        except Exception:
            pass

        # 高级规则：保留原能力
        advanced_card = self._card(advanced_tab)
        advanced_card.grid_columnconfigure((0, 1), weight=1, uniform="cfg_fields")
        ctk.CTkLabel(advanced_card, text="高级规则配置（保留原能力）", text_color=COLOR_TEXT, font=self._font(16, "bold"), anchor="w").grid(row=0, column=0, columnspan=2, sticky="ew", padx=24, pady=(18, 8))
        ctk.CTkLabel(advanced_card, text="适合熟悉字段规则的用户；新手优先使用“快捷校验 / 类型行自动校验 / 复合字段校验”。", text_color=COLOR_MUTED, font=self._font(12), anchor="w").grid(row=1, column=0, columnspan=2, sticky="ew", padx=24, pady=(0, 8))
        self._small_entry(advanced_card, 2, 0, "必填字段", self.cfg_required_fields_var, "字段1, 字段2")
        self._small_entry(advanced_card, 2, 1, "必须存在字段", self.cfg_required_columns_var, "字段1, 字段2")
        self._small_entry(advanced_card, 3, 0, "数值字段", self.cfg_numeric_fields_var, "level, price")
        self._small_entry(advanced_card, 3, 1, "数值范围", self.cfg_range_rules_var, "level:1:100，每行一个")
        self._small_entry(advanced_card, 4, 0, "布尔字段", self.cfg_boolean_fields_var, "enable, is_open")
        self._small_entry(advanced_card, 4, 1, "枚举规则", self.cfg_enum_rules_var, "type=1,2,3，每行一个")
        self._small_entry(advanced_card, 5, 0, "时间字段", self.cfg_time_fields_var, "start_time, end_time")
        self._small_entry(advanced_card, 5, 1, "时间范围", self.cfg_time_range_rules_var, "start_time,end_time，每行一个")
        self._small_entry(advanced_card, 6, 0, "奖励字段", self.cfg_reward_fields_var, "reward, rewards")
        self._small_entry(advanced_card, 6, 1, "区间规则", self.cfg_interval_rules_var, "min,max，每行一个")
        self._small_entry(advanced_card, 7, 0, "引用 ID 规则", self.cfg_reference_rules_var, "item_id=道具表.xlsx|ID|Sheet1|1|2")
        switch_card = self._card(switch_tab)
        switch_card.grid_columnconfigure((0, 1, 2, 3), weight=1)
        ctk.CTkLabel(
            switch_card,
            text="启用的检查项目",
            text_color=COLOR_TEXT,
            font=self._font(16, "bold"),
            anchor="w",
        ).grid(row=0, column=0, columnspan=4, sticky="ew", padx=24, pady=(18, 10))
        for col in range(4):
            switch_card.grid_columnconfigure(col, weight=1)
        for idx, (label, var) in enumerate([
            ("主键重复", self.cfg_check_duplicate_var), ("主键为空", self.cfg_check_empty_key_var), ("必填", self.cfg_check_required_var), ("数值", self.cfg_check_numeric_var),
            ("范围", self.cfg_check_range_var), ("布尔", self.cfg_check_boolean_var), ("枚举", self.cfg_check_enum_var), ("时间", self.cfg_check_time_var),
            ("时间范围", self.cfg_check_time_range_var), ("奖励", self.cfg_check_reward_var), ("引用", self.cfg_check_reference_var), ("区间", self.cfg_check_interval_var),
            ("空行", self.cfg_check_empty_rows_var), ("字段缺失", self.cfg_check_required_columns_var),
        ]):
            ctk.CTkCheckBox(switch_card, text=label, variable=var, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(12)).grid(row=1 + idx // 4, column=idx % 4, sticky="w", padx=14, pady=8)

        # 校验结果
        result_tab.grid_rowconfigure(0, weight=1)
        result_card = ctk.CTkFrame(result_tab, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
        result_card.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        result_card.grid_columnconfigure(0, weight=1)
        result_card.grid_rowconfigure(2, weight=1)
        result_action = ctk.CTkFrame(result_card, fg_color=COLOR_SURFACE)
        result_action.grid(row=0, column=0, sticky="ew", padx=18, pady=(16, 8))
        result_action.grid_columnconfigure(3, weight=1)
        self.cfg_check_button = ctk.CTkButton(result_action, text="开始校验", width=130, height=40, corner_radius=12, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color="#FFFFFF", font=self._font(15, "bold"), command=self.start_config_validation)
        self.cfg_check_button.grid(row=0, column=0, padx=(0, 10), sticky="w")
        self.cfg_stop_button = ctk.CTkButton(result_action, text="停止", width=78, height=40, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(14, "bold"), command=self.stop_config_validation)
        self.cfg_stop_button.grid(row=0, column=1, padx=(0, 12), sticky="w")
        search_entry = ctk.CTkEntry(result_action, textvariable=self.cfg_search_var, placeholder_text="搜索问题类型 / 字段名 / 主键 ID / 当前值 / 问题说明", height=40, corner_radius=10, fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(14))
        search_entry.grid(row=0, column=3, sticky="ew", padx=(0, 10))
        ctk.CTkButton(result_action, text="复制结果", width=100, height=40, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self.copy_config_validation_results).grid(row=0, column=4, padx=(0, 10))
        ctk.CTkButton(result_action, text="导出 Excel", width=110, height=40, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=self.export_config_validation_excel).grid(row=0, column=5)
        ctk.CTkLabel(result_card, textvariable=self.cfg_summary_var, text_color=COLOR_TEXT, font=self._font(14, "bold"), anchor="w").grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 6))

        table_frame = ctk.CTkFrame(result_card, fg_color=COLOR_SURFACE_2, corner_radius=12, border_width=1, border_color=COLOR_BORDER)
        table_frame.grid(row=2, column=0, sticky="nsew", padx=18, pady=(0, 18))
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)
        columns = ("index", "row", "item_id", "field", "field_type", "value", "issue_type", "remark", "suggestion", "file_name", "sheet")
        self.cfg_tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=16, style="Xiao.Treeview")
        headings = {"index": "序号", "row": "行号", "item_id": "主键 ID", "field": "字段名", "field_type": "字段类型", "value": "当前值", "issue_type": "问题类型", "remark": "问题说明", "suggestion": "建议处理", "file_name": "文件名", "sheet": "Sheet"}
        widths = {"index": 60, "row": 80, "item_id": 130, "field": 150, "field_type": 100, "value": 220, "issue_type": 150, "remark": 380, "suggestion": 320, "file_name": 170, "sheet": 100}
        for col in columns:
            self.cfg_tree.heading(col, text=headings[col])
            self.cfg_tree.column(col, width=widths[col], anchor="center" if col in {"index", "row", "field_type", "issue_type"} else "w", stretch=False)
        self.cfg_tree.bind("<Double-1>", self.show_config_validation_issue_detail)
        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.cfg_tree.yview)
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.cfg_tree.xview)
        self.cfg_tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.cfg_tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        pager = self._build_pager(result_card, self.cfg_page_status_var, self.cfg_page_jump_var, self.cfg_prev_page, self.cfg_next_page, self.cfg_jump_page)
        pager.grid(row=3, column=0, sticky="ew", padx=18, pady=(0, 14))

        if not self.cfg_search_trace_bound:
            self.cfg_search_var.trace_add("write", lambda *_: self.apply_config_validation_search_filter())
            self.cfg_search_trace_bound = True
        self._refresh_config_validator_button()
        self.update_config_quick_preview()
        self._refresh_config_field_choices([])


    def _refresh_config_field_choices(self, headers):
        headers = list(headers or [])
        self.cfg_current_headers = headers
        values = [""] + headers if headers else [""]
        for attr in ("cfg_quick_field_menu", "cfg_compound_field_menu"):
            widget = getattr(self, attr, None)
            if widget is not None:
                try:
                    widget.configure(values=values)
                except Exception:
                    pass

        quick_field = self.cfg_quick_field_var.get().strip()
        if quick_field and quick_field not in headers:
            self.cfg_quick_field_var.set("")
        if not self.cfg_quick_field_var.get().strip() and headers:
            preferred = self._find_quick_template_field(self.cfg_quick_template_var.get().strip(), headers)
            if preferred:
                self.cfg_quick_field_var.set(preferred)

        # 复合字段必须由用户明确选择或模板明确匹配，禁止自动回退到第一个字段（尤其是 #id）。
        compound_field = self.cfg_compound_field_var.get().strip()
        if compound_field and (compound_field not in headers or is_key_index_field(compound_field)):
            self.cfg_compound_field_var.set("")
        self.update_config_quick_preview()

    def _find_quick_template_field(self, template, headers=None):
        headers = list(headers if headers is not None else getattr(self, "cfg_current_headers", []))
        safe_headers = [h for h in headers if str(h or "").strip() and not is_key_index_field(h)]
        template = (template or "").strip()
        if template in {"允许为空或 x|y", "必须为 x|y", "禁止单值 0", "restrict 条件格式校验"}:
            for header in safe_headers:
                if str(header).strip().lower() == "restrict":
                    return header
            return ""
        if template in {"reward 四参数校验", "reward 三参数校验"}:
            return find_reward_like_field(safe_headers)
        if template == "condition 条件字段校验":
            for header in safe_headers:
                if str(header).strip().lower() == "condition":
                    return header
            return ""
        return ""

    def _selected_compound_rule_text(self):
        field = self.cfg_compound_field_var.get().strip()
        if not field:
            return ""
        allow_empty = "是" if self.cfg_compound_allow_empty_var.get() else "否"
        allow_zero = "是" if self.cfg_compound_allow_zero_var.get() else "否"
        group_sep = self.cfg_compound_group_sep_var.get().strip() or "无"
        param_sep = self.cfg_compound_param_sep_var.get().strip() or "无"
        count = self.cfg_compound_count_var.get().strip() or "4"
        etype = self.cfg_compound_type_var.get().strip() or "整数"
        forbid = self.cfg_compound_forbid_values_var.get().strip()
        return f"字段={field};允许空={allow_empty};一级={group_sep};二级={param_sep};数量={count};类型={etype};允许0={allow_zero};禁止值={forbid}"

    def _build_config_rule_preview(self, options=None):
        path = self.cfg_file_path_var.get().strip()
        if not path:
            return ""
        table = parse_config_table(
            path,
            header_row=safe_int(self.cfg_header_row_var.get(), 1),
            data_start_row=safe_int(self.cfg_data_start_row_var.get(), 2),
            key_column=self.cfg_key_column_var.get().strip(),
            sheet_name=self.cfg_sheet_var.get().strip(),
            type_row=safe_int(self.cfg_type_row_var.get(), 0),
            description_row=safe_int(self.cfg_description_row_var.get(), 0),
        )
        self.cfg_key_column_var.set(table.key_column)
        self._refresh_config_field_choices(table.headers)
        if options is None:
            options = self.get_config_validation_options()
        quick_template = self.cfg_quick_template_var.get().strip()
        quick_field = self.cfg_quick_field_var.get().strip()
        if quick_template in {"reward 四参数校验", "reward 三参数校验"}:
            if quick_field and is_key_index_field(quick_field):
                raise ValueError(f"该字段是主键 / 索引字段，不适用于复合字段参数个数校验，请选择 {COMPOUND_FIELD_EXAMPLES} 等复合字段")
            if quick_field and not is_reward_like_field(quick_field):
                matched = find_reward_like_field(table.headers)
                if matched:
                    self.cfg_quick_field_var.set(matched)
                    options = self.get_config_validation_options()
                else:
                    raise ValueError("未找到 reward 类字段，请手动选择需要校验的字段")
            elif not find_reward_like_field(table.headers) and not (quick_field and is_reward_like_field(quick_field)):
                raise ValueError("未找到 reward 类字段，请手动选择需要校验的字段")
        lines = []
        if options.check_type_fields:
            lines.append(f"类型行自动校验：将根据第 {options.type_row} 行类型检查全表，预计检查行数：{len(table.records)} 行")
        if options.check_compound_rules:
            rules_text = self._get_compound_rules_text_from_ui()
            rule_lines = [line.strip() for line in rules_text.splitlines() if line.strip()]
            if not rule_lines:
                raise ValueError(f"请先选择要校验的复合字段，例如 {COMPOUND_FIELD_EXAMPLES}")
            # 使用底层解析函数不暴露在外，因此按规则文本做轻量预览；真正执行仍由 validator 兜底校验。
            for line in rule_lines:
                data = {}
                for part in line.split(";"):
                    if "=" in part:
                        k, v = part.split("=", 1)
                        data[k.strip()] = v.strip()
                field = data.get("字段") or data.get("field") or ""
                if not field:
                    raise ValueError(f"请先选择要校验的复合字段，例如 {COMPOUND_FIELD_EXAMPLES}")
                actual_field = next((h for h in table.headers if str(h).strip().lower() == field.strip().lower()), field)
                if is_key_index_field(actual_field):
                    raise ValueError(f"该字段是主键 / 索引字段，不适用于复合字段参数个数校验，请选择 {COMPOUND_FIELD_EXAMPLES} 等复合字段")
                ftype = table.field_types.get(actual_field, "") or "未识别"
                group_sep = data.get("一级") or "无"
                param_sep = data.get("二级") or "无"
                count = data.get("数量") or ""
                allow_empty = data.get("允许空") or "是"
                allow_zero = data.get("允许0") or "是"
                lines.append(
                    f"本次将检查字段：{actual_field}\n"
                    f"字段类型：{ftype}\n"
                    f"一级分隔符：{group_sep}；二级分隔符：{param_sep}；每组参数数量：{count}\n"
                    f"是否允许为空：{allow_empty}；是否允许 0：{allow_zero}\n"
                    f"预计检查行数：{len(table.records)} 行"
                )
        if not lines:
            raise ValueError("请先在快捷校验、类型行自动校验或复合字段校验中选择校验规则")
        return "\n\n".join(lines)

    def _confirm_config_rule_preview(self, options):
        preview = self._build_config_rule_preview(options)
        return messagebox.askokcancel("规则预览", preview + "\n\n确认开始校验？")


    def detect_config_structure_for_current_file(self):
        path = self.cfg_file_path_var.get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择配置表")
            return
        try:
            info = detect_config_structure(path, self.cfg_sheet_var.get().strip())
            self.cfg_description_row_var.set(str(info.get("description_row") or 0))
            self.cfg_type_row_var.set(str(info.get("type_row") or 0))
            self.cfg_header_row_var.set(str(info.get("header_row") or 1))
            self.cfg_data_start_row_var.set(str(info.get("data_start_row") or 2))
            self.cfg_key_column_var.set(info.get("key_column") or "")
            if info.get("sheet_name") and info.get("sheet_name") != "CSV":
                self.cfg_sheet_var.set(info.get("sheet_name"))
            self._refresh_config_field_choices(info.get("headers") or [])
            summary = (
                f"已识别表结构：Sheet={info.get('sheet_name')}\n"
                f"说明行：{info.get('description_row')}｜类型行：{info.get('type_row')}｜字段名行：{info.get('header_row')}｜数据起始行：{info.get('data_start_row')}\n"
                f"主键字段：{info.get('key_column')}｜字段数量：{len(info.get('headers') or [])}｜类型字段：{info.get('type_field_count')}｜数据行数：{info.get('data_rows')}｜空单元格：{info.get('empty_cell_count')}\n"
                f"字段预览：{', '.join((info.get('headers') or [])[:40])}"
            )
            self.cfg_detect_summary_var.set(summary)
            if hasattr(self, "cfg_detect_box") and self.cfg_detect_box.winfo_exists():
                self._set_textbox_text(self.cfg_detect_box, summary)
            self.cfg_file_status_var.set(f"自动识别成功：字段名行 {info.get('header_row')}，类型行 {info.get('type_row')}，数据起始行 {info.get('data_start_row')}，主键 {info.get('key_column')}")
        except Exception as exc:
            messagebox.showerror("识别失败", str(exc))

    def update_config_quick_preview(self, *args):
        template = self.cfg_quick_template_var.get().strip()
        field = self.cfg_quick_field_var.get().strip()
        if not field:
            preferred = self._find_quick_template_field(template)
            if preferred:
                self.cfg_quick_field_var.set(preferred)
                field = preferred
        help_text = TEMPLATE_HELP.get(template, "")
        if field and is_key_index_field(field) and template in {"reward 四参数校验", "reward 三参数校验", "允许为空或 x|y", "必须为 x|y", "禁止单值 0", "restrict 条件格式校验", "condition 条件字段校验"}:
            text = f"模板：{template}\n说明：{help_text}\n\n该字段是主键 / 索引字段，不适用于复合字段参数个数校验，请选择 {COMPOUND_FIELD_EXAMPLES} 等复合字段。"
        elif not field:
            text = f"模板：{template}\n说明：{help_text}\n\n请先选择字段。"
        else:
            text = f"本次将检查字段：{field}\n模板：{template}\n说明：{help_text}\n\n"
            text += self._build_config_quick_rule_preview(field, template)
        self.cfg_quick_preview_var.set(text)
        if hasattr(self, "cfg_quick_preview_box") and self.cfg_quick_preview_box.winfo_exists():
            self._set_textbox_text(self.cfg_quick_preview_box, text)


    def _build_config_quick_rule_preview(self, field, template):
        if template == "必填字段":
            extras = []
            if self.cfg_quick_required_no_zero_var.get():
                extras.append("0 视为无效")
            if self.cfg_quick_required_no_minus_one_var.get():
                extras.append("-1 视为无效")
            return "规则：字段不能为空" + ("；" + "；".join(extras) if extras else "")
        if template in {"允许为空或 x|y", "restrict 条件格式校验"}:
            return "规则：允许为空；如果不为空，必须为 2 个整数并使用 | 分隔；禁止填写 0。"
        if template == "必须为 x|y":
            return "规则：不能为空；必须为 2 个整数并使用 | 分隔。"
        if template == "禁止单值 0":
            return "规则：字段不能只填写 0。"
        if template == "reward 四参数校验":
            return "规则：多组用 | 分隔；每组用 * 分隔；每组必须 4 个整数参数。"
        if template == "reward 三参数校验":
            return "规则：多组用 | 分隔；每组用 * 分隔；每组必须 3 个整数参数。"
        if template == "condition 条件字段校验":
            return "规则：允许为空；如果不为空，必须为 2 或 3 个整数参数。"
        if template == "按类型行自动校验":
            return "规则：根据类型行自动检查 uint / int / float / bool / 数组字段。"
        return "规则：按模板生成快捷校验规则。"

    def _build_config_quick_required_fields(self):
        field = self.cfg_quick_field_var.get().strip()
        template = self.cfg_quick_template_var.get().strip()
        if template == "必填字段" and field:
            return field
        return ""

    def _build_config_quick_compound_rules(self):
        field = self.cfg_quick_field_var.get().strip()
        template = self.cfg_quick_template_var.get().strip()
        rules = []
        if not field:
            field = self._find_quick_template_field(template)
            if field:
                self.cfg_quick_field_var.set(field)
        if not field or is_key_index_field(field):
            return ""
        if template in {"允许为空或 x|y", "restrict 条件格式校验"}:
            if template == "restrict 条件格式校验" and field.strip().lower() != "restrict":
                return ""
            rules.append(f"字段={field};允许空=是;一级=无;二级=|;数量=2;类型=uint;允许0=是;禁止值=0")
        elif template == "必须为 x|y":
            rules.append(f"字段={field};允许空=否;一级=无;二级=|;数量=2;类型=uint;允许0=是;禁止值=")
        elif template == "禁止单值 0":
            rules.append(f"字段={field};允许空=是;一级=无;二级=无;数量=1;类型=string;允许0=是;禁止值=0")
        elif template == "reward 四参数校验":
            if not is_reward_like_field(field):
                field = find_reward_like_field(getattr(self, "cfg_current_headers", []))
            if not field:
                return ""
            rules.append(f"字段={field};允许空=是;一级=|;二级=*;数量=4;类型=整数;允许0=是;禁止值=")
        elif template == "reward 三参数校验":
            if not is_reward_like_field(field):
                field = find_reward_like_field(getattr(self, "cfg_current_headers", []))
            if not field:
                return ""
            rules.append(f"字段={field};允许空=是;一级=|;二级=*;数量=3;类型=整数;允许0=是;禁止值=")
        elif template == "condition 条件字段校验":
            if field.strip().lower() != "condition":
                return ""
            # 默认先按两个参数检查；高级场景可在复合字段校验页手动调整为 3。
            rules.append(f"字段={field};允许空=是;一级=无;二级=|;数量=2;类型=uint;允许0=是;禁止值=")
        return "\n".join(rules)


    def _get_compound_rules_text_from_ui(self):
        lines = []
        if self.cfg_compound_enabled_var.get():
            try:
                if hasattr(self, "cfg_compound_rules_box") and self.cfg_compound_rules_box.winfo_exists():
                    content = self.cfg_compound_rules_box.get("1.0", "end").strip()
                else:
                    content = self.cfg_compound_rules_var.get().strip()
                for line in content.splitlines():
                    s = line.strip()
                    if s and not s.startswith("规则格式"):
                        lines.append(s)
            except Exception:
                pass
            selected = self._selected_compound_rule_text()
            if selected:
                lines.append(selected)
        quick = self._build_config_quick_compound_rules()
        if quick:
            lines.extend(quick.splitlines())
        return "\n".join(dict.fromkeys(lines))


    def add_config_compound_rule(self):
        field = self.cfg_compound_field_var.get().strip()
        if not field:
            messagebox.showwarning("提示", f"请先选择要校验的复合字段，例如 {COMPOUND_FIELD_EXAMPLES}")
            return
        if is_key_index_field(field):
            messagebox.showwarning("提示", f"该字段是主键 / 索引字段，不适用于复合字段参数个数校验，请选择 {COMPOUND_FIELD_EXAMPLES} 等复合字段")
            return
        line = self._selected_compound_rule_text()
        existing = self._get_compound_rules_text_from_ui()
        lines = [l for l in existing.splitlines() if l.strip() and f"字段={field};" not in l]
        lines.append(line)
        text = "\n".join(dict.fromkeys(lines))
        self.cfg_compound_rules_var.set(text)
        if hasattr(self, "cfg_compound_rules_box") and self.cfg_compound_rules_box.winfo_exists():
            self._set_textbox_text(self.cfg_compound_rules_box, text)
            try:
                self.cfg_compound_rules_box.configure(state="normal")
            except Exception:
                pass


    def clear_config_compound_rules(self):
        self.cfg_compound_rules_var.set("")
        if hasattr(self, "cfg_compound_rules_box") and self.cfg_compound_rules_box.winfo_exists():
            self._set_textbox_text(self.cfg_compound_rules_box, "")
            try:
                self.cfg_compound_rules_box.configure(state="normal")
            except Exception:
                pass

    def apply_compound_template(self, kind):
        if kind == "reward4":
            field = find_reward_like_field(getattr(self, "cfg_current_headers", []))
            if not field:
                current = self.cfg_compound_field_var.get().strip()
                field = current if current and is_reward_like_field(current) and not is_key_index_field(current) else ""
            if not field:
                messagebox.showwarning("提示", "未找到 reward 类字段，请手动选择需要校验的字段")
                return
            self.cfg_compound_enabled_var.set(True)
            self.cfg_compound_field_var.set(field)
            self.cfg_compound_allow_empty_var.set(True)
            self.cfg_compound_group_sep_var.set("|")
            self.cfg_compound_param_sep_var.set("*")
            self.cfg_compound_count_var.set("4")
            self.cfg_compound_type_var.set("整数")
            self.cfg_compound_allow_zero_var.set(True)
            self.cfg_compound_forbid_values_var.set("")
        elif kind == "restrict":
            headers = getattr(self, "cfg_current_headers", [])
            field = next((h for h in headers if str(h).strip().lower() == "restrict"), "")
            if not field:
                current = self.cfg_compound_field_var.get().strip()
                field = current if current.strip().lower() == "restrict" else ""
            if not field:
                messagebox.showwarning("提示", "未找到 restrict 字段，请手动选择需要校验的字段")
                return
            self.cfg_compound_enabled_var.set(True)
            self.cfg_compound_field_var.set(field)
            self.cfg_compound_allow_empty_var.set(True)
            self.cfg_compound_group_sep_var.set("")
            self.cfg_compound_param_sep_var.set("|")
            self.cfg_compound_count_var.set("2")
            self.cfg_compound_type_var.set("uint")
            self.cfg_compound_allow_zero_var.set(True)
            self.cfg_compound_forbid_values_var.set("0")
        self.add_config_compound_rule()


    def show_config_validation_issue_detail(self, event=None):
        if not hasattr(self, "cfg_tree") or not self.cfg_tree.winfo_exists():
            return
        item_id = self.cfg_tree.focus()
        values = self.cfg_tree.item(item_id, "values") if item_id else None
        if not values:
            return
        labels = ["序号", "行号", "主键 ID", "字段名", "字段类型", "当前值", "问题类型", "问题说明", "建议处理", "文件名", "Sheet"]
        detail = "\n".join(f"{label}：{value}" for label, value in zip(labels, values))
        top = ctk.CTkToplevel(self)
        top.title("配置表校验结果详情")
        top.geometry("920x620")
        top.minsize(760, 480)
        top.configure(fg_color=COLOR_BG)
        try:
            top.transient(self)
            top.grab_set()
        except Exception:
            pass
        top.grid_columnconfigure(0, weight=1)
        top.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(top, text="配置表校验结果详情", text_color=COLOR_TEXT, font=self._font(18, "bold"), anchor="w").grid(row=0, column=0, sticky="ew", padx=18, pady=(16, 8))
        box = ctk.CTkTextbox(top, wrap="word", fg_color=COLOR_SURFACE_2, border_color=COLOR_BORDER, border_width=1, text_color=COLOR_TEXT, font=self._font(13), corner_radius=12)
        box.grid(row=1, column=0, sticky="nsew", padx=18, pady=(0, 12))
        box.insert("end", detail)
        box.configure(state="disabled")
        btns = ctk.CTkFrame(top, fg_color="transparent")
        btns.grid(row=2, column=0, sticky="ew", padx=18, pady=(0, 16))
        btns.grid_columnconfigure(0, weight=1)
        def copy_detail():
            self.clipboard_clear()
            self.clipboard_append(detail)
            messagebox.showinfo("已复制", "已复制当前问题详情")
        ctk.CTkButton(btns, text="复制详情", width=110, height=34, corner_radius=12, fg_color=COLOR_ACCENT, hover_color=COLOR_ACCENT_HOVER, text_color="#FFFFFF", font=self._font(13, "bold"), command=copy_detail).grid(row=0, column=1, padx=(0, 10))
        ctk.CTkButton(btns, text="关闭", width=90, height=34, corner_radius=12, fg_color=COLOR_SURFACE_2, hover_color=COLOR_HOVER, border_width=1, border_color=COLOR_BORDER, text_color=COLOR_TEXT, font=self._font(13, "bold"), command=top.destroy).grid(row=0, column=2)


    def select_config_file(self):
        path = filedialog.askopenfilename(title="选择配置表", filetypes=[("配置表", "*.xlsx *.csv *.json"), ("所有文件", "*.*")])
        if not path:
            return
        self.cfg_file_path_var.set(path)
        self.cfg_file_status_var.set(f"已选择：{get_file_info(path)}")
        self.cfg_all_issues = []
        self.cfg_visible_issues = []
        self.cfg_summary_var.set("已选择配置表，建议先自动识别表结构")
        if hasattr(self, "cfg_tree") and self.cfg_tree.winfo_exists():
            self.clear_config_validation_table()
        try:
            self.detect_config_structure_for_current_file()
        except Exception:
            pass

    def preview_config_fields(self):
        path = self.cfg_file_path_var.get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择配置表")
            return
        try:
            table = parse_config_table(
                path,
                header_row=safe_int(self.cfg_header_row_var.get(), 1),
                data_start_row=safe_int(self.cfg_data_start_row_var.get(), 2),
                key_column=self.cfg_key_column_var.get().strip(),
                sheet_name=self.cfg_sheet_var.get().strip(),
                type_row=safe_int(self.cfg_type_row_var.get(), 0),
                description_row=safe_int(self.cfg_description_row_var.get(), 0),
            )
            self.cfg_key_column_var.set(table.key_column)
            self._refresh_config_field_choices(table.headers)
            preview_lines = []
            for h in table.headers[:80]:
                comment = table.field_comments.get(h, "")
                ftype = table.field_types.get(h, "")
                suffix = []
                if comment:
                    suffix.append(comment)
                if ftype:
                    suffix.append(ftype)
                preview_lines.append(f"{h}" + (f"（{' / '.join(suffix)}）" if suffix else ""))
            preview = "\n".join(preview_lines)
            self.cfg_file_status_var.set(f"字段预览：Sheet={table.sheet_name} / 记录 {len(table.records)} 条 / 主键 {table.key_column} / 字段 {len(table.headers)} 个")
            messagebox.showinfo("字段预览", preview or "未识别到字段")
        except Exception as exc:
            messagebox.showerror("预览失败", str(exc))

    def _toggle_config_rules(self):
        if self.cfg_rules_frame is None:
            return
        self.cfg_rules_visible = not self.cfg_rules_visible
        if self.cfg_rules_visible:
            self.cfg_rules_frame.grid()
            if self.cfg_rules_toggle_button is not None:
                self.cfg_rules_toggle_button.configure(text="收起规则配置")
        else:
            self.cfg_rules_frame.grid_remove()
            if self.cfg_rules_toggle_button is not None:
                self.cfg_rules_toggle_button.configure(text="展开规则配置")

    def get_config_validation_options(self):
        required_fields_text = self.cfg_required_fields_var.get().strip()
        quick_required = self._build_config_quick_required_fields()
        if quick_required:
            required_fields_text = ", ".join([v for v in [required_fields_text, quick_required] if v])
        compound_rules_text = self._get_compound_rules_text_from_ui()
        check_compound_rules = bool(compound_rules_text.strip()) or self.cfg_compound_enabled_var.get()
        return ConfigValidationOptions(
            header_row=safe_int(self.cfg_header_row_var.get(), 1),
            description_row=safe_int(self.cfg_description_row_var.get(), 0),
            type_row=safe_int(self.cfg_type_row_var.get(), 0),
            data_start_row=safe_int(self.cfg_data_start_row_var.get(), 2),
            key_column=self.cfg_key_column_var.get().strip(),
            sheet_name=self.cfg_sheet_var.get().strip(),
            check_type_fields=self.cfg_auto_type_enabled_var.get(),
            check_compound_rules=check_compound_rules,
            compound_rules_text=compound_rules_text,
            required_fields_text=required_fields_text,
            numeric_fields_text=self.cfg_numeric_fields_var.get(),
            range_rules_text=self.cfg_range_rules_var.get(),
            boolean_fields_text=self.cfg_boolean_fields_var.get(),
            enum_rules_text=self.cfg_enum_rules_var.get(),
            time_fields_text=self.cfg_time_fields_var.get(),
            time_range_rules_text=self.cfg_time_range_rules_var.get(),
            reward_fields_text=self.cfg_reward_fields_var.get(),
            reference_rules_text=self.cfg_reference_rules_var.get(),
            interval_rules_text=self.cfg_interval_rules_var.get(),
            required_columns_text=self.cfg_required_columns_var.get(),
            check_duplicate_key=self.cfg_check_duplicate_var.get(),
            check_empty_key=self.cfg_check_empty_key_var.get(),
            check_required_fields=self.cfg_check_required_var.get(),
            check_numeric_fields=self.cfg_check_numeric_var.get(),
            check_range=self.cfg_check_range_var.get(),
            check_boolean_fields=self.cfg_check_boolean_var.get(),
            check_enum_fields=self.cfg_check_enum_var.get(),
            check_time_fields=self.cfg_check_time_var.get(),
            check_time_range=self.cfg_check_time_range_var.get(),
            check_reward_fields=self.cfg_check_reward_var.get(),
            check_reference_ids=self.cfg_check_reference_var.get(),
            check_interval=self.cfg_check_interval_var.get(),
            check_empty_rows=self.cfg_check_empty_rows_var.get(),
            check_required_columns=self.cfg_check_required_columns_var.get(),
            ui_limit=UI_MAX_RENDER_ROWS,
        )

    def apply_config_validation_options(self, options):
        self.cfg_header_row_var.set(str(options.header_row))
        self.cfg_description_row_var.set(str(getattr(options, "description_row", 1)))
        self.cfg_type_row_var.set(str(getattr(options, "type_row", 2)))
        self.cfg_data_start_row_var.set(str(options.data_start_row))
        self.cfg_key_column_var.set(options.key_column or "")
        self.cfg_sheet_var.set(options.sheet_name or "")
        self.cfg_auto_type_enabled_var.set(bool(getattr(options, "check_type_fields", False)))
        self.cfg_compound_enabled_var.set(bool(getattr(options, "check_compound_rules", False)))
        self.cfg_compound_rules_var.set(getattr(options, "compound_rules_text", "") or "")
        if not self.cfg_compound_rules_var.get().strip():
            self.cfg_compound_field_var.set("")
        if hasattr(self, "cfg_compound_rules_box") and self.cfg_compound_rules_box.winfo_exists():
            self._set_textbox_text(self.cfg_compound_rules_box, self.cfg_compound_rules_var.get())
        self.cfg_required_fields_var.set(options.required_fields_text or "")
        self.cfg_numeric_fields_var.set(options.numeric_fields_text or "")
        self.cfg_range_rules_var.set(options.range_rules_text or "")
        self.cfg_boolean_fields_var.set(options.boolean_fields_text or "")
        self.cfg_enum_rules_var.set(options.enum_rules_text or "")
        self.cfg_time_fields_var.set(options.time_fields_text or "")
        self.cfg_time_range_rules_var.set(options.time_range_rules_text or "")
        self.cfg_reward_fields_var.set(options.reward_fields_text or "")
        self.cfg_reference_rules_var.set(options.reference_rules_text or "")
        self.cfg_interval_rules_var.set(options.interval_rules_text or "")
        self.cfg_required_columns_var.set(options.required_columns_text or "")
        self.cfg_check_duplicate_var.set(options.check_duplicate_key)
        self.cfg_check_empty_key_var.set(options.check_empty_key)
        self.cfg_check_required_var.set(options.check_required_fields)
        self.cfg_check_numeric_var.set(options.check_numeric_fields)
        self.cfg_check_range_var.set(options.check_range)
        self.cfg_check_boolean_var.set(options.check_boolean_fields)
        self.cfg_check_enum_var.set(options.check_enum_fields)
        self.cfg_check_time_var.set(options.check_time_fields)
        self.cfg_check_time_range_var.set(options.check_time_range)
        self.cfg_check_reward_var.set(options.check_reward_fields)
        self.cfg_check_reference_var.set(options.check_reference_ids)
        self.cfg_check_interval_var.set(options.check_interval)
        self.cfg_check_empty_rows_var.set(options.check_empty_rows)
        self.cfg_check_required_columns_var.set(options.check_required_columns)

    def save_config_rules(self):
        output_path = filedialog.asksaveasfilename(title="保存校验规则", defaultextension=".json", filetypes=[("JSON 文件", "*.json")], initialfile="配置表校验规则.json")
        if not output_path:
            return
        try:
            data = config_options_to_dict(self.get_config_validation_options())
            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("保存成功", f"已保存：{output_path}")
        except Exception as exc:
            messagebox.showerror("保存失败", str(exc))

    def load_config_rules(self):
        path = filedialog.askopenfilename(title="加载校验规则", filetypes=[("JSON 文件", "*.json"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.apply_config_validation_options(config_options_from_dict(data))
            self.cfg_summary_var.set(f"已加载规则：{os.path.basename(path)}")
        except Exception as exc:
            messagebox.showerror("加载失败", str(exc))

    def clear_config_rules(self):
        self.apply_config_validation_options(ConfigValidationOptions())
        self.cfg_summary_var.set("已清空规则配置")

    def start_config_validation(self):
        if self.cfg_is_working:
            return
        path = self.cfg_file_path_var.get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择配置表")
            return
        try:
            preview = self._build_config_rule_preview()
            options = self.get_config_validation_options()
            if not messagebox.askokcancel("规则预览", preview + "\n\n确认开始校验？"):
                return
        except Exception as exc:
            messagebox.showwarning("提示", str(exc))
            return
        self.cfg_is_working = True
        self.cfg_stop_event.clear()
        self._refresh_config_validator_button()
        self._set_badge(True)
        self.cfg_summary_var.set("正在校验配置表...")
        self.clear_config_validation_table()
        threading.Thread(target=self._config_validation_worker, args=(path, options), daemon=True).start()

    def stop_config_validation(self):
        if self.cfg_is_working:
            self.cfg_stop_event.set()
            self.cfg_summary_var.set("正在请求停止校验...")

    def _config_validation_worker(self, path, options):
        try:
            issues, table = validate_config_table(path, options, self.cfg_stop_event)
            self.after(0, lambda: self._on_config_validation_success(issues, table))
        except Exception as exc:
            self.after(0, lambda msg=f"配置表校验失败：{exc}": self._on_config_validation_error(msg))

    def _on_config_validation_success(self, issues, table):
        self.cfg_is_working = False
        self._refresh_config_validator_button()
        self._set_badge(False)
        self.cfg_all_issues = issues
        self.cfg_visible_issues = issues
        self.cfg_key_column_var.set(table.key_column)
        self.cfg_file_status_var.set(f"读取成功：Sheet={table.sheet_name} / 记录 {len(table.records)} 条 / 字段 {len(table.headers)} 个 / 主键 {table.key_column}")
        self.apply_config_validation_search_filter()
        counts = {}
        for issue in issues:
            counts[issue.issue_type] = counts.get(issue.issue_type, 0) + 1
        if not issues:
            self.cfg_summary_var.set("校验完成：未发现问题")
        else:
            top = " | ".join(f"{k} {v}" for k, v in sorted(counts.items())[:8])
            self.cfg_summary_var.set(f"校验完成：共 {len(issues)} 条问题 | {top}")

    def _on_config_validation_error(self, message):
        self.cfg_is_working = False
        self._refresh_config_validator_button()
        self._set_badge(False)
        self.cfg_summary_var.set("校验失败")
        messagebox.showerror("校验失败", message)

    def apply_config_validation_search_filter(self):
        keyword = self.cfg_search_var.get().strip().lower()
        if not keyword:
            self.cfg_visible_issues = self.cfg_all_issues
        else:
            self.cfg_visible_issues = [item for item in self.cfg_all_issues if keyword in self._config_validation_issue_to_text(item).lower()]
        self.cfg_page_index = 1
        self.refresh_config_validation_table(self.cfg_visible_issues)

    def refresh_config_validation_table(self, issues):
        if not hasattr(self, "cfg_tree") or not self.cfg_tree.winfo_exists():
            return
        self.clear_config_validation_table()
        self.cfg_page_index, total_pages, start, end = self._get_page_bounds(len(issues), self.cfg_page_index)
        self.cfg_page_jump_var.set(str(self.cfg_page_index))
        self.cfg_page_status_var.set(f"第 {self.cfg_page_index}/{total_pages} 页，每页 {UI_PAGE_SIZE} 条，当前筛选 {len(issues)} 条")
        for item in list(issues)[start:end]:
            self.cfg_tree.insert("", "end", values=(
                item.index,
                item.row_number,
                self._shorten_compare_text(item.item_id, 120),
                self._shorten_compare_text(item.field_name, 120),
                self._shorten_compare_text(getattr(item, "field_type", ""), 80),
                self._shorten_compare_text(item.current_value, 180),
                item.issue_type,
                self._shorten_compare_text(item.remark, 260),
                self._shorten_compare_text(item.suggestion, 220),
                self._shorten_compare_text(item.file_name, 120),
                self._shorten_compare_text(item.sheet, 80),
            ))

    def cfg_prev_page(self):
        self.cfg_page_index = max(1, self.cfg_page_index - 1)
        self.refresh_config_validation_table(self.cfg_visible_issues)

    def cfg_next_page(self):
        self.cfg_page_index += 1
        self.refresh_config_validation_table(self.cfg_visible_issues)

    def cfg_jump_page(self):
        try:
            self.cfg_page_index = int(self.cfg_page_jump_var.get().strip() or "1")
        except Exception:
            self.cfg_page_index = 1
        self.refresh_config_validation_table(self.cfg_visible_issues)

    def clear_config_validation_table(self):
        if not hasattr(self, "cfg_tree") or not self.cfg_tree.winfo_exists():
            return
        for item in self.cfg_tree.get_children():
            self.cfg_tree.delete(item)

    @staticmethod
    def _config_validation_issue_to_text(item):
        return "\t".join([str(item.index), item.row_number, item.item_id, item.field_name, getattr(item, "field_type", ""), item.current_value, item.issue_type, item.remark, item.suggestion, item.file_name, item.sheet])

    def copy_config_validation_results(self):
        if not self.cfg_visible_issues:
            messagebox.showinfo("提示", "当前没有可复制的校验结果")
            return
        text = validation_issues_to_tsv(self.cfg_visible_issues)
        self.clipboard_clear()
        self.clipboard_append(text)
        messagebox.showinfo("已复制", f"已复制 {len(self.cfg_visible_issues)} 条校验结果")

    def export_config_validation_excel(self):
        if not self.cfg_visible_issues:
            messagebox.showinfo("提示", "当前没有可导出的校验结果")
            return
        output_path = filedialog.asksaveasfilename(
            title="导出配置表校验报告",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile="配置表校验报告.xlsx",
        )
        if not output_path:
            return
        try:
            export_validation_issues_to_excel(self.cfg_visible_issues, output_path)
            messagebox.showinfo("导出成功", f"已导出：{output_path}")
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))

    def _refresh_config_validator_button(self):
        if hasattr(self, "cfg_check_button") and self.cfg_check_button.winfo_exists():
            self.cfg_check_button.configure(
                state="disabled" if self.cfg_is_working else "normal",
                text="校验中..." if self.cfg_is_working else "开始校验",
                fg_color=COLOR_DANGER if self.cfg_is_working else COLOR_ACCENT,
                hover_color=COLOR_DANGER_HOVER if self.cfg_is_working else COLOR_ACCENT_HOVER,
            )
        if hasattr(self, "cfg_stop_button") and self.cfg_stop_button.winfo_exists():
            self.cfg_stop_button.configure(state="normal" if self.cfg_is_working else "disabled")

    # ---------- System tray ----------
    def _setup_tray(self):
        if pystray is None:
            self.tray_available = False
            self.tray_error = "pystray 未安装"
            return
        try:
            menu = pystray.Menu(
                pystray.MenuItem("显示主窗口", self._tray_show_main_window, default=True),
                pystray.MenuItem("隐藏主窗口", self._tray_hide_main_window),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("退出程序", self._tray_quit_app),
            )
            self.tray_icon = pystray.Icon(
                "XiaoLinAssistant",
                self._load_tray_image(),
                APP_NAME,
                menu,
            )
            self.tray_thread = threading.Thread(target=self._run_tray_icon, daemon=True)
            self.tray_thread.start()
        except Exception as exc:
            self.tray_available = False
            self.tray_error = str(exc)
            self.tray_icon = None

    def _load_tray_image(self):
        for icon_path in (get_icon_path("png"), get_icon_path("ico")):
            if os.path.exists(icon_path):
                try:
                    return Image.open(icon_path).convert("RGBA")
                except Exception:
                    pass
        return Image.new("RGBA", (64, 64), (47, 111, 237, 255))

    def _run_tray_icon(self):
        try:
            if self.tray_icon:
                self.tray_icon.run()
        except Exception as exc:
            self.tray_available = False
            self.tray_error = str(exc)

    def _tray_show_main_window(self, icon=None, item=None):
        try:
            self.after(0, self.show_main_window)
        except Exception:
            pass

    def _tray_hide_main_window(self, icon=None, item=None):
        try:
            self.after(0, self.hide_to_tray)
        except Exception:
            pass

    def _tray_quit_app(self, icon=None, item=None):
        try:
            self.after(0, self.quit_app)
        except Exception:
            self.quit_app()

    def on_close_window(self):
        """Handle the window close button only. Minimize/taskbar events must stay normal."""
        if self.is_exiting:
            return
        dialog = getattr(self, "_close_dialog", None)
        try:
            if dialog is not None and dialog.winfo_exists():
                self._restore_window_for_modal()
                dialog.deiconify()
                dialog.lift()
                dialog.focus_force()
                return
        except Exception:
            self._close_dialog = None
        if self._close_prompt_active:
            return
        self._close_prompt_active = True
        try:
            choice = self._ask_close_behavior()
        finally:
            self._close_prompt_active = False
        if choice == "tray":
            self.hide_to_tray()
        elif choice == "exit":
            self.quit_app()

    def _restore_window_for_modal(self):
        """Make the owner visible before creating a modal child window.

        A modal dialog whose transient owner is minimized can become unreachable
        on Windows: the dialog owns the input grab while both taskbar thumbnails
        remain inactive. Restoring the owner first avoids that modal deadlock.
        """
        try:
            state = str(self.state()).lower()
        except Exception:
            state = ""
        try:
            if state in {"iconic", "withdrawn"} or not self.winfo_viewable():
                self.deiconify()
                self.update_idletasks()
                if str(self.state()).lower() == "iconic":
                    self.state("normal")
            self.lift()
            self.attributes("-topmost", True)
            self.update_idletasks()
            self.after(300, lambda: self.attributes("-topmost", False) if not self.is_exiting else None)
        except Exception:
            pass

    def _ask_close_behavior(self):
        """Ask whether to hide to tray, exit, or cancel closing."""
        result = {"choice": "cancel"}
        dialog = None
        self._restore_window_for_modal()
        try:
            dialog = ctk.CTkToplevel(self)
            self._close_dialog = dialog
            dialog.title("关闭测试助手")
            dialog.resizable(False, False)
            dialog.transient(self)
            dialog.grab_set()
            dialog.configure(fg_color=COLOR_BG)

            width, height = 430, 190
            try:
                self.update_idletasks()
                x = self.winfo_rootx() + max(0, int((self.winfo_width() - width) / 2))
                y = self.winfo_rooty() + max(0, int((self.winfo_height() - height) / 2))
                dialog.geometry(f"{width}x{height}+{x}+{y}")
            except Exception:
                dialog.geometry(f"{width}x{height}")

            body = ctk.CTkFrame(dialog, fg_color=COLOR_SURFACE, corner_radius=16, border_width=1, border_color=COLOR_BORDER)
            body.pack(fill="both", expand=True, padx=14, pady=14)
            msg = "是否隐藏到系统托盘？\n隐藏后可通过右下角托盘图标重新打开。"
            if not self.tray_available or self.tray_icon is None:
                msg = "系统托盘当前不可用，无法隐藏到托盘。\n你可以直接退出，或取消并保持窗口打开。"
            ctk.CTkLabel(
                body,
                text=msg,
                text_color=COLOR_TEXT,
                justify="left",
                wraplength=360,
                font=self._font(14),
            ).pack(anchor="w", padx=18, pady=(18, 12))

            buttons = ctk.CTkFrame(body, fg_color="transparent")
            buttons.pack(fill="x", padx=18, pady=(8, 16))
            buttons.grid_columnconfigure((0, 1, 2), weight=1)

            def choose(value):
                result["choice"] = value
                self._close_dialog = None
                try:
                    dialog.grab_release()
                except Exception:
                    pass
                dialog.destroy()

            ctk.CTkButton(
                buttons,
                text="隐藏到托盘",
                width=110,
                height=34,
                corner_radius=12,
                fg_color=COLOR_ACCENT,
                hover_color=COLOR_ACCENT_HOVER,
                text_color="white",
                state="normal" if self.tray_available and self.tray_icon is not None else "disabled",
                command=lambda: choose("tray"),
            ).grid(row=0, column=0, padx=(0, 8), sticky="ew")
            ctk.CTkButton(
                buttons,
                text="直接退出",
                width=110,
                height=34,
                corner_radius=12,
                fg_color=COLOR_DANGER,
                hover_color=COLOR_DANGER_HOVER,
                text_color="white",
                command=lambda: choose("exit"),
            ).grid(row=0, column=1, padx=8, sticky="ew")
            ctk.CTkButton(
                buttons,
                text="取消",
                width=110,
                height=34,
                corner_radius=12,
                fg_color=COLOR_SURFACE_2,
                hover_color=COLOR_HOVER,
                border_width=1,
                border_color=COLOR_BORDER,
                text_color=COLOR_TEXT,
                command=lambda: choose("cancel"),
            ).grid(row=0, column=2, padx=(8, 0), sticky="ew")

            dialog.protocol("WM_DELETE_WINDOW", lambda: choose("cancel"))
            dialog.bind("<Escape>", lambda _event: choose("cancel"))
            dialog.update_idletasks()
            dialog.attributes("-topmost", True)
            dialog.lift()
            dialog.focus_force()
            dialog.after(300, lambda: dialog.attributes("-topmost", False) if dialog.winfo_exists() else None)
            self.wait_window(dialog)
        except Exception:
            self._close_dialog = None
            if dialog is not None:
                try:
                    dialog.grab_release()
                except Exception:
                    pass
                try:
                    dialog.destroy()
                except Exception:
                    pass
            self._restore_window_for_modal()
            answer = messagebox.askyesnocancel(
                "关闭测试助手",
                "是否隐藏到系统托盘？\n是：隐藏到托盘\n否：直接退出\n取消：保持窗口打开",
                parent=self,
            )
            if answer is True:
                result["choice"] = "tray"
            elif answer is False:
                result["choice"] = "exit"
            else:
                result["choice"] = "cancel"
        finally:
            self._close_dialog = None
        return result["choice"]

    def hide_to_tray(self, show_tip=True):
        if self.is_exiting:
            return
        if not self.tray_available or self.tray_icon is None:
            msg = "系统托盘不可用"
            if self.tray_error:
                msg += f"：{self.tray_error}"
            msg += "\n是否直接退出程序？"
            if messagebox.askyesno("退出程序", msg):
                self.quit_app()
            return
        try:
            self.withdraw()
            if show_tip and not self._tray_hide_tip_shown:
                self._tray_hide_tip_shown = True
                try:
                    self.tray_icon.notify("测试助手已隐藏到托盘，右键可退出。", APP_NAME)
                except Exception:
                    pass
        except Exception as exc:
            messagebox.showerror("隐藏失败", str(exc))

    def show_main_window(self):
        if self.is_exiting:
            return
        try:
            self.deiconify()
            self.state("normal")
            self.lift()
            self.focus_force()
            self.attributes("-topmost", True)
            self.after(250, lambda: self.attributes("-topmost", False))
            dialog = getattr(self, "_close_dialog", None)
            if dialog is not None and dialog.winfo_exists():
                dialog.deiconify()
                dialog.lift()
                dialog.focus_force()
        except Exception:
            pass

    def _stop_tray_icon(self):
        icon = self.tray_icon
        self.tray_icon = None
        if icon is None:
            return
        try:
            icon.stop()
        except Exception:
            pass

    # ---------- Common ----------
    @staticmethod
    def _format_network_duration(seconds):
        total = max(0, int(seconds or 0))
        hours, remainder = divmod(total, 3600)
        minutes, seconds = divmod(remainder, 60)
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}" if hours else f"{minutes:02d}:{seconds:02d}"

    @staticmethod
    def _format_network_rate(bits_per_second):
        value = max(0.0, float(bits_per_second or 0))
        if value >= 1000000:
            return f"{value / 1000000:.2f} Mbps"
        if value >= 1000:
            return f"{value / 1000:.1f} Kbps"
        return f"{value:.0f} bps"

    @staticmethod
    def _format_network_bytes(byte_count):
        value = max(0.0, float(byte_count or 0))
        if value >= 1024 * 1024:
            return f"{value / (1024 * 1024):.1f} MB"
        if value >= 1024:
            return f"{value / 1024:.1f} KB"
        return f"{value:.0f} B"

    def _tick_stats(self):
        if self.is_exiting:
            return
        stats = self.network.stats_snapshot()
        duration = self._format_network_duration(stats["elapsed_seconds"])
        remaining = (
            self._format_network_duration(stats["remaining_seconds"])
            if stats["auto_stop_enabled"]
            else "--"
        )
        self.network_stats_var.set(
            f"捕获 {stats['captured']} ｜ 放行 {stats['sent']} ｜ 随机丢包 {stats['random_drop']} ｜ "
            f"突发丢包 {stats['burst_drop']} ｜ 乱序 {stats['reordered']}\n"
            f"限速等待 {stats['bandwidth_wait']} ｜ 队列 {stats['queue_count']}/峰值 {stats['peak_queue_count']} "
            f"（{self._format_network_bytes(stats['queue_bytes'])}）｜ 溢出 {stats['queue_overflow']} ｜ "
            f"发送失败 {stats['send_failure']} ｜ 停止取消 {stats['cancelled']}\n"
            f"上行 {self._format_network_rate(stats['upload_bps'])} ｜ "
            f"下行 {self._format_network_rate(stats['download_bps'])} ｜ "
            f"丢包率 {stats['loss_rate']:.2f}% ｜ 剩余 {remaining} ｜ 运行 {duration}"
        )
        self._sync_legacy_task_states()
        self._refresh_current_task_ui()

        self.after(500, self._tick_stats)

    def quit_app(self):
        if self.is_exiting:
            return
        self.is_exiting = True
        try:
            self.clicker.close()
        except Exception:
            pass
        try:
            self.network.stop()
        except Exception:
            pass
        try:
            self.adb_log.stop()
        except Exception:
            pass
        try:
            self.adb_tools.stop_recording_safely()
        except Exception:
            pass
        try:
            self.log_monitor.clear()
            self.crash_stream_detector.clear()
        except Exception:
            pass
        try:
            self.loc_stop_event.set()
        except Exception:
            pass
        try:
            self.cfg_stop_event.set()
        except Exception:
            pass
        try:
            self.ios_stop_log()
        except Exception:
            pass
        try:
            self._stop_tray_icon()
        except Exception:
            pass
        self.destroy()


if __name__ == "__main__":
    if os.name != "nt":
        messagebox.showerror("不支持的系统", "该工具仅支持 Windows")
        sys.exit(1)

    app = XiaoXinAssistant()
    app.mainloop()
