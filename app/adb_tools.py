from __future__ import annotations

import os
import re
import subprocess
import sys
import time
from dataclasses import asdict, dataclass
from typing import Dict, List, Optional, Sequence


DEFAULT_TIMEOUT = 12


ADB_DEVICE_COMMANDS = {
    "shell",
    "logcat",
    "install",
    "uninstall",
    "pull",
    "push",
    "reboot",
    "bugreport",
    "backup",
    "restore",
    "root",
    "unroot",
    "remount",
    "sideload",
    "forward",
    "reverse",
}


ADB_GLOBAL_COMMANDS = {
    "devices",
    "start-server",
    "kill-server",
    "version",
    "help",
    "connect",
    "disconnect",
}


def _creationflags() -> int:
    return subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0


def resource_path(relative_path: str) -> str:
    try:
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def safe_filename(text: str, fallback: str = "unknown") -> str:
    value = str(text or "").strip() or fallback
    return re.sub(r"[^0-9A-Za-z._\-\u4e00-\u9fff]+", "_", value).strip("_") or fallback


@dataclass
class CommandResult:
    command: List[str]
    returncode: int
    stdout: str = ""
    stderr: str = ""
    timed_out: bool = False
    serial: str = ""

    @property
    def ok(self) -> bool:
        return self.returncode == 0 and not self.timed_out

    @property
    def output(self) -> str:
        text = "\n".join(part for part in [self.stdout.strip(), self.stderr.strip()] if part)
        return text.strip()


@dataclass
class ADBDeviceInfo:
    serial: str
    status: str = ""
    model: str = ""
    product: str = ""
    device: str = ""
    transport_id: str = ""
    brand: str = ""
    android_version: str = ""
    sdk: str = ""
    foreground_package: str = ""
    foreground_activity: str = ""
    is_emulator: bool = False
    connection: str = "USB"
    last_refresh: str = ""

    @property
    def label(self) -> str:
        model = self.model or self.device or self.product or "未知设备"
        version = f"Android {self.android_version}" if self.android_version else "Android ?"
        return f"{self.serial}｜{model}｜{version}｜{self.status}"

    def to_dict(self) -> Dict[str, object]:
        data = asdict(self)
        data["label"] = self.label
        return data


@dataclass
class RecordingSession:
    process: subprocess.Popen
    device_path: str
    local_path: str
    serial: str


class ADBTools:
    """ADB helper with multi-device serial binding.

    All device-specific operations go through :meth:`run`, which injects
    ``adb -s <serial>`` automatically. In multi-device scenarios, callers must
    select a target serial before running device commands.
    """

    def __init__(self, base_dir: str):
        self.base_dir = base_dir
        self.screenshot_dir = os.path.join(base_dir, "Screenshots")
        self.recording_dir = os.path.join(base_dir, "Recordings")
        self._recording: Optional[RecordingSession] = None
        self.selected_serial: str = ""
        self._last_devices: List[ADBDeviceInfo] = []

    def adb_path(self) -> str:
        local_adb = resource_path("adb.exe")
        if os.path.exists(local_adb):
            return local_adb
        app_adb = os.path.join(self.base_dir, "adb.exe")
        if os.path.exists(app_adb):
            return app_adb
        return "adb"

    def _run_raw(self, args: Sequence[str], timeout: int = DEFAULT_TIMEOUT, serial: str = "") -> CommandResult:
        cmd = [self.adb_path()] + list(args)
        try:
            completed = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="ignore",
                timeout=timeout,
                creationflags=_creationflags(),
            )
            return CommandResult(cmd, completed.returncode, completed.stdout or "", completed.stderr or "", serial=serial)
        except FileNotFoundError as exc:
            return CommandResult(cmd, 127, "", f"未找到 adb.exe：{exc}", serial=serial)
        except subprocess.TimeoutExpired as exc:
            stdout = exc.stdout or ""
            stderr = exc.stderr or ""
            if isinstance(stdout, bytes):
                stdout = stdout.decode("utf-8", "ignore")
            if isinstance(stderr, bytes):
                stderr = stderr.decode("utf-8", "ignore")
            return CommandResult(cmd, 124, stdout, stderr or "命令执行超时", timed_out=True, serial=serial)
        except Exception as exc:
            return CommandResult(cmd, 1, "", str(exc), serial=serial)

    @staticmethod
    def _is_device_command(args: Sequence[str]) -> bool:
        if not args:
            return False
        first = str(args[0]).strip()
        if first in ADB_GLOBAL_COMMANDS:
            return False
        if first in ADB_DEVICE_COMMANDS:
            return True
        # Conservative default: unknown adb subcommands that are not explicitly global
        # should not run against an arbitrary device in multi-device environments.
        return True

    def run(self, args: List[str], timeout: int = DEFAULT_TIMEOUT, serial: Optional[str] = None) -> CommandResult:
        if self._is_device_command(args):
            target = self.resolve_serial(serial)
            result = self._run_raw(["-s", target] + list(args), timeout=timeout, serial=target)
        else:
            result = self._run_raw(args, timeout=timeout)
        return self._normalize_adb_error(result)

    def run_adb(self, serial: Optional[str], args: List[str], timeout: int = DEFAULT_TIMEOUT) -> CommandResult:
        return self.run(args, timeout=timeout, serial=serial)

    def run_shell(self, shell_cmd: str, timeout: int = DEFAULT_TIMEOUT, serial: Optional[str] = None) -> CommandResult:
        return self.run(["shell", shell_cmd], timeout=timeout, serial=serial)

    @staticmethod
    def _normalize_adb_error(result: CommandResult) -> CommandResult:
        text = (result.output or "").lower()
        if "more than one device" in text:
            result.stderr = "检测到多台 Android 设备，请先选择目标设备后再执行操作。"
        elif "device unauthorized" in text or "unauthorized" in text:
            result.stderr = "设备未授权，请在手机上允许 USB 调试授权。"
        elif "device offline" in text or "offline" in text:
            result.stderr = "设备处于 offline 状态，请重新插拔设备或重启 ADB 服务。"
        elif "no devices" in text or "no device" in text:
            result.stderr = "未检测到可用 Android 设备，请检查连接、USB 调试和授权弹窗。"
        return result

    @staticmethod
    def _parse_devices_output(text: str) -> List[ADBDeviceInfo]:
        devices: List[ADBDeviceInfo] = []
        for line in (text or "").splitlines():
            raw = line.strip()
            if not raw or raw.startswith("List of devices"):
                continue
            parts = raw.split()
            if len(parts) < 2:
                continue
            info = ADBDeviceInfo(serial=parts[0], status=parts[1])
            for token in parts[2:]:
                if ":" not in token:
                    continue
                key, value = token.split(":", 1)
                if key == "model":
                    info.model = value
                elif key == "product":
                    info.product = value
                elif key == "device":
                    info.device = value
                elif key == "transport_id":
                    info.transport_id = value
            serial_lower = info.serial.lower()
            info.is_emulator = serial_lower.startswith("emulator-") or serial_lower.startswith("127.0.0.1:") or serial_lower.startswith("localhost:")
            info.connection = "TCP/IP" if ":" in info.serial else ("模拟器" if info.is_emulator else "USB")
            info.last_refresh = time.strftime("%Y-%m-%d %H:%M:%S")
            devices.append(info)
        return devices

    def list_devices(self, detailed: bool = True) -> Dict[str, object]:
        result = self._run_raw(["devices", "-l"], timeout=8)
        result = self._normalize_adb_error(result)
        devices = self._parse_devices_output(result.stdout)
        if detailed:
            for device in devices:
                if device.status == "device":
                    self._fill_device_details(device)
        self._last_devices = devices
        ready = [item for item in devices if item.status == "device"]
        if self.selected_serial and not any(item.serial == self.selected_serial and item.status == "device" for item in devices):
            self.selected_serial = ""
        if not self.selected_serial and len(ready) == 1:
            self.selected_serial = ready[0].serial
        return {
            "result": result,
            "devices": [device.to_dict() for device in devices],
            "selected_serial": self.selected_serial,
            "ready_count": len(ready),
        }

    def _fill_device_details(self, device: ADBDeviceInfo) -> None:
        def prop(name: str) -> str:
            res = self._run_raw(["-s", device.serial, "shell", "getprop", name], timeout=6, serial=device.serial)
            return (res.stdout or "").strip()

        device.model = device.model or prop("ro.product.model")
        device.brand = prop("ro.product.brand")
        device.android_version = prop("ro.build.version.release")
        device.sdk = prop("ro.build.version.sdk")
        fg = self.get_foreground_app(serial=device.serial, ensure=False)
        device.foreground_package = fg.get("package", "")
        device.foreground_activity = fg.get("activity", "")

    def format_devices_text(self, data: Optional[Dict[str, object]] = None) -> str:
        data = data or self.list_devices()
        result = data.get("result")
        devices = data.get("devices", []) or []
        lines: List[str] = []
        if result and not getattr(result, "ok", False):
            lines.append(getattr(result, "output", "") or "ADB 设备检测失败")
        if not devices:
            lines.append("未检测到设备。请检查 USB 调试、授权弹窗或 adb 配置。")
            return "\n".join(lines)
        selected = data.get("selected_serial") or self.selected_serial or "未选择"
        lines.append(f"当前选中设备：{selected}")
        lines.append("已连接设备：")
        for item in devices:
            prefix = "*" if item.get("serial") == selected else "-"
            lines.append(
                f"{prefix} {item.get('serial')} | {item.get('status')} | "
                f"{item.get('model') or item.get('device') or '未知型号'} | "
                f"Android {item.get('android_version') or '?'} | "
                f"{item.get('connection') or ''} | 前台：{item.get('foreground_package') or '-'}"
            )
        ready = [item for item in devices if item.get("status") == "device"]
        if len(ready) > 1 and not self.selected_serial:
            lines.append("提示：检测到多台 Android 设备，请先选择目标设备后再执行操作。")
        for item in devices:
            if item.get("status") == "unauthorized":
                lines.append(f"提示：{item.get('serial')} 未授权，请在手机上允许 USB 调试授权。")
            elif item.get("status") == "offline":
                lines.append(f"提示：{item.get('serial')} offline，请重新插拔设备或重启 ADB 服务。")
        return "\n".join(lines)

    def set_selected_serial(self, serial: str) -> None:
        serial = (serial or "").strip()
        self.selected_serial = serial

    def get_selected_serial(self) -> str:
        return self.selected_serial

    def resolve_serial(self, serial: Optional[str] = None) -> str:
        serial = (serial or self.selected_serial or "").strip()
        data = self.list_devices(detailed=False)
        result = data.get("result")
        if result and not getattr(result, "ok", False):
            raise RuntimeError(getattr(result, "output", "") or "ADB 设备检测失败")
        devices = data.get("devices", []) or []
        if serial:
            for item in devices:
                if item.get("serial") == serial:
                    status = item.get("status")
                    if status == "device":
                        return serial
                    if status == "unauthorized":
                        raise RuntimeError("设备未授权，请在手机上允许 USB 调试授权。")
                    if status == "offline":
                        raise RuntimeError("设备处于 offline 状态，请重新插拔设备或重启 ADB 服务。")
                    raise RuntimeError(f"目标设备不可用：{serial}={status}")
            raise RuntimeError("设备 serial 已变化，请刷新设备列表后重新选择。")
        ready = [item for item in devices if item.get("status") == "device"]
        if not devices:
            raise RuntimeError("未检测到设备，请检查 USB 调试、授权弹窗或 adb 配置。")
        if not ready:
            status_text = ", ".join(f"{item.get('serial')}={item.get('status')}" for item in devices)
            if any(item.get("status") == "unauthorized" for item in devices):
                raise RuntimeError("设备未授权，请在手机上允许 USB 调试授权。")
            if any(item.get("status") == "offline" for item in devices):
                raise RuntimeError("设备处于 offline 状态，请重新插拔设备或重启 ADB 服务。")
            raise RuntimeError(f"未检测到可用 device 状态设备：{status_text}")
        if len(ready) == 1:
            self.selected_serial = ready[0].get("serial", "")
            return self.selected_serial
        raise RuntimeError("检测到多台 Android 设备，请先选择目标设备后再执行操作。")

    def has_ready_device(self) -> bool:
        data = self.list_devices(detailed=False)
        return any(item.get("status") == "device" for item in data.get("devices", []))

    def ensure_ready_device(self) -> None:
        self.resolve_serial()

    def get_prop(self, prop: str, serial: Optional[str] = None) -> str:
        return self.run(["shell", "getprop", prop], timeout=8, serial=serial).stdout.strip()

    _FOREGROUND_LINE_KEYS = (
        "mCurrentFocus",
        "mFocusedApp",
        "topResumedActivity",
        "ResumedActivity",
        "mResumedActivity",
        "ACTIVITY",
    )

    _COMPONENT_RE = re.compile(
        r"(?P<package>[A-Za-z0-9_]+(?:\.[A-Za-z0-9_]+)+)/(?P<activity>[^\s}\])]+)"
    )

    @staticmethod
    def _clean_activity_name(activity: str) -> str:
        return (activity or "").strip().strip("}])>").strip()

    @classmethod
    def _parse_component_from_line(cls, line: str) -> Optional[Dict[str, str]]:
        if not line or "null" in line.lower() and "/" not in line:
            return None
        match = cls._COMPONENT_RE.search(line)
        if not match:
            return None
        package = match.group("package").strip()
        activity = cls._clean_activity_name(match.group("activity"))
        if not package or not activity:
            return None
        if activity.startswith("."):
            activity = f"{package}{activity}"
        full_activity = f"{package}/{activity}"
        source = "dumpsys"
        for key in cls._FOREGROUND_LINE_KEYS:
            if key in line:
                source = key
                break
        return {
            "package": package,
            "activity": activity,
            "full_activity": full_activity,
            "source": source,
            "raw": line.strip(),
        }

    @classmethod
    def parse_foreground_app_text(cls, text: str) -> Optional[Dict[str, str]]:
        if not text:
            return None
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        for key in cls._FOREGROUND_LINE_KEYS:
            for line in lines:
                if key in line:
                    parsed = cls._parse_component_from_line(line)
                    if parsed:
                        return parsed
        for line in lines:
            parsed = cls._parse_component_from_line(line)
            if parsed:
                parsed["source"] = parsed.get("source") or "fallback"
                return parsed
        return None

    def get_foreground_app(self, serial: Optional[str] = None, ensure: bool = True) -> Dict[str, str]:
        target = self.resolve_serial(serial) if ensure else (serial or "")
        if not target:
            return {"package": "", "activity": "", "full_activity": "", "source": "未识别", "command": "", "raw": ""}
        attempts = [
            ("dumpsys window", ["shell", "dumpsys", "window"], 10),
            ("dumpsys activity activities", ["shell", "dumpsys", "activity", "activities"], 10),
            ("dumpsys activity top", ["shell", "dumpsys", "activity", "top"], 10),
        ]
        outputs = []
        for command_name, args, timeout in attempts:
            result = self.run(args, timeout=timeout, serial=target)
            text = result.stdout or result.stderr or ""
            outputs.append(f"[{command_name}]\n{text[:1200].strip()}")
            parsed = self.parse_foreground_app_text(text)
            if parsed:
                parsed["command"] = command_name
                parsed["serial"] = target
                return parsed
        summary = "\n\n".join(part for part in outputs if part.strip())
        return {
            "serial": target,
            "package": "",
            "activity": "",
            "full_activity": "",
            "source": "未识别",
            "command": "",
            "raw": summary[:1800].strip(),
        }

    def get_device_info(self, serial: Optional[str] = None) -> Dict[str, str]:
        target = self.resolve_serial(serial)
        info = {
            "设备 serial": target,
            "设备型号": self.get_prop("ro.product.model", serial=target),
            "设备品牌": self.get_prop("ro.product.brand", serial=target),
            "Android 版本": self.get_prop("ro.build.version.release", serial=target),
            "SDK 版本": self.get_prop("ro.build.version.sdk", serial=target),
            "CPU ABI": self.get_prop("ro.product.cpu.abi", serial=target),
        }
        size = self.run(["shell", "wm", "size"], timeout=8, serial=target).stdout.strip()
        info["分辨率"] = size.replace("Physical size:", "").strip() if size else ""
        battery = self.run(["shell", "dumpsys", "battery"], timeout=8, serial=target).stdout
        level = re.search(r"level:\s*(\d+)", battery)
        status = re.search(r"status:\s*(\d+)", battery)
        info["电量"] = (level.group(1) + "%") if level else ""
        info["电池状态码"] = status.group(1) if status else ""
        fg = self.get_foreground_app(serial=target)
        info["前台包名"] = fg.get("package", "")
        info["前台 Activity"] = fg.get("activity", "")
        return info

    def _serial_suffix(self, serial: str) -> str:
        return safe_filename(serial, "device")

    def take_screenshot(self, serial: Optional[str] = None) -> str:
        target = self.resolve_serial(serial)
        os.makedirs(self.screenshot_dir, exist_ok=True)
        name = f"screenshot_{self._serial_suffix(target)}_{time.strftime('%Y%m%d_%H%M%S')}.png"
        local_path = os.path.join(self.screenshot_dir, name)
        device_path = f"/sdcard/{name}"
        result = self.run(["shell", "screencap", "-p", device_path], timeout=15, serial=target)
        if not result.ok:
            raise RuntimeError(result.output or "截图失败")
        pull = self.run(["pull", device_path, local_path], timeout=30, serial=target)
        self.run(["shell", "rm", device_path], timeout=8, serial=target)
        if not pull.ok:
            raise RuntimeError(pull.output or "截图拉取失败")
        if not os.path.exists(local_path):
            raise RuntimeError("截图文件未生成")
        return local_path

    def start_recording(self, serial: Optional[str] = None) -> RecordingSession:
        target = self.resolve_serial(serial)
        if self._recording and self._recording.process.poll() is None:
            raise RuntimeError("录屏已在进行中")
        os.makedirs(self.recording_dir, exist_ok=True)
        name = f"record_{self._serial_suffix(target)}_{time.strftime('%Y%m%d_%H%M%S')}.mp4"
        local_path = os.path.join(self.recording_dir, name)
        device_path = f"/sdcard/{name}"
        cmd = [self.adb_path(), "-s", target, "shell", "screenrecord", device_path]
        try:
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="ignore",
                creationflags=_creationflags(),
            )
        except FileNotFoundError as exc:
            raise RuntimeError(f"未找到 adb.exe：{exc}") from exc
        except Exception as exc:
            raise RuntimeError(f"启动录屏失败：{exc}") from exc
        time.sleep(0.5)
        if process.poll() is not None:
            output = ""
            try:
                output = process.stdout.read() if process.stdout else ""
            except Exception:
                output = ""
            raise RuntimeError(output.strip() or "启动录屏失败，请检查设备是否支持 screenrecord")
        self._recording = RecordingSession(process=process, device_path=device_path, local_path=local_path, serial=target)
        return self._recording

    def stop_recording(self) -> str:
        session = self._recording
        if not session:
            return ""
        if session.process.poll() is None:
            try:
                session.process.terminate()
                session.process.wait(timeout=5)
            except Exception:
                try:
                    session.process.kill()
                except Exception:
                    pass
        time.sleep(0.6)
        pull = self.run(["pull", session.device_path, session.local_path], timeout=60, serial=session.serial)
        self.run(["shell", "rm", session.device_path], timeout=8, serial=session.serial)
        self._recording = None
        if not pull.ok:
            raise RuntimeError(pull.output or "录屏拉取失败")
        if not os.path.exists(session.local_path):
            raise RuntimeError("录屏文件未生成")
        return session.local_path

    def is_recording(self) -> bool:
        return bool(self._recording and self._recording.process.poll() is None)

    def stop_recording_safely(self) -> None:
        try:
            if self._recording:
                self.stop_recording()
        except Exception:
            self._recording = None

    def install_apk(self, apk_path: str, serial: Optional[str] = None) -> CommandResult:
        return self.run(["install", "-r", apk_path], timeout=180, serial=serial)

    def start_app(self, package: str, serial: Optional[str] = None) -> CommandResult:
        return self.run(["shell", "monkey", "-p", package, "-c", "android.intent.category.LAUNCHER", "1"], timeout=15, serial=serial)

    def stop_app(self, package: str, serial: Optional[str] = None) -> CommandResult:
        return self.run(["shell", "am", "force-stop", package], timeout=15, serial=serial)

    def restart_app(self, package: str, serial: Optional[str] = None) -> CommandResult:
        stop = self.stop_app(package, serial=serial)
        if not stop.ok:
            return stop
        time.sleep(0.5)
        return self.start_app(package, serial=serial)

    def clear_app_data(self, package: str, serial: Optional[str] = None) -> CommandResult:
        return self.run(["shell", "pm", "clear", package], timeout=30, serial=serial)

    def uninstall_app(self, package: str, serial: Optional[str] = None) -> CommandResult:
        return self.run(["uninstall", package], timeout=120, serial=serial)

    def open_dir(self, path: str) -> None:
        os.makedirs(path, exist_ok=True)
        if os.name == "nt":
            os.startfile(path)  # type: ignore[attr-defined]
        else:
            subprocess.Popen(["xdg-open", path])
