from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional

IOS_TOOL_NAMES = [
    "idevice_id.exe",
    "ideviceinfo.exe",
    "idevicesyslog.exe",
    "idevicecrashreport.exe",
]

REQUIRED_IOS_TOOL_NAMES = ["idevice_id.exe", "ideviceinfo.exe", "idevicesyslog.exe"]

DEFAULT_IOS_KEYWORDS = [
    "error",
    "exception",
    "crash",
    "fatal",
    "warning",
    "assert",
    "Unity",
    "Lua",
    "il2cpp",
    "objc",
    "terminated",
    "memory",
    "SIGABRT",
    "SIGSEGV",
]


@dataclass
class CommandResult:
    command: List[str]
    returncode: int
    stdout: str = ""
    stderr: str = ""
    timeout: bool = False

    @property
    def ok(self) -> bool:
        return self.returncode == 0 and not self.timeout

    @property
    def output(self) -> str:
        text = "\n".join(part for part in [self.stdout.strip(), self.stderr.strip()] if part)
        return text.strip()


@dataclass
class IOSToolSearchResult:
    name: str
    path: str = ""
    exists: bool = False
    source: str = "未找到"
    note: str = ""
    checked_locations: List[str] = field(default_factory=list)


@dataclass
class IOSDeviceInfo:
    udid: str
    name: str = ""
    version: str = ""
    product_type: str = ""
    state: str = "connected"
    raw: str = ""


@dataclass
class IOSLogFilter:
    keywords: List[str] = field(default_factory=list)
    text_filter: str = ""
    only_hits: bool = False

    def should_display(self, line: str) -> bool:
        text = line or ""
        lower = text.lower()
        text_filter = (self.text_filter or "").strip().lower()
        keyword_list = [k.strip() for k in self.keywords if str(k).strip()]
        text_ok = not text_filter or text_filter in lower
        keyword_ok = True
        if keyword_list:
            keyword_ok = any(k.lower() in lower for k in keyword_list)
        if self.only_hits:
            return text_ok and keyword_ok
        return text_ok

    def is_hit(self, line: str) -> bool:
        lower = (line or "").lower()
        return any(k.strip().lower() in lower for k in self.keywords if str(k).strip())


def safe_filename(text: str, default: str = "ios_device") -> str:
    cleaned = "".join(ch if ch.isalnum() or ch in "._-" else "_" for ch in (text or ""))
    return cleaned.strip("._-") or default


def run_command(command: List[str], timeout: int = 10) -> CommandResult:
    try:
        result = subprocess.run(
            [str(item) for item in command],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            timeout=timeout,
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
        )
        return CommandResult(command, result.returncode, result.stdout or "", result.stderr or "")
    except subprocess.TimeoutExpired as exc:
        return CommandResult(command, -1, exc.stdout or "", exc.stderr or "", timeout=True)
    except FileNotFoundError as exc:
        return CommandResult(command, -2, "", str(exc))
    except Exception as exc:
        return CommandResult(command, -3, "", str(exc))


def get_app_base_dir() -> Path:
    """Return the real folder where the running app/exe lives."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def get_project_root() -> Path:
    """Return source project root without relying on os.getcwd()."""
    return Path(__file__).resolve().parents[1]


def _norm_path(value: str | os.PathLike | None) -> str:
    if not value:
        return ""
    return str(Path(str(value)).expanduser())


def _is_file(path: str | os.PathLike | None) -> bool:
    if not path:
        return False
    try:
        return Path(str(path)).expanduser().is_file()
    except Exception:
        return False


class IOSLogTools:
    def __init__(self, app_base_dir: str | os.PathLike | None = None):
        # app_base_dir is kept for compatibility with older UI code, but detection no
        # longer relies only on it. The executable/source/cwd/PATH are all checked.
        self.app_base_dir = Path(app_base_dir).resolve() if app_base_dir else get_app_base_dir()
        self.exe_base_dir = get_app_base_dir()
        self.project_root = get_project_root()
        self.cwd_base_dir = Path.cwd().resolve()
        self.tools_dir = self.exe_base_dir / "tools" / "ios"
        self.config_path = self.app_base_dir / "ios_tools_config.json"
        self.log_dir = self.app_base_dir / "IOS_Logs"
        self.tool_paths: Dict[str, str] = {name: "" for name in IOS_TOOL_NAMES}
        self.saved_tool_paths: Dict[str, str] = {name: "" for name in IOS_TOOL_NAMES}
        self.load_config()

    def load_config(self) -> None:
        if not self.config_path.exists():
            return
        try:
            data = json.loads(self.config_path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                for name in IOS_TOOL_NAMES:
                    value = _norm_path(data.get(name, "") or "")
                    if value:
                        self.saved_tool_paths[name] = value
                        self.tool_paths[name] = value
        except Exception:
            pass

    def save_config(self) -> None:
        self.config_path.parent.mkdir(parents=True, exist_ok=True)
        clean_data = {name: _norm_path(self.tool_paths.get(name, "")) for name in IOS_TOOL_NAMES}
        self.config_path.write_text(json.dumps(clean_data, ensure_ascii=False, indent=2), encoding="utf-8")
        self.saved_tool_paths.update(clean_data)

    def set_tool_path(self, tool_name: str, path: str) -> None:
        if tool_name not in self.tool_paths:
            self.tool_paths[tool_name] = ""
        self.tool_paths[tool_name] = _norm_path(path)

    def _candidate_locations(self, tool_name: str) -> List[tuple[str, str]]:
        """Return path candidates and human-readable sources in priority order."""
        candidates: List[tuple[str, str]] = []

        configured = _norm_path(self.tool_paths.get(tool_name, ""))
        saved = _norm_path(self.saved_tool_paths.get(tool_name, ""))
        if configured:
            source = "本地配置" if saved and configured == saved else "手动配置"
            candidates.append((configured, source))
        if saved and saved != configured:
            candidates.append((saved, "本地配置"))

        # App base folder passed by the UI. In EXE/installer mode this is the EXE
        # folder; in source mode this is the project root. This is the most important
        # non-configured search location and fixes toolbox folders on D/E/USB drives.
        app_candidate = self.app_base_dir / "tools" / "ios" / tool_name
        candidates.append((str(app_candidate), "程序目录 tools/ios"))

        # Real executable folder: works after PyInstaller build, installer install,
        # and when the whole toolbox is moved to another drive.
        exe_candidate = self.exe_base_dir / "tools" / "ios" / tool_name
        if str(exe_candidate) not in [p for p, _ in candidates]:
            candidates.append((str(exe_candidate), "程序目录 tools/ios"))

        # Source project folder: works when running python xiaoxin_assistant.py.
        project_candidate = self.project_root / "tools" / "ios" / tool_name
        if str(project_candidate) not in [p for p, _ in candidates]:
            candidates.append((str(project_candidate), "源码目录 tools/ios"))

        # Compatibility fallback for users launching from a different working dir.
        cwd_candidate = self.cwd_base_dir / "tools" / "ios" / tool_name
        if str(cwd_candidate) not in [p for p, _ in candidates]:
            candidates.append((str(cwd_candidate), "当前目录 tools/ios"))

        which = shutil.which(tool_name)
        if which:
            candidates.append((which, "PATH"))
        if tool_name.lower().endswith(".exe"):
            which_no_ext = shutil.which(tool_name[:-4])
            if which_no_ext and which_no_ext not in [p for p, _ in candidates]:
                candidates.append((which_no_ext, "PATH"))
        return candidates

    def find_ios_tool(self, tool_name: str) -> IOSToolSearchResult:
        candidates = self._candidate_locations(tool_name)
        checked = [path for path, _ in candidates if path]
        configured = _norm_path(self.tool_paths.get(tool_name, ""))
        saved = _norm_path(self.saved_tool_paths.get(tool_name, ""))
        invalid_notes: List[str] = []
        if configured and not _is_file(configured):
            invalid_notes.append(f"已配置路径不存在：{configured}")
        if saved and saved != configured and not _is_file(saved):
            invalid_notes.append(f"本地配置路径失效：{saved}")

        for path, source in candidates:
            if _is_file(path):
                return IOSToolSearchResult(
                    name=tool_name,
                    path=str(Path(path).expanduser()),
                    exists=True,
                    source=source,
                    note="；".join(invalid_notes),
                    checked_locations=checked,
                )
        return IOSToolSearchResult(
            name=tool_name,
            path=configured or str(self.exe_base_dir / "tools" / "ios" / tool_name),
            exists=False,
            source="未找到",
            note="；".join(invalid_notes),
            checked_locations=checked,
        )

    def find_ios_tools(self) -> Dict[str, IOSToolSearchResult]:
        return {name: self.find_ios_tool(name) for name in IOS_TOOL_NAMES}

    def resolve_tool(self, tool_name: str) -> str:
        result = self.find_ios_tool(tool_name)
        return result.path

    def detect_tools(self) -> Dict[str, Dict[str, str | bool | List[str]]]:
        results: Dict[str, Dict[str, str | bool | List[str]]] = {}
        for name, item in self.find_ios_tools().items():
            results[name] = {
                "path": item.path,
                "exists": item.exists,
                "source": item.source,
                "note": item.note,
                "checked_locations": item.checked_locations,
            }
        return results

    def checked_locations_summary(self) -> str:
        dirs = [
            (self.app_base_dir / "tools" / "ios", "程序目录 tools/ios"),
            (self.exe_base_dir / "tools" / "ios", "程序目录 tools/ios"),
            (self.project_root / "tools" / "ios", "源码目录 tools/ios"),
            (self.cwd_base_dir / "tools" / "ios", "当前目录 tools/ios"),
        ]
        seen = set()
        rows = []
        for path, label in dirs:
            key = str(path)
            if key in seen:
                continue
            seen.add(key)
            rows.append(f"- {label}：{path}")
        rows.append("- PATH：系统环境变量中的 libimobiledevice 工具")
        return "\n".join(rows)

    def missing_required_tools(self) -> List[str]:
        results = self.find_ios_tools()
        return [name for name in REQUIRED_IOS_TOOL_NAMES if not results.get(name) or not results[name].exists]

    def list_device_udids(self) -> tuple[List[str], CommandResult]:
        exe = self.resolve_tool("idevice_id.exe")
        result = run_command([exe, "-l"], timeout=8)
        if not result.ok:
            return [], result
        udids = [line.strip() for line in result.stdout.splitlines() if line.strip()]
        return udids, result

    def get_device_info(self, udid: str) -> IOSDeviceInfo:
        exe = self.resolve_tool("ideviceinfo.exe")
        cmd = [exe]
        if udid:
            cmd += ["-u", udid]
        result = run_command(cmd, timeout=10)
        info = IOSDeviceInfo(udid=udid or "", raw=result.output)
        if not result.ok:
            info.state = "unavailable"
            return info
        parsed: Dict[str, str] = {}
        for line in result.stdout.splitlines():
            if ":" not in line:
                continue
            key, value = line.split(":", 1)
            parsed[key.strip()] = value.strip()
        info.name = parsed.get("DeviceName", "")
        info.version = parsed.get("ProductVersion", "")
        info.product_type = parsed.get("ProductType", "")
        info.state = "connected"
        return info

    def list_devices(self) -> tuple[List[IOSDeviceInfo], CommandResult]:
        udids, result = self.list_device_udids()
        devices: List[IOSDeviceInfo] = []
        if result.ok:
            for udid in udids:
                devices.append(self.get_device_info(udid))
        return devices, result

    def start_syslog_process(self, udid: str = "") -> subprocess.Popen:
        exe = self.resolve_tool("idevicesyslog.exe")
        if not _is_file(exe):
            raise FileNotFoundError(f"未找到 idevicesyslog.exe：{exe}")
        cmd = [exe]
        if udid:
            cmd += ["-u", udid]
        return subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="ignore",
            creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
        )

    def make_log_paths(self, udid: str = "") -> tuple[Path, Path]:
        self.log_dir.mkdir(parents=True, exist_ok=True)
        stamp = time.strftime("%Y%m%d_%H%M%S")
        safe_udid = safe_filename(udid, "ios_device")
        full_path = self.log_dir / f"{safe_udid}_{stamp}_full.txt"
        filtered_path = self.log_dir / f"{safe_udid}_{stamp}_filtered.txt"
        return full_path, filtered_path

    def open_log_dir(self) -> None:
        self.log_dir.mkdir(parents=True, exist_ok=True)
        if os.name == "nt":
            os.startfile(str(self.log_dir))

    def open_tools_dir(self) -> None:
        tools_dir = self.app_base_dir / "tools" / "ios"
        tools_dir.mkdir(parents=True, exist_ok=True)
        if os.name == "nt":
            os.startfile(str(tools_dir))


def parse_keywords(text: str) -> List[str]:
    keywords: List[str] = []
    seen = set()
    for line in str(text or "").replace(",", "\n").splitlines():
        item = line.strip()
        if not item or item.startswith("#"):
            continue
        low = item.lower()
        if low not in seen:
            seen.add(low)
            keywords.append(item)
    return keywords


def format_devices(devices: Iterable[IOSDeviceInfo], command_result: Optional[CommandResult] = None) -> str:
    devices = list(devices)
    if devices:
        rows = ["已检测到 iOS 设备："]
        for item in devices:
            rows.append(
                f"- UDID：{item.udid}\n  名称：{item.name or '未获取'}\n  系统：{item.version or '未获取'}\n  型号：{item.product_type or '未获取'}\n  状态：{item.state}"
            )
        return "\n".join(rows)
    msg = [
        "未检测到 iOS 设备。",
        "请检查：",
        "1. 请用数据线连接 iPhone / iPad。",
        "2. 请在设备上点击“信任此电脑”。",
        "3. 请确认已安装 Apple Mobile Device Support / iTunes 驱动。",
        "4. 请确认 libimobiledevice 工具路径配置正确。",
    ]
    if command_result and command_result.output:
        msg.append("\n命令输出：")
        msg.append(command_result.output)
    return "\n".join(msg)
