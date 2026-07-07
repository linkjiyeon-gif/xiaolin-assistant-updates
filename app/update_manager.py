from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from threading import Event
from typing import Callable, Optional


UPDATE_MANIFEST_URL = (
    "https://raw.githubusercontent.com/"
    "linkjiyeon-gif/xiaolin-assistant-updates/main/latest.json"
)
MAX_MANIFEST_BYTES = 1024 * 1024
ALLOWED_DOWNLOAD_HOSTS = {
    "github.com",
    "objects.githubusercontent.com",
    "release-assets.githubusercontent.com",
}


class UpdateError(RuntimeError):
    pass


class DownloadCancelled(UpdateError):
    pass


@dataclass(frozen=True)
class UpdateInfo:
    version: str
    build: int
    published_at: str
    installer_url: str
    size: int
    sha256: str
    release_notes: tuple[str, ...]

    @property
    def display_notes(self) -> str:
        return "\n".join(f"• {item}" for item in self.release_notes) or "本次更新未提供详细说明。"


def _version_key(value: str) -> tuple[int, ...]:
    numbers = [int(part) for part in re.findall(r"\d+", str(value or ""))]
    return tuple((numbers + [0, 0, 0, 0])[:4])


def is_newer_version(remote_version: str, local_version: str) -> bool:
    return _version_key(remote_version) > _version_key(local_version)


def _validate_https_url(value: str, allowed_hosts: Optional[set[str]] = None) -> str:
    url = str(value or "").strip()
    parsed = urllib.parse.urlparse(url)
    if parsed.scheme.lower() != "https" or not parsed.hostname:
        raise UpdateError("更新地址必须使用有效的 HTTPS 链接")
    if allowed_hosts and parsed.hostname.lower() not in allowed_hosts:
        raise UpdateError(f"更新下载域名不在允许列表中：{parsed.hostname}")
    return url


def _request(url: str) -> urllib.request.Request:
    return urllib.request.Request(
        url,
        headers={
            "Accept": "application/json",
            "Cache-Control": "no-cache",
            "User-Agent": "XiaoLinAssistant-Updater/1.0",
        },
    )


def fetch_update_info(
    manifest_url: str = UPDATE_MANIFEST_URL,
    timeout: int = 12,
) -> UpdateInfo:
    manifest_url = _validate_https_url(manifest_url)
    separator = "&" if "?" in manifest_url else "?"
    request = _request(f"{manifest_url}{separator}_={int(time.time())}")
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            raw = response.read(MAX_MANIFEST_BYTES + 1)
    except Exception as exc:
        raise UpdateError(f"无法获取更新信息：{exc}") from exc
    if len(raw) > MAX_MANIFEST_BYTES:
        raise UpdateError("更新清单超过允许大小")
    try:
        data = json.loads(raw.decode("utf-8-sig"))
    except Exception as exc:
        raise UpdateError("更新清单不是有效的 UTF-8 JSON") from exc

    try:
        version = str(data["version"]).strip()
        build = int(data.get("build", 0))
        installer_url = _validate_https_url(str(data["installer_url"]), ALLOWED_DOWNLOAD_HOSTS)
        size = int(data["size"])
        sha256 = str(data["sha256"]).strip().lower()
    except (KeyError, TypeError, ValueError) as exc:
        raise UpdateError("更新清单缺少必要字段") from exc
    if not re.fullmatch(r"\d+(?:\.\d+){2,3}", version):
        raise UpdateError("更新清单中的版本号格式无效")
    if size <= 0:
        raise UpdateError("更新清单中的安装包大小无效")
    if not re.fullmatch(r"[0-9a-f]{64}", sha256):
        raise UpdateError("更新清单中的 SHA-256 无效")

    notes_value = data.get("release_notes", [])
    if isinstance(notes_value, str):
        notes = (notes_value.strip(),) if notes_value.strip() else ()
    elif isinstance(notes_value, list):
        notes = tuple(str(item).strip() for item in notes_value if str(item).strip())
    else:
        notes = ()
    return UpdateInfo(
        version=version,
        build=build,
        published_at=str(data.get("published_at", "")).strip(),
        installer_url=installer_url,
        size=size,
        sha256=sha256,
        release_notes=notes,
    )


def file_sha256(path: str) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as file:
        while True:
            chunk = file.read(1024 * 1024)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def verify_installer(path: str, info: UpdateInfo) -> None:
    if not os.path.isfile(path):
        raise UpdateError("下载的安装包不存在")
    actual_size = os.path.getsize(path)
    if actual_size != info.size:
        raise UpdateError(f"安装包大小校验失败：期望 {info.size}，实际 {actual_size}")
    actual_hash = file_sha256(path)
    if actual_hash.lower() != info.sha256.lower():
        raise UpdateError("安装包 SHA-256 校验失败，已阻止安装")


def update_download_dir() -> str:
    base = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
    path = os.path.join(base, "QA测试工具盒", "Update_Data")
    os.makedirs(path, exist_ok=True)
    return path


def download_installer(
    info: UpdateInfo,
    cancel_event: Optional[Event] = None,
    progress_callback: Optional[Callable[[int, int], None]] = None,
    timeout: int = 30,
) -> str:
    destination = os.path.join(update_download_dir(), f"XiaoLinAssistant_Setup_v{info.version}.exe")
    partial = destination + ".part"
    if os.path.isfile(destination):
        try:
            verify_installer(destination, info)
            if progress_callback:
                progress_callback(info.size, info.size)
            return destination
        except UpdateError:
            try:
                os.remove(destination)
            except OSError:
                pass

    request = urllib.request.Request(
        info.installer_url,
        headers={"User-Agent": "XiaoLinAssistant-Updater/1.0"},
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            final_url = response.geturl()
            _validate_https_url(final_url, ALLOWED_DOWNLOAD_HOSTS)
            total = int(response.headers.get("Content-Length") or info.size)
            downloaded = 0
            with open(partial, "wb") as file:
                while True:
                    if cancel_event and cancel_event.is_set():
                        raise DownloadCancelled("用户已取消下载")
                    chunk = response.read(1024 * 256)
                    if not chunk:
                        break
                    file.write(chunk)
                    downloaded += len(chunk)
                    if progress_callback:
                        progress_callback(downloaded, total)
        os.replace(partial, destination)
        verify_installer(destination, info)
        return destination
    except DownloadCancelled:
        try:
            os.remove(partial)
        except OSError:
            pass
        raise
    except Exception as exc:
        try:
            os.remove(partial)
        except OSError:
            pass
        if isinstance(exc, UpdateError):
            raise
        raise UpdateError(f"下载安装包失败：{exc}") from exc


def launch_update_helper(installer_path: str, target_version: str) -> None:
    if os.name != "nt":
        raise UpdateError("在线安装目前仅支持 Windows")
    app_path = os.path.abspath(sys.executable)
    app_dir = os.path.dirname(app_path)
    updater_source = os.path.join(app_dir, "XiaoLinUpdater.exe")
    if not getattr(sys, "frozen", False) or not os.path.isfile(updater_source):
        raise UpdateError("当前为源码运行或缺少 XiaoLinUpdater.exe，无法执行自动安装")

    helper_dir = os.path.join(tempfile.gettempdir(), "XiaoLinAssistant_Update_Helper")
    os.makedirs(helper_dir, exist_ok=True)
    helper_path = os.path.join(helper_dir, f"XiaoLinUpdater_{target_version}.exe")
    shutil.copy2(updater_source, helper_path)
    creation_flags = (
        getattr(subprocess, "CREATE_NO_WINDOW", 0)
        | getattr(subprocess, "DETACHED_PROCESS", 0)
        | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    )
    subprocess.Popen(
        [
            helper_path,
            "--pid",
            str(os.getpid()),
            "--installer",
            os.path.abspath(installer_path),
            "--app",
            app_path,
            "--version",
            target_version,
        ],
        cwd=app_dir,
        close_fds=True,
        creationflags=creation_flags,
    )
