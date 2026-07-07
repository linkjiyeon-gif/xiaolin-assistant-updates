from __future__ import annotations

import ctypes
import os
import sys
from dataclasses import dataclass
from typing import Sequence


@dataclass
class AdminLaunchResult:
    ok: bool
    message: str
    code: int = 0


def is_windows_admin() -> bool:
    if os.name != "nt":
        return False
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def _quote(arg: str) -> str:
    text = str(arg or "")
    return '"' + text.replace('"', '\\"') + '"'


def build_relaunch_command(argv: Sequence[str] | None = None, frozen: bool | None = None) -> tuple[str, str]:
    """Return executable and parameter string for UAC relaunch.

    Frozen EXE: executable is the current EXE and params are original extra args.
    Source mode: executable is python.exe and params include script path + extra args.
    """
    args = list(argv if argv is not None else sys.argv)
    if frozen is None:
        frozen = bool(getattr(sys, "frozen", False))
    if frozen:
        executable = sys.executable
        params = " ".join(_quote(arg) for arg in args[1:])
    else:
        executable = sys.executable
        params = " ".join(_quote(arg) for arg in args)
    return executable, params


def relaunch_current_process_as_admin() -> AdminLaunchResult:
    if os.name != "nt":
        return AdminLaunchResult(False, "管理员重启仅支持 Windows", 0)
    if is_windows_admin():
        return AdminLaunchResult(False, "当前已是管理员权限", 0)
    try:
        executable, params = build_relaunch_command()
        cwd = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else sys.argv[0])) or None
        code = int(ctypes.windll.shell32.ShellExecuteW(None, "runas", executable, params, cwd, 1))
        if code > 32:
            return AdminLaunchResult(True, "已请求管理员权限重启", code)
        if code == 1223:
            return AdminLaunchResult(False, "用户取消了 UAC 授权", code)
        return AdminLaunchResult(False, f"管理员重启失败，ShellExecute 返回码：{code}", code)
    except Exception as exc:
        return AdminLaunchResult(False, f"管理员重启失败：{exc}", 0)
