from __future__ import annotations

import argparse
import ctypes
import datetime as dt
import os
import subprocess
import sys
import time


def log_path() -> str:
    base = os.environ.get("LOCALAPPDATA") or os.path.dirname(os.path.abspath(__file__))
    folder = os.path.join(base, "QA测试工具盒", "Update_Logs")
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, "update.log")


def write_log(message: str) -> None:
    line = f"{dt.datetime.now():%Y-%m-%d %H:%M:%S} {message}\n"
    try:
        with open(log_path(), "a", encoding="utf-8") as file:
            file.write(line)
    except Exception:
        pass


def wait_for_process(pid: int, timeout_seconds: int = 180) -> bool:
    synchronize = 0x00100000
    wait_object_0 = 0x00000000
    handle = ctypes.windll.kernel32.OpenProcess(synchronize, False, int(pid))
    if not handle:
        return True
    try:
        result = ctypes.windll.kernel32.WaitForSingleObject(handle, timeout_seconds * 1000)
        return result == wait_object_0
    finally:
        ctypes.windll.kernel32.CloseHandle(handle)


def launch_app(path: str) -> None:
    if not os.path.isfile(path):
        write_log(f"主程序不存在，无法重启：{path}")
        return
    flags = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    subprocess.Popen(
        [path, "--post-update"],
        cwd=os.path.dirname(path),
        close_fds=True,
        creationflags=flags,
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pid", type=int, required=True)
    parser.add_argument("--installer", required=True)
    parser.add_argument("--app", required=True)
    parser.add_argument("--version", default="")
    args = parser.parse_args()

    write_log(f"开始更新到 v{args.version}，等待主程序 PID={args.pid} 退出")
    if not wait_for_process(args.pid):
        write_log("等待主程序退出超时，取消更新")
        return 10
    if not os.path.isfile(args.installer):
        write_log(f"安装包不存在：{args.installer}")
        return 11

    command = [
        args.installer,
        "/VERYSILENT",
        "/SUPPRESSMSGBOXES",
        "/NORESTART",
        "/CLOSEAPPLICATIONS",
        "/SP-",
    ]
    try:
        write_log("启动静默安装")
        completed = subprocess.run(command, check=False)
        write_log(f"安装程序退出码：{completed.returncode}")
    except Exception as exc:
        write_log(f"启动安装程序失败：{exc}")
        launch_app(os.path.abspath(args.app))
        return 12
    if completed.returncode != 0:
        write_log("安装未成功，重新启动现有主程序")
        launch_app(os.path.abspath(args.app))
        return completed.returncode or 13

    try:
        os.remove(args.installer)
    except OSError:
        pass
    time.sleep(1)
    launch_app(os.path.abspath(args.app))
    write_log("更新完成，已重新启动主程序")
    return 0


if __name__ == "__main__":
    if os.name != "nt":
        sys.exit(20)
    sys.exit(main())
