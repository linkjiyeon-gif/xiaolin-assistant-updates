import json
import os
import sys


def _load_version_info():
    if getattr(sys, "frozen", False):
        base_dir = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    else:
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    path = os.path.join(base_dir, "version.json")
    try:
        with open(path, "r", encoding="utf-8") as file:
            data = json.load(file)
        return str(data["version"]).strip(), int(data["build"])
    except Exception:
        return "0.0.0", 0


APP_NAME = "测试助手"
APP_SUBTITLE = "QA 测试工具箱"
APP_VERSION, APP_BUILD = _load_version_info()

SUPPORTED_EXTENSIONS = {
    ".txt", ".md", ".log", ".json", ".csv", ".xlsx", ".xml", ".yaml", ".yml",
    ".py", ".js", ".ts", ".java", ".cs", ".cpp", ".c", ".h", ".lua",
    ".ini", ".cfg", ".conf", ".properties", ".docx", ".pdf"
}

TEXT_EXTENSIONS = {
    ".txt", ".md", ".log", ".py", ".js", ".ts", ".java", ".cs", ".cpp", ".c", ".h",
    ".lua", ".ini", ".cfg", ".conf", ".properties"
}

COMPARE_MODES = ["普通文本 / 文件比对", "策划文档 vs 配置文件", "数值参考表 vs 游戏配置表", "翻译源文件 vs 游戏翻译配置表"]
