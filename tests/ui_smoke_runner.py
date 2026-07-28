"""Repeatable real-window smoke test for the cached main pages.

Run locally with:
    python tests/ui_smoke_runner.py
"""

from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from xiaoxin_assistant import XiaoXinAssistant


PAGES = (
    "home",
    "clicker",
    "network",
    "adb_log",
    "log_analysis",
    "adb_tools",
    "ios_log",
    "compare",
    "localization",
    "config_validator",
    "time_tools",
)


def main() -> int:
    app = XiaoXinAssistant()
    app.withdraw()
    try:
        for page in PAGES:
            app.show_page(page)
            app.update_idletasks()
            print(f"OK {page}")
    finally:
        app.is_exiting = True
        app.destroy()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
