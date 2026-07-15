import hashlib
import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from app.constants import APP_BUILD, APP_VERSION
from app.update_manager import (
    UpdateInfo,
    download_installer,
    fetch_update_info,
    is_newer_version,
    verify_installer,
)


class FakeResponse:
    def __init__(self, data: bytes, url: str, content_type: str = "application/octet-stream"):
        self._data = data
        self._offset = 0
        self._url = url
        self.headers = {
            "Content-Length": str(len(data)),
            "Content-Type": content_type,
        }

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        return False

    def read(self, size=-1):
        if size is None or size < 0:
            size = len(self._data) - self._offset
        chunk = self._data[self._offset:self._offset + size]
        self._offset += len(chunk)
        return chunk

    def geturl(self):
        return self._url


class UpdateManagerTests(unittest.TestCase):
    def test_single_version_source(self):
        version_file = Path(__file__).resolve().parents[1] / "version.json"
        expected = json.loads(version_file.read_text(encoding="utf-8"))
        self.assertEqual(expected["version"], APP_VERSION)
        self.assertEqual(expected["build"], APP_BUILD)

    def test_version_comparison(self):
        self.assertTrue(is_newer_version("2.4.33", "2.4.32"))
        self.assertTrue(is_newer_version("2.5.0", "2.4.99"))
        self.assertFalse(is_newer_version("2.4.32", "2.4.32"))
        self.assertFalse(is_newer_version("2.4.31", "2.4.32"))

    def test_fetch_update_info(self):
        payload = {
            "version": "2.4.33",
            "build": 2433,
            "published_at": "2026-07-06",
            "installer_url": (
                "https://github.com/linkjiyeon-gif/xiaolin-assistant-updates/"
                "releases/download/v2.4.33/XiaoLinAssistant_Setup_v2.4.33.exe"
            ),
            "size": 123,
            "sha256": "a" * 64,
            "release_notes": ["修复问题", "提升稳定性"],
        }
        response = FakeResponse(
            json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            "https://raw.githubusercontent.com/linkjiyeon-gif/xiaolin-assistant-updates/main/latest.json",
            "application/json",
        )
        with mock.patch("app.update_manager.urllib.request.urlopen", return_value=response):
            info = fetch_update_info()
        self.assertEqual("2.4.33", info.version)
        self.assertEqual(2433, info.build)
        self.assertEqual(("修复问题", "提升稳定性"), info.release_notes)

    def test_download_and_verify_installer(self):
        data = b"verified installer bytes"
        digest = hashlib.sha256(data).hexdigest()
        info = UpdateInfo(
            version="2.4.33",
            build=2433,
            published_at="2026-07-06",
            installer_url=(
                "https://github.com/linkjiyeon-gif/xiaolin-assistant-updates/"
                "releases/download/v2.4.33/XiaoLinAssistant_Setup_v2.4.33.exe"
            ),
            size=len(data),
            sha256=digest,
            release_notes=(),
        )
        response = FakeResponse(
            data,
            "https://release-assets.githubusercontent.com/example/XiaoLinAssistant_Setup_v2.4.33.exe",
        )
        progress = []
        with tempfile.TemporaryDirectory() as folder:
            with (
                mock.patch("app.update_manager.update_download_dir", return_value=folder),
                mock.patch("app.update_manager.urllib.request.urlopen", return_value=response),
            ):
                path = download_installer(info, progress_callback=lambda done, total: progress.append((done, total)))
            self.assertTrue(os.path.isfile(path))
            verify_installer(path, info)
            with open(path, "rb") as file:
                self.assertEqual(data, file.read())
            self.assertTrue(progress)


if __name__ == "__main__":
    unittest.main()
