from __future__ import annotations

import tempfile
import unittest
from unittest.mock import patch

from app.adb_tools import ADBTools, CommandResult, parse_device_resolution


def command_result(args, stdout="", stderr="", returncode=0, *, timed_out=False, serial=""):
    return CommandResult(
        command=["adb", *list(args)],
        returncode=returncode,
        stdout=stdout,
        stderr=stderr,
        timed_out=timed_out,
        serial=serial,
    )


class ADBResolutionTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.tools = ADBTools(self.temp_dir.name)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_parse_physical_and_override_resolution(self):
        self.assertEqual(parse_device_resolution("Physical size: 1080x2400"), "1080 × 2400")
        self.assertEqual(
            parse_device_resolution("Physical size: 1080x2400\nOverride size: 720x1600"),
            "720 × 1600",
        )

    def test_parse_resolution_ignores_case_and_extra_spaces(self):
        self.assertEqual(parse_device_resolution("  pHySiCaL   SIZE :  1440 X 3200  "), "1440 × 3200")
        self.assertEqual(parse_device_resolution("override size = 900 × 2000"), "900 × 2000")

    def test_parse_common_dumpsys_display_formats(self):
        samples = {
            "mBaseDisplayInfo=DisplayInfo{real 1440 x 3120, app 1440 x 2992}": "1440 × 3120",
            "logicalWidth=1080, logicalHeight=2400": "1080 × 2400",
            "mCurrentDisplayRect=Rect(0, 0 - 720, 1600)": "720 × 1600",
            "resolution: 1200x1920": "1200 × 1920",
        }
        for raw, expected in samples.items():
            with self.subTest(raw=raw):
                self.assertEqual(parse_device_resolution(raw), expected)

    def test_wm_failure_uses_dumpsys_fallback(self):
        calls = []

        def fake_run(args, timeout=12, serial=""):
            calls.append((list(args), serial, timeout))
            if args[-2:] == ["wm", "size"]:
                return command_result(args, stderr="wm: inaccessible", returncode=1, serial=serial)
            return command_result(args, stdout="DisplayDeviceInfo{real 1080 x 2340}", serial=serial)

        with patch.object(self.tools, "_run_raw", side_effect=fake_run):
            self.assertEqual(self.tools.get_device_resolution("SERIAL-A"), "1080 × 2340")

        self.assertEqual(calls[0][0][:2], ["-s", "SERIAL-A"])
        self.assertEqual(calls[1][0][:2], ["-s", "SERIAL-A"])
        self.assertEqual(calls[1][0][-2:], ["dumpsys", "display"])

    def test_timeout_returns_failure_and_is_cached_until_refresh(self):
        calls = []

        def fake_run(args, timeout=12, serial=""):
            calls.append(list(args))
            return command_result(args, stderr="命令执行超时", returncode=124, timed_out=True, serial=serial)

        with patch.object(self.tools, "_run_raw", side_effect=fake_run):
            self.assertEqual(self.tools.get_device_resolution("SERIAL-T"), "获取失败")
            self.assertEqual(self.tools.get_device_resolution("SERIAL-T"), "获取失败")
            self.assertEqual(len(calls), 2)  # wm + dumpsys; second read uses cache.
            self.assertEqual(self.tools.get_device_resolution("SERIAL-T", force_refresh=True), "获取失败")
            self.assertEqual(len(calls), 4)

    def test_unauthorized_and_offline_devices_do_not_query_details(self):
        output = "List of devices attached\nUNAUTH unauthorized\nOFFLINE offline\n"
        with patch.object(
            self.tools,
            "_run_raw",
            return_value=command_result(["devices", "-l"], stdout=output),
        ), patch.object(self.tools, "_fill_device_details") as fill_details:
            data = self.tools.list_devices(refresh_resolutions=True)

        fill_details.assert_not_called()
        self.assertEqual([item["resolution"] for item in data["devices"]], ["未知", "未知"])

    def test_multiple_devices_are_always_queried_by_their_own_serial(self):
        calls = []

        def fake_run(args, timeout=12, serial=""):
            calls.append((list(args), serial))
            output = "Physical size: 1080x2400" if serial == "SERIAL-A" else "Physical size: 720x1600"
            return command_result(args, stdout=output, serial=serial)

        with patch.object(self.tools, "_run_raw", side_effect=fake_run):
            self.assertEqual(self.tools.get_device_resolution("SERIAL-A"), "1080 × 2400")
            self.assertEqual(self.tools.get_device_resolution("SERIAL-B"), "720 × 1600")

        self.assertEqual([serial for _, serial in calls], ["SERIAL-A", "SERIAL-B"])
        self.assertTrue(all(args[:2] == ["-s", serial] for args, serial in calls))

    def test_disconnect_clears_cache_and_reconnect_refreshes(self):
        scans = iter(
            [
                "List of devices attached\nSERIAL-A device\nSERIAL-B device\n",
                "List of devices attached\nSERIAL-A device\n",
                "List of devices attached\nSERIAL-A device\nSERIAL-B device\n",
            ]
        )
        wm_calls = []

        def fake_run(args, timeout=12, serial=""):
            if list(args) == ["devices", "-l"]:
                return command_result(args, stdout=next(scans))
            wm_calls.append(serial)
            output = "Physical size: 1080x2400" if serial == "SERIAL-A" else "Physical size: 720x1600"
            return command_result(args, stdout=output, serial=serial)

        def fill_details(device, refresh_resolution=False):
            device.resolution = self.tools.get_device_resolution(device.serial, force_refresh=refresh_resolution)

        with patch.object(self.tools, "_run_raw", side_effect=fake_run), patch.object(
            self.tools, "_fill_device_details", side_effect=fill_details
        ):
            first = self.tools.list_devices(refresh_resolutions=True)
            self.assertEqual([item["resolution"] for item in first["devices"]], ["1080 × 2400", "720 × 1600"])
            self.assertEqual(self.tools.get_device_resolution("SERIAL-B"), "720 × 1600")
            self.assertEqual(wm_calls.count("SERIAL-B"), 1)

            self.tools.list_devices(detailed=False)
            self.assertEqual(self.tools.get_cached_device_resolution("SERIAL-B"), "未知")

            third = self.tools.list_devices(refresh_resolutions=True)
            self.assertEqual(third["devices"][1]["resolution"], "720 × 1600")
            self.assertEqual(wm_calls.count("SERIAL-B"), 2)


if __name__ == "__main__":
    unittest.main()
