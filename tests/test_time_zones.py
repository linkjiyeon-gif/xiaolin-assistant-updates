from __future__ import annotations

import unittest
from datetime import datetime, timedelta, timezone

from app.time_tools import (
    TIMEZONE_OPTIONS,
    countdown,
    datetime_to_timestamps,
    parse_datetime_text,
    time_difference,
    timestamp_to_datetime,
)


def utc_of(text, zone):
    return parse_datetime_text(text, zone).astimezone(timezone.utc).strftime("%Y-%m-%d %H:%M")


class NamedTimezoneTests(unittest.TestCase):
    def test_options_keep_old_names(self):
        for name in ("本地时区", "UTC", "UTC+8", "自定义 UTC 偏移"):
            self.assertIn(name, TIMEZONE_OPTIONS)

    def test_us_pacific_summer_and_winter(self):
        self.assertEqual("2026-07-01 19:00", utc_of("2026-07-01 12:00:00", "美西（洛杉矶）"))
        self.assertEqual("2026-01-15 20:00", utc_of("2026-01-15 12:00:00", "美西（洛杉矶）"))
        self.assertEqual("2026-07-01 16:00", utc_of("2026-07-01 12:00:00", "美东（纽约）"))

    def test_us_switch_days_2026(self):
        # 2026 年美国 3 月 8 日开始、11 月 1 日结束夏令时
        self.assertEqual("2026-03-08 09:59", utc_of("2026-03-08 01:59:00", "美西（洛杉矶）"))
        self.assertEqual("2026-03-08 10:00", utc_of("2026-03-08 03:00:00", "美西（洛杉矶）"))
        self.assertEqual("2026-11-01 07:30", utc_of("2026-11-01 00:30:00", "美西（洛杉矶）"))  # 回拨前，夏令时 UTC-7
        self.assertEqual("2026-11-01 10:30", utc_of("2026-11-01 02:30:00", "美西（洛杉矶）"))  # 回拨后，标准时间
        # 01:00～02:00 会出现两次，这里约定按标准时间计算
        self.assertEqual("2026-11-01 09:30", utc_of("2026-11-01 01:30:00", "美西（洛杉矶）"))

    def test_europe_and_sydney(self):
        self.assertEqual("2026-07-01 11:00", utc_of("2026-07-01 12:00:00", "英国（伦敦）"))
        self.assertEqual("2026-01-01 12:00", utc_of("2026-01-01 12:00:00", "英国（伦敦）"))
        self.assertEqual("2026-07-01 10:00", utc_of("2026-07-01 12:00:00", "中欧（巴黎 / 柏林）"))
        self.assertEqual("2026-01-15 01:00", utc_of("2026-01-15 12:00:00", "澳洲东部（悉尼）"))
        self.assertEqual("2026-07-15 02:00", utc_of("2026-07-15 12:00:00", "澳洲东部（悉尼）"))
        self.assertEqual("2026-07-15 03:00", utc_of("2026-07-15 12:00:00", "日本 / 韩国（UTC+9）"))

    def test_timestamp_to_named_zone(self):
        ts = int(datetime(2026, 7, 1, 19, 0, tzinfo=timezone.utc).timestamp())
        data = timestamp_to_datetime(str(ts), "秒", "美西（洛杉矶）")
        self.assertEqual("2026-07-01 12:00:00", data["zone"])
        self.assertIn("UTC-07:00，夏令时", data["zone_label"])
        ts_winter = int(datetime(2026, 1, 15, 20, 0, tzinfo=timezone.utc).timestamp())
        self.assertEqual("2026-01-15 12:00:00", timestamp_to_datetime(str(ts_winter), "秒", "美西（洛杉矶）")["zone"])

    def test_round_trip_and_label(self):
        data = datetime_to_timestamps("2026-07-01 12:00:00", "美东（纽约）")
        self.assertIn("UTC-04:00，夏令时", data["timezone"])
        back = timestamp_to_datetime(data["seconds"], "秒", "美东（纽约）")
        self.assertEqual("2026-07-01 12:00:00", back["zone"])

    def test_difference_across_dst_switch_uses_real_elapsed_time(self):
        # 美西 3 月 8 日 00:00 到 04:00，墙上时间差 4 小时，实际只过了 3 小时
        data = time_difference("2026-03-08 00:00:00", "2026-03-08 04:00:00", "美西（洛杉矶）")
        self.assertEqual(3 * 3600, data["seconds"])
        data = countdown("2026-03-08 00:00:00", "2026-03-08 04:00:00", "美西（洛杉矶）")
        self.assertEqual(3 * 3600, data["seconds"])

    def test_boundary_with_empty_base_uses_selected_zone(self):
        from app.time_tools import boundary_times
        data = boundary_times("", "UTC")
        day_start = data["当天 00:00:00"]
        self.assertEqual(timedelta(0), day_start.utcoffset())
        self.assertEqual((0, 0, 0), (day_start.hour, day_start.minute, day_start.second))

    def test_fixed_offsets_unchanged(self):
        self.assertEqual("2026-07-01 04:00", utc_of("2026-07-01 12:00:00", "UTC+8"))
        self.assertEqual(4 * 3600, time_difference("2026-03-08 00:00:00", "2026-03-08 04:00:00", "UTC+8")["seconds"])


if __name__ == "__main__":
    unittest.main()
