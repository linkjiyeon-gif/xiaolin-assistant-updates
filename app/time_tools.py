from __future__ import annotations

import math
import re
from datetime import datetime, timedelta, timezone, tzinfo
from typing import Callable, Dict, Optional, Tuple

DATETIME_FORMAT = "%Y-%m-%d %H:%M:%S"

_TIME_UNITS_IN_SECONDS = {
    "秒": 1,
    "分钟": 60,
    "小时": 3600,
    "天": 86400,
    "周": 604800,
    "月": 2592000,  # 30 days, estimate
    "年": 31536000,  # 365 days, estimate
}

_SAFE_FORMATS = [
    "%Y-%m-%d %H:%M:%S",
    "%Y/%m/%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%Y/%m/%d %H:%M",
    "%Y-%m-%d",
    "%Y/%m/%d",
]


def local_timezone():
    return datetime.now().astimezone().tzinfo or timezone.utc


def format_datetime(dt: datetime) -> str:
    return dt.strftime(DATETIME_FORMAT)


def format_timestamp(ms_or_seconds: float) -> str:
    if float(ms_or_seconds).is_integer():
        return str(int(ms_or_seconds))
    return f"{ms_or_seconds:.3f}".rstrip("0").rstrip(".")


def detect_timestamp_unit(value: str, manual_unit: str = "自动识别") -> str:
    text = str(value).strip()
    if not text:
        raise ValueError("请输入时间戳")
    if manual_unit in {"秒", "毫秒"}:
        return manual_unit
    normalized = re.sub(r"[^0-9]", "", text)
    if len(normalized) >= 13:
        return "毫秒"
    return "秒"


def timestamp_to_datetime(value: str, unit: str = "自动识别", tz_option: str = "", custom_offset: str = "+8") -> Dict[str, str]:
    text = str(value).strip()
    if not text:
        raise ValueError("请输入时间戳")
    try:
        raw = float(text)
    except ValueError as exc:
        raise ValueError("时间戳只能输入数字") from exc
    actual_unit = detect_timestamp_unit(text, unit)
    seconds = raw / 1000 if actual_unit == "毫秒" else raw
    try:
        utc_dt = datetime.fromtimestamp(seconds, tz=timezone.utc)
        local_dt = datetime.fromtimestamp(seconds, tz=local_timezone())
    except Exception as exc:
        raise ValueError("时间戳超出可转换范围") from exc
    result = {
        "unit": actual_unit,
        "local": format_datetime(local_dt),
        "utc": format_datetime(utc_dt),
        "seconds": str(int(seconds)),
        "milliseconds": str(int(seconds * 1000)),
    }
    # 额外给出所选时区的时间（本地 / UTC 已经在上面）。
    if tz_option and tz_option not in {"本地时区", "UTC"}:
        zone_dt = datetime.fromtimestamp(seconds, tz=timezone_from_option(tz_option, custom_offset))
        label = tz_option if tz_option != "自定义 UTC 偏移" else f"UTC{custom_offset}"
        result["zone_label"] = f"{label}（{describe_offset(zone_dt)}）"
        result["zone"] = format_datetime(zone_dt)
    return result


def parse_utc_offset(offset_text: str) -> timezone:
    text = str(offset_text).strip().upper().replace("UTC", "")
    if not text:
        raise ValueError("请输入自定义 UTC 偏移，例如 +8 或 +08:00")
    text = text.replace("：", ":")
    match = re.fullmatch(r"([+-]?)(\d{1,2})(?::?(\d{1,2}))?", text)
    if not match:
        raise ValueError("UTC 偏移格式错误，可输入 +8、+08:00、-05:30")
    sign_text, hour_text, minute_text = match.groups()
    sign = -1 if sign_text == "-" else 1
    hours = int(hour_text)
    minutes = int(minute_text or "0")
    if hours > 14 or minutes >= 60:
        raise ValueError("UTC 偏移超出常用范围")
    return timezone(sign * timedelta(hours=hours, minutes=minutes))


# ---------------- 常用时区（含夏令时） ----------------
# Windows 版 Python 没有自带时区数据库（zoneinfo 需要额外的 tzdata 包），
# 这里直接实现几个常用时区现行的夏令时规则：
#   美国：3 月第 2 个周日 02:00 至 11 月第 1 个周日 02:00（当地时间，2007 年起）
#   欧盟 / 英国：3 月最后一个周日 01:00 UTC 至 10 月最后一个周日 01:00 UTC
#   澳洲东部：10 月第 1 个周日 02:00 至次年 4 月第 1 个周日 03:00（当地时间）


def _nth_weekday(year: int, month: int, weekday: int, n: int) -> datetime:
    """n-th weekday (0=Mon) of a month; n=-1 means the last one. Returns a naive date at 00:00."""
    if n > 0:
        first = datetime(year, month, 1)
        return first + timedelta(days=(weekday - first.weekday()) % 7 + 7 * (n - 1))
    next_month = datetime(year + (month == 12), month % 12 + 1, 1)
    last = next_month - timedelta(days=1)
    return last - timedelta(days=(last.weekday() - weekday) % 7)


def _us_dst_utc(year: int, std: timedelta) -> Tuple[datetime, datetime]:
    start_local = _nth_weekday(year, 3, 6, 2) + timedelta(hours=2)  # 标准时间 02:00
    end_local = _nth_weekday(year, 11, 6, 1) + timedelta(hours=2)  # 夏令时 02:00
    return start_local - std, end_local - (std + timedelta(hours=1))


def _eu_dst_utc(year: int, std: timedelta) -> Tuple[datetime, datetime]:
    return _nth_weekday(year, 3, 6, -1) + timedelta(hours=1), _nth_weekday(year, 10, 6, -1) + timedelta(hours=1)


def _au_dst_utc(year: int, std: timedelta) -> Tuple[datetime, datetime]:
    # 南半球：夏令时跨年。返回本年 10 月开始、本年 4 月结束，判断时用“>= 开始 或 < 结束”。
    start_local = _nth_weekday(year, 10, 6, 1) + timedelta(hours=2)
    end_local = _nth_weekday(year, 4, 6, 1) + timedelta(hours=3)
    return start_local - std, end_local - (std + timedelta(hours=1))


class RuleTimezone(tzinfo):
    """Fixed standard offset plus a yearly DST rule, enough for common game-server regions."""

    def __init__(self, label: str, std_hours: float, rule: Optional[Callable] = None, southern: bool = False):
        self.label = label
        self.std = timedelta(hours=std_hours)
        self.rule = rule
        self.southern = southern

    def _is_dst_utc(self, utc_naive: datetime) -> bool:
        if self.rule is None:
            return False
        start, end = self.rule(utc_naive.year, self.std)
        if self.southern:
            return utc_naive >= start or utc_naive < end
        return start <= utc_naive < end

    def utcoffset(self, dt):
        return self.std + self.dst(dt)

    def dst(self, dt):
        if dt is None or self.rule is None:
            return timedelta(0)
        # 按当地墙上时间判断：先当作标准时间换算到 UTC。重复的那一小时按标准时间处理。
        return timedelta(hours=1) if self._is_dst_utc(dt.replace(tzinfo=None) - self.std) else timedelta(0)

    def fromutc(self, dt):
        utc_naive = dt.replace(tzinfo=None)
        offset = self.std + (timedelta(hours=1) if self._is_dst_utc(utc_naive) else timedelta(0))
        return (utc_naive + offset).replace(tzinfo=self)

    def tzname(self, dt):
        return self.label

    def __repr__(self):
        return f"RuleTimezone({self.label!r})"


# 下拉框选项（保持旧选项名不变，旧配置和代码仍然可用）。
NAMED_TIMEZONES: Dict[str, tzinfo] = {
    "美西（洛杉矶）": RuleTimezone("美西（洛杉矶）", -8, _us_dst_utc),
    "美东（纽约）": RuleTimezone("美东（纽约）", -5, _us_dst_utc),
    "英国（伦敦）": RuleTimezone("英国（伦敦）", 0, _eu_dst_utc),
    "中欧（巴黎 / 柏林）": RuleTimezone("中欧（巴黎 / 柏林）", 1, _eu_dst_utc),
    "日本 / 韩国（UTC+9）": RuleTimezone("日本 / 韩国（UTC+9）", 9),
    "澳洲东部（悉尼）": RuleTimezone("澳洲东部（悉尼）", 10, _au_dst_utc, southern=True),
}
TIMEZONE_OPTIONS = ["本地时区", "UTC", "UTC+8"] + list(NAMED_TIMEZONES) + ["自定义 UTC 偏移"]


def describe_offset(dt: datetime) -> str:
    """e.g. 'UTC-07:00，夏令时' for an aware datetime."""
    offset = dt.utcoffset() or timedelta(0)
    total = int(offset.total_seconds())
    sign = "-" if total < 0 else "+"
    hours, minutes = divmod(abs(total) // 60, 60)
    text = f"UTC{sign}{hours:02d}:{minutes:02d}"
    if isinstance(dt.tzinfo, RuleTimezone) and dt.tzinfo.rule is not None:
        text += "，夏令时" if dt.dst() else "，标准时间"
    return text


def timezone_from_option(option: str, custom_offset: str = "+8") -> tzinfo:
    option = (option or "本地时区").strip()
    if option == "UTC":
        return timezone.utc
    if option == "UTC+8":
        return timezone(timedelta(hours=8))
    if option == "自定义 UTC 偏移":
        return parse_utc_offset(custom_offset)
    if option in NAMED_TIMEZONES:
        return NAMED_TIMEZONES[option]
    return local_timezone()


def _seconds_between(start: datetime, end: datetime) -> int:
    # 不用 end - start：同一个 tzinfo 时 Python 只按墙上时间相减，跨夏令时切换会差 1 小时。
    return int(round(end.timestamp() - start.timestamp()))


def parse_datetime_text(text: str, tz_option: str = "本地时区", custom_offset: str = "+8") -> datetime:
    raw = str(text).strip()
    if not raw:
        raise ValueError("请输入日期时间")
    normalized = raw.replace("T", " ").strip()
    # Try ISO first for inputs with timezone such as 2026-06-09T15:30:00+08:00.
    try:
        iso_candidate = raw.replace("/", "-")
        dt = datetime.fromisoformat(iso_candidate)
    except Exception:
        dt = None
    if dt is None:
        for fmt in _SAFE_FORMATS:
            try:
                dt = datetime.strptime(normalized, fmt)
                break
            except ValueError:
                continue
    if dt is None:
        raise ValueError("日期时间格式错误，示例：2026-06-09 15:30:00")
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone_from_option(tz_option, custom_offset))
    return dt


def datetime_to_timestamps(text: str, tz_option: str = "本地时区", custom_offset: str = "+8") -> Dict[str, str]:
    dt = parse_datetime_text(text, tz_option, custom_offset)
    seconds = int(dt.timestamp())
    label = tz_option if tz_option != "自定义 UTC 偏移" else f"UTC{custom_offset}"
    return {
        "datetime": format_datetime(dt),
        "timezone": f"{label}（{describe_offset(dt)}）",
        "seconds": str(seconds),
        "milliseconds": str(seconds * 1000),
        "utc": format_datetime(dt.astimezone(timezone.utc)),
        "local": format_datetime(dt.astimezone(local_timezone())),
    }


def convert_time_units(value: str, source_unit: str = "秒") -> Dict[str, float]:
    try:
        number = float(str(value).strip())
    except ValueError as exc:
        raise ValueError("时间单位换算只能输入数字") from exc
    if not math.isfinite(number):
        raise ValueError("请输入有效数字")
    if source_unit not in _TIME_UNITS_IN_SECONDS:
        raise ValueError("请选择有效单位")
    total_seconds = number * _TIME_UNITS_IN_SECONDS[source_unit]
    return {unit: total_seconds / seconds for unit, seconds in _TIME_UNITS_IN_SECONDS.items()}


def split_seconds(total_seconds: int) -> Tuple[int, int, int, int]:
    total_seconds = abs(int(total_seconds))
    days, remain = divmod(total_seconds, 86400)
    hours, remain = divmod(remain, 3600)
    minutes, seconds = divmod(remain, 60)
    return days, hours, minutes, seconds


def readable_duration(total_seconds: int) -> str:
    days, hours, minutes, seconds = split_seconds(total_seconds)
    parts = []
    if days:
        parts.append(f"{days} 天")
    if hours or parts:
        parts.append(f"{hours} 小时")
    if minutes or parts:
        parts.append(f"{minutes} 分")
    parts.append(f"{seconds} 秒")
    return " ".join(parts)


def time_difference(start_text: str, end_text: str, tz_option: str = "本地时区", custom_offset: str = "+8") -> Dict[str, object]:
    start = parse_datetime_text(start_text, tz_option, custom_offset)
    end = parse_datetime_text(end_text, tz_option, custom_offset)
    seconds = _seconds_between(start, end)
    abs_seconds = abs(seconds)
    return {
        "is_negative": seconds < 0,
        "seconds": seconds,
        "abs_seconds": abs_seconds,
        "minutes": abs_seconds / 60,
        "hours": abs_seconds / 3600,
        "days": abs_seconds / 86400,
        "readable": readable_duration(abs_seconds),
    }


_CUSTOM_DURATION_UNITS = {
    "周": 604800,
    "天": 86400,
    "小时": 3600,
    "时": 3600,
    "分钟": 60,
    "分": 60,
    "秒": 1,
}

_SUPPORTED_CUSTOM_DURATION_UNITS_TEXT = "周、天、小时、分钟、秒"


def _format_duration_number(value: float) -> str:
    if float(value).is_integer():
        return str(int(value))
    return f"{value:.6f}".rstrip("0").rstrip(".")


def parse_custom_duration_text(text: str) -> Dict[str, object]:
    raw = str(text).strip()
    if not raw:
        raise ValueError("请输入自定义时长")
    normalized = re.sub(r"\s+", "", raw).replace("，", "").replace(",", "")
    if "-" in normalized or "－" in normalized or "负" in normalized:
        raise ValueError("自定义时长不支持负数，请通过增加 / 减少方向选择计算方式")

    token_pattern = re.compile(r"(\d+(?:\.\d+)?)(周|天|小时|时|分钟|分|秒)")
    pos = 0
    total_seconds = 0.0
    parsed_parts = []
    matched = False

    for match in token_pattern.finditer(normalized):
        gap = normalized[pos:match.start()]
        if gap:
            if re.search(r"\d", gap):
                raise ValueError(f"仅支持{_SUPPORTED_CUSTOM_DURATION_UNITS_TEXT}")
            raise ValueError("自定义时长格式错误，请输入如：1天2小时30分5秒")
        number_text, unit_text = match.groups()
        amount = float(number_text)
        if not math.isfinite(amount):
            raise ValueError("请输入有效数字")
        total_seconds += amount * _CUSTOM_DURATION_UNITS[unit_text]
        parsed_parts.append(f"{_format_duration_number(amount)}{unit_text}")
        matched = True
        pos = match.end()

    tail = normalized[pos:]
    if tail:
        if re.search(r"\d", tail) or (matched and re.search(r"[年月日毫厘a-zA-Z]", tail)):
            raise ValueError(f"仅支持{_SUPPORTED_CUSTOM_DURATION_UNITS_TEXT}")
        raise ValueError("自定义时长格式错误，请输入如：1天2小时30分5秒")
    if not matched:
        if re.search(r"\d", normalized) and re.search(r"[年月日毫厘a-zA-Z]", normalized):
            raise ValueError(f"仅支持{_SUPPORTED_CUSTOM_DURATION_UNITS_TEXT}")
        raise ValueError("自定义时长格式错误，请输入如：1天2小时30分5秒")
    if total_seconds <= 0:
        raise ValueError("自定义时长必须大于 0")

    return {
        "display": "".join(parsed_parts),
        "seconds": total_seconds,
    }


def add_or_subtract_time(base_text: str, direction: str, value: str, unit: str, tz_option: str = "本地时区", custom_offset: str = "+8") -> Dict[str, str]:
    base = parse_datetime_text(base_text, tz_option, custom_offset)
    result_data: Dict[str, str] = {}
    if unit == "自定义":
        duration_info = parse_custom_duration_text(value)
        total_seconds = float(duration_info["seconds"])
        result_data["custom_duration"] = str(duration_info["display"])
        result_data["duration_seconds"] = format_timestamp(total_seconds)
    else:
        try:
            amount = float(str(value).strip())
        except ValueError as exc:
            raise ValueError("加减数值只能输入数字") from exc
        if not math.isfinite(amount):
            raise ValueError("请输入有效数字")
        if unit not in {"秒", "分钟", "小时", "天", "周"}:
            raise ValueError("请选择有效单位")
        total_seconds = amount * _TIME_UNITS_IN_SECONDS[unit]

    delta = timedelta(seconds=total_seconds)
    result = base - delta if direction == "减少" else base + delta
    seconds = int(result.timestamp())
    result_data.update({
        "datetime": format_datetime(result),
        "seconds": str(seconds),
        "milliseconds": str(seconds * 1000),
    })
    return result_data



def _parse_duration_part(value: str, field_name: str) -> int:
    raw = str(value).strip()
    if raw == "":
        return 0
    if "-" in raw or "－" in raw or raw.startswith("负"):
        raise ValueError("时长输入不支持负数")
    if not re.fullmatch(r"\d+", raw):
        raise ValueError("时长输入仅支持数字")
    return int(raw)


def parse_duration_parts(days: str, hours: str, minutes: str, seconds: str) -> Dict[str, object]:
    day_value = _parse_duration_part(days, "天")
    hour_value = _parse_duration_part(hours, "时")
    minute_value = _parse_duration_part(minutes, "分")
    second_value = _parse_duration_part(seconds, "秒")
    total_seconds = day_value * 86400 + hour_value * 3600 + minute_value * 60 + second_value
    if total_seconds <= 0:
        raise ValueError("请输入有效时长")
    return {
        "days": day_value,
        "hours": hour_value,
        "minutes": minute_value,
        "seconds": second_value,
        "total_seconds": total_seconds,
        "display": f"{day_value}天{hour_value}小时{minute_value}分{second_value}秒",
    }


def add_or_subtract_time_parts(base_text: str, direction: str, days: str, hours: str, minutes: str, seconds_text: str, tz_option: str = "本地时区", custom_offset: str = "+8") -> Dict[str, str]:
    base = parse_datetime_text(base_text, tz_option, custom_offset)
    duration_info = parse_duration_parts(days, hours, minutes, seconds_text)
    total_seconds = int(duration_info["total_seconds"])
    delta = timedelta(seconds=total_seconds)
    result = base - delta if direction == "减少" else base + delta
    timestamp_seconds = int(result.timestamp())
    return {
        "datetime": format_datetime(result),
        "seconds": str(timestamp_seconds),
        "milliseconds": str(timestamp_seconds * 1000),
        "duration_display": str(duration_info["display"]),
        "duration_seconds": str(total_seconds),
    }

def countdown(current_text: str, end_text: str, tz_option: str = "本地时区", custom_offset: str = "+8") -> Dict[str, object]:
    current = datetime.now(local_timezone()) if not str(current_text).strip() else parse_datetime_text(current_text, tz_option, custom_offset)
    end = parse_datetime_text(end_text, tz_option, custom_offset)
    remaining = _seconds_between(current, end)
    ended = remaining <= 0
    display_seconds = max(0, remaining)
    days, hours, minutes, seconds = split_seconds(display_seconds)
    return {
        "ended": ended,
        "seconds": display_seconds,
        "minutes": display_seconds / 60,
        "hours": display_seconds / 3600,
        "days": display_seconds / 86400,
        "dd_hhmmss": f"{days:02d} 天 {hours:02d}:{minutes:02d}:{seconds:02d}",
        "hhmmss": f"{days * 24 + hours:02d}:{minutes:02d}:{seconds:02d}",
    }


def boundary_times(base_text: str, tz_option: str = "本地时区", custom_offset: str = "+8") -> Dict[str, datetime]:
    # 留空时取“此刻”，但要换算到所选时区，否则跨天边界会按本机时区的 0 点算。
    base = datetime.now(timezone_from_option(tz_option, custom_offset)) if not str(base_text).strip() else parse_datetime_text(base_text, tz_option, custom_offset)
    day_start = base.replace(hour=0, minute=0, second=0, microsecond=0)
    day_end = base.replace(hour=23, minute=59, second=59, microsecond=0)
    next_day_start = day_start + timedelta(days=1)
    week_start = day_start - timedelta(days=day_start.weekday())
    next_week_start = week_start + timedelta(days=7)
    month_start = day_start.replace(day=1)
    if month_start.month == 12:
        next_month_start = month_start.replace(year=month_start.year + 1, month=1)
    else:
        next_month_start = month_start.replace(month=month_start.month + 1)
    return {
        "当天 00:00:00": day_start,
        "当天 23:59:59": day_end,
        "次日 00:00:00": next_day_start,
        "本周一 00:00:00": week_start,
        "下周一 00:00:00": next_week_start,
        "本月 1 日 00:00:00": month_start,
        "下月 1 日 00:00:00": next_month_start,
    }


def timestamp_pair(dt: datetime) -> Tuple[int, int]:
    seconds = int(dt.timestamp())
    return seconds, seconds * 1000
