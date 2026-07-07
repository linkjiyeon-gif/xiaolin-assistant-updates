from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Dict, Iterable, List, Tuple

from .file_readers import ParsedContent, clean_key, split_lines
from .normalizer import NormalizeOptions, normalize_key, normalize_text, should_ignore_key


@dataclass(frozen=True)
class PlanningItem:
    category: str
    key: str
    value: str
    location: str
    raw: str = ""


REWARD_KEYWORDS = (
    "奖励", "reward", "rewards", "item", "items", "drop", "drops", "道具", "礼包", "资源",
    "金币", "金幣", "钻石", "鑽石", "宝石", "粮食", "木材", "石头", "铁矿", "加速", "碎片",
)
TIER_KEYWORDS = ("档位", "檔位", "等级", "等級", "level", "lv", "vip", "tier", "stage", "rank", "段位")
SWITCH_KEYWORDS = (
    "开关", "開關", "开启", "開啟", "关闭", "關閉", "启用", "啟用", "禁用",
    "enable", "enabled", "disable", "disabled", "open", "opened", "close", "closed", "switch", "isopen", "status",
)
CONFIG_KEYWORDS = ("配置", "字段", "field", "config", "参数", "參數", "value", "数值", "數值")
STRUCTURED_EXTENSIONS = {".json", ".yaml", ".yml", ".xml", ".csv", ".xlsx"}

BOOL_TRUE = {"1", "true", "yes", "y", "on", "open", "opened", "enable", "enabled", "开启", "開啟", "启用", "啟用", "打开", "開放"}
BOOL_FALSE = {"0", "false", "no", "n", "off", "close", "closed", "disable", "disabled", "关闭", "關閉", "禁用", "未开启", "未開啟"}

KEY_VALUE_PATTERNS = [
    re.compile(r"^\s*([\w\-.\[\]/\u4e00-\u9fa5（）()]{1,80})\s*[:=：]\s*(.+?)\s*[,;；，]?\s*$"),
    re.compile(r"(?:字段|配置|参数|參數|field|config)\s*[:：=]?\s*([\w\-.\[\]/\u4e00-\u9fa5（）()]{1,80})\s*(?:为|爲|=|:|：)\s*(.+?)\s*$", re.I),
]

# 奖励文本常见写法：金币 x100、金币*100、金币：100、100金币、item_id=1001,count=10
REWARD_PAIR_PATTERNS = [
    re.compile(r"([\u4e00-\u9fa5A-Za-z_][\u4e00-\u9fa5A-Za-z0-9_\-]{0,30})\s*(?:x|X|×|\*|:|：|=)\s*([\d.]+)"),
    re.compile(r"([\d.]+)\s*(?:个|個|份|枚|点|點|次)?\s*([\u4e00-\u9fa5A-Za-z_][\u4e00-\u9fa5A-Za-z0-9_\-]{0,30})"),
    re.compile(r"(?:item[_-]?id|道具id|奖励id|reward[_-]?id)\s*[:=：]\s*([\w\-]+).*?(?:count|num|数量|數量)\s*[:=：]\s*([\d.]+)", re.I),
]

TIER_PATTERNS = [
    re.compile(r"(VIP\s*\d+)", re.I),
    re.compile(r"(?:第\s*)?(\d+)\s*(?:档|檔|档位|檔位|级|級|等级|等級)", re.I),
    re.compile(r"(?:level|lv|tier|stage)\s*[_:\-= ]?\s*(\d+)", re.I),
]


def extract_planning_items(content: ParsedContent, options_normalize: NormalizeOptions, ignore_fields: Iterable[str]) -> List[PlanningItem]:
    """把策划文档或实际配置抽取成可比对的业务项。

    这是确定性启发式解析，不执行源文件代码，也不修改原文件。
    """
    ignore_list = list(ignore_fields)
    items: List[PlanningItem] = []

    for key, value in content.fields.items():
        if _skip_low_value_field_key(key):
            continue
        if _should_ignore_item(key, str(value), ignore_list, options_normalize):
            continue
        items.extend(_items_from_field(key, str(value), f"字段：{key}"))

    if content.ext not in STRUCTURED_EXTENSIONS:
        for line_num, line in enumerate(content.lines, start=1):
            clean_line = line.strip()
            if not clean_line or len(clean_line) > 1000:
                continue
            if _should_ignore_item(clean_line, clean_line, ignore_list, options_normalize):
                continue
            items.extend(_items_from_line(clean_line, f"第 {line_num} 行"))

    return _deduplicate_items(items, options_normalize)


def _items_from_field(key: str, value: str, location: str) -> List[PlanningItem]:
    clean_field_key = _clean_field_key(key)
    clean_value = _clean_value(value)
    lower_key = clean_field_key.lower()
    lower_value = clean_value.lower()
    merged = f"{clean_field_key} {clean_value}"
    merged_lower = merged.lower()

    items: List[PlanningItem] = []
    is_reward = _contains_any(merged_lower, REWARD_KEYWORDS)
    is_tier = _contains_any(lower_key, TIER_KEYWORDS) or _contains_any(lower_value, TIER_KEYWORDS)
    is_switch = _looks_like_switch(clean_field_key, clean_value)

    if is_reward:
        reward_items = _extract_rewards(merged, location)
        if reward_items:
            items.extend(reward_items)
        else:
            items.append(PlanningItem("奖励内容", clean_field_key, normalize_business_value(clean_value), location, f"{key}: {value}"))

    if is_tier:
        if _is_generic_tier_field(clean_field_key):
            tier_key = _canonical_tier_field_key(clean_field_key)
            tier_value = clean_value
        else:
            tier_key = _extract_tier_key(merged) or clean_field_key
            tier_value = clean_value
        items.append(PlanningItem("档位", tier_key, normalize_business_value(tier_value), location, f"{key}: {value}"))

    if is_switch:
        items.append(PlanningItem("开关状态", clean_field_key, normalize_switch_value(clean_value), location, f"{key}: {value}"))

    if not is_reward and not is_tier and not is_switch:
        # 通用字段/配置值：用于检查普通字段和配置值是否一致。
        items.append(PlanningItem("配置值", clean_field_key, normalize_business_value(clean_value), location, f"{key}: {value}"))

    return items

def _items_from_line(line: str, location: str) -> List[PlanningItem]:
    items: List[PlanningItem] = []
    lower_line = line.lower()
    line_is_reward = _contains_any(lower_line, REWARD_KEYWORDS)
    line_is_tier = _contains_any(lower_line, TIER_KEYWORDS)
    line_is_switch = _contains_any(lower_line, SWITCH_KEYWORDS)

    kv = _extract_key_value_from_line(line)
    if kv:
        key, value = kv
        if _looks_like_switch(key, value):
            items.append(PlanningItem("开关状态", key, normalize_switch_value(value), location, line))
        elif not line_is_reward and not line_is_tier:
            items.append(PlanningItem("配置值", key, normalize_business_value(value), location, line))

    if line_is_reward:
        reward_items = _extract_rewards(line, location)
        if reward_items:
            items.extend(reward_items)
        else:
            key = _extract_key_before_colon(line) or _short_key_from_line(line)
            items.append(PlanningItem("奖励内容", key, normalize_business_value(line), location, line))

    if line_is_tier:
        tier_key = _extract_tier_key(line)
        if tier_key:
            items.append(PlanningItem("档位", tier_key, normalize_business_value(line), location, line))

    if line_is_switch and not kv:
        key = _short_key_from_line(line)
        items.append(PlanningItem("开关状态", key, normalize_switch_value(line), location, line))

    return items

def _extract_rewards(text: str, location: str) -> List[PlanningItem]:
    items: List[PlanningItem] = []
    seen = set()
    for pattern in REWARD_PAIR_PATTERNS:
        for match in pattern.finditer(text):
            if pattern.pattern.startswith("(?:item"):
                reward_key = f"item_{match.group(1)}"
                amount = match.group(2)
            else:
                first, second = match.group(1), match.group(2)
                if _is_number(first):
                    reward_key, amount = second, first
                else:
                    reward_key, amount = first, second
            reward_key = _clean_field_key(reward_key)
            amount = _clean_value(amount)
            if not reward_key or reward_key in {"奖励", "reward", "count", "num", "数量", "數量", "档", "檔", "级", "級", "等级", "等級", "档位", "檔位"}:
                continue
            normalized = (reward_key.lower(), amount)
            if normalized in seen:
                continue
            seen.add(normalized)
            items.append(PlanningItem("奖励内容", reward_key, _normalize_amount(amount), location, text))
    return items


def _extract_key_value_from_line(line: str) -> Tuple[str, str] | None:
    for pattern in KEY_VALUE_PATTERNS:
        match = pattern.search(line)
        if match:
            return _clean_field_key(match.group(1)), _clean_value(match.group(2))
    return None


def _extract_key_before_colon(line: str) -> str:
    match = re.match(r"^\s*([^:：=]{1,60})\s*[:：=]", line)
    return _clean_field_key(match.group(1)) if match else ""


def _extract_tier_key(text: str) -> str:
    for pattern in TIER_PATTERNS:
        match = pattern.search(text)
        if match:
            value = re.sub(r"\s+", "", match.group(1))
            if value.isdigit():
                return f"档位{value}"
            return value.upper()
    return ""


def _looks_like_switch(key: str, value: str) -> bool:
    text = f"{key} {value}".lower()
    if _contains_any(text, SWITCH_KEYWORDS):
        return True
    normalized = _clean_value(value).lower()
    return normalized in BOOL_TRUE or normalized in BOOL_FALSE



def _is_generic_tier_field(key: str) -> bool:
    normalized = key.strip().lower()
    return normalized in {"等级", "等級", "level", "lv", "tier", "stage", "rank", "档位", "檔位"}


def _canonical_tier_field_key(key: str) -> str:
    normalized = key.strip().lower()
    if normalized in {"等级", "等級", "level", "lv"}:
        return "等级"
    if normalized in {"tier", "stage", "rank", "档位", "檔位"}:
        return "档位"
    return key


def _normalize_amount(value: str) -> str:
    value = _clean_value(value)
    if re.fullmatch(r"-?\d+\.0+", value):
        return value.split(".", 1)[0]
    return value

def normalize_switch_value(value: str) -> str:
    raw = _clean_value(value)
    low = raw.lower()
    if low in BOOL_TRUE or raw in BOOL_TRUE:
        return "开启"
    if low in BOOL_FALSE or raw in BOOL_FALSE:
        return "关闭"
    if any(word in raw for word in ("未开启", "未開啟", "关闭", "關閉", "禁用")):
        return "关闭"
    if any(word in raw for word in ("开启", "開啟", "启用", "啟用", "打开")):
        return "开启"
    return normalize_business_value(raw)


def normalize_business_value(value: str) -> str:
    value = _clean_value(value)
    low = value.lower()
    # 普通配置值里 1/0 可能是数值，不直接当成开关；开关统一由 normalize_switch_value 处理。
    if low in {"true", "yes", "y", "on", "open", "opened", "enable", "enabled"} or value in {"开启", "開啟", "启用", "啟用", "打开", "開放"}:
        return "开启"
    if low in {"false", "no", "n", "off", "close", "closed", "disable", "disabled"} or value in {"关闭", "關閉", "禁用", "未开启", "未開啟"}:
        return "关闭"
    # 统一纯数字表现：100.0 -> 100
    if re.fullmatch(r"-?\d+\.0+", value):
        return value.split(".", 1)[0]
    return value


def _clean_field_key(key: str) -> str:
    key = clean_key(key)
    key = key.strip().strip('"\'`')
    key = re.sub(r"\s+", "", key)
    key = key.replace("（", "(").replace("）", ")")
    return key


def _clean_value(value: str) -> str:
    value = "" if value is None else str(value)
    value = value.strip().strip(",，;；")
    value = value.strip().strip('"\'`')
    value = re.sub(r"\s+", " ", value)
    return value


def _short_key_from_line(line: str) -> str:
    line = re.sub(r"\s+", "", line.strip())
    return line[:40]


def _contains_any(text: str, keywords: Iterable[str]) -> bool:
    low = text.lower()
    return any(keyword.lower() in low for keyword in keywords)


def _is_number(value: str) -> bool:
    return bool(re.fullmatch(r"\d+(?:\.\d+)?", value.strip()))


def _should_ignore_item(key: str, value: str, ignore_fields: Iterable[str], normalize_options: NormalizeOptions) -> bool:
    if should_ignore_key(key, ignore_fields, normalize_options):
        return True
    normalized_text = normalize_text(f"{key} {value}", normalize_options)
    for ignore in ignore_fields:
        if ignore and ignore in normalized_text:
            return True
    return False


def _skip_low_value_field_key(key: str) -> bool:
    # 表格坐标字段通常只说明单元格位置，不适合作为业务项参与比对。
    if re.fullmatch(r".+!R\d+C\d+", key):
        return True
    if re.fullmatch(r"table_\d+\.R\d+C\d+", key):
        return True
    return False


def _deduplicate_items(items: List[PlanningItem], normalize_options: NormalizeOptions) -> List[PlanningItem]:
    result: Dict[Tuple[str, str, str], PlanningItem] = {}
    for item in items:
        key = normalize_key(item.key, normalize_options)
        value = normalize_text(item.value, normalize_options)
        dedup_key = (item.category, key, value)
        if dedup_key not in result:
            result[dedup_key] = item
    return list(result.values())
