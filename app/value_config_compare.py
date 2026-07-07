from __future__ import annotations

import csv
import json
import re
from difflib import SequenceMatcher
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

from .comparer import DiffItem
from .file_readers import read_text_with_fallback


@dataclass
class ValueCompareOptions:
    # 旧专项选项保留，避免 UI / 旧调用回退。
    manual_ids: str = ""
    keywords: str = ""
    compare_rewards: bool = True
    compare_price: bool = True
    compare_pcid: bool = True
    compare_duration: bool = True
    compare_score: bool = False
    compare_limit_time: bool = True

    # 通用结构化比对选项。左侧为游戏配置表，右侧为数值参考表。
    reference_sheet: str = ""
    config_sheet: str = ""
    reference_header_row: str = "auto"
    reference_data_start_row: str = "auto"
    config_header_row: str = "auto"
    config_data_start_row: str = "auto"
    key_fields: str = "auto"
    reference_filter: str = ""
    config_filter: str = ""
    template_name: str = ""
    field_mappings: str = ""
    ignore_fields: str = "备注,说明,comment,desc,client_note,dev_note,程序不读,展示用字段,临时字段"
    zero_equal_fields: str = "reward"
    percent_fields: str = "rate,discount,percent"
    bool_fields: str = ""
    case_insensitive_fields: str = ""
    enable_general_compare: bool = True


@dataclass
class RewardItem:
    item_id: str
    quantity: int
    name: str = ""
    reward_type: str = ""
    source: str = ""
    raw: str = ""

    @property
    def key(self) -> str:
        # ID 优先；钻石统一为 item_id=1，避免 1000钻石道具 ID 与 mall reward 原子配置误报。
        return normalize_id(self.item_id)


@dataclass
class ValueRecord:
    record_id: str = ""
    name: str = ""
    sheet: str = ""
    row: int = 0
    price: Optional[int] = None
    recharge_price: Optional[int] = None
    pcid: str = ""
    duration: Optional[int] = None
    score: Optional[int] = None
    limit_time: Optional[int] = None
    rewards: List[RewardItem] = field(default_factory=list)
    raw: Dict[str, str] = field(default_factory=dict)

    def display_name(self) -> str:
        parts = []
        if self.name:
            parts.append(self.name)
        if self.record_id:
            parts.append(f"ID={self.record_id}")
        if self.sheet:
            parts.append(self.sheet)
        return " / ".join(parts) or f"第 {self.row} 行"


@dataclass
class ParseReport:
    reference_sheets: List[str] = field(default_factory=list)
    config_fields: List[str] = field(default_factory=list)
    reference_records: int = 0
    config_records: int = 0
    reference_rewards: int = 0
    matched_records: int = 0
    high_confidence_matches: int = 0
    medium_confidence_matches: int = 0
    low_confidence_matches: int = 0
    confirmed_field_diffs: int = 0
    confirmed_missing: int = 0
    confirmed_extra: int = 0
    pending_matches: int = 0
    filter_risks: int = 0
    notes: List[str] = field(default_factory=list)

    def summary(self) -> str:
        sheet_text = "、".join(self.reference_sheets[:8]) or "未识别"
        field_text = "、".join(self.config_fields[:10]) or "未识别"
        notes = "；".join(self.notes[:3])
        return (
            f"已识别参考 Sheet：{sheet_text} | 已识别配置字段：{field_text} | "
            f"参考记录 {self.reference_records} 条，参考奖励 {self.reference_rewards} 项，配置记录 {self.config_records} 条，匹配 {self.matched_records} 条"
            + (f" | {notes}" if notes else "")
        )


VALUE_COMPARE_MODE = "数值参考表 vs 游戏配置表"

ID_HEADERS = {"商品id", "商品ID", "goodsid", "goods_id", "#goods_id", "礼包对应id", "礼包id", "档位id", "id"}
NAME_HEADERS = {"礼包名称", "商品名称", "名称", "name", "comment", "^comment", "策划注释点此查看goods_id配置规范"}
PRICE_HEADERS = {"美刀定价", "代币价格", "price", "recharge_price", "真实充值"}
SCORE_HEADERS = {"累充积分值", "score", "积分", "分数"}
DURATION_HEADERS = {"持续时间", "duration", "周月卡的持续时间"}
LIMIT_HEADERS = {"限购次数", "limit_time"}
PCID_HEADERS = {"pcid", "平台商品id(有价格礼包不可填0)"}
REWARD_HEADER_RE = re.compile(r"奖励\s*\d*|装扮奖励|每日奖励|立即获得奖励|档位奖励|累充奖励|免费奖励|付费奖励", re.I)
REWARD_TOKEN_RE = re.compile(r"^\s*([^|*\s]+)\*([^|*\s]+)\*([^|*\s]+)\s*$")
NUMBER_RE = re.compile(r"^-?\d+(?:\.\d+)?$")


def compare_value_reference_to_config(config_path: str, reference_path: str, options: Optional[ValueCompareOptions] = None) -> Tuple[List[DiffItem], ParseReport]:
    """通用“数值参考表 vs 游戏配置表”结构化比对。

    左侧文件约定为游戏配置表，右侧文件约定为数值参考表。
    新逻辑默认按 Sheet、字段名行、主键、过滤条件、字段映射进行结构化比对，
    不再扫描整本 Excel / 不按行号比对 / 不把未映射字段直接当错误。
    """
    options = prepare_value_compare_options(options or ValueCompareOptions(), config_path, reference_path)
    # mall/商业化礼包参考表通常包含多段奖励区、嵌入式配置块和空表头列，
    # 并不是一张可直接按通用主键关联的平面表。该模板必须使用已验证的
    # 商业化专项解析器；通用结构化比对继续用于真正的平面配置表。
    use_mall_specialized_compare = options.template_name == "mall 礼包类"
    if getattr(options, "enable_general_compare", True) and not use_mall_specialized_compare:
        return compare_general_value_config(config_path, reference_path, options)

    # 兜底保留旧专项逻辑入口，原则上新版 UI 不再走到这里。
    config_records, config_fields = parse_config_records(config_path)
    reference_records, reference_sheets, reference_notes = parse_reference_records(reference_path)

    report = ParseReport(
        reference_sheets=reference_sheets,
        config_fields=config_fields,
        reference_records=len(reference_records),
        config_records=len(config_records),
        reference_rewards=sum(len(r.rewards) for r in reference_records),
        notes=reference_notes,
    )

    manual_ids = parse_user_list(options.manual_ids)
    manual_keywords = parse_user_list(options.keywords)
    matched_pairs = match_records(reference_records, config_records, manual_ids, manual_keywords)
    report.matched_records = len(matched_pairs)

    diffs: List[DiffItem] = []
    idx = 1

    if not reference_records:
        return [DiffItem(1, "参考表无法识别奖励", "未识别到参考表奖励或档位", "", "参考表", "请确认参考表包含 奖励/id/数量、商品ID、价格等字段")], report
    if not config_records:
        return [DiffItem(1, "配置表缺少对应商品", "", "未识别到配置表商品记录", "配置表", "请确认配置表存在 #goods_id、reward 等程序字段行")], report

    for ref in reference_records:
        cfg = matched_pairs.get(record_signature(ref))
        if not cfg:
            if manual_ids or manual_keywords or ref.record_id or ref.rewards:
                diffs.append(DiffItem(idx, "配置表缺少对应商品", ref.display_name(), "未匹配", ref.sheet, "参考表存在该礼包/档位，但配置表未匹配到对应商品"))
                idx += 1
            continue

        if options.compare_price:
            idx = compare_number_field(diffs, idx, ref, cfg, "价格", ref.price, cfg.price or cfg.recharge_price, allow_none=True)
        if options.compare_pcid:
            idx = compare_text_field(diffs, idx, ref, cfg, "pcid", ref.pcid, cfg.pcid, allow_none=True)
        if options.compare_duration:
            idx = compare_number_field(diffs, idx, ref, cfg, "持续时间", ref.duration, cfg.duration, allow_none=True)
        if options.compare_score:
            idx = compare_number_field(diffs, idx, ref, cfg, "积分/分数", ref.score, cfg.score, allow_none=True)
        if options.compare_limit_time:
            idx = compare_number_field(diffs, idx, ref, cfg, "限购次数", ref.limit_time, cfg.limit_time, allow_none=True)
        if options.compare_rewards:
            idx = compare_rewards(diffs, idx, ref, cfg)

    return reindex(diffs), report


# ---------------- 通用结构化配置比对 ----------------

KEY_CANDIDATES = ["#id", "id", "key", "#key", "config_id", "uid", "商品ID", "礼包ID", "配置ID", "name", "pcid"]
DEFAULT_IGNORE_FIELD_NAMES = {"备注", "说明", "comment", "desc", "client_note", "dev_note", "程序不读", "展示用字段", "临时字段"}
SHOP_TEMPLATE_NAMES = {"商店商品配置", "神秘商店", "商店", "商品配置", "shop_basic"}

# 通用结构化比对的字段近似映射。这里不改变字段值比对算法，只用于：
# 1) 自动识别中文/英文字段名；2) 自动建立字段映射；3) 自动主键匹配；4) 给用户更明确的预览。
FIELD_ALIAS_GROUPS: Dict[str, List[str]] = {
    "#id": ["#id", "id", "商品id", "商品ID", "礼包id", "礼包ID", "配置id", "配置ID", "档位id", "档位ID", "goodsid", "goods_id", "mallid", "mall_id", "productid", "product_id"],
    "name": ["name", "$name", "名称", "礼包名称", "商品名称", "显示名称", "档位名称", "标题", "title", "comment", "^comment"],
    "price": ["price", "price_practical", "recharge_price", "价格", "售价", "美刀定价", "美元价格", "充值金额", "真实充值", "代币价格", "原价", "现价"],
    "pcid": ["pcid", "支付id", "支付ID", "平台商品id", "平台商品ID", "档位id", "档位ID", "商品档位", "支付档位", "payid", "pay_id"],
    "reward": ["reward", "item", "奖励", "奖励内容", "礼包内容", "道具", "道具id", "道具ID", "奖励道具", "奖励配置", "内容"],
    "trigger_level": ["trigger_level", "解锁等级", "触发等级", "开启等级", "主堡等级", "城堡等级", "城镇等级", "等级条件", "openlevel", "open_level", "level", "lv"],
    "open_end_conditions": ["open_end_conditions", "open_condition", "open_conditions", "condition", "条件", "触发条件", "解锁条件", "开启条件", "open_end_conditions1", "open_end_conditions2", "open_end_conditions3"],
    "type": ["type", "类型", "分类", "活动ID", "活动id", "父活动", "父活动ID", "父活动id", "activity_id", "activityid", "event_id", "eventid", "group", "分组", "category", "activity_type"],
    "order": ["order", "sort", "排序", "显示顺序", "权重排序"],
    "limit_time": ["limit_time", "limit", "限购", "限购次数", "购买次数", "购买限制"],
    "duration": ["duration", "持续时间", "有效期", "时长", "礼包持续时间"],
    "count": ["count", "num", "number", "quantity", "数量", "个数", "次数", "份数"],
    "weight": ["weight", "rate", "probability", "权重", "概率", "掉率", "返利比", "价值", "总价值", "原价值"],
}

FIELD_ALIAS_LOOKUP: Dict[str, str] = {}
for _canon, _aliases in FIELD_ALIAS_GROUPS.items():
    for _alias in _aliases:
        FIELD_ALIAS_LOOKUP[re.sub(r"\s+", "", str(_alias)).strip().lower()] = _canon


def normalize_value_template_name(template_name: str, config_path: str = "", reference_path: str = "") -> str:
    text = str(template_name or "").strip()
    compact = re.sub(r"\s+", "", text).lower()
    names_compact = re.sub(r"\s+", "", (Path(str(config_path or "")).name + " " + Path(str(reference_path or "")).name)).lower()
    if any(k in compact for k in ["mall", "礼包", "商业化", "月卡", "周卡", "弹窗", "自选", "整页", "vip", "首选"]):
        return "mall 礼包类"
    if "mall" in names_compact and any(k in names_compact for k in ["礼包", "商业化", "月卡", "周卡", "弹窗", "自选", "整页", "vip", "首选"]):
        return "mall 礼包类"
    if any(k.lower() in compact for k in ["神秘商店", "商店商品", "商品配置", "shop_basic", "mysteryshop"]):
        return "商店商品配置"
    # 兼容旧包没有模板下拉框时的神秘商店三表场景：只有明确文件名命中时才自动推荐。
    names = (Path(str(config_path or "")).name + " " + Path(str(reference_path or "")).name).lower()
    if ("shop_basic" in names) and ("神秘商店" in names or "mystery" in names):
        return "商店商品配置"
    if "活动" in text or "activity" in compact or "event" in compact:
        return "活动配置"
    return text or "通用配置表"


def prepare_value_compare_options(options: Optional[ValueCompareOptions], config_path: str = "", reference_path: str = "") -> ValueCompareOptions:
    """Return one canonical compare config used by preview, execution and export.

    This prevents UI/preview/export from each rebuilding their own filter values.
    """
    base = options or ValueCompareOptions()
    prepared = replace(base)
    prepared.template_name = normalize_value_template_name(getattr(base, "template_name", ""), config_path, reference_path)
    if prepared.template_name == "商店商品配置":
        if not str(prepared.reference_filter or "").strip():
            prepared.reference_filter = "type=3"
        if not str(prepared.config_filter or "").strip():
            prepared.config_filter = "type=3"
    elif prepared.template_name == "mall 礼包类":
        # mall.csv 同时承载 type=3 月周卡、type=13 弹窗、type=2000 活动商品等业务。
        # 不再预填单一 type 过滤，也不强制 name 主键；这两个默认值会排除 type=3
        # 月周卡，并把同系列的多个价格档位折叠成重复主键。
        mall_ignored = "备注,说明,comment,desc,client_note,dev_note,程序不读,展示用字段,临时字段,价值,原价值,总价值,返利比,折扣价值,性价比,展示文案"
        existing_ignore = str(prepared.ignore_fields or "").strip()
        prepared.ignore_fields = (existing_ignore + "," + mall_ignored).strip(",") if existing_ignore else mall_ignored
    return prepared


def build_filter_risk_messages(ref_preview: FilterPreview, cfg_previews: List[FilterPreview]) -> List[str]:
    messages: List[str] = []
    cfg_total_after = sum(p.after_count for p in cfg_previews)
    cfg_total_before = sum(p.before_count for p in cfg_previews)
    ref_has_filter = bool(str(ref_preview.expression or "").strip())
    cfg_has_filter = any(str(p.expression or "").strip() for p in cfg_previews)
    if ref_has_filter and not cfg_has_filter:
        messages.append(
            f"参考表已设置过滤条件 {ref_preview.expression}，但配置表过滤条件为空；配置表将按全量 {cfg_total_before} 条参与比对，可能产生非目标范围多配。"
        )
    if cfg_has_filter and not ref_has_filter:
        messages.append(
            f"配置表已设置过滤条件，但参考表过滤条件为空；参考表将按全量 {ref_preview.before_count} 条参与比对，请确认范围是否一致。"
        )
    if not cfg_has_filter and cfg_total_after > ref_preview.after_count and ref_preview.after_count > 0:
        messages.append(
            f"配置表过滤条件为空且配置表数据量 {cfg_total_after} 明显大于参考表 {ref_preview.after_count}，多配结果可能来自未过滤全表。"
        )
    if ref_preview.after_count and cfg_total_after and abs(cfg_total_after - ref_preview.after_count) >= max(10, int(ref_preview.after_count * 0.1)):
        messages.append(
            f"参考表过滤后 {ref_preview.after_count} 条，配置表过滤后 {cfg_total_after} 条，数量差异较大，请确认过滤条件和主键范围。"
        )
    for p in [ref_preview] + cfg_previews:
        if "数量未变化" in p.status and str(p.expression or "").strip():
            messages.append(f"{p.source_type} {p.file_name} 过滤前后数量一致，请确认过滤条件 {p.expression} 是否符合预期。")
    return messages



def normalize_match_text(value: str) -> str:
    s = normalize_name_key(value)
    s = re.sub(r"\s+", "", s)
    return s


def row_best_field(row: StructuredRow, canonical_names: Iterable[str]) -> str:
    # Fast path: add_standard_fields_to_rows injects canonical fields such as name/pcid/price.
    for canon in canonical_names:
        if canon in row and stringify(row.get(canon, "")):
            return stringify(row.get(canon, ""))
    keys = list(row.keys())
    for canon in canonical_names:
        actual = first_matching_field(keys, canon)
        if actual and stringify(row.get(actual, "")):
            return stringify(row.get(actual, ""))
    return ""


def row_name_value(row: StructuredRow) -> str:
    return row_best_field(row, ["name", "$name", "^comment", "名称", "礼包名称"])


def row_pcid_value(row: StructuredRow) -> str:
    return normalize_basic_value(row_best_field(row, ["pcid", "支付ID", "档位ID"]))


def row_price_value(row: StructuredRow) -> str:
    return normalize_price_value(row_best_field(row, ["price", "价格", "售价", "price_practical"]))


def row_id_value(row: StructuredRow) -> str:
    return normalize_basic_value(row_best_field(row, ["#id", "id", "#goods_id", "goods_id", "商品ID", "配置ID", "礼包ID"]))


def row_section_value(row: StructuredRow) -> str:
    return stringify(row.get("section_name") or row.get("__section") or row.get("__region") or "")


def match_name_similarity(a: str, b: str) -> float:
    aa = normalize_match_text(a)
    bb = normalize_match_text(b)
    if not aa or not bb:
        return 0.0
    if aa == bb:
        return 1.0
    if aa in bb or bb in aa:
        return 0.82
    # 防止包含长说明/规则文本时 SequenceMatcher 过慢。长文本用词元重合度近似。
    if len(aa) > 80 or len(bb) > 80:
        ta = set(re.findall(r"[A-Za-z0-9]+|[\u4e00-\u9fff]{1,4}", aa))
        tb = set(re.findall(r"[A-Za-z0-9]+|[\u4e00-\u9fff]{1,4}", bb))
        if ta and tb:
            return len(ta & tb) / max(1, len(ta | tb))
        aa = aa[:80]
        bb = bb[:80]
    return SequenceMatcher(None, aa, bb).ratio()


def critical_series_token_conflict(ref_name: str, cfg_name: str) -> bool:
    ref_norm = normalize_match_text(ref_name)
    cfg_norm = normalize_match_text(cfg_name)
    critical_tokens = ["宝石", "宠物", "英雄", "城建", "成长"]
    for token in critical_tokens:
        if token in ref_norm and token not in cfg_norm:
            return True
    return False


def assess_match_confidence(ref_row: StructuredRow, cfg_row: StructuredRow, key_fields: List[str]) -> Tuple[str, str]:
    """Return 高 / 中 / 低 and a human-readable reason.

    The goal is to avoid turning weak name-only / price-only matches into hundreds
    of false field differences. High-confidence pairs are safe for field-by-field
    comparison; medium/low pairs should be reviewed by QA.
    """
    ref_id = row_id_value(ref_row)
    cfg_id = row_id_value(cfg_row)
    ref_name = row_name_value(ref_row)
    cfg_name = row_name_value(cfg_row)
    ref_pcid = row_pcid_value(ref_row)
    cfg_pcid = row_pcid_value(cfg_row)
    ref_price = row_price_value(ref_row)
    cfg_price = row_price_value(cfg_row)
    name_score = match_name_similarity(ref_name, cfg_name)
    key_canons = {canonical_field_name(k) for k in key_fields}
    token_conflict = critical_series_token_conflict(ref_name, cfg_name)

    if ref_id and cfg_id and ref_id == cfg_id:
        return "高", f"id 完全一致：{ref_id}"
    if name_score >= 0.98 and ref_name and cfg_name and not token_conflict:
        return "高", "name 完全一致"
    if ref_pcid and cfg_pcid and ref_pcid == cfg_pcid and ref_pcid not in {"0", "0.0"}:
        if (ref_price and cfg_price and ref_price == cfg_price) or name_score >= 0.55:
            return "高", "pcid 一致，且价格或名称也匹配"
        return "中", "pcid 一致，但名称/价格不足以唯一确认"
    if name_score >= 0.75 and ref_price and cfg_price and ref_price == cfg_price and not token_conflict:
        return "高", "name 高相似且价格一致"
    if name_score >= 0.60:
        return "中", "name 相似，但存在系列关键词缺失或缺少稳定主键/价格/pcid佐证"
    if ref_price and cfg_price and ref_price == cfg_price:
        return "低", "仅价格一致"
    if "name" in key_canons:
        return "中", "使用 name 作为主键，但匹配证据不足"
    return "低", "缺少稳定匹配依据"


def best_candidate_match(ref_row: StructuredRow, cfg_rows: List[StructuredRow]) -> Tuple[Optional[StructuredRow], str, str]:
    ref_name = row_name_value(ref_row)
    ref_pcid = row_pcid_value(ref_row)
    ref_price = row_price_value(ref_row)
    best = None
    best_score = 0.0
    best_reason = ""
    for cfg_row in cfg_rows:
        cfg_name = row_name_value(cfg_row)
        cfg_pcid = row_pcid_value(cfg_row)
        cfg_price = row_price_value(cfg_row)
        name_score = match_name_similarity(ref_name, cfg_name)
        score = name_score
        reason_parts = []
        if ref_pcid and cfg_pcid and ref_pcid == cfg_pcid and ref_pcid not in {"0", "0.0"}:
            score += 0.45
            reason_parts.append("pcid 一致")
        if ref_price and cfg_price and ref_price == cfg_price:
            score += 0.25
            reason_parts.append("价格一致")
        if critical_series_token_conflict(ref_name, cfg_name):
            score -= 0.60
            reason_parts.append("系列关键词不完整")
        if name_score >= 0.98:
            score += 0.6
            reason_parts.append("名称一致")
        elif name_score >= 0.60:
            score += 0.25
            reason_parts.append("名称相似")
        if score > best_score:
            best_score = score
            best = cfg_row
            best_reason = "，".join(reason_parts) or "名称弱相似"
    if best is None or best_score < 0.60:
        return None, "低", "未找到可信候选"
    # 名称高度相似且有 pcid/价格佐证，可作为高可信候选进入字段比对；弱匹配仍仅待确认。
    best_name_score = match_name_similarity(ref_name, row_name_value(best))
    best_pcid = row_pcid_value(best)
    best_price = row_price_value(best)
    if best_name_score >= 0.78 and not critical_series_token_conflict(ref_name, row_name_value(best)) and ((ref_pcid and best_pcid and ref_pcid == best_pcid and ref_pcid not in {"0", "0.0"}) or (ref_price and best_price and ref_price == best_price)):
        return best, "高", "主键未完全一致，但名称高度相似且 pcid/价格佐证；" + best_reason
    if best_score >= 1.10:
        return best, "中", "存在候选但未通过当前主键确认；" + best_reason
    return best, "低", "存在低可信候选；" + best_reason


def has_strong_filter_risk(ref_preview: FilterPreview, cfg_previews: List[FilterPreview], options: ValueCompareOptions) -> bool:
    cfg_after = sum(p.after_count for p in cfg_previews)
    ref_after = max(1, ref_preview.after_count)
    # “填写了过滤条件”不代表过滤范围就合理。过滤后配置量仍远大于参考量时，
    # 应合并为一条过滤风险，而不是逐条输出数百条疑似多配。
    if cfg_after >= max(ref_after * 2, ref_after + 20):
        return True
    return not str(options.config_filter or "").strip() and cfg_after > ref_after

def compare_general_value_config(config_path: str, reference_path: str, options: ValueCompareOptions) -> Tuple[List[DiffItem], ParseReport]:
    options = prepare_value_compare_options(options, config_path, reference_path)
    ref_table = load_structured_table(reference_path, options.reference_sheet, options.reference_header_row, options.reference_data_start_row, role="reference")
    cfg_tables = load_config_tables(config_path, options.config_sheet, options.config_header_row, options.config_data_start_row)
    if not cfg_tables:
        diag = build_structured_table_diagnostics(config_path, options.config_sheet, role="config")
        report = ParseReport(reference_sheets=[ref_table.sheet_name] if ref_table else [], config_fields=[], notes=["未读取到配置表数据", diag])
        return [DiffItem(1, "执行条件不足", "", "未读取到配置表数据", "配置表", diag)], report

    if ref_table is None or not ref_table.headers:
        diag = build_structured_table_diagnostics(reference_path, options.reference_sheet, role="reference")
        report = ParseReport(config_fields=merge_structured_tables(cfg_tables).headers if cfg_tables else [], notes=["未读取到参考表数据", diag])
        return [DiffItem(1, "执行条件不足", "未读取到参考表字段", "", "参考表", diag)], report

    try:
        ref_rows, ref_preview = filter_table_rows(ref_table, options.reference_filter, "参考表")
        filtered_cfg_tables: List[StructuredTable] = []
        cfg_previews: List[FilterPreview] = []
        for table in cfg_tables:
            filtered_rows, preview = filter_table_rows(table, options.config_filter, "配置表")
            filtered_cfg_tables.append(replace(table, rows=filtered_rows))
            cfg_previews.append(preview)
    except FilterValidationError as exc:
        report = ParseReport(reference_sheets=[ref_table.sheet_name], config_fields=merge_structured_tables(cfg_tables).headers, notes=[str(exc)])
        return [DiffItem(1, "执行条件不足", "过滤条件未生效", "过滤条件校验失败", "过滤条件", str(exc))], report

    cfg_table = merge_structured_tables(filtered_cfg_tables)
    cfg_rows = cfg_table.rows

    key_fields = resolve_key_fields(options.key_fields, ref_table.headers, cfg_table.headers)
    report = ParseReport(
        reference_sheets=[ref_table.sheet_name],
        config_fields=cfg_table.headers,
        reference_records=len(ref_rows),
        config_records=len(cfg_rows),
        notes=[],
    )
    for note in getattr(ref_table, "notes", []) or []:
        report.notes.append(note)
    report.notes.extend([
        f"结构化比对：参考表 Sheet={ref_table.sheet_name}，配置表={cfg_table.display_name}",
        f"字段名行：参考表第 {ref_table.header_row + 1} 行，配置表第 {cfg_table.header_row + 1} 行",
        f"过滤后数据：参考表 {len(ref_rows)} 条，配置表 {len(cfg_rows)} 条",
        f"参考表过滤状态：{ref_preview.status}；" + "；".join(f"{p.file_name}:{p.status}" for p in cfg_previews[:4]),
    ])

    if not key_fields:
        return [DiffItem(1, "执行条件不足", "未识别到主键字段", "未识别到主键字段", "主键设置", "请手动填写主键字段，例如 #id、id、key；新版默认不按行号比对")], report

    mapping = build_field_mapping(ref_table.headers, cfg_table.headers, options.field_mappings, options.ignore_fields, key_fields)
    if not mapping:
        return [DiffItem(1, "执行条件不足", "字段映射为空", "字段映射为空", "字段映射", "请确认两表有同名字段，或手动填写字段映射，例如 item=>reward_id")], report

    diffs: List[DiffItem] = []
    idx = 1

    ref_index, ref_dups = index_rows_by_key(ref_rows, key_fields)
    cfg_index, cfg_dups = index_rows_by_key(cfg_rows, key_fields)
    report.matched_records = len(set(ref_index) & set(cfg_index))

    filter_risks = build_filter_risk_messages(ref_preview, cfg_previews)
    strong_filter_risk = has_strong_filter_risk(ref_preview, cfg_previews, options)
    report.filter_risks = len(filter_risks)
    summary_lines = [
        f"模板：{options.template_name or '通用配置表'}",
        f"参考表：{Path(reference_path).name} / Sheet={ref_table.sheet_name}",
        f"配置表：{Path(config_path).name} / {cfg_table.display_name}",
        f"主键字段：{'+'.join(key_fields)}",
        f"参考表过滤条件：{options.reference_filter or '无'}；配置表过滤条件：{options.config_filter or '无'}",
        f"字段映射数量：{len(mapping)}；忽略字段：{options.ignore_fields or '默认'}",
        "本次按主键关联，不按行号关联；低可信匹配不计入确认字段差异，会进入待确认匹配项。",
    ]
    diffs.append(DiffItem(idx, "比对摘要", "\n".join(summary_lines), "", "结构化配置比对", "摘要信息，不作为错误")); idx += 1
    diffs.append(DiffItem(idx, "比对摘要", ref_preview.to_summary_text(), "", "过滤预览 / 参考表", "参考表过滤条件已校验")); idx += 1
    for preview in cfg_previews:
        diffs.append(DiffItem(idx, "比对摘要", "", preview.to_summary_text(), f"过滤预览 / 配置表 / {preview.file_name}", "配置表过滤条件已校验；多配置表逐个校验")); idx += 1
    for note in getattr(ref_table, "notes", []) or []:
        diffs.append(DiffItem(idx, "解析诊断", note, "", "参考表 Sheet 自动回退", "文件切换后旧 Sheet 不存在时自动回退 auto 扫描")); idx += 1
    for risk in filter_risks:
        diffs.append(DiffItem(idx, "过滤风险", "", risk, "过滤条件", "该风险不代表字段值不一致；请确认是否需要补充过滤条件后再判断漏配/多配")); idx += 1
        report.notes.append(risk)

    for key, rows in sorted(ref_dups.items()):
        diffs.append(DiffItem(idx, "待确认匹配项", f"参考表主键重复：{key}", "", f"参考表第 {','.join(str(r.__rownum__) for r in rows)} 行", "同一主键存在多行，请确认是否应使用联合主键或过滤条件")); idx += 1
        report.pending_matches += 1
    for key, rows in sorted(cfg_dups.items()):
        diffs.append(DiffItem(idx, "待确认匹配项", "", f"配置表主键重复：{key}", f"配置表第 {','.join(str(r.__rownum__) for r in rows)} 行", "同一主键存在多行，请确认是否应使用联合主键或过滤条件")); idx += 1
        report.pending_matches += 1

    zero_equal_fields = parse_field_set(options.zero_equal_fields)
    percent_fields = parse_field_set(options.percent_fields)
    bool_fields = parse_field_set(options.bool_fields)
    case_insensitive_fields = parse_field_set(options.case_insensitive_fields)

    no_baseline_fields: Dict[Tuple[str, str], int] = {}
    confirmed_keys = set()
    pending_keys = set()

    for key in sorted(set(ref_index) & set(cfg_index), key=natural_key):
        ref_row = ref_index[key]
        cfg_row = cfg_index[key]
        confidence, reason = assess_match_confidence(ref_row, cfg_row, key_fields)
        if confidence == "高":
            confirmed_keys.add(key)
            report.high_confidence_matches += 1
        elif confidence == "中":
            pending_keys.add(key)
            report.medium_confidence_matches += 1
            report.pending_matches += 1
            diffs.append(DiffItem(idx, "待确认匹配项", row_to_preview(ref_row), row_to_preview(cfg_row), f"主键={key}", f"匹配可信度：中；匹配依据：{reason}。未计入确认字段差异，请人工确认是否同一礼包/档位。")); idx += 1
            continue
        else:
            pending_keys.add(key)
            report.low_confidence_matches += 1
            report.pending_matches += 1
            diffs.append(DiffItem(idx, "待确认匹配项", row_to_preview(ref_row), row_to_preview(cfg_row), f"主键={key}", f"匹配可信度：低；匹配依据：{reason}。未计入确认字段差异。")); idx += 1
            continue

        for ref_field, cfg_field in mapping:
            ref_val = ref_row.get(ref_field, "")
            cfg_val = cfg_row.get(cfg_field, "")
            ref_norm = normalize_compare_cell(ref_val, ref_field, zero_equal_fields, percent_fields, bool_fields, case_insensitive_fields)
            cfg_norm = normalize_compare_cell(cfg_val, cfg_field, zero_equal_fields, percent_fields, bool_fields, case_insensitive_fields)
            if ref_norm != cfg_norm:
                if not ref_norm and cfg_norm:
                    no_baseline_fields[(ref_field, cfg_field)] = no_baseline_fields.get((ref_field, cfg_field), 0) + 1
                    continue
                diff_type, remark_suffix, confirmed = classify_structured_field_difference(ref_field, cfg_field, ref_val, cfg_val, ref_row, cfg_row)
                diffs.append(DiffItem(
                    idx,
                    diff_type,
                    f"{ref_field}: {ref_val}",
                    f"{cfg_field}: {cfg_val}",
                    f"主键={key} / 参考表第 {ref_row.__rownum__} 行 / 配置表第 {cfg_row.__rownum__} 行",
                    f"匹配可信度：高；匹配依据：{reason}。标准化后参考值：{ref_norm or '空'}；配置值：{cfg_norm or '空'}。{remark_suffix}",
                ))
                idx += 1
                if confirmed:
                    report.confirmed_field_diffs += 1
                else:
                    report.pending_matches += 1
            elif ref_norm == cfg_norm and stringify(ref_val).strip() != stringify(cfg_val).strip() and canonical_field_name(ref_field) in {"weight"}:
                diffs.append(DiffItem(idx, "字段格式差异但语义一致", f"{ref_field}: {ref_val}", f"{cfg_field}: {cfg_val}", f"主键={key} / 参考表第 {ref_row.__rownum__} 行 / 配置表第 {cfg_row.__rownum__} 行", "格式不同但标准化后语义一致，不计入确认差异")); idx += 1

    cfg_all_rows = list(cfg_rows)
    # 参考表缺失：如果存在候选项但未高可信确认，则放入待确认；否则才算确认漏配。
    for key in sorted(set(ref_index) - set(cfg_index), key=natural_key):
        row = ref_index[key]
        candidate, confidence, reason = best_candidate_match(row, cfg_all_rows)
        if candidate is not None and confidence == "高":
            report.high_confidence_matches += 1
            diffs.append(DiffItem(idx, "高可信匹配", row_to_preview(row), row_to_preview(candidate), f"参考表主键={key} / 参考表第 {row.__rownum__} 行 / 配置表第 {candidate.__rownum__} 行", f"主键未完全一致，但候选匹配可信度高；匹配依据：{reason}。已继续执行字段比对。")); idx += 1
            for ref_field, cfg_field in mapping:
                ref_val = row.get(ref_field, "")
                cfg_val = candidate.get(cfg_field, "")
                ref_norm = normalize_compare_cell(ref_val, ref_field, zero_equal_fields, percent_fields, bool_fields, case_insensitive_fields)
                cfg_norm = normalize_compare_cell(cfg_val, cfg_field, zero_equal_fields, percent_fields, bool_fields, case_insensitive_fields)
                if ref_norm != cfg_norm:
                    if not ref_norm and cfg_norm:
                        no_baseline_fields[(ref_field, cfg_field)] = no_baseline_fields.get((ref_field, cfg_field), 0) + 1
                        continue
                    diff_type, remark_suffix, confirmed = classify_structured_field_difference(ref_field, cfg_field, ref_val, cfg_val, row, candidate)
                    diffs.append(DiffItem(idx, diff_type, f"{ref_field}: {ref_val}", f"{cfg_field}: {cfg_val}", f"高可信候选 / 参考表主键={key} / 参考表第 {row.__rownum__} 行 / 配置表第 {candidate.__rownum__} 行", f"匹配依据：{reason}。标准化后参考值：{ref_norm or '空'}；配置值：{cfg_norm or '空'}。{remark_suffix}")); idx += 1
                    if confirmed:
                        report.confirmed_field_diffs += 1
                    else:
                        report.pending_matches += 1
        elif candidate is not None and confidence == "中":
            report.pending_matches += 1
            diffs.append(DiffItem(idx, "待确认匹配项", row_to_preview(row), row_to_preview(candidate), f"参考表主键={key} / 参考表第 {row.__rownum__} 行", f"未找到同主键记录，但存在候选配置；匹配可信度：{confidence}；匹配依据：{reason}。请人工确认是否同一礼包/档位。")); idx += 1
        else:
            report.confirmed_missing += 1
            diffs.append(DiffItem(idx, "疑似漏配", row_to_preview(row), "未配置", f"主键={key} / 参考表第 {row.__rownum__} 行", "参考表存在该记录，配置表过滤后未找到高可信或中可信候选项；请结合业务过滤条件确认是否真漏配")); idx += 1

    extra_keys = sorted(set(cfg_index) - set(ref_index), key=natural_key)
    if extra_keys and strong_filter_risk:
        sample = extra_keys[:20]
        report.filter_risks += 1
        diffs.append(DiffItem(
            idx,
            "过滤风险",
            "参考表未覆盖这些配置记录",
            "；".join(row_to_preview(cfg_index[k], limit=4) for k in sample),
            "疑似多配 / 配置表未过滤",
            f"配置表未设置业务范围过滤，存在 {len(extra_keys)} 条未匹配配置。该类记录暂不计入确认多配；请按 type/sub_type/goods_group/^comment/name 等字段过滤后再判断。",
        ))
        idx += 1
    else:
        for key in extra_keys:
            row = cfg_index[key]
            report.confirmed_extra += 1
            diffs.append(DiffItem(idx, "疑似多配", "未填写", row_to_preview(row), f"主键={key} / 配置表第 {row.__rownum__} 行", "配置表存在该记录，但参考表过滤后未找到同主键记录；若未设置配置表过滤，请先确认业务范围")); idx += 1

    for (ref_field, cfg_field), count in sorted(no_baseline_fields.items()):
        diffs.append(DiffItem(idx, "待确认匹配项", f"参考表字段 {ref_field} 未提供明确值", f"配置表字段 {cfg_field} 存在值", "参考表无基准字段", f"共 {count} 条记录参考表为空但配置表有值，未计入确认差异；如需比对，请补充参考值或取消该字段映射"))
        idx += 1
        report.pending_matches += 1

    ref_mapped = {m[0] for m in mapping} | set(key_fields)
    cfg_mapped = {m[1] for m in mapping} | set(key_fields)
    ref_mapped_norm = {normalize_field_name(x) for x in ref_mapped}
    cfg_mapped_norm = {normalize_field_name(x) for x in cfg_mapped}
    ref_mapped_canon = {canonical_field_name(x) for x in ref_mapped}
    cfg_mapped_canon = {canonical_field_name(x) for x in cfg_mapped}
    ignored = build_ignore_set(options.ignore_fields)
    reward_source_skipped = 0
    for field in ref_table.headers:
        if not clean_header(field):
            continue
        if normalize_field_name(field) in ref_mapped_norm or canonical_field_name(field) in ref_mapped_canon or normalize_field_name(field) in ignored:
            continue
        if is_reference_reward_source_field(field):
            reward_source_skipped += 1
            continue
        diffs.append(DiffItem(idx, "未映射字段", f"参考表字段：{field}", "", "参考表", "未参与确认差异统计；如需比对，请在字段映射中补充")); idx += 1
    if reward_source_skipped:
        diffs.append(DiffItem(idx, "解析诊断", f"参考表 reward 源列 {reward_source_skipped} 个已参与奖励生成或被识别为价值/数量辅助列", "", "参考表 reward 解析", "这些列不会逐列生成未映射字段，避免把价值列当作 reward 或字段差异；未静默丢弃，已在此诊断汇总。")); idx += 1
    for field in cfg_table.headers:
        if not clean_header(field):
            continue
        if normalize_field_name(field) in cfg_mapped_norm or canonical_field_name(field) in cfg_mapped_canon or normalize_field_name(field) in ignored:
            continue
        diffs.append(DiffItem(idx, "未映射字段", "", f"配置表字段：{field}", "配置表", "参考表未建立明确映射，不作为错误；如为类型级/额外配置，请人工确认")); idx += 1

    diffs.append(DiffItem(idx, "解析诊断", "解析字段", f"参考表字段：{', '.join(ref_table.headers[:30])}\n配置表字段：{', '.join(cfg_table.headers[:30])}", "解析诊断", "该行用于定位字段识别和映射情况，不作为错误")); idx += 1
    diffs.append(DiffItem(idx, "已忽略项/规则说明", "未扫描未选择的 Sheet；未按行号关联；低可信匹配不算确认差异", f"空值标准化：None/NaN/空白一致；数字 1 与 1.0 一致；字段 {options.zero_equal_fields or '无'} 允许空值与 0 等价；字段 {options.percent_fields or '无'} 支持小数/百分比等价", "规则说明", "该行不作为错误"))
    return reindex(diffs), report


@dataclass
class StructuredRow(dict):
    __rownum__: int = 0


@dataclass
class StructuredTable:
    source_path: str
    sheet_name: str
    display_name: str
    headers: List[str]
    field_types: Dict[str, str]
    rows: List[StructuredRow]
    header_row: int
    data_start_row: int
    notes: List[str] = field(default_factory=list)






@dataclass
class FilterCondition:
    field: str
    op: str
    values: List[str] = field(default_factory=list)


@dataclass
class FilterPreview:
    source_type: str
    file_name: str
    sheet_name: str
    header_row: int
    data_start_row: int
    expression: str
    field_check: str
    before_count: int
    after_count: int
    status: str
    message: str = ""

    @property
    def filtered_count(self) -> int:
        return max(0, self.before_count - self.after_count)

    def to_summary_text(self) -> str:
        expr = self.expression or "无"
        msg = f"；{self.message}" if self.message else ""
        return (
            f"{self.source_type}：{self.file_name} / Sheet={self.sheet_name}\n"
            f"字段名行：第 {self.header_row + 1} 行；数据起始行：第 {self.data_start_row + 1} 行\n"
            f"过滤条件：{expr}；字段检查：{self.field_check}\n"
            f"过滤前：{self.before_count}；过滤后：{self.after_count}；过滤掉：{self.filtered_count}\n"
            f"过滤状态：{self.status}{msg}"
        )


class FilterValidationError(ValueError):
    pass

def merge_structured_tables(tables: List[StructuredTable]) -> StructuredTable:
    if len(tables) == 1:
        return tables[0]
    headers: List[str] = []
    field_types: Dict[str, str] = {}
    rows: List[StructuredRow] = []
    names = []
    for table in tables:
        names.append(table.display_name)
        for h in table.headers:
            if h not in headers:
                headers.append(h)
            if h not in field_types and h in table.field_types:
                field_types[h] = table.field_types[h]
        for row in table.rows:
            merged = StructuredRow()
            merged.__rownum__ = row.__rownum__
            for h in headers:
                merged[h] = row.get(h, "")
            # 保留来源，方便导出定位。
            merged["__source_file"] = table.display_name
            rows.append(merged)
    return StructuredTable(
        source_path=";".join(t.source_path for t in tables),
        sheet_name="多配置表",
        display_name=";".join(names),
        headers=headers,
        field_types=field_types,
        rows=rows,
        header_row=tables[0].header_row,
        data_start_row=tables[0].data_start_row,
    )

def load_config_tables(path: str, sheet_name: str, header_row: str, data_start_row: str) -> List[StructuredTable]:
    # 当前 UI 支持单文件/多文件；多文件用 ; 分隔。
    paths = [p.strip() for p in re.split(r"[;；]\s*", str(path or "")) if p.strip()]
    tables: List[StructuredTable] = []
    for single in paths:
        table = load_structured_table(single, sheet_name, header_row, data_start_row, role="config")
        if table is not None and table.headers:
            tables.append(table)
    return tables


def load_structured_table(path: str, sheet_name: str = "", header_row: str = "auto", data_start_row: str = "auto", role: str = "config") -> Optional[StructuredTable]:
    p = Path(path)
    suffix = p.suffix.lower()
    if suffix == ".csv":
        raw = read_text_with_fallback(p)
        rows = list(csv.reader(raw.splitlines()))
        return rows_to_structured_table(str(p), p.name, p.name, rows, header_row, data_start_row, role=role)
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        from openpyxl import load_workbook
        wb = load_workbook(p, data_only=True, read_only=True)
        try:
            requested = str(sheet_name or "").strip()
            selected_note = ""
            selected = choose_sheet_for_structured_compare(wb, requested, role)
            if selected is None and requested:
                selected = choose_sheet_for_structured_compare(wb, "", role)
                if selected:
                    selected_note = f"旧/指定 Sheet「{requested}」在当前文件中不存在，已自动回退为 auto 并重新扫描当前文件；当前文件包含 Sheet：{', '.join(wb.sheetnames)}。"
            if selected is None:
                return None
            ws = wb[selected]
            rows = [[cell for cell in row] for row in ws.iter_rows(values_only=True)]
            table = rows_to_structured_table(str(p), selected, selected, rows, header_row, data_start_row, role=role)
            if table is not None and selected_note:
                table.notes.append(selected_note)
            return table
        finally:
            wb.close()
    if suffix == ".json":
        data = json.loads(read_text_with_fallback(p))
        rows = json_to_rows(data)
        return rows_to_structured_table(str(p), p.name, p.name, rows, header_row, data_start_row, role=role)
    raw = read_text_with_fallback(p)
    rows = list(csv.reader(raw.splitlines()))
    return rows_to_structured_table(str(p), p.name, p.name, rows, header_row, data_start_row, role=role)


def choose_sheet_for_structured_compare(wb, requested: str, role: str) -> Optional[str]:
    names = list(wb.sheetnames)
    if requested and requested.strip():
        req = requested.strip()
        for n in names:
            if n == req or n.strip().lower() == req.lower():
                return n
        # 请求了不存在的 Sheet 时，宁可返回 None，不默认全表扫描。
        return None

    # 自动扫描所有 Sheet，而不是只读取第一个 Sheet。优先选择“字段结构最像数值表/配置表”的 Sheet。
    # 对参考表保留常见 Sheet 名加分，但如果该 Sheet 没有有效字段，不再强行选择它。
    best = None
    best_score = -1
    for ws in wb.worksheets:
        max_row = min(ws.max_row or 0, 80)
        sample = [[cell for cell in row] for row in ws.iter_rows(min_row=1, max_row=max_row, values_only=True)]
        candidates = detect_structured_header_rows(sample, max_rows=80)
        score = 0
        if candidates:
            score = candidates[0][1]
        title = ws.title.strip()
        if role == "reference" and title in {"配置参考表", "参考表", "数值参考表"}:
            score += 8
        if re.search(r"自选|弹窗|礼包|商城|mall|配置|参考", title, re.I):
            score += 4
        if re.search(r"引用|索引|说明|备注|字典", title, re.I):
            score -= 40
        # 只有识别到表头才认为是有效 Sheet，避免说明/引用 Sheet 抢占。
        if candidates and score > best_score:
            best_score = score
            best = ws.title
    return best


def rows_to_structured_table(source_path: str, sheet_name: str, display_name: str, rows: List[List[object]], header_row: str, data_start_row: str, role: str = "config") -> Optional[StructuredTable]:
    if not rows:
        return None
    h_idx = parse_row_number(header_row)
    if h_idx is not None:
        # 用户手动指定字段名行时，严格按该行解析。
        return build_structured_table_region(source_path, sheet_name, display_name, rows, h_idx, parse_row_number(data_start_row), len(rows), role=role)

    # 自动识别：允许前面有标题、说明、空行、合并单元格。参考表允许同一 Sheet 存在多个表格区域。
    candidates = detect_structured_header_rows(rows, max_rows=80)
    if not candidates:
        return None

    # 配置表通常只有一个字段行，优先选择评分最高的程序字段行；参考表可能有多个区域，按行顺序拆区。
    if role != "reference" or len(candidates) == 1:
        h_idx = max(candidates, key=lambda item: item[1])[0] if role != "reference" else candidates[0][0]
        return build_structured_table_region(source_path, sheet_name, display_name, rows, h_idx, parse_row_number(data_start_row), len(rows), role=role)

    regions: List[StructuredTable] = []
    candidate_indices = [idx for idx, _score in candidates]
    for pos, h in enumerate(candidate_indices):
        next_h = candidate_indices[pos + 1] if pos + 1 < len(candidate_indices) else len(rows)
        # 如果相邻候选行很近，保留高分行，避免字段说明/类型行被误当新表。
        if pos > 0 and h - candidate_indices[pos - 1] <= 2:
            continue
        table = build_structured_table_region(source_path, sheet_name, f"{display_name}#区域{pos+1}", rows, h, None, next_h, role=role)
        if table and table.headers and table.rows:
            regions.append(table)
    if not regions:
        h_idx = candidates[0][0]
        return build_structured_table_region(source_path, sheet_name, display_name, rows, h_idx, parse_row_number(data_start_row), len(rows), role=role)
    if len(regions) == 1:
        return regions[0]
    return merge_structured_regions(source_path, sheet_name, display_name, regions)



def compose_reference_headers(rows: List[List[object]], h_idx: int, end_idx: int) -> List[str]:
    """Build more useful headers for non-standard reference sheets.

    Many planning sheets use a title row plus a field row, for example row 1 says
    “价格($)” while row 2 holds the section title and reward option headers.  This
    helper combines nearby header clues and infers common fields from sample data.
    """
    cur = rows[h_idx] if 0 <= h_idx < len(rows) else []
    prev = rows[h_idx - 1] if h_idx > 0 else []
    max_cols = max(len(cur), len(prev))
    out: List[str] = []
    sample_rows = rows[h_idx + 1:min(end_idx, h_idx + 8, len(rows))]
    for c in range(max_cols):
        v = clean_header(cur[c]) if c < len(cur) else ""
        pv = clean_header(prev[c]) if c < len(prev) else ""
        header = v or pv
        # 常见策划表第一列是分区标题，但数据行实际是礼包名/配置名。
        if c == 0 and (not header or re.search(r"系列|礼包|配置|活动", header)):
            first_values = [stringify(r[c]) for r in sample_rows if c < len(r) and stringify(r[c])]
            if any(re.search(r"礼包|活动|系统|主堡|获得|后续|档位|等级", x) for x in first_values):
                header = "名称"
        # 第二列常见为 pcid，表头为空但数据为 40989/40990 这类档位。
        if not header and looks_like_pcid_column(sample_rows, c):
            header = "pcid"
        # 第三列常见为价格，上一行/当前行未提供时根据小数价格推断。
        if (not header or normalize_field_name(header).startswith("__col")) and looks_like_price_column(sample_rows, c):
            header = "价格"
        out.append(header or f"__col_{c+1}")
    return out


def looks_like_pcid_column(rows: List[List[object]], col: int) -> bool:
    # pcid 通常在表格前几列；奖励区域中的道具 ID 也像 5 位数字，不能误推为 pcid。
    if col > 3:
        return False
    vals = [stringify(r[col]) for r in rows if col < len(r) and stringify(r[col])]
    if len(vals) < 2:
        return False
    hits = 0
    for v in vals[:8]:
        if re.fullmatch(r"\d{5,6}", normalize_basic_value(v)):
            hits += 1
    return hits >= max(2, len(vals[:8]) // 2)


def looks_like_price_column(rows: List[List[object]], col: int) -> bool:
    vals = [stringify(r[col]) for r in rows if col < len(r) and stringify(r[col])]
    if len(vals) < 2:
        return False
    hits = 0
    for v in vals[:8]:
        s = normalize_basic_value(v)
        if re.fullmatch(r"\d+(?:\.\d{1,2})", s):
            try:
                if 0 < float(s) < 1000:
                    hits += 1
            except Exception:
                pass
    return hits >= max(2, len(vals[:8]) // 2)


def find_section_title_above(rows: List[List[object]], h_idx: int) -> str:
    # 当前字段行第一格若是“xxx系列/xxx礼包”，通常就是分区名。
    if 0 <= h_idx < len(rows):
        vals = [stringify(v) for v in rows[h_idx] if stringify(v)]
        if vals and re.search(r"系列|礼包|活动|配置", vals[0]) and len(vals[0]) <= 40:
            return vals[0]
    for r in range(h_idx - 1, max(-1, h_idx - 6), -1):
        vals = [stringify(v) for v in rows[r] if stringify(v)]
        if len(vals) == 1 and len(vals[0]) <= 40 and re.search(r"系列|礼包|活动|配置", vals[0]):
            return vals[0]
    return ""


def add_standard_fields_to_rows(table: StructuredTable, role: str) -> StructuredTable:
    if table is None:
        return table
    headers = list(table.headers)
    standard_headers = ["id", "name", "pcid", "price", "reward", "trigger_level", "condition", "type", "section_name"]
    for h in standard_headers:
        if h not in headers:
            headers.append(h)
    section_name = find_section_title_above_cached(table)
    new_rows: List[StructuredRow] = []
    for row in table.rows:
        item = StructuredRow()
        item.__rownum__ = row.__rownum__
        for h in headers:
            item[h] = row.get(h, "")
        if not item.get("section_name"):
            item["section_name"] = row.get("__section", "") or section_name or row.get("__region", "")
        # 从原始字段别名归一化到标准字段。
        for h, v in row.items():
            if h.startswith("__") or not stringify(v):
                continue
            canon = canonical_field_name(h)
            if canon == "#id" and not item.get("id"):
                item["id"] = v
            elif canon == "name" and not item.get("name"):
                item["name"] = v
            elif canon == "pcid" and not item.get("pcid"):
                item["pcid"] = v
            elif canon == "price" and not item.get("price"):
                item["price"] = v
            elif canon == "reward" and not item.get("reward"):
                item["reward"] = v
            elif canon == "trigger_level" and not item.get("trigger_level"):
                item["trigger_level"] = v
            elif canon == "open_end_conditions" and not item.get("condition"):
                item["condition"] = v
                item.setdefault("open_end_conditions", v)
            elif canon == "type" and not item.get("type"):
                item["type"] = v
        # 策划参考表常见前几列无标准表头，用数据形态兜底。
        if role == "reference":
            if not stringify(item.get("name")):
                item["name"] = infer_reference_name(row, item.get("section_name", ""))
            if not stringify(item.get("pcid")):
                item["pcid"] = infer_first_value_by_predicate(row, lambda x: bool(re.fullmatch(r"\d{5,6}", normalize_basic_value(x))))
            if not stringify(item.get("price")):
                item["price"] = infer_first_value_by_predicate(row, lambda x: bool(re.fullmatch(r"\d+(?:\.\d{1,2})", normalize_basic_value(x))))
            if stringify(item.get("section_name")) and re.fullmatch(r"礼包\d+", stringify(item.get("name"))):
                item["name"] = f"{item.get('section_name')}{item.get('name')}"
            name_range_level = trigger_level_from_text(item.get("name") or row_name_value(item))
            if name_range_level:
                item["trigger_level"] = name_range_level
            elif not stringify(item.get("trigger_level")):
                item["trigger_level"] = infer_reference_trigger_level(row)
            if not stringify(item.get("reward")):
                item["reward"] = aggregate_reward_like_values(row)
        # 配置表保留 open_end_conditions1/2/3 的首个有效条件，但触发等级单独抽取，避免把 condition 类型码当等级。
        if role == "config":
            if not stringify(item.get("trigger_level")):
                item["trigger_level"] = infer_config_trigger_level(row)
            if not stringify(item.get("condition")):
                item["condition"] = infer_config_condition_summary(row)
        new_rows.append(item)
    # open_end_conditions 作为兼容字段也加入 headers，便于旧映射逻辑。
    if "open_end_conditions" not in headers:
        headers.append("open_end_conditions")
    # 去重保持顺序；自动合成标准字段若全为空则不作为候选主键/映射字段，避免 id 空列抢占主键。
    final_headers: List[str] = []
    synthetic = {"id", "name", "pcid", "price", "reward", "trigger_level", "condition", "type", "section_name", "open_end_conditions"}
    for h in headers:
        if not h or h in final_headers:
            continue
        if h in synthetic and not any(stringify(r.get(h, "")) for r in new_rows):
            continue
        final_headers.append(h)
    return StructuredTable(table.source_path, table.sheet_name, table.display_name, final_headers, table.field_types, new_rows, table.header_row, table.data_start_row)


def find_section_title_above_cached(table: StructuredTable) -> str:
    # 表级 section 通过 display_name 中的区域名或原始行上方标题获得；这里保持轻量，不访问源文件。
    text = str(table.display_name or "")
    m = re.search(r"([^#（]+(?:系列|礼包|活动|配置)[^#（]*)", text)
    return m.group(1).strip() if m else ""


def infer_reference_name(row: StructuredRow, section: str = "") -> str:
    vals = []
    for k, v in row.items():
        if k.startswith("__"):
            continue
        s = stringify(v)
        if s and not re.fullmatch(r"\d+(?:\.\d+)?", s):
            vals.append(s)
            break
    if vals:
        base = vals[0]
        if section and re.fullmatch(r"礼包\d+", base):
            return f"{section}{base}"
        return base
    return ""


def infer_first_value_by_predicate(row: StructuredRow, pred) -> str:
    # 跳过第一个名称字段，避免礼包1中的数字误判。
    for i, (k, v) in enumerate(row.items()):
        if k.startswith("__") or i == 0:
            continue
        s = stringify(v)
        if s and pred(s):
            return s
    return ""


def trigger_level_from_text(value) -> str:
    text = stringify(value).replace("～", "~").replace("-", "~")
    m = re.search(r"(\d+)\s*~\s*(\d+)\s*级", text)
    if m:
        return f"{int(m.group(1))}~{int(m.group(2))}"
    m = re.search(r"主堡\s*(\d+)\s*级", text)
    if m:
        return str(int(m.group(1)))
    return ""


def infer_reference_trigger_level(row: StructuredRow) -> str:
    name_level = trigger_level_from_text(row_name_value(row))
    if name_level:
        return name_level
    for k, v in row.items():
        if k.startswith("__"):
            continue
        if canonical_field_name(k) == "trigger_level" and stringify(v):
            return normalize_basic_value(v)
    return ""


def infer_config_trigger_level(row: StructuredRow) -> str:
    """Extract semantic trigger/unlock level from mall-style condition params.

    For mall.csv, open_end_conditions1 often stores a trigger type code such as 118/121,
    while open_end_conditions2 may store values like 1000|24 or range params like 1|16|24.
    We expose only the semantic level/range here so 解锁等级 is not compared to type codes.
    """
    keys = list(row.keys())
    candidates: List[Tuple[str, str]] = []
    for k in keys:
        nk = normalize_field_name(k)
        if nk in {"open_end_conditions2", "open_end_conditions3", "open_condition", "open_conditions", "buy_conditions1", "buy_conditions2"}:
            v = stringify(row.get(k, ""))
            if v:
                candidates.append((nk, v))
    # Prefer explicit 1000|level style level params.
    for _k, v in candidates:
        parts = [normalize_basic_value(x) for x in re.split(r"[|,;，；]", v) if normalize_basic_value(x)]
        if len(parts) >= 2 and parts[0] in {"1000", "level", "lv"}:
            return parts[-1]
    # Range-style params: 1|16|24 -> 16~24. This should not be equal to a single level unless exact.
    for _k, v in candidates:
        parts = [normalize_basic_value(x) for x in re.split(r"[|,;，；]", v) if normalize_basic_value(x)]
        nums = [p for p in parts if re.fullmatch(r"-?\d+", p)]
        if len(nums) >= 3:
            return f"{nums[-2]}~{nums[-1]}"
        if len(nums) == 2 and nums[0] in {"1000"}:
            return nums[-1]
    # A bare numeric param may be hero id / object id / condition code, so keep it empty and let the row enter 待确认 instead of false-confirming a level diff.
    return ""


def infer_config_condition_summary(row: StructuredRow) -> str:
    vals = []
    for k, v in row.items():
        if k.startswith("__"):
            continue
        nk = normalize_field_name(k)
        if nk in {"open_end_type", "open_end_conditions1", "open_end_conditions2", "open_end_conditions3", "condition", "open_condition"} and stringify(v):
            vals.append(f"{k}={v}")
    return ";".join(vals)


def aggregate_reward_like_values(row: StructuredRow) -> str:
    """Generate a reference reward string only from real item-id + quantity columns.

    It intentionally ignores value columns such as 原价值/总价值/价值/返利比.  Reference
    sheets commonly use repeated groups like 道具1, 空列(item_id), 数量, 价值 or
    选项1, 空列(item_id), 数量, 价值.  We keep the option item ids but do not use
    any value/price columns as reward atoms.
    """
    values = [(k, stringify(v)) for k, v in row.items() if not k.startswith("__meta")]
    items: List[str] = []
    used_id_positions: set[int] = set()
    def is_reward_name_header(h: str) -> bool:
        hn = normalize_field_name(h)
        return bool(re.search(r"^(道具|奖励|选项)\d*", hn)) or hn in {"item", "reward"}
    def is_value_header(h: str) -> bool:
        hn = normalize_field_name(h)
        return bool(re.search(r"价值|返利|折扣|性价比|总价值|原价值", hn))
    def is_qty_header(h: str) -> bool:
        hn = normalize_field_name(h)
        return canonical_field_name(h) == "count" or bool(re.search(r"数量|个数|count|num|quantity", hn))
    for i, (h, name_value) in enumerate(values):
        if not is_reward_name_header(h) or not name_value:
            continue
        if is_value_header(h):
            continue
        id_pos = None
        for j in range(i + 1, min(i + 4, len(values))):
            hj, vj = values[j]
            if is_value_header(hj):
                continue
            if re.fullmatch(r"\d{4,7}", normalize_basic_value(vj)):
                id_pos = j
                break
        if id_pos is None or id_pos in used_id_positions:
            continue
        qty = ""
        for j in range(id_pos + 1, min(id_pos + 5, len(values))):
            hj, vj = values[j]
            if is_value_header(hj):
                continue
            if is_qty_header(hj) and re.fullmatch(r"\d+(?:\.0)?", normalize_basic_value(vj)):
                qty = normalize_basic_value(vj)
                break
        if not qty:
            continue
        item_id = normalize_basic_value(values[id_pos][1])
        items.append(f"{item_id}*{qty}")
        used_id_positions.add(id_pos)
    return "|".join(items)

def build_structured_table_region(source_path: str, sheet_name: str, display_name: str, rows: List[List[object]], h_idx: int, d_idx: Optional[int], end_idx: int, role: str = "config") -> Optional[StructuredTable]:
    if h_idx is None or h_idx < 0 or h_idx >= len(rows):
        return None
    if role == "reference":
        headers = uniquify_headers(compose_reference_headers(rows, h_idx, end_idx))
    else:
        headers = uniquify_headers([clean_header(v) for v in rows[h_idx]])
    if not any(clean_header(h) for h in headers):
        return None
    type_row = rows[h_idx - 1] if h_idx > 0 else []
    field_types = {headers[i]: stringify(type_row[i]) if i < len(type_row) else "" for i in range(len(headers)) if headers[i]}
    if d_idx is None:
        d_idx = h_idx + 1
    d_idx = max(h_idx + 1, d_idx)
    data: List[StructuredRow] = []
    section_name = find_section_title_above(rows, h_idx) if role == "reference" else ""
    for excel_row_num, row in enumerate(rows[d_idx:end_idx], start=d_idx + 1):
        if not any(stringify(v) for v in row):
            continue
        # 自动区域解析时，参考表跳过像下一段小标题/说明的行；配置表数据值本身可能像字段，不做此拦截。
        if role == "reference" and looks_like_section_or_note_row(row, headers):
            continue
        item = StructuredRow()
        item.__rownum__ = excel_row_num
        for i, h in enumerate(headers):
            if not h:
                continue
            item[h] = stringify(row[i]) if i < len(row) else ""
        if section_name:
            item["__section"] = section_name
        if row_has_structured_signal(item) or role == "config":
            data.append(item)
    cleaned_headers = [h for h in headers if h]
    table = StructuredTable(source_path, sheet_name, display_name, cleaned_headers, field_types, data, h_idx, d_idx)
    return add_standard_fields_to_rows(table, role)


def merge_structured_regions(source_path: str, sheet_name: str, display_name: str, regions: List[StructuredTable]) -> StructuredTable:
    """Merge multiple table blocks in one reference sheet.

    If two regions use near-synonym headers, such as “商品ID” and “礼包ID”, values are
    normalized into the first seen header so downstream key matching/mapping can work as one table.
    """
    headers: List[str] = []
    field_types: Dict[str, str] = {}
    canonical_primary: Dict[str, str] = {}
    for region in regions:
        for h in region.headers:
            canon = canonical_field_name(h)
            if canon and canon in FIELD_ALIAS_GROUPS:
                if canon not in canonical_primary:
                    canonical_primary[canon] = h
                    headers.append(h)
                    if h in region.field_types:
                        field_types[h] = region.field_types[h]
                continue
            if h not in headers:
                headers.append(h)
                if h in region.field_types:
                    field_types[h] = region.field_types[h]

    rows: List[StructuredRow] = []
    for region in regions:
        for row in region.rows:
            merged = StructuredRow()
            merged.__rownum__ = row.__rownum__
            for h in headers:
                val = row.get(h, "")
                if not stringify(val):
                    canon = canonical_field_name(h)
                    if canon:
                        for source_h in row.keys():
                            if source_h.startswith("__"):
                                continue
                            if canonical_field_name(source_h) == canon and stringify(row.get(source_h, "")):
                                val = row.get(source_h, "")
                                break
                merged[h] = val
            merged["__region"] = region.display_name
            rows.append(merged)
    table = StructuredTable(source_path, sheet_name, f"{display_name}（{len(regions)} 个区域）", headers, field_types, rows, regions[0].header_row, regions[0].data_start_row)
    return add_standard_fields_to_rows(table, "reference")


def detect_structured_header_rows(rows: List[List[object]], max_rows: int = 80) -> List[Tuple[int, int]]:
    scored: List[Tuple[int, int]] = []
    limit = min(len(rows), max_rows)
    for i in range(limit):
        score = score_structured_header_row(rows[i])
        if score >= 12:
            scored.append((i, score))
    # 同一小范围内可能有字段类型行/重复行，只保留更像字段名的高分行。
    scored.sort(key=lambda item: (-item[1], item[0]))
    chosen: List[Tuple[int, int]] = []
    for idx, score in scored:
        if any(abs(idx - old_idx) <= 1 for old_idx, _ in chosen):
            continue
        chosen.append((idx, score))
    chosen.sort(key=lambda item: item[0])
    return chosen


def detect_structured_header_row(rows: List[List[object]]) -> Optional[int]:
    candidates = detect_structured_header_rows(rows, max_rows=80)
    return candidates[0][0] if candidates else None


def score_structured_header_row(row: List[object]) -> int:
    cleaned = [clean_header(v) for v in row]
    non_empty_headers = [v for v in cleaned if v]
    if len(non_empty_headers) < 2:
        return 0
    score = 0
    canonical_hits = set()
    for h in non_empty_headers:
        n = normalize_field_name(h)
        canon = canonical_field_name(h)
        if canon != n:
            canonical_hits.add(canon)
        elif n in {"#id", "id", "type", "reward", "item", "price", "pcid", "name", "$name", "open_end_conditions"}:
            canonical_hits.add(n)
    # 关键字段组合越完整，越可能是真正字段名行。
    if "#id" in canonical_hits:
        score += 24
    if "name" in canonical_hits:
        score += 8
    if "price" in canonical_hits:
        score += 8
    if "pcid" in canonical_hits:
        score += 8
    if "reward" in canonical_hits:
        score += 8
    if "open_end_conditions" in canonical_hits:
        score += 6
    if "trigger_level" in canonical_hits:
        score += 6
    if "type" in canonical_hits:
        score += 5
    # 字段数量只能作为辅助分，不允许“数据行字段多”直接变成表头。
    if canonical_hits:
        score += min(len(non_empty_headers), 8)
    # 配置表程序字段行通常包含短英文/下划线/#/$/^字段，额外加分。
    program_like = sum(1 for h in non_empty_headers if re.fullmatch(r"[#\$\^]?[A-Za-z_][A-Za-z0-9_]*", h))
    score += min(program_like * 3, 18)
    # 纯说明行/标题行往往很长但字段少，降分。
    joined = "".join(non_empty_headers)
    avg_len = sum(len(x) for x in non_empty_headers) / max(1, len(non_empty_headers))
    if avg_len > 18 and program_like <= 2:
        score -= 18
    if len(non_empty_headers) <= 3 and len(joined) > 40:
        score -= 12
    if any(word in joined for word in ["说明", "注意", "需求", "规则", "备注"]):
        score -= 4
    return max(0, score)


def looks_like_section_or_note_row(row: List[object], headers: List[str]) -> bool:
    values = [stringify(v) for v in row if stringify(v)]
    if not values:
        return True
    if len(values) == 1 and len(values[0]) <= 30 and re.search(r"自选礼包|弹窗礼包|礼包|说明|备注|需求|规则", values[0]):
        return True
    return False


def row_has_structured_signal(row: StructuredRow) -> bool:
    # 参考表区域内跳过纯说明文字；有效配置行至少应有名称，并且有 pcid/价格/等级/reward/ID 之一。
    name = ""
    has_anchor = False
    for k, v in row.items():
        if k.startswith("__") or not stringify(v):
            continue
        canon = canonical_field_name(k)
        val = normalize_basic_value(v)
        if canon == "name":
            name = stringify(v)
        elif canon in {"#id", "pcid"} and re.fullmatch(r"\d{4,9}", val):
            has_anchor = True
        elif canon == "price" and re.fullmatch(r"\d+(?:\.\d{1,2})?", val):
            has_anchor = True
        elif canon == "trigger_level" and re.fullmatch(r"\d+(?:\.0)?", val):
            has_anchor = True
        elif canon == "reward" and stringify(v):
            has_anchor = True
    if not name and has_anchor:
        return True
    if name and has_anchor:
        return True
    return False

def structured_table_preview_lines(table: Optional[StructuredTable], title: str = "参考表") -> List[str]:
    if table is None:
        return [f"{title}识别结果：未识别到有效表格。"]
    lines = [
        f"{title}识别结果：Sheet={table.sheet_name}；字段名行=第 {table.header_row + 1} 行；数据起始行=第 {table.data_start_row + 1} 行；有效数据={len(table.rows)} 条",
        f"{title}字段列表：{', '.join(table.headers[:30]) or '无'}" + (" ..." if len(table.headers) > 30 else ""),
    ]
    if getattr(table, "notes", None):
        lines.extend([f"{title}提示：{note}" for note in table.notes])
    if table.rows:
        section_counts: Dict[str, int] = {}
        for row in table.rows:
            sec = stringify(row.get("section_name") or row.get("__region") or row.get("__source_file") or "未分区")
            section_counts[sec] = section_counts.get(sec, 0) + 1
        if section_counts:
            lines.append(f"{title}分区预览：" + "；".join(f"{k}={v}条" for k, v in list(section_counts.items())[:12]))
        lines.append(f"{title}前几条数据：")
        for row in table.rows[:3]:
            lines.append(f"- 第 {row.__rownum__} 行：{row_to_preview(row, limit=10)}")
    return lines


def build_structured_table_diagnostics(path: str, requested_sheet: str = "", role: str = "reference") -> str:
    p = Path(path)
    lines = [f"当前读取文件：{p.name}"]
    try:
        suffix = p.suffix.lower()
        if suffix in {".xlsx", ".xlsm", ".xls"}:
            from openpyxl import load_workbook
            wb = load_workbook(p, data_only=True, read_only=True)
            try:
                lines.append(f"扫描 Sheet：{', '.join(wb.sheetnames)}")
                if requested_sheet and requested_sheet.strip() and requested_sheet.strip() not in wb.sheetnames:
                    lines.append(f"指定 Sheet 不存在：{requested_sheet}")
                for ws in wb.worksheets:
                    if requested_sheet and requested_sheet.strip() and ws.title.strip().lower() != requested_sheet.strip().lower():
                        continue
                    max_row = min(ws.max_row or 0, 80)
                    sample = [[cell for cell in row] for row in ws.iter_rows(min_row=1, max_row=max_row, values_only=True)]
                    candidates = detect_structured_header_rows(sample, max_rows=80)
                    if candidates:
                        desc = []
                        for idx, score in candidates[:5]:
                            headers = [clean_header(v) for v in sample[idx] if clean_header(v)]
                            desc.append(f"第 {idx + 1} 行(评分 {score})：{', '.join(headers[:12])}")
                        lines.append(f"Sheet【{ws.title}】候选字段名行：" + "；".join(desc))
                    else:
                        lines.append(f"Sheet【{ws.title}】前 {max_row} 行未识别到有效字段名行。")
            finally:
                wb.close()
        else:
            raw = read_text_with_fallback(p)
            rows = list(csv.reader(raw.splitlines()))
            candidates = detect_structured_header_rows(rows, max_rows=40)
            lines.append(f"扫描 CSV 前 {min(len(rows), 40)} 行")
            if candidates:
                lines.append("候选字段名行：" + "；".join(f"第 {idx + 1} 行(评分 {score})" for idx, score in candidates[:5]))
            else:
                lines.append("未识别到有效字段名行。")
    except Exception as exc:
        lines.append(f"诊断失败：{exc}")
    lines.extend([
        "关键字段建议至少包含：商品ID/礼包ID/配置ID/id/#id、名称/礼包名称、价格/售价、pcid/支付ID、奖励/reward/item、触发等级/解锁等级/open_end_conditions。",
        "请在界面手动指定参考表 Sheet、字段名行、数据起始行后再刷新预览。",
    ])
    return "\n".join(lines)

def parse_row_number(value: str) -> Optional[int]:
    text = str(value or "").strip().lower()
    if not text or text == "auto" or text == "自动":
        return None
    try:
        num = int(float(text))
        return max(0, num - 1)
    except Exception:
        return None


def uniquify_headers(headers: List[str]) -> List[str]:
    seen: Dict[str, int] = {}
    out: List[str] = []
    for i, h in enumerate(headers):
        h = h or f"__col_{i+1}"
        if h in seen:
            seen[h] += 1
            out.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 1
            out.append(h)
    return out


def json_to_rows(data) -> List[List[object]]:
    if isinstance(data, dict):
        values = list(data.values())
    elif isinstance(data, list):
        values = data
    else:
        return []
    dicts = [v for v in values if isinstance(v, dict)]
    headers = []
    for item in dicts:
        for k in item.keys():
            if k not in headers:
                headers.append(str(k))
    rows = [headers]
    for item in dicts:
        rows.append([item.get(h, "") for h in headers])
    return rows


def resolve_key_fields(text: str, ref_headers: List[str], cfg_headers: List[str]) -> List[str]:
    if text and str(text).strip().lower() not in {"auto", "自动"}:
        candidates = [v.strip() for v in re.split(r"[+,，,;；]\s*", text) if v.strip()]
        return [find_actual_header(c, ref_headers, cfg_headers) or c for c in candidates]
    for cand in KEY_CANDIDATES:
        ref_actual = first_matching_field(ref_headers, cand)
        cfg_actual = first_matching_field(cfg_headers, cand)
        if ref_actual and cfg_actual:
            # 返回参考表实际字段，build_row_key 会用别名在配置表中找到对应字段。
            return [ref_actual]
    # 兜底：如果两边存在同一语义字段，也可作为主键候选。
    ref_canons = {canonical_field_name(h): h for h in ref_headers if canonical_field_name(h)}
    cfg_canons = {canonical_field_name(h): h for h in cfg_headers if canonical_field_name(h)}
    for canon in ["#id", "name", "pcid"]:
        if canon in ref_canons and canon in cfg_canons:
            return [ref_canons[canon]]
    return []


def find_actual_header(field: str, ref_headers: List[str], cfg_headers: List[str]) -> str:
    n = normalize_field_name(field)
    for h in ref_headers + cfg_headers:
        if normalize_field_name(h) == n:
            return h
    canon = canonical_field_name(field)
    for h in ref_headers + cfg_headers:
        if canonical_field_name(h) == canon:
            return h
    return ""


def build_field_mapping(ref_headers: List[str], cfg_headers: List[str], mapping_text: str, ignore_text: str, key_fields: List[str]) -> List[Tuple[str, str]]:
    ignored = build_ignore_set(ignore_text)
    ref_norm = {normalize_field_name(h): h for h in ref_headers}
    cfg_norm = {normalize_field_name(h): h for h in cfg_headers}
    ref_canon: Dict[str, str] = {}
    cfg_canon: Dict[str, str] = {}
    for h in ref_headers:
        c = canonical_field_name(h)
        if c and c not in ref_canon:
            ref_canon[c] = h
    for h in cfg_headers:
        c = canonical_field_name(h)
        if c and c not in cfg_canon:
            cfg_canon[c] = h
    mapping: List[Tuple[str, str]] = []
    used_ref = set()
    used_cfg = set()
    used_canon = set()
    for left, right in parse_mapping_text(mapping_text):
        lf = ref_norm.get(normalize_field_name(left)) or first_matching_field(ref_headers, left) or left
        rf = cfg_norm.get(normalize_field_name(right)) or first_matching_field(cfg_headers, right) or right
        if normalize_field_name(lf) in ignored or normalize_field_name(rf) in ignored:
            continue
        if lf in ref_headers and rf in cfg_headers:
            canon = canonical_field_name(lf) or normalize_field_name(lf)
            mapping.append((lf, rf)); used_ref.add(lf); used_cfg.add(rf); used_canon.add(canon)
    key_canons = {canonical_field_name(k) for k in key_fields}
    key_norms = {normalize_field_name(k) for k in key_fields}
    # 先按完全同名字段自动匹配。
    for n, rf in ref_norm.items():
        cf = cfg_norm.get(n)
        canon = canonical_field_name(rf) or n
        if cf and n not in ignored and n not in key_norms and canon not in key_canons and rf not in used_ref and cf not in used_cfg and canon not in used_canon:
            mapping.append((rf, cf)); used_ref.add(rf); used_cfg.add(cf); used_canon.add(canon)
    # 再按语义别名匹配中文参考表与程序配置表，例如 商品ID=>#id、价格=>price、触发等级=>open_end_conditions。
    for c, rf in ref_canon.items():
        cf = cfg_canon.get(c)
        if not cf or c in key_canons:
            continue
        if normalize_field_name(rf) in ignored or normalize_field_name(cf) in ignored:
            continue
        if rf in used_ref or cf in used_cfg or c in used_canon:
            continue
        mapping.append((rf, cf)); used_ref.add(rf); used_cfg.add(cf); used_canon.add(c)
    return mapping


def parse_mapping_text(text: str) -> List[Tuple[str, str]]:
    pairs = []
    for line in str(text or "").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if "=>" in line:
            left, right = line.split("=>", 1)
        elif ":" in line:
            left, right = line.split(":", 1)
        elif "=" in line:
            left, right = line.split("=", 1)
        else:
            continue
        left = left.strip(); right = right.strip()
        if left and right:
            pairs.append((left, right))
    return pairs


def build_ignore_set(text: str) -> set[str]:
    values = set(DEFAULT_IGNORE_FIELD_NAMES)
    values.update(parse_user_list(text))
    return {normalize_field_name(v) for v in values if v}


def filter_table_rows(table: StructuredTable, expression: str, source_type: str) -> Tuple[List[StructuredRow], FilterPreview]:
    expr = str(expression or "").strip()
    before = len(table.rows)
    if not expr:
        preview = FilterPreview(
            source_type=source_type,
            file_name=Path(table.source_path).name or table.display_name,
            sheet_name=table.sheet_name,
            header_row=table.header_row,
            data_start_row=table.data_start_row,
            expression="",
            field_check="无过滤字段",
            before_count=before,
            after_count=before,
            status="无过滤",
            message="未填写过滤条件，按当前数据全集参与比对",
        )
        return list(table.rows), preview

    if not table.headers:
        raise FilterValidationError(f"{source_type}过滤条件未生效：未识别到字段名行。请检查字段名行是否正确。")
    if not table.rows:
        raise FilterValidationError(f"{source_type}过滤条件未生效：数据起始行后没有有效数据。请检查数据起始行。")

    conditions = parse_filter_conditions(expr)
    missing: List[str] = []
    actual_fields: List[str] = []
    for cond in conditions:
        actual = first_matching_field(table.headers, cond.field)
        if not actual:
            missing.append(cond.field)
        else:
            actual_fields.append(actual)
    if missing:
        raise FilterValidationError(
            f"{source_type}过滤条件未生效：未找到字段 {', '.join(missing)}。"
            f"请检查字段名行是否正确，或修改过滤字段名。当前识别字段：{', '.join(table.headers[:20])}"
        )

    filtered = [row for row in table.rows if all(match_filter_condition(row, cond) for cond in conditions)]
    status = "已生效"
    message = ""
    if len(filtered) == before:
        status = "已生效（数量未变化）"
        message = "过滤前后数量一致，请确认该条件是否符合预期"
    preview = FilterPreview(
        source_type=source_type,
        file_name=Path(table.source_path).name or table.display_name,
        sheet_name=table.sheet_name,
        header_row=table.header_row,
        data_start_row=table.data_start_row,
        expression=expr,
        field_check="、".join(f"{field} 存在" for field in actual_fields) or "无",
        before_count=before,
        after_count=len(filtered),
        status=status,
        message=message,
    )
    return filtered, preview


def build_value_compare_rule_preview(config_path: str, reference_path: str, options: ValueCompareOptions) -> str:
    """Build execution preview and validate filters before starting UI worker."""
    options = prepare_value_compare_options(options, config_path, reference_path)
    if options.template_name == "mall 礼包类":
        config_records, config_fields = parse_config_records(config_path)
        reference_records, reference_sheets, reference_notes = parse_reference_records(reference_path)
        if not config_records:
            raise FilterValidationError("配置表未识别到 mall 商品记录，请确认存在 #goods_id、pcid、reward 等字段。")
        if not reference_records:
            raise FilterValidationError("参考表未识别到商业化礼包/档位，请确认包含礼包名称、商品 ID、价格或奖励字段。")
        lines = [
            "本次将执行商业化礼包专项比对：",
            "",
            f"参考表：{Path(reference_path).name}",
            f"配置表：{Path(config_path).name}",
            f"已识别参考 Sheet：{'、'.join(reference_sheets[:8]) or '未识别'}",
            f"参考礼包/档位：{len(reference_records)} 条",
            f"配置商品记录：{len(config_records)} 条",
            f"配置字段：{'、'.join(config_fields[:12])}",
            "",
            "匹配规则：优先商品 ID，其次完整礼包名、pcid、价格和奖励内容；不会使用单一 name 主键折叠不同价格档位。",
            "范围规则：不预设 type in 2000,13，避免排除 type=3 月周卡或引入无关业务。",
            "输出规则：仅输出已匹配礼包的价格、pcid、持续时间、限购次数和奖励差异。",
        ]
        if reference_notes:
            lines.extend(["", "解析提示："] + [f"- {note}" for note in reference_notes[:8]])
        return "\n".join(lines)

    ref_table = load_structured_table(reference_path, options.reference_sheet, options.reference_header_row, options.reference_data_start_row, role="reference")
    if ref_table is None or not ref_table.headers:
        raise FilterValidationError(
            "参考表未识别到有效字段，无法进入字段映射和结构化比对。\n"
            + build_structured_table_diagnostics(reference_path, options.reference_sheet, role="reference")
        )
    cfg_tables = load_config_tables(config_path, options.config_sheet, options.config_header_row, options.config_data_start_row)
    if not cfg_tables:
        raise FilterValidationError(
            "配置表未识别到有效数据，请检查配置表文件、Sheet、字段名行和数据起始行。\n"
            + build_structured_table_diagnostics(config_path, options.config_sheet, role="config")
        )
    ref_rows, ref_preview = filter_table_rows(ref_table, options.reference_filter, "参考表")
    cfg_previews: List[FilterPreview] = []
    filtered_cfg_tables: List[StructuredTable] = []
    for table in cfg_tables:
        rows, preview = filter_table_rows(table, options.config_filter, "配置表")
        cfg_previews.append(preview)
        filtered_cfg_tables.append(replace(table, rows=rows))
    cfg_table = merge_structured_tables(filtered_cfg_tables)
    key_fields = resolve_key_fields(options.key_fields, ref_table.headers, cfg_table.headers)
    mapping = build_field_mapping(ref_table.headers, cfg_table.headers, options.field_mappings, options.ignore_fields, key_fields) if key_fields else []
    filter_risks = build_filter_risk_messages(ref_preview, cfg_previews)
    lines = [
        "本次将执行结构化配置比对：",
        "",
        f"模板：{options.template_name or '通用配置表'}",
        f"参考表：{Path(reference_path).name} / Sheet：{ref_table.sheet_name}",
        f"配置表：{Path(config_path).name} / {cfg_table.display_name}",
        f"主键字段：{'+'.join(key_fields) if key_fields else '未识别'}",
        f"字段映射数量：{len(mapping)}",
        "",
        "参考表字段识别预览：",
        *structured_table_preview_lines(ref_table, "参考表"),
        "",
        "配置表字段识别预览：",
        *structured_table_preview_lines(cfg_table, "配置表"),
        "",
        "字段映射预览：",
        *(build_field_mapping_preview_lines(ref_table.headers, cfg_table.headers, mapping, key_fields)),
        "",
        "过滤预览：",
        ref_preview.to_summary_text(),
    ]
    for preview in cfg_previews:
        lines.append(preview.to_summary_text())
    if filter_risks:
        lines.extend(["", "过滤风险提示："] + [f"- {msg}" for msg in filter_risks])
    lines.extend([
        "",
        "说明：本模式按主键关联，不按行号比对；漏配、多配、字段差异都基于过滤后的数据集计算。",
    ])
    if not key_fields:
        lines.append("\n警告：未识别到主键字段，请手动填写主键字段后再执行。")
    if not mapping:
        lines.append("\n警告：字段映射为空，请确认两表有同名字段，或手动填写字段映射。")
    return "\n".join(lines)


def build_field_mapping_preview_lines(ref_headers: List[str], cfg_headers: List[str], mapping: List[Tuple[str, str]], key_fields: List[str]) -> List[str]:
    lines: List[str] = []
    if not mapping:
        return ["- 未建立字段映射；请手动填写映射，例如 价格=>price。"]
    key_canons = {canonical_field_name(k) or normalize_field_name(k) for k in key_fields}
    for ref_field, cfg_field in mapping[:30]:
        canon = canonical_field_name(ref_field) or normalize_field_name(ref_field)
        conf = "高" if normalize_field_name(ref_field) == normalize_field_name(cfg_field) or canonical_field_name(ref_field) == canonical_field_name(cfg_field) else "中"
        participate = "否（主键字段）" if canon in key_canons else "是"
        need = "否" if conf == "高" else "建议确认"
        lines.append(f"- {ref_field} => {cfg_field}；置信度={conf}；参与比对={participate}；需要确认={need}")
    if len(mapping) > 30:
        lines.append(f"- 其余 {len(mapping)-30} 个映射略。")
    return lines


def apply_filters(rows: List[StructuredRow], expression: str) -> List[StructuredRow]:
    # 旧兼容函数：无表头上下文时只做宽松过滤。新结构化比对请使用 filter_table_rows。
    try:
        filters = parse_filter_conditions(expression)
    except FilterValidationError:
        return rows
    if not filters:
        return rows
    return [row for row in rows if all(match_filter_condition(row, cond) for cond in filters)]


def parse_filter_conditions(expression: str) -> List[FilterCondition]:
    text = str(expression or "").strip()
    if not text:
        return []
    if re.search(r"\bor\b|\bOR\b|\s+或\s+", text):
        raise FilterValidationError("过滤条件暂不支持 OR，请使用分号 ; 连接 AND 条件。")
    parts = [p.strip() for p in re.split(r"[;；\n]+", text) if p.strip()]
    result: List[FilterCondition] = []
    for part in parts:
        m = re.match(r"^(.+?)\s+(not\s+in|in)\s+(.+)$", part, re.I)
        if m:
            field, op, values = m.groups()
            vals = [v.strip() for v in re.split(r"[,，|]", values) if v.strip()]
            if not vals:
                raise FilterValidationError(f"过滤条件语法错误：{part} 缺少可匹配的值。")
            result.append(FilterCondition(field.strip(), op.lower().replace("  ", " "), vals))
            continue
        m = re.match(r"^(.+?)\s+(not\s+contains|contains|不包含|包含)\s+(.+)$", part, re.I)
        if m:
            field, op, value = m.groups()
            op_norm = op.lower().replace("  ", " ")
            if op_norm == "包含":
                op_norm = "contains"
            elif op_norm == "不包含":
                op_norm = "not contains"
            result.append(FilterCondition(field.strip(), op_norm, [value.strip()]))
            continue
        m = re.match(r"^(.+?)\s*(==|=|!=)\s*(.+)$", part)
        if m:
            field, op, value = m.groups()
            if not field.strip():
                raise FilterValidationError(f"过滤条件语法错误：{part} 缺少字段名。")
            result.append(FilterCondition(field.strip(), op, [value.strip()]))
            continue
        if part.endswith("非空"):
            field = part[:-2].strip()
            if not field:
                raise FilterValidationError(f"过滤条件语法错误：{part} 缺少字段名。")
            result.append(FilterCondition(field, "not_empty", [])); continue
        if part.endswith("为空"):
            field = part[:-2].strip()
            if not field:
                raise FilterValidationError(f"过滤条件语法错误：{part} 缺少字段名。")
            result.append(FilterCondition(field, "empty", [])); continue
        raise FilterValidationError(f"过滤条件语法错误：{part}。支持：字段=值、字段!=值、字段 in a,b、字段 contains 关键词、字段为空、字段非空；多个条件用 ; 连接。")
    return result


def parse_filter_expression(expression: str) -> List[Tuple[str, str, List[str]]]:
    # 兼容旧调用，优先使用 parse_filter_conditions。
    return [(c.field, c.op, c.values) for c in parse_filter_conditions(expression)]


def match_filter_condition(row: StructuredRow, cond: FilterCondition) -> bool:
    actual_field = first_matching_field(row.keys(), cond.field)
    value = row.get(actual_field, "") if actual_field else ""
    norm = normalize_filter_value(value)
    norms = [normalize_filter_value(v) for v in cond.values]
    if cond.op in {"=", "=="}:
        return norm == (norms[0] if norms else "")
    if cond.op == "!=":
        return norm != (norms[0] if norms else "")
    if cond.op == "in":
        return norm in norms
    if cond.op == "not in":
        return norm not in norms
    if cond.op == "contains":
        return bool(norms) and norms[0] in norm
    if cond.op == "not contains":
        return (not norms) or norms[0] not in norm
    if cond.op == "not_empty":
        return bool(norm)
    if cond.op == "empty":
        return not norm
    return True


def match_filter(row: StructuredRow, flt: Tuple[str, str, List[str]]) -> bool:
    field, op, values = flt
    return match_filter_condition(row, FilterCondition(field, op, values))


def normalize_filter_value(value) -> str:
    s = stringify(value).replace("\u3000", " ").strip()
    s = re.sub(r"\s+", " ", s)
    if s.lower() in {"none", "nan", "null"}:
        return ""
    if re.fullmatch(r"[+-]?0\d+", s):
        # ID / code 类字段常见前导 0，默认不把 001 与 1 视为一致。
        return s
    if re.fullmatch(r"[+-]?0\d+\.\d+", s):
        return s.rstrip("0").rstrip(".")
    if re.fullmatch(r"[+-]?\d{1,3}(?:,\d{3})+(?:\.\d+)?", s):
        s = s.replace(",", "")
    if NUMBER_RE.match(s):
        try:
            num = float(s)
            if num.is_integer():
                return str(int(num))
            return ("%f" % num).rstrip("0").rstrip(".")
        except Exception:
            pass
    return s


def first_matching_field(fields: Iterable[str], target: str) -> str:
    n = normalize_field_name(target)
    fields_list = list(fields)
    for f in fields_list:
        if normalize_field_name(f) == n:
            return f
    canon = canonical_field_name(target)
    if canon:
        for f in fields_list:
            if canonical_field_name(f) == canon:
                return f
    return ""


def index_rows_by_key(rows: List[StructuredRow], key_fields: List[str]) -> Tuple[Dict[str, StructuredRow], Dict[str, List[StructuredRow]]]:
    indexed: Dict[str, StructuredRow] = {}
    dups: Dict[str, List[StructuredRow]] = {}
    for row in rows:
        key = build_row_key(row, key_fields)
        if not key:
            continue
        if key in indexed:
            dups.setdefault(key, [indexed[key]]).append(row)
        else:
            indexed[key] = row
    for key in dups:
        indexed.pop(key, None)
    return indexed, dups


def build_row_key(row: StructuredRow, key_fields: List[str]) -> str:
    parts = []
    for field in key_fields:
        actual = first_matching_field(row.keys(), field)
        value = row.get(actual, "") if actual else ""
        # 多区域参考表合并前后可能同时存在“商品ID / 礼包ID / 配置ID”等近义列；
        # 如果首个匹配列为空，则继续找同语义且有值的列，避免后续区域匹配不到配置表 #id。
        if not stringify(value):
            canon = canonical_field_name(field)
            for k in row.keys():
                if k.startswith("__"):
                    continue
                if canonical_field_name(k) == canon and stringify(row.get(k, "")):
                    actual = k
                    value = row.get(k, "")
                    break
        if not actual or not stringify(value):
            return ""
        if canonical_field_name(field) == "name":
            parts.append(normalize_name_key(value))
        else:
            parts.append(normalize_filter_value(value))
    return "+".join(parts)


def normalize_name_key(value) -> str:
    s = normalize_basic_value(value)
    s = re.sub(r"\s+", "", s)
    # 常见分区词不影响名称匹配；“礼包系列礼包1”和“礼包1”需可归一到同一档位。
    s = re.sub(r"系列", "", s)
    s = re.sub(r"礼包", "", s)
    s = re.sub(r"弹窗", "", s)
    return s


def normalize_compare_cell(value, field: str, zero_equal_fields: set[str], percent_fields: set[str], bool_fields: set[str], case_insensitive_fields: set[str]) -> str:
    nfield = normalize_field_name(field)
    raw = stringify(value).replace("\u3000", " ").strip()
    raw = re.sub(r"\s+", " ", raw)
    if raw.lower() in {"none", "nan", "null"}:
        raw = ""
    if nfield in zero_equal_fields and raw in {"0", "0.0"}:
        raw = ""
    if nfield in bool_fields:
        low = raw.lower()
        if low in {"true", "1", "yes", "y", "是"}:
            return "true"
        if low in {"false", "0", "no", "n", "否"}:
            return "false"
    if canonical_field_name(field) == "reward":
        return normalize_reward_compare_value(raw)
    if canonical_field_name(field) == "trigger_level":
        return normalize_trigger_level_value(raw)
    if canonical_field_name(field) == "open_end_conditions" or nfield == "condition":
        return normalize_trigger_condition_value(raw)
    if canonical_field_name(field) == "price" or nfield == "price":
        return normalize_price_value(raw)
    if nfield in percent_fields:
        return normalize_percent_value(raw)
    # 复合字段只规范空格和数字 token，不改分隔符结构。
    if any(sep in raw for sep in ["|", "*", ";", ","]):
        return normalize_composite_value(raw, nfield in zero_equal_fields, nfield in percent_fields)
    norm = normalize_basic_value(raw)
    if nfield in case_insensitive_fields:
        norm = norm.lower()
    return norm


def normalize_reward_compare_value(value: str) -> str:
    raw = stringify(value).strip()
    if not raw or raw in {"0", "0.0"}:
        return ""
    merged: Dict[str, int] = {}
    unknown: List[str] = []
    for token in re.split(r"[|；;]+", raw):
        token = token.strip()
        if not token:
            continue
        parts = [normalize_basic_value(p) for p in token.split("*") if normalize_basic_value(p)]
        item_id = ""
        qty = ""
        if len(parts) >= 3 and parts[0] == "12":
            # mall reward type 12 is usually a display/extra presentation binding and should not create reward quantity diffs.
            continue
        if len(parts) >= 3 and re.fullmatch(r"\d+", parts[-2]) and re.fullmatch(r"-?\d+", parts[-1]):
            item_id, qty = parts[-2], parts[-1]
        elif len(parts) >= 2 and re.fullmatch(r"\d+", parts[0]) and re.fullmatch(r"-?\d+", parts[1]):
            item_id, qty = parts[0], parts[1]
        if item_id and qty:
            if len(item_id) == 1 and item_id in {"1", "2", "12"} and len(parts) >= 3:
                item_id = parts[-2]
            try:
                merged[item_id] = merged.get(item_id, 0) + int(float(qty))
            except Exception:
                unknown.append(token)
        else:
            unknown.append(normalize_composite_value(token, False, False))
    chunks = [f"{k}*{merged[k]}" for k in sorted(merged, key=natural_key)]
    chunks.extend(sorted([x for x in unknown if x]))
    return "|".join(chunks)


def normalize_trigger_level_value(value: str) -> str:
    raw = stringify(value).strip()
    if not raw:
        return ""
    raw = raw.replace("～", "~").replace("-", "~")
    nums = [normalize_basic_value(x) for x in re.findall(r"\d+(?:\.0)?", raw)]
    if "~" in raw and len(nums) >= 2:
        return f"{nums[-2]}~{nums[-1]}"
    if nums:
        return nums[-1]
    return normalize_basic_value(raw)


def normalize_basic_value(value) -> str:
    s = stringify(value).replace("\u3000", " ").strip()
    s = re.sub(r"\s+", " ", s)
    if s.lower() in {"none", "nan", "null"}:
        return ""
    # 去千分位，保留复合字段分隔符场景外的普通数字。
    if re.fullmatch(r"[+-]?\d{1,3}(?:,\d{3})+(?:\.\d+)?", s):
        s = s.replace(",", "")
    if NUMBER_RE.match(s):
        try:
            num = float(s)
            if num.is_integer():
                return str(int(num))
            return ("%f" % num).rstrip("0").rstrip(".")
        except Exception:
            pass
    return s


def normalize_price_value(value: str) -> str:
    raw = stringify(value).strip()
    if not raw:
        return ""
    s = normalize_basic_value(raw)
    try:
        num = float(s)
        # 参考表常写 4.99，配置表常写 499，统一成分。
        if "." in s and abs(num) < 10000:
            return str(int(round(num * 100)))
        if num.is_integer():
            return str(int(num))
        return ("%f" % num).rstrip("0").rstrip(".")
    except Exception:
        return s


def normalize_percent_value(value: str) -> str:
    raw = stringify(value).strip()
    if not raw:
        return ""
    # 支持配置中常见展示格式：-5%|small、10 %、-5%|xxx。
    m = re.search(r"([+-]?\d+(?:\.\d+)?)\s*%", raw)
    if m:
        try:
            return normalize_basic_value(str(float(m.group(1)) / 100.0))
        except Exception:
            pass
    s = normalize_basic_value(raw)
    # Excel 可能把 5% 读作 0.05；CSV 可能直接写 0.05。保持小数语义。
    return s


def normalize_trigger_condition_value(value: str) -> str:
    raw = stringify(value).replace("\u3000", " ").strip()
    if not raw:
        return ""
    # 礼包参考表常写“主堡10级 / 解锁等级 10”，配置表可能只写 10 或条件字段。
    # 提取等级类数字用于同口径比较，避免所有“主堡X级 vs X”都误报。
    nums = re.findall(r"-?\d+(?:\.\d+)?", raw)
    if nums and re.search(r"级|等级|主堡|城镇|level|lv|open|condition", raw, re.I):
        return "|".join(normalize_basic_value(n) for n in nums)
    return normalize_basic_value(raw)


def normalize_composite_value(value: str, zero_equal: bool, percent_equal: bool) -> str:
    text = stringify(value).replace("\u3000", " ").strip()
    def repl(m):
        token = m.group(0)
        if percent_equal and token.strip().endswith("%"):
            return normalize_percent_value(token)
        return normalize_basic_value(token)
    normalized = re.sub(r"[+-]?\d{1,3}(?:,\d{3})+(?:\.\d+)?%?|[+-]?\d+(?:\.\d+)?%?", repl, text)
    normalized = re.sub(r"\s*([|*;,])\s*", r"\1", normalized).strip()
    if zero_equal and normalized in {"0", "0.0"}:
        return ""
    return normalized


def parse_field_set(text: str) -> set[str]:
    return {normalize_field_name(v) for v in parse_user_list(text)}


def normalize_field_name(value: str) -> str:
    return clean_header(value).strip().lower()


def canonical_field_name(value: str) -> str:
    """Return a stable semantic field name for Chinese/English near-synonym headers."""
    n = normalize_field_name(value)
    if not n:
        return ""
    if n in {"section_name", "__section", "__region"}:
        return n
    # 纯数字/小数是数据值，不是字段名；避免 open_end_conditions2 这类别名把“2”误识别成字段。
    if re.fullmatch(r"-?\d+(?:\.\d+)?", n):
        return n
    if n in FIELD_ALIAS_LOOKUP:
        return FIELD_ALIAS_LOOKUP[n]
    # 部分表头会带括号说明，例如“平台商品ID(有价格礼包不可填0)”。
    compact = re.sub(r"[（(].*?[）)]", "", n)
    if compact in FIELD_ALIAS_LOOKUP:
        return FIELD_ALIAS_LOOKUP[compact]
    # 长说明文本（例如配置表第一行中文说明）不做包含式匹配，避免被误判为字段名行。
    if len(n) > 30:
        return n
    for alias_norm, canon in FIELD_ALIAS_LOOKUP.items():
        if alias_norm and (alias_norm in n or n in alias_norm):
            # 避免过短 id/name/数字误伤长文本说明或数据行。
            if len(n) < 2 and n not in {"#id"}:
                continue
            if len(alias_norm) >= 3 or alias_norm in {"#id", "pcid", "type"}:
                return canon
    return n


def is_reference_reward_source_field(field: str) -> bool:
    n = normalize_field_name(field)
    if not n:
        return True
    if n.startswith("__col_"):
        return True
    if re.search(r"^(选项|道具|奖励)\d*", n):
        return True
    if re.search(r"数量|价值|原价值|总价值|返利|折扣|性价比", n):
        return True
    return False


def row_to_preview(row: StructuredRow, limit: int = 8) -> str:
    items = []
    for k, v in row.items():
        if k.startswith("__"):
            continue
        if stringify(v):
            items.append(f"{k}={v}")
        if len(items) >= limit:
            break
    return "；".join(items)


def natural_key(text: str):
    return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", str(text))]


def parse_config_records(path: str) -> Tuple[List[ValueRecord], List[str]]:
    p = Path(path)
    if p.suffix.lower() == ".csv":
        raw = read_text_with_fallback(p)
        rows = list(csv.reader(raw.splitlines()))
    elif p.suffix.lower() == ".xlsx":
        from openpyxl import load_workbook
        wb = load_workbook(p, data_only=True, read_only=True)
        ws = wb.worksheets[0]
        rows = [[cell for cell in row] for row in ws.iter_rows(values_only=True)]
        wb.close()
    elif p.suffix.lower() == ".json":
        data = json.loads(read_text_with_fallback(p))
        return parse_json_config_records(data), ["json"]
    else:
        rows = list(csv.reader(read_text_with_fallback(p).splitlines()))

    if not rows:
        return [], []
    header_index = detect_config_header_row(rows)
    headers = [clean_header(v) for v in rows[header_index]]
    columns = {h: i for i, h in enumerate(headers) if h}
    fields = [h for h in headers if h]
    id_col = find_col(columns, ["#goods_id", "goods_id", "商品id", "商品ID", "id"])
    if id_col is None:
        return [], fields

    records: List[ValueRecord] = []
    for row_num, row in enumerate(rows[header_index + 1 :], start=header_index + 2):
        goods_id = get_cell(row, id_col)
        if not goods_id:
            continue
        raw_map = {h: get_cell(row, i) for h, i in columns.items()}
        reward_str = first_by_keys(raw_map, ["reward", "礼包内容"])
        reward = stringify(row[26]) if len(row) >= 27 else ""
        price_value = to_int_or_none(row[18] if len(row) >= 19 else None)
        if not ("*" in reward or price_value is not None):
            continue
        rec = ValueRecord(
            record_id=goods_id,
            name=first_by_keys(raw_map, ["^comment", "comment", "name", "策划注释点此查看goods_id配置规范", "礼包名称"]),
            sheet="配置表",
            row=row_num,
            price=to_int_or_none(first_by_keys(raw_map, ["price", "价格", "代币价格"])),
            recharge_price=to_int_or_none(first_by_keys(raw_map, ["recharge_price", "末日充值金额"])),
            pcid=first_by_keys(raw_map, ["pcid"]),
            duration=to_int_or_none(first_by_keys(raw_map, ["duration", "周月卡的持续时间"])),
            score=to_int_or_none(first_by_keys(raw_map, ["score", "充值可累计的积分"])),
            limit_time=to_int_or_none(first_by_keys(raw_map, ["limit_time", "限购次数"])),
            rewards=parse_config_reward_string(reward_str),
            raw=raw_map,
        )
        records.append(rec)
    return records, fields


def parse_json_config_records(data) -> List[ValueRecord]:
    if isinstance(data, dict):
        values = data.values()
    elif isinstance(data, list):
        values = data
    else:
        return []
    records: List[ValueRecord] = []
    for index, item in enumerate(values, start=1):
        if not isinstance(item, dict):
            continue
        raw_map = {str(k): "" if v is None else str(v) for k, v in item.items()}
        goods_id = first_by_keys(raw_map, ["#goods_id", "goods_id", "id", "商品ID"])
        if not goods_id:
            continue
        records.append(ValueRecord(
            record_id=goods_id,
            name=first_by_keys(raw_map, ["^comment", "comment", "name", "商品名称"]),
            sheet="配置表", row=index, price=to_int_or_none(first_by_keys(raw_map, ["price"])),
            recharge_price=to_int_or_none(first_by_keys(raw_map, ["recharge_price"])),
            pcid=first_by_keys(raw_map, ["pcid"]), duration=to_int_or_none(first_by_keys(raw_map, ["duration"])),
            score=to_int_or_none(first_by_keys(raw_map, ["score"])),
            limit_time=to_int_or_none(first_by_keys(raw_map, ["limit_time"])),
            rewards=parse_config_reward_string(first_by_keys(raw_map, ["reward"])), raw=raw_map
        ))
    return records


def parse_reference_records(path: str) -> Tuple[List[ValueRecord], List[str], List[str]]:
    from openpyxl import load_workbook
    p = Path(path)
    wb = load_workbook(p, data_only=True, read_only=True)
    item_map = parse_item_maps(wb)
    records: List[ValueRecord] = []
    used_sheets: List[str] = []
    notes: List[str] = []

    # 优先识别参考表中直接粘贴的游戏配置行；这类行通常最可靠，且能避免同一工作簿中说明区/示例区再次产生误报。
    direct_records: List[ValueRecord] = []
    direct_sheets: List[str] = []
    sheet_rows = []
    for ws in wb.worksheets:
        title = ws.title.strip()
        rows = [[cell for cell in row] for row in ws.iter_rows(values_only=True)]
        if is_mapping_sheet(title) or looks_like_item_map_sheet(rows):
            continue
        sheet_rows.append((title, rows))
        parsed_direct = parse_reference_mall_like_rows(title, rows)
        if parsed_direct:
            direct_records.extend(parsed_direct)
            direct_sheets.append(title)
    if direct_records:
        wb.close()
        return direct_records, direct_sheets, notes

    for title, rows in sheet_rows:
        if not rows:
            continue
        header_index = detect_reference_header_row(rows)
        if header_index is None:
            # 仍尝试按整张表扫描奖励；双子星这类非标准表至少能识别到一部分 id/数量组合。
            header_index = 0
        sheet_records = records_from_reference_sheet(title, rows, header_index, item_map)
        if sheet_records:
            used_sheets.append(title)
            records.extend(sheet_records)
    wb.close()
    if not used_sheets:
        notes.append("未自动识别到带奖励/id/数量的 Sheet")
    return records, used_sheets, notes


def parse_item_maps(wb) -> Dict[str, Tuple[str, int]]:
    result: Dict[str, Tuple[str, int]] = {}
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not (is_mapping_sheet(ws.title.strip()) or looks_like_item_map_sheet(rows)):
            continue
        for row in rows:
            if len(row) < 2:
                continue
            name = str(row[0]).strip() if row[0] is not None else ""
            item_id = normalize_id(row[1]) if row[1] is not None else ""
            value = to_int_or_none(row[2] if len(row) > 2 else None) or 0
            if name and item_id and item_id.lower() not in {"道具id", "#item_id", "uint"}:
                result[normalize_name(name)] = (item_id, value)
    return result


def records_from_reference_sheet(title: str, rows: List[List[object]], header_index: int, item_map: Dict[str, Tuple[str, int]]) -> List[ValueRecord]:
    direct_records = parse_reference_mall_like_rows(title, rows)
    if direct_records:
        return direct_records
    header = [clean_header(v) for v in rows[header_index]] if header_index < len(rows) else []
    records: List[ValueRecord] = []
    for r_index in range(header_index + 1, len(rows)):
        row = rows[r_index]
        if not any(v not in (None, "") for v in row):
            continue
        rec = build_reference_record(title, r_index + 1, row, header, item_map)
        if rec.rewards or rec.record_id or rec.price is not None or rec.score is not None:
            records.append(rec)
    # 对月卡/周卡这类 Sheet：如果只有少量有效行，按 Sheet 聚合，减少分散匹配。
    if title in {"周卡", "月卡"}:
        useful = [r for r in records if r.rewards or r.price is not None]
        if useful:
            merged = ValueRecord(sheet=title, row=useful[0].row, name=best_sheet_name(title, useful), price=first_non_none([r.price for r in useful]), duration=7 if "周" in title else 30)
            for r in useful:
                merged.rewards.extend(r.rewards)
            return [merged]
    return records



def parse_reference_mall_like_rows(title: str, rows: List[List[object]]) -> List[ValueRecord]:
    """Parse reference rows that directly paste mall.csv style config rows.

    常见数值参考表会把 mall.csv 的部分行粘到 Sheet 末尾；此时第一列是 goods_id，
    第 2 列 pcid，第 3 列名称，第 19 列 price，第 22 列 duration，第 27 列 reward。
    直接使用这些行可大幅减少整表泛比对误报。
    """
    records: List[ValueRecord] = []
    for row_num, row in enumerate(rows, start=1):
        if len(row) < 3:
            continue
        goods_id = normalize_id(row[0])
        if not (goods_id.isdigit() and len(goods_id) >= 7):
            continue
        name = stringify(row[2]) if len(row) >= 3 else ""
        reward = stringify(row[26]) if len(row) >= 27 else ""
        reward = stringify(row[26]) if len(row) >= 27 else ""
        price_value = to_int_or_none(row[18] if len(row) >= 19 else None)
        if not ("*" in reward or price_value is not None):
            continue
        rec = ValueRecord(
            record_id=goods_id,
            name=name or title,
            sheet=title,
            row=row_num,
            pcid=stringify(row[1]) if len(row) >= 2 else "",
            price=price_value,
            duration=to_int_or_none(row[21] if len(row) >= 22 else None),
            limit_time=to_int_or_none(row[14] if len(row) >= 15 else None),
            score=None,
            rewards=parse_config_reward_string(reward),
            raw={"reward": reward},
        )
        records.append(rec)
    return records

def build_reference_record(title: str, row_num: int, row: List[object], header: List[str], item_map: Dict[str, Tuple[str, int]]) -> ValueRecord:
    raw: Dict[str, str] = {header[i]: stringify(row[i]) for i in range(min(len(header), len(row))) if header[i]}
    name = first_by_keys(raw, ["礼包名称", "商品名称", "名称", "name"])
    if not name:
        # 周卡/月卡常见名称放在表上方，当前行可用 Sheet 名兜底。
        name = title
    rec = ValueRecord(
        record_id=first_by_keys(raw, ["商品ID", "商品id", "礼包对应id", "礼包id", "档位", "id"]),
        name=name,
        sheet=title,
        row=row_num,
        price=reference_price_to_cents(first_by_keys(raw, ["美刀定价", "代币价格", "真实充值", "price"])),
        pcid=first_by_keys(raw, ["pcid"]),
        duration=to_int_or_none(first_by_keys(raw, ["持续时间", "duration"])),
        score=to_int_or_none(first_by_keys(raw, ["累充积分值", "积分", "score"])),
        limit_time=to_int_or_none(first_by_keys(raw, ["限购次数", "limit_time"])),
        raw=raw,
    )
    rec.rewards.extend(extract_reference_rewards_from_row(row, header, item_map, title))
    return rec


def extract_reference_rewards_from_row(row: List[object], header: List[str], item_map: Dict[str, Tuple[str, int]], sheet: str) -> List[RewardItem]:
    rewards: List[RewardItem] = []
    max_len = max(len(row), len(header))
    for col in range(max_len):
        h = header[col] if col < len(header) else ""
        if not h:
            continue
        h_norm = h.lower()
        is_reward_name_col = bool(REWARD_HEADER_RE.search(h)) or h in {"装扮奖励", "联盟礼物"}
        if not is_reward_name_col:
            continue
        name = stringify(row[col]) if col < len(row) else ""
        if not name or name in {"0", "None"}:
            continue
        id_col = find_neighbor_header(header, col + 1, {"id", "ID", "道具ID", "道具id"})
        qty_col = find_neighbor_header(header, col + 1, {"数量", "num", "count"})
        item_id = stringify(row[id_col]) if id_col is not None and id_col < len(row) else ""
        qty = to_int_or_none(row[qty_col] if qty_col is not None and qty_col < len(row) else None)
        if item_id and not normalize_id(item_id).isdigit():
            continue
        if not item_id:
            mapped = item_map.get(normalize_name(name))
            if mapped:
                item_id = mapped[0]
        if qty is None:
            qty = 1
        reward = normalize_reference_reward(item_id, qty, name, item_map, source=f"{sheet}:{h}")
        if reward:
            rewards.append(reward)
    # 兼容“奖励6 / 数量 / 价值”缺 id 的配置，尝试通过道具引用名称映射 ID。
    return rewards


def normalize_reference_reward(item_id: str, qty: int, name: str, item_map: Dict[str, Tuple[str, int]], source: str = "") -> Optional[RewardItem]:
    item_id = normalize_id(item_id)
    name_key = normalize_name(name)
    mapped_value = 0
    mapped = item_map.get(name_key)
    if mapped:
        if not item_id:
            item_id = mapped[0]
        mapped_value = mapped[1]
    if not item_id or not normalize_id(item_id).isdigit():
        return None
    if qty <= 0:
        return None
    # 钻石道具转换：参考表常用 500钻石/1000钻石道具 ID，mall reward 常用 1*1*钻石数。
    if "钻石" in name or item_id in {"10603", "10604", "10605", "10606"}:
        diamond_value = mapped_value or diamond_value_from_name(name) or diamond_value_from_id(item_id)
        if diamond_value:
            return RewardItem(item_id="1", quantity=diamond_value * qty, name=name, reward_type="1", source=source, raw=f"{name}({item_id})*{qty}")
    # VIP 时间转换：参考表常用 VIP时间30天，mall reward 常用 1*13*秒数。
    if "vip" in name.lower() or "VIP" in name or item_id in {"34003"}:
        seconds = vip_seconds_from_name(name) or (mapped_value if mapped_value and mapped_value > 100000 else 0)
        if seconds:
            return RewardItem(item_id="13", quantity=seconds * qty, name=name, reward_type="1", source=source, raw=f"{name}({item_id})*{qty}")
    return RewardItem(item_id=item_id, quantity=qty, name=name, reward_type="2", source=source, raw=f"{name}({item_id})*{qty}")


def parse_config_reward_string(value: str) -> List[RewardItem]:
    rewards: List[RewardItem] = []
    for token in str(value or "").split("|"):
        token = token.strip()
        if not token:
            continue
        m = REWARD_TOKEN_RE.match(token)
        if not m:
            continue
        reward_type, item_id, qty = m.groups()
        qty_int = to_int_or_none(qty) or 0
        if qty_int <= 0:
            continue
        rewards.append(RewardItem(item_id=normalize_id(item_id), quantity=qty_int, reward_type=str(reward_type), raw=token))
    return rewards


def match_records(reference_records: List[ValueRecord], config_records: List[ValueRecord], manual_ids: List[str], manual_keywords: List[str]) -> Dict[str, ValueRecord]:
    result: Dict[str, ValueRecord] = {}
    config_by_id = {normalize_id(r.record_id): r for r in config_records if r.record_id}
    filtered_config = config_records
    if manual_ids:
        id_set = {normalize_id(v) for v in manual_ids}
        filtered_config = [r for r in config_records if normalize_id(r.record_id) in id_set]
    if manual_keywords:
        filtered_config = [r for r in filtered_config if any(k.lower() in (r.name or "").lower() or k.lower() in normalize_id(r.record_id).lower() for k in manual_keywords)]

    for ref in reference_records:
        match = None
        if ref.record_id and normalize_id(ref.record_id) in config_by_id:
            match = config_by_id[normalize_id(ref.record_id)]
        if match is None:
            match = best_config_match(ref, filtered_config)
        if match is not None:
            result[record_signature(ref)] = match
    return result


def best_config_match(ref: ValueRecord, config_records: List[ValueRecord]) -> Optional[ValueRecord]:
    candidates: List[Tuple[int, ValueRecord]] = []
    ref_tokens = tokens_for_match(ref.name + " " + ref.sheet)
    for cfg in config_records:
        score = 0
        cfg_text = (cfg.name or "") + " " + (cfg.record_id or "")
        cfg_tokens = tokens_for_match(cfg_text)
        common = ref_tokens & cfg_tokens
        score += len(common) * 20
        if ref.price is not None and cfg.price == ref.price:
            score += 30
        if ref.duration is not None and cfg.duration == ref.duration:
            score += 15
        if ref.score is not None and cfg.score == ref.score:
            score += 30
        # 周/月卡专项：排除 BP，优先 900 开头的周月卡商品。
        if ("周卡" in ref.name or "周卡" in ref.sheet or "月卡" in ref.name or "月卡" in ref.sheet) and "BP" in cfg.name.upper():
            score -= 25
        if cfg.record_id.startswith("900"):
            score += 10
        if ref.name and normalize_name(ref.name) and normalize_name(ref.name) in normalize_name(cfg.name):
            score += 50
        if score > 0:
            candidates.append((score, cfg))
    if not candidates:
        return None
    candidates.sort(key=lambda item: item[0], reverse=True)
    # 阈值避免错把无关商品硬匹配。
    return candidates[0][1] if candidates[0][0] >= 30 else None


def compare_rewards(diffs: List[DiffItem], idx: int, ref: ValueRecord, cfg: ValueRecord) -> int:
    ref_rewards = merge_rewards(ref.rewards)
    cfg_rewards = merge_rewards([r for r in cfg.rewards if r.reward_type != "12"])  # 联盟礼物默认不和参考奖励混比，避免误报。
    ref_keys = set(ref_rewards)
    cfg_keys = set(cfg_rewards)
    location = f"{cfg.record_id} / {cfg.name or ref.display_name()}"
    for key in sorted(ref_keys - cfg_keys):
        rr = ref_rewards[key]
        diffs.append(DiffItem(idx, "奖励缺失", format_reward(rr), "未配置", location, "参考表有该奖励，但配置表 reward 未找到对应道具"))
        idx += 1
    for key in sorted(cfg_keys - ref_keys):
        cr = cfg_rewards[key]
        # 只在参考有奖励时才提示多配；否则很多配置行会因为参考表结构未识别而误报。
        if ref_rewards:
            diffs.append(DiffItem(idx, "奖励多配", "未填写", format_reward(cr), location, "配置表 reward 存在参考表未识别到的奖励，请确认是否为额外奖励/联盟礼物/配置多配"))
            idx += 1
    for key in sorted(ref_keys & cfg_keys):
        rr = ref_rewards[key]
        cr = cfg_rewards[key]
        if rr.quantity != cr.quantity:
            diffs.append(DiffItem(idx, "数量不一致", format_reward(rr), format_reward(cr), location, "同一道具 ID 的奖励数量与参考表不一致"))
            idx += 1
    return idx


def compare_number_field(diffs: List[DiffItem], idx: int, ref: ValueRecord, cfg: ValueRecord, label: str, ref_val: Optional[int], cfg_val: Optional[int], allow_none: bool = True) -> int:
    if allow_none and ref_val is None:
        return idx
    if ref_val != cfg_val:
        diffs.append(DiffItem(idx, f"{label}不一致", f"{label}: {ref_val if ref_val is not None else '未填写'}", f"{label}: {cfg_val if cfg_val is not None else '未配置'}", f"{cfg.record_id} / {cfg.name or ref.display_name()}", f"请确认配置表 {label} 是否与参考表一致"))
        return idx + 1
    return idx


def compare_text_field(diffs: List[DiffItem], idx: int, ref: ValueRecord, cfg: ValueRecord, label: str, ref_val: str, cfg_val: str, allow_none: bool = True) -> int:
    if allow_none and not ref_val:
        return idx
    if normalize_id(ref_val) != normalize_id(cfg_val):
        diffs.append(DiffItem(idx, f"{label}不一致", f"{label}: {ref_val or '未填写'}", f"{label}: {cfg_val or '未配置'}", f"{cfg.record_id} / {cfg.name or ref.display_name()}", f"请确认配置表 {label} 是否与参考表一致"))
        return idx + 1
    return idx


def classify_structured_field_difference(ref_field: str, cfg_field: str, ref_val, cfg_val, ref_row: StructuredRow, cfg_row: StructuredRow) -> Tuple[str, str, bool]:
    canon = canonical_field_name(ref_field) or canonical_field_name(cfg_field)
    ref_text = stringify(ref_val)
    cfg_text = stringify(cfg_val)
    section = row_section_value(ref_row)
    cfg_price = row_price_value(cfg_row)
    if canon == "pcid":
        if normalize_basic_value(ref_text) not in {"", "0"} and normalize_basic_value(cfg_text) in {"", "0"} and cfg_price not in {"", "0"}:
            return "确认字段差异", "付费礼包参考表有 pcid，但配置表 pcid 为 0，属于高风险支付档位差异，需修复或确认商品是否免费/资源购买。", True
        return "确认字段差异", "pcid 不一致；价格一致不能代表 pcid 一致，请按支付档位确认。", True
    if canon == "price":
        return "确认字段差异", "价格不一致；已按美元小数与分单位整数做 9.99↔999 换算后比较。", True
    if canon == "trigger_level":
        if ("自选" in section or "自选" in row_name_value(ref_row)) and normalize_basic_value(cfg_text) in {"", "0"}:
            return "待确认字段差异", "自选礼包触发等级可能配置在父活动或其他关联表，当前子商品未抽取到等级；不直接计入确认差异。", False
        if normalize_basic_value(cfg_text) in {"", "0"}:
            return "待确认字段差异", "配置表未抽取到明确触发等级，可能是触发对象/类型码而非等级；请人工确认 open_end_conditions/open_condition。", False
        return "确认字段差异", "触发/解锁等级不一致；已避免将 condition 类型码直接当等级比较。", True
    if canon == "reward":
        if "自选" in section or "自选" in row_name_value(ref_row):
            return "待确认字段差异", "自选礼包奖励包含选项结构，已仅使用真实道具ID+数量生成 reward，未使用价值字段；该差异需确认选项配置口径后处理。", False
        return "确认字段差异", "reward 不一致；参考表 reward 仅由道具ID+数量生成，未使用价值/原价值/总价值/返利比字段。", True
    if canon == "open_end_conditions":
        return "待确认字段差异", "条件字段含义可能是触发类型/参数组合，无法确认时不直接判定为确认差异。", False
    return "确认字段差异", f"请确认字段 {ref_field} => {cfg_field} 是否一致。", True


def merge_rewards(rewards: Iterable[RewardItem]) -> Dict[str, RewardItem]:
    merged: Dict[str, RewardItem] = {}
    for r in rewards:
        if not r.key:
            continue
        if r.key not in merged:
            merged[r.key] = RewardItem(item_id=r.item_id, quantity=0, name=r.name, reward_type=r.reward_type, source=r.source, raw=r.raw)
        merged[r.key].quantity += int(r.quantity or 0)
        if not merged[r.key].name and r.name:
            merged[r.key].name = r.name
    return merged


def format_reward(r: RewardItem) -> str:
    name = f"{r.name} / " if r.name else ""
    return f"{name}ID {normalize_id(r.item_id)} × {r.quantity}" + (f"（原始：{r.raw}）" if r.raw else "")


def detect_config_header_row(rows: List[List[object]]) -> int:
    for i, row in enumerate(rows[:8]):
        cleaned = {clean_header(v) for v in row}
        if "#goods_id" in cleaned or "reward" in cleaned:
            return i
    return 0


def detect_reference_header_row(rows: List[List[object]]) -> Optional[int]:
    best_idx = None
    best_score = 0
    for i, row in enumerate(rows[:15]):
        text = "|".join(clean_header(v) for v in row)
        score = 0
        for kw in ["奖励", "id", "数量", "礼包", "美刀定价", "累充积分", "商品ID", "礼包对应id"]:
            if kw.lower() in text.lower():
                score += 1
        if score > best_score:
            best_idx = i
            best_score = score
    return best_idx if best_score >= 2 else None


def find_neighbor_header(header: List[str], start_col: int, candidates: set[str]) -> Optional[int]:
    cand_norm = {clean_header(c).lower() for c in candidates}
    for col in range(start_col, min(len(header), start_col + 5)):
        if clean_header(header[col]).lower() in cand_norm:
            return col
    return None


def find_col(columns: Dict[str, int], candidates: Iterable[str]) -> Optional[int]:
    lookup = {clean_header(k).lower(): v for k, v in columns.items()}
    for cand in candidates:
        key = clean_header(cand).lower()
        if key in lookup:
            return lookup[key]
    # fuzzy
    for cand in candidates:
        key = clean_header(cand).lower().replace("#", "")
        for col_name, idx in lookup.items():
            if col_name.replace("#", "") == key:
                return idx
    return None


def first_by_keys(raw: Dict[str, str], keys: Iterable[str]) -> str:
    lower_map = {clean_header(k).lower(): v for k, v in raw.items()}
    for key in keys:
        ck = clean_header(key).lower()
        if ck in lower_map and str(lower_map[ck]).strip():
            return str(lower_map[ck]).strip()
    for key in keys:
        ck = clean_header(key).lower()
        for k, v in lower_map.items():
            if ck and (ck in k or k in ck) and str(v).strip():
                return str(v).strip()
    return ""


def first_non_none(values: Iterable[Optional[int]]) -> Optional[int]:
    for v in values:
        if v is not None:
            return v
    return None


def get_cell(row: List[object], index: int) -> str:
    if index < 0 or index >= len(row):
        return ""
    return stringify(row[index])


def stringify(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def clean_header(value) -> str:
    return re.sub(r"\s+", "", stringify(value)).strip()


def normalize_id(value) -> str:
    s = stringify(value)
    if s.endswith(".0") and NUMBER_RE.match(s):
        s = s[:-2]
    return s.strip()


def normalize_name(value) -> str:
    return re.sub(r"[\s\-_/（）()【】\[\]：:，,]+", "", stringify(value)).lower()


def tokens_for_match(text: str) -> set[str]:
    text = stringify(text)
    tokens = set(re.findall(r"[A-Za-z0-9]+|[\u4e00-\u9fff]{2,}", text))
    # 加入常用中文关键词，避免整串中文只有一个大 token。
    for kw in ["月卡", "周卡", "超级月卡", "常规礼包", "无尽礼包", "无限礼包", "商城累充", "双子星", "宠物", "加速", "文森特", "斯嘉丽"]:
        if kw in text:
            tokens.add(kw)
    return {t.lower() for t in tokens if t}


def parse_user_list(text: str) -> List[str]:
    return [v.strip() for v in re.split(r"[,，\s\n;；]+", text or "") if v.strip()]


def to_int_or_none(value) -> Optional[int]:
    s = stringify(value)
    if not s:
        return None
    try:
        return int(float(s))
    except Exception:
        return None


def reference_price_to_cents(value) -> Optional[int]:
    s = stringify(value)
    if not s:
        return None
    try:
        num = float(s)
    except Exception:
        return None
    # 美刀 4.99 -> 499；代币价格 499 则保持 499。
    if 0 < num < 100:
        return int(round(num * 100))
    return int(round(num))


def diamond_value_from_name(name: str) -> int:
    m = re.search(r"(\d+)\s*钻石", stringify(name))
    return int(m.group(1)) if m else 0


def diamond_value_from_id(item_id: str) -> int:
    return {"10603": 100, "10604": 500, "10605": 1000, "10606": 10000}.get(normalize_id(item_id), 0)


def vip_seconds_from_name(name: str) -> int:
    m = re.search(r"(\d+)\s*天", stringify(name))
    return int(m.group(1)) * 86400 if m else 0



def looks_like_item_map_sheet(rows: List[List[object]]) -> bool:
    if not rows:
        return False
    head = "|".join(clean_header(v) for v in rows[0][:5])
    return ("道具ID" in head or "#item_id" in head or "备注,程序不读" in head) and len(rows) > 20

def is_mapping_sheet(title: str) -> bool:
    return any(k in title for k in ["道具引用", "道具索引", "道具ID", "item"])


def best_sheet_name(title: str, records: List[ValueRecord]) -> str:
    names = [r.name for r in records if r.name and r.name != title]
    return names[0] if names else title


def record_signature(record: ValueRecord) -> str:
    return f"{record.sheet}:{record.row}:{record.record_id}:{record.name}"


def reindex(diffs: List[DiffItem]) -> List[DiffItem]:
    for i, d in enumerate(diffs, start=1):
        d.index = i
    return diffs
