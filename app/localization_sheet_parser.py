from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, Iterable, List, Sequence, Tuple

from .file_readers import FileReadError
from .translation_compare import find_column, normalize_header

LANGUAGE_SHEET_MAP: Dict[str, str] = {
    "EN": "英语",
    "DE": "德语",
    "FR": "法语",
    "RU": "俄语",
    "ES": "西语",
    "JP": "日语",
    "JA": "日语",
    "ID": "印尼语",
    "KR": "韩语",
    "KO": "韩语",
    "TH": "泰语",
    "IT": "意大利语",
    "PT": "葡语",
    "PT_BR": "巴西葡语",
    "PTBR": "巴西葡语",
    "TR": "土耳其语",
    "VN": "越南语",
    "VI": "越南语",
    "ARB": "阿语",
    "AR": "阿语",
    "HIN": "印地语",
    "HI": "印地语",
    "UR": "乌尔都语",
    "FA": "波斯语",
    "MS": "马来语",
    "UA": "乌克兰语",
    "UK": "乌克兰语",
    "TW": "中文繁体",
    "CN": "中文简体",
}

LANGUAGE_PRIORITY = ["EN", "TW", "RU", "KR", "DE", "TR", "TH", "FR", "ARB", "HIN", "ID", "UR", "FA", "MS", "ES", "JP", "IT", "PT_BR", "PTBR", "PT", "VN", "UA"]

ID_CANDIDATES = ["ID", "Text ID", "字串编号", "字符 ID", "字串ID", "文本ID", "text_id", "tid", "#tid", "key"]
SOURCE_CANDIDATES = [
    "繁体", "繁中", "原文", "CN 原文", "CN", "中文", "中文原文", "Source", "Source Text",
    "Traditional", "Traditional CN", "Text", "源文本",
]
TARGET_CANDIDATES = ["翻译", "Translation", "Target", "Target Text", "译文", "多语言", "文本"]

MULTI_SHEET_STRUCTURE = "multi_sheet"
SINGLE_SHEET_STRUCTURE = "single_sheet"
UNKNOWN_STRUCTURE = "unknown"


def normalize_localization_mode(mode: str) -> str:
    """Normalize UI structure text to multi_sheet / single_sheet / auto.

    这里必须精确判断，不能再用 "多" in mode 这类模糊判断，否则
    “单 Sheet 多语言列”会被误判成“多 Sheet 语言页”。
    """
    text = str(mode or "").strip().lower()
    text = text.replace(" ", "").replace("\u3000", "").replace("_", "").replace("-", "")
    if text in {"多sheet语言页", "多sheet", "multisheet", "multisheetlanguage", "multisheetlanguages"}:
        return MULTI_SHEET_STRUCTURE
    if text in {"单sheet多语言列", "单sheet", "singlesheet", "singlesheetcolumns", "singlesheetlanguagecolumns"}:
        return SINGLE_SHEET_STRUCTURE
    return "auto"


@dataclass
class LanguageSheetInfo:
    language: str
    language_name: str
    sheet_name: str
    header_row: int = 1
    data_start_row: int = 2
    id_col: int = -1
    source_col: int = -1
    target_col: int = -1
    headers: List[str] = field(default_factory=list)
    error: str = ""


@dataclass
class NormalizedLocalizationEntry:
    text_id: str
    source_text: str
    language: str
    language_name: str
    translation: str
    sheet_name: str
    row_number: int
    column_name: str


@dataclass
class NormalizedLocalizationDataset:
    structure: str
    path: str
    sheet_infos: List[LanguageSheetInfo]
    entries: List[NormalizedLocalizationEntry]
    records_by_language: Dict[str, Dict[str, NormalizedLocalizationEntry]]
    duplicate_ids: Dict[str, Dict[str, List[int]]]
    invalid_rows: List[Tuple[str, int]]
    header_errors: List[str]
    base_language: str = ""

    @property
    def language_codes(self) -> List[str]:
        return [info.language for info in self.sheet_infos if not info.error]

    @property
    def text_ids(self) -> set[str]:
        ids: set[str] = set()
        for records in self.records_by_language.values():
            ids.update(records.keys())
        return ids


def canonical_language_code(sheet_name: str) -> str:
    raw = (sheet_name or "").strip()
    normalized = normalize_header(raw).upper().replace("_", "")
    aliases = {
        "ENGLISH": "EN", "GERMAN": "DE", "FRENCH": "FR", "RUSSIAN": "RU", "SPANISH": "ES",
        "JAPANESE": "JP", "KOREAN": "KR", "THAI": "TH", "ITALIAN": "IT", "PORTUGUESE": "PT", "PTBR": "PT_BR", "PT_BR": "PT_BR", "BRAZILIANPORTUGUESE": "PT_BR",
        "TURKISH": "TR", "VIETNAMESE": "VN", "ARABIC": "ARB", "INDONESIAN": "ID",
        "HINDI": "HIN", "URDU": "UR", "FARSI": "FA", "PERSIAN": "FA", "MALAY": "MS",
        "繁体": "TW", "繁體": "TW", "中文繁体": "TW", "中文繁體": "TW", "简体": "CN", "簡體": "CN",
    }
    if normalized in aliases:
        return aliases[normalized]
    if normalized in LANGUAGE_SHEET_MAP:
        return normalized
    # 支持 EN_翻译 / JP翻译 这类表名。
    for code in sorted(LANGUAGE_SHEET_MAP, key=len, reverse=True):
        if normalized == code or normalized.startswith(code) or normalized.endswith(code):
            return code
    return ""


def is_language_sheet(sheet_name: str) -> bool:
    return bool(canonical_language_code(sheet_name))


def _load_xlsx_rows(path: str) -> Dict[str, List[List[str]]]:
    file_path = Path(path)
    if file_path.suffix.lower() != ".xlsx":
        raise FileReadError("多 Sheet 语言页暂只支持 xlsx 文件")
    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        raise FileReadError("缺少依赖 openpyxl，请先安装 requirements.txt") from exc
    wb = load_workbook(file_path, data_only=True, read_only=True)
    try:
        result: Dict[str, List[List[str]]] = {}
        for ws in wb.worksheets:
            try:
                ws.reset_dimensions()
            except Exception:
                pass
            rows: List[List[str]] = []
            for row in ws.iter_rows(values_only=True):
                rows.append(["" if cell is None else str(cell).strip() for cell in row])
            result[ws.title] = rows
        return result
    finally:
        wb.close()


def list_language_sheets(path: str) -> List[Tuple[str, str, str]]:
    rows_by_sheet = _load_xlsx_rows(path)
    items: List[Tuple[str, str, str]] = []
    for sheet_name in rows_by_sheet:
        code = canonical_language_code(sheet_name)
        # 多 Sheet 语言页只按 Sheet 名判断语言。ID 在这里是印尼语 Sheet，
        # ID 在表头里仍然由字段检测逻辑当作主键列处理。
        if code and code in LANGUAGE_PRIORITY:
            items.append((code, LANGUAGE_SHEET_MAP.get(code, code), sheet_name))
    return sorted(items, key=lambda item: LANGUAGE_PRIORITY.index(item[0]) if item[0] in LANGUAGE_PRIORITY else 999)


def _has_multi_sheet_required_columns(rows: List[List[str]], sheet_name: str) -> bool:
    info = detect_language_sheet_columns(rows, sheet_name, 0)
    return not info.error


def _looks_like_single_sheet_rows(rows: List[List[str]]) -> bool:
    # 只检查前 10 行，找到“ID + 源文本 + 至少 1 个语言列”即可视为单 Sheet 多语言列。
    language_codes = set(LANGUAGE_PRIORITY)
    for row in rows[:10]:
        headers = [str(cell).strip() for cell in row]
        if not any(headers):
            continue
        id_col = find_column(headers, "", ID_CANDIDATES)
        source_col = find_column(headers, "", SOURCE_CANDIDATES)
        if id_col < 0 or source_col < 0:
            continue
        lang_count = 0
        for idx, header in enumerate(headers):
            if idx in {id_col, source_col}:
                continue
            code = canonical_language_code(header)
            if code in language_codes:
                lang_count += 1
        if lang_count >= 1:
            return True
    return False


def detect_localization_structure(path: str) -> str:
    """Return multi_sheet / single_sheet / unknown.

    轻量识别：只读取 xlsx workbook 元数据，避免大表在“识别结构”和“开始检查”阶段重复整表加载。
    多 Sheet 语言页仍按 EN/DE/FR/RU 等 Sheet 名识别；其他 xlsx/csv 默认交给单 Sheet 解析器做精细诊断。
    """
    file_path = Path(path)
    ext = file_path.suffix.lower()
    if ext == ".xlsx":
        try:
            with zipfile.ZipFile(file_path) as zf:
                xml = zf.read("xl/workbook.xml")
            root = ET.fromstring(xml)
            ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            sheet_names = [elem.attrib.get("name", "") for elem in root.findall(".//main:sheet", ns)]
            language_sheets = [name for name in sheet_names if canonical_language_code(name)]
            if len(language_sheets) >= 2:
                return MULTI_SHEET_STRUCTURE
            return SINGLE_SHEET_STRUCTURE if sheet_names else UNKNOWN_STRUCTURE
        except Exception:
            return UNKNOWN_STRUCTURE
    if ext == ".csv":
        return SINGLE_SHEET_STRUCTURE
    return UNKNOWN_STRUCTURE

def _row_score_for_header(row: Sequence[str]) -> int:
    values = [str(cell).strip() for cell in row]
    if not any(values):
        return 0
    score = 0
    if find_column(values, "", ID_CANDIDATES) >= 0:
        score += 4
    if find_column(values, "", SOURCE_CANDIDATES) >= 0:
        score += 3
    if find_column(values, "", TARGET_CANDIDATES) >= 0:
        score += 3
    normalized = [normalize_header(v) for v in values]
    if any(v in {"id", "text_id", "字串编号", "tid", "#tid"} for v in normalized):
        score += 2
    return score


def detect_language_sheet_columns(rows: List[List[str]], sheet_name: str, forced_header_row: int = 0) -> LanguageSheetInfo:
    code = canonical_language_code(sheet_name) or sheet_name
    info = LanguageSheetInfo(language=code, language_name=LANGUAGE_SHEET_MAP.get(code, code), sheet_name=sheet_name)
    if not rows:
        info.error = "Sheet 为空"
        return info

    if forced_header_row and 1 <= forced_header_row <= len(rows):
        header_index = forced_header_row - 1
    else:
        candidates = [(idx, _row_score_for_header(row)) for idx, row in enumerate(rows[:10])]
        candidates.sort(key=lambda item: item[1], reverse=True)
        header_index = candidates[0][0] if candidates and candidates[0][1] > 0 else 0
    headers = [str(cell).strip() for cell in rows[header_index]]
    info.header_row = header_index + 1
    info.data_start_row = header_index + 2
    info.headers = headers
    info.id_col = find_column(headers, "", ID_CANDIDATES)
    info.source_col = find_column(headers, "", SOURCE_CANDIDATES)
    info.target_col = find_column(headers, "", TARGET_CANDIDATES)

    # 翻译列兜底：优先使用 ID/源文本之外的最后一个非空列。
    if info.target_col < 0:
        excluded = {info.id_col, info.source_col}
        for idx in range(len(headers) - 1, -1, -1):
            if idx not in excluded and str(headers[idx]).strip():
                info.target_col = idx
                break
    if info.source_col < 0:
        # 源文本列兜底：ID 后的第一个非翻译列。
        excluded = {info.id_col, info.target_col}
        for idx, header in enumerate(headers):
            if idx not in excluded and str(header).strip():
                info.source_col = idx
                break

    missing = []
    if info.id_col < 0:
        missing.append("ID 列")
    if info.source_col < 0:
        missing.append("源文本列")
    if info.target_col < 0:
        missing.append("翻译列")
    if missing:
        info.error = f"未识别到{'、'.join(missing)}"
    return info


def parse_multi_sheet_localization(path: str, header_row: int = 0, enabled_languages: Iterable[str] | None = None) -> NormalizedLocalizationDataset:
    rows_by_sheet = _load_xlsx_rows(path)
    enabled = {str(item).strip().upper() for item in (enabled_languages or []) if str(item).strip()}
    sheet_items = []
    for sheet_name in rows_by_sheet:
        code = canonical_language_code(sheet_name)
        if not code:
            continue
        if enabled and code.upper() not in enabled and sheet_name.upper() not in enabled:
            continue
        sheet_items.append((code, LANGUAGE_SHEET_MAP.get(code, code), sheet_name))
    sheet_items.sort(key=lambda item: LANGUAGE_PRIORITY.index(item[0]) if item[0] in LANGUAGE_PRIORITY else 999)
    if len(sheet_items) < 1:
        raise FileReadError("未识别到多 Sheet 语言页结构，请确认 Sheet 名是否为 EN、DE、FR、RU、ES、JP、ID、KR、TH、IT、PT、TR、VN、ARB 等语言代码")

    sheet_infos: List[LanguageSheetInfo] = []
    entries: List[NormalizedLocalizationEntry] = []
    records_by_language: Dict[str, Dict[str, NormalizedLocalizationEntry]] = {}
    duplicate_ids: Dict[str, Dict[str, List[int]]] = {}
    invalid_rows: List[Tuple[str, int]] = []
    header_errors: List[str] = []

    for code, language_name, sheet_name in sheet_items:
        rows = rows_by_sheet.get(sheet_name, [])
        info = detect_language_sheet_columns(rows, sheet_name, header_row)
        info.language = code
        info.language_name = language_name
        sheet_infos.append(info)
        if info.error:
            header_errors.append(f"{sheet_name}：{info.error}")
            continue
        records: Dict[str, NormalizedLocalizationEntry] = {}
        id_rows: Dict[str, List[int]] = {}
        for row_index in range(info.data_start_row - 1, len(rows)):
            row = rows[row_index]
            if not any(str(cell).strip() for cell in row):
                continue
            text_id = row[info.id_col].strip() if info.id_col < len(row) else ""
            if not text_id:
                invalid_rows.append((sheet_name, row_index + 1))
                continue
            source_text = row[info.source_col].strip() if info.source_col < len(row) else ""
            translation = row[info.target_col].strip() if info.target_col < len(row) else ""
            column_name = info.headers[info.target_col] if info.target_col < len(info.headers) else "翻译"
            id_rows.setdefault(text_id, []).append(row_index + 1)
            if text_id not in records:
                entry = NormalizedLocalizationEntry(
                    text_id=text_id,
                    source_text=source_text,
                    language=code,
                    language_name=language_name,
                    translation=translation,
                    sheet_name=sheet_name,
                    row_number=row_index + 1,
                    column_name=column_name,
                )
                records[text_id] = entry
                entries.append(entry)
        records_by_language[code] = records
        dup = {text_id: rows for text_id, rows in id_rows.items() if len(rows) > 1}
        if dup:
            duplicate_ids[code] = dup

    # 基准语言优先取 EN，否则取 ID 数量最多的语言。
    base_language = ""
    for candidate in LANGUAGE_PRIORITY:
        if candidate in records_by_language:
            base_language = candidate
            break
    if not base_language and records_by_language:
        base_language = max(records_by_language, key=lambda lang: len(records_by_language[lang]))

    return NormalizedLocalizationDataset(
        structure="multi_sheet",
        path=path,
        sheet_infos=sheet_infos,
        entries=entries,
        records_by_language=records_by_language,
        duplicate_ids=duplicate_ids,
        invalid_rows=invalid_rows,
        header_errors=header_errors,
        base_language=base_language,
    )


def build_multi_sheet_preview(dataset: NormalizedLocalizationDataset) -> str:
    valid_infos = [info for info in dataset.sheet_infos if not info.error]
    language_codes = [info.language for info in valid_infos]
    id_counts = {lang: len(records) for lang, records in dataset.records_by_language.items()}
    first = valid_infos[0] if valid_infos else None
    lines = [
        "已识别结构：多 Sheet 语言页",
        f"已识别语言页：{', '.join(language_codes) if language_codes else '未识别'}",
    ]
    if first:
        def col_name(index: int) -> str:
            if index < 0:
                return "未识别"
            # A, B, C...
            n = index + 1
            s = ""
            while n:
                n, r = divmod(n - 1, 26)
                s = chr(65 + r) + s
            header = first.headers[index] if index < len(first.headers) else ""
            return f"{s}列 / {header or '空表头'}"
        lines.extend([
            f"ID 列：{col_name(first.id_col)}",
            f"源文本列：{col_name(first.source_col)}",
            f"翻译列：{col_name(first.target_col)}",
            f"数据起始行：第 {first.data_start_row} 行",
        ])
    lines.append(f"已识别文本 ID 数量：{len(dataset.text_ids)}")
    lines.append(f"已识别翻译条目数量：{len(dataset.entries)}")
    if id_counts:
        lines.append("各语言 ID 数量：" + "；".join(f"{k}={v}" for k, v in id_counts.items()))
    if dataset.header_errors:
        lines.append("表头异常：" + "；".join(dataset.header_errors[:8]))
    return "\n".join(lines)
