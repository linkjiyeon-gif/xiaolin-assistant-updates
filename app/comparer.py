from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Tuple

from .file_readers import ParsedContent
from .normalizer import (
    NormalizeOptions,
    normalize_key,
    normalize_line,
    normalize_text,
    normalize_value,
    should_ignore_key,
)
from .planning_rules import PlanningItem, extract_planning_items, normalize_business_value
from .translation_compare import (
    DEFAULT_CONFIG_KEY_CANDIDATES,
    DEFAULT_SOURCE_KEY_CANDIDATES,
    TranslationCompareOptions,
    first_existing_column,
    normalize_translation_text,
    parse_language_mappings,
    parse_translation_table,
    safe_int,
)


@dataclass
class CompareOptions:
    mode: str
    normalize: NormalizeOptions
    ignore_fields: List[str]
    translation_options: Optional[TranslationCompareOptions] = None


@dataclass
class DiffItem:
    index: int
    diff_type: str
    document_content: str
    file_content: str
    location: str
    remark: str = ""
    item_id: str = ""
    language: str = ""
    source_row: str = ""
    config_row: str = ""


class ContentComparer:
    def compare(self, actual: ParsedContent, document: ParsedContent, options: CompareOptions) -> List[DiffItem]:
        mode = options.mode
        if mode == "全文文本比对":
            return self.compare_full_text(actual, document, options)
        if mode in {"按行比对", "普通文本 / 文件比对"}:
            return self.compare_by_lines(actual, document, options)
        if mode == "按字段名比对":
            return self.compare_by_fields(actual, document, options, check_value=True, include_extra=True)
        if mode == "只检查文档字段存在":
            return self.compare_document_fields_exist(actual, document, options)
        if mode == "策划文档 vs 配置文件":
            return self.compare_planning_document_to_config(actual, document, options)
        if mode == "翻译源文件 vs 游戏翻译配置表":
            return self.compare_translation_source_to_config(actual, document, options)
        return self.compare_by_lines(actual, document, options)

    def compare_full_text(self, actual: ParsedContent, document: ParsedContent, options: CompareOptions) -> List[DiffItem]:
        actual_text = normalize_text(actual.text, options.normalize)
        document_text = normalize_text(document.text, options.normalize)
        if actual_text == document_text:
            return []

        diffs = [
            DiffItem(
                index=1,
                diff_type="值不一致",
                document_content=f"文档全文长度：{len(document_text)}",
                file_content=f"文件全文长度：{len(actual_text)}",
                location="全文",
                remark="全文内容不一致，下面列出行级差异",
            )
        ]
        diffs.extend(self.compare_by_lines(actual, document, options, start_index=2))
        return self._reindex(diffs)

    def compare_by_lines(
        self,
        actual: ParsedContent,
        document: ParsedContent,
        options: CompareOptions,
        start_index: int = 1,
    ) -> List[DiffItem]:
        actual_pairs = self._normalized_line_pairs(actual.lines, options.normalize)
        document_pairs = self._normalized_line_pairs(document.lines, options.normalize)

        actual_counter = Counter(key for key, _line, _num in actual_pairs)
        document_counter = Counter(key for key, _line, _num in document_pairs)

        actual_examples = self._line_examples(actual_pairs)
        document_examples = self._line_examples(document_pairs)

        diffs: List[DiffItem] = []
        current_index = start_index

        for key in sorted(document_counter.keys() - actual_counter.keys()):
            line, line_num = document_examples[key][0]
            diffs.append(
                DiffItem(current_index, "缺失", line, "", f"文档第 {line_num} 行", "文档有但文件没有")
            )
            current_index += 1

        for key in sorted(actual_counter.keys() - document_counter.keys()):
            line, line_num = actual_examples[key][0]
            diffs.append(
                DiffItem(current_index, "多余", "", line, f"文件第 {line_num} 行", "文件有但文档没有")
            )
            current_index += 1

        common_keys = document_counter.keys() & actual_counter.keys()
        for key in sorted(common_keys):
            delta = document_counter[key] - actual_counter[key]
            if delta <= 0:
                continue
            line, line_num = document_examples[key][0]
            diffs.append(
                DiffItem(current_index, "缺失", line, "", f"文档第 {line_num} 行", f"相同行数量少 {delta} 条")
            )
            current_index += 1

        for key in sorted(common_keys):
            delta = actual_counter[key] - document_counter[key]
            if delta <= 0:
                continue
            line, line_num = actual_examples[key][0]
            diffs.append(
                DiffItem(current_index, "多余", "", line, f"文件第 {line_num} 行", f"相同行数量多 {delta} 条")
            )
            current_index += 1

        return self._reindex(diffs)

    def compare_by_fields(
        self,
        actual: ParsedContent,
        document: ParsedContent,
        options: CompareOptions,
        check_value: bool = True,
        include_extra: bool = True,
    ) -> List[DiffItem]:
        actual_fields = self._filter_and_normalize_fields(actual.fields, options)
        document_fields = self._filter_and_normalize_fields(document.fields, options)

        diffs: List[DiffItem] = []
        index = 1

        actual_keys = set(actual_fields.keys())
        document_keys = set(document_fields.keys())

        for key in sorted(document_keys - actual_keys):
            original_key, value = document_fields[key]
            diffs.append(
                DiffItem(index, "缺失", f"{original_key}: {value}", "", original_key, "文档字段在文件中不存在")
            )
            index += 1

        if include_extra:
            for key in sorted(actual_keys - document_keys):
                original_key, value = actual_fields[key]
                diffs.append(
                    DiffItem(index, "多余", "", f"{original_key}: {value}", original_key, "文件字段在文档中不存在")
                )
                index += 1

        if check_value:
            for key in sorted(document_keys & actual_keys):
                doc_key, doc_value = document_fields[key]
                actual_key, actual_value = actual_fields[key]
                if normalize_value(doc_value, options.normalize) != normalize_value(actual_value, options.normalize):
                    diffs.append(
                        DiffItem(
                            index,
                            "值不一致",
                            f"{doc_key}: {doc_value}",
                            f"{actual_key}: {actual_value}",
                            doc_key,
                            "字段名一致但值不一致",
                        )
                    )
                    index += 1

        return diffs

    def compare_document_fields_exist(
        self,
        actual: ParsedContent,
        document: ParsedContent,
        options: CompareOptions,
    ) -> List[DiffItem]:
        document_fields = self._filter_and_normalize_fields(document.fields, options)
        actual_fields = self._filter_and_normalize_fields(actual.fields, options)
        actual_keys = set(actual_fields.keys())
        actual_text = normalize_text(actual.text, options.normalize)

        diffs: List[DiffItem] = []
        index = 1
        for key, (original_key, value) in sorted(document_fields.items()):
            key_exists = key in actual_keys or normalize_key(original_key, options.normalize) in actual_text
            if not key_exists:
                diffs.append(
                    DiffItem(index, "缺失", f"{original_key}: {value}", "", original_key, "只检查字段是否存在，不校验字段值")
                )
                index += 1
        return diffs



    def compare_translation_source_to_config(
        self,
        actual: ParsedContent,
        document: ParsedContent,
        options: CompareOptions,
    ) -> List[DiffItem]:
        """通用模式：比对翻译源文件与游戏实际翻译配置表。

        约定：左侧“待比对文件”为游戏配置表，右侧“参考文档”为翻译源文件。
        通过主键 ID + 语言列映射逐项比对，适配不同项目的字段命名。
        """
        trans_options = options.translation_options or TranslationCompareOptions()

        source_header_row = safe_int(trans_options.source_header_row, 1)
        source_data_start_row = safe_int(trans_options.source_data_start_row, 2)
        config_header_row = safe_int(trans_options.config_real_header_row, 0) or safe_int(trans_options.config_header_row, 1)
        config_data_start_row = safe_int(trans_options.config_data_start_row, 2)

        source_table = parse_translation_table(
            path=document.path,
            sheet_name=trans_options.source_sheet,
            header_row=source_header_row,
            data_start_row=source_data_start_row,
            key_column=trans_options.source_key_column,
            key_candidates=DEFAULT_SOURCE_KEY_CANDIDATES,
        )
        config_table = parse_translation_table(
            path=actual.path,
            sheet_name=trans_options.config_sheet,
            header_row=config_header_row,
            data_start_row=config_data_start_row,
            key_column=trans_options.config_key_column,
            key_candidates=DEFAULT_CONFIG_KEY_CANDIDATES,
        )

        mappings = parse_language_mappings(trans_options.language_mapping_text)
        enabled = {name.strip() for name in (trans_options.enabled_languages or []) if name.strip()}
        if enabled:
            mappings = [mapping for mapping in mappings if mapping.name in enabled]

        diffs: List[DiffItem] = []
        index = 1

        for row_number in source_table.invalid_rows:
            diffs.append(self._translation_diff(index, "无效 ID", "", "", row_number, "", "", "", "源文件存在非空行，但主键 ID 为空"))
            index += 1
        for row_number in config_table.invalid_rows:
            diffs.append(self._translation_diff(index, "无效 ID", "", "", "", row_number, "", "", "配置表存在非空行，但主键 ID 为空"))
            index += 1

        numeric_id_mode = self._mostly_numeric_ids(source_table.records.keys())

        for text_id, rows in sorted(source_table.duplicate_ids.items()):
            diffs.append(self._translation_diff(index, "重复 ID", text_id, "", ",".join(map(str, rows)), "", "", "", "源文件存在重复 ID"))
            index += 1
        for text_id, rows in sorted(config_table.duplicate_ids.items()):
            if numeric_id_mode and not str(text_id).strip().isdigit():
                continue
            diffs.append(self._translation_diff(index, "重复 ID", text_id, "", "", ",".join(map(str, rows)), "", "", "配置表存在重复 ID"))
            index += 1

        # 如果源文件的 ID 明显是纯数字，自动把配置表中的非数字“ID”视为格式异常行。
        # 这能兼容部分游戏 CSV 中未规范转义的换行文本，避免续行被误判成配置多余 ID。
        if numeric_id_mode:
            bad_config_ids = [text_id for text_id in config_table.records.keys() if not str(text_id).strip().isdigit()]
            for text_id in sorted(bad_config_ids):
                config_record = config_table.records.pop(text_id)
                diffs.append(self._translation_diff(index, "无效 ID", text_id, "", "", config_record.row_number, "", "", "配置表主键格式异常：源文件 ID 基本为数字，但该配置表 ID 不是数字"))
                index += 1

        source_ids = set(source_table.records.keys())
        config_ids = set(config_table.records.keys())

        for text_id in sorted(source_ids - config_ids):
            source_record = source_table.records[text_id]
            diffs.append(self._translation_diff(index, "ID 缺失", text_id, "", source_record.row_number, "", "", "", "源文件有该 ID，配置表没有"))
            index += 1

        if trans_options.check_extra_ids:
            for text_id in sorted(config_ids - source_ids):
                config_record = config_table.records[text_id]
                diffs.append(self._translation_diff(index, "ID 多余", text_id, "", "", config_record.row_number, "", "", "配置表有该 ID，源文件没有"))
                index += 1

        usable_mappings = []
        for mapping in mappings:
            source_col = first_existing_column(source_table.headers, mapping.source_candidates)
            config_col = first_existing_column(config_table.headers, mapping.config_candidates)
            if source_col < 0 or config_col < 0:
                missing_side = []
                if source_col < 0:
                    missing_side.append("源文件")
                if config_col < 0:
                    missing_side.append("配置表")
                diffs.append(self._translation_diff(
                    index,
                    "缺少语言列",
                    "",
                    mapping.name,
                    "",
                    "",
                    " / ".join(mapping.source_candidates),
                    " / ".join(mapping.config_candidates),
                    f"{mapping.name} 在{'、'.join(missing_side)}中未识别到语言列",
                ))
                index += 1
                continue
            usable_mappings.append((mapping, source_col, config_col))

        for text_id in sorted(source_ids & config_ids):
            source_record = source_table.records[text_id]
            config_record = config_table.records[text_id]
            for mapping, source_col, config_col in usable_mappings:
                source_text = source_record.values.get(source_col, "")
                config_text = config_record.values.get(config_col, "")
                source_normalized = normalize_translation_text(source_text, trans_options)
                config_normalized = normalize_translation_text(config_text, trans_options)

                if not source_normalized and not config_normalized:
                    continue
                if source_normalized and not config_normalized:
                    diffs.append(self._translation_diff(index, "翻译缺失", text_id, mapping.name, source_record.row_number, config_record.row_number, source_text, config_text, "源文件有值，配置表为空"))
                    index += 1
                    continue
                if not source_normalized and config_normalized:
                    diffs.append(self._translation_diff(index, "配置多余", text_id, mapping.name, source_record.row_number, config_record.row_number, source_text, config_text, "源文件为空，配置表有值"))
                    index += 1
                    continue
                if source_normalized != config_normalized:
                    diffs.append(self._translation_diff(index, "文本不一致", text_id, mapping.name, source_record.row_number, config_record.row_number, source_text, config_text, "同 ID、同语言下文本不同"))
                    index += 1

        return self._reindex(diffs)

    @staticmethod
    def _mostly_numeric_ids(ids) -> bool:
        values = [str(text_id).strip() for text_id in ids if str(text_id).strip()]
        if len(values) < 5:
            return False
        numeric_count = sum(1 for value in values if value.isdigit())
        return numeric_count / len(values) >= 0.8

    @staticmethod
    def _translation_diff(index, diff_type, text_id, language, source_row, config_row, source_text, config_text, remark):
        source_row_text = "" if source_row is None else str(source_row)
        config_row_text = "" if config_row is None else str(config_row)
        location_parts = []
        if text_id:
            location_parts.append(f"ID={text_id}")
        if language:
            location_parts.append(f"语言={language}")
        if source_row_text:
            location_parts.append(f"源行={source_row_text}")
        if config_row_text:
            location_parts.append(f"配置行={config_row_text}")
        return DiffItem(
            index=index,
            diff_type=diff_type,
            document_content="" if source_text is None else str(source_text),
            file_content="" if config_text is None else str(config_text),
            location=" / ".join(location_parts),
            remark=remark,
            item_id="" if text_id is None else str(text_id),
            language="" if language is None else str(language),
            source_row=source_row_text,
            config_row=config_row_text,
        )

    def compare_planning_document_to_config(
        self,
        actual: ParsedContent,
        document: ParsedContent,
        options: CompareOptions,
    ) -> List[DiffItem]:
        """专用模式：比对策划/配置说明与实际配置/导出文件。

        该模式会把两边内容抽取成“配置值、奖励内容、档位、开关状态”等业务项，
        再按类别 + 业务键进行缺失、多余、值不一致判断。
        """
        actual_items = extract_planning_items(actual, options.normalize, options.ignore_fields)
        document_items = extract_planning_items(document, options.normalize, options.ignore_fields)

        actual_map = self._planning_item_map(actual_items, options)
        document_map = self._planning_item_map(document_items, options)

        diffs: List[DiffItem] = []
        index = 1

        actual_keys = set(actual_map.keys())
        document_keys = set(document_map.keys())

        for item_key in sorted(document_keys - actual_keys):
            doc_item = document_map[item_key]
            diffs.append(
                DiffItem(
                    index=index,
                    diff_type="缺失",
                    document_content=self._format_planning_item(doc_item),
                    file_content="",
                    location=doc_item.location,
                    remark=f"{doc_item.category}：文档有要求，但实际配置中未找到",
                )
            )
            index += 1

        for item_key in sorted(actual_keys - document_keys):
            actual_item = actual_map[item_key]
            diffs.append(
                DiffItem(
                    index=index,
                    diff_type="多余",
                    document_content="",
                    file_content=self._format_planning_item(actual_item),
                    location=actual_item.location,
                    remark=f"{actual_item.category}：实际配置存在，但文档中未找到对应说明",
                )
            )
            index += 1

        for item_key in sorted(document_keys & actual_keys):
            doc_item = document_map[item_key]
            actual_item = actual_map[item_key]
            doc_value = self._normalize_planning_value(doc_item.value, options)
            actual_value = self._normalize_planning_value(actual_item.value, options)
            if doc_value != actual_value:
                diffs.append(
                    DiffItem(
                        index=index,
                        diff_type="值不一致",
                        document_content=self._format_planning_item(doc_item),
                        file_content=self._format_planning_item(actual_item),
                        location=f"文档：{doc_item.location} / 文件：{actual_item.location}",
                        remark=f"{doc_item.category}：同一配置项的值不一致",
                    )
                )
                index += 1

        return self._reindex(diffs)

    def _planning_item_map(self, items: List[PlanningItem], options: CompareOptions) -> Dict[Tuple[str, str], PlanningItem]:
        result: Dict[Tuple[str, str], PlanningItem] = {}
        for item in items:
            if should_ignore_key(item.key, options.ignore_fields, options.normalize):
                continue
            normalized_key = normalize_key(item.key, options.normalize)
            if not normalized_key:
                continue
            map_key = (item.category, normalized_key)
            # 同一业务键重复出现时，优先保留信息更完整的值。
            if map_key not in result or len(str(item.value)) > len(str(result[map_key].value)):
                result[map_key] = item
        return result

    @staticmethod
    def _format_planning_item(item: PlanningItem) -> str:
        return f"[{item.category}] {item.key}: {item.value}"

    @staticmethod
    def _normalize_planning_value(value: str, options: CompareOptions) -> str:
        value = normalize_business_value(value)
        return normalize_text(value, options.normalize)

    def _filter_and_normalize_fields(
        self, fields: Dict[str, str], options: CompareOptions
    ) -> Dict[str, Tuple[str, str]]:
        result: Dict[str, Tuple[str, str]] = {}
        for key, value in fields.items():
            if should_ignore_key(key, options.ignore_fields, options.normalize):
                continue
            normalized_key = normalize_key(key, options.normalize)
            if not normalized_key:
                continue
            result[normalized_key] = (key, "" if value is None else str(value))
        return result

    @staticmethod
    def _normalized_line_pairs(lines: Iterable[str], options: NormalizeOptions) -> List[Tuple[str, str, int]]:
        pairs: List[Tuple[str, str, int]] = []
        for line_num, line in enumerate(lines, start=1):
            normalized = normalize_line(line, options)
            if not normalized:
                continue
            pairs.append((normalized, line, line_num))
        return pairs

    @staticmethod
    def _line_examples(pairs: Iterable[Tuple[str, str, int]]):
        result = defaultdict(list)
        for key, line, line_num in pairs:
            result[key].append((line, line_num))
        return result

    @staticmethod
    def _reindex(diffs: List[DiffItem]) -> List[DiffItem]:
        for index, item in enumerate(diffs, start=1):
            item.index = index
        return diffs
