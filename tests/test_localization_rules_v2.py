from __future__ import annotations

import collections
import tempfile
import threading
import unittest
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook

from app.localization_checker import (
    LocalizationCheckOptions,
    check_localization_file,
    export_localization_issues_to_excel,
)


class LocalizationRulesV2Tests(unittest.TestCase):
    def _run(self, rows, *, aggregate=True):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        path = Path(temp.name) / "rules.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Text"
        sheet.append(["#tid", "CN", "EN", "DE"])
        for row in rows:
            sheet.append(row)
        workbook.save(path)
        issues, _table, _columns = check_localization_file(
            str(path),
            LocalizationCheckOptions(
                table_structure="单 Sheet 多语言列",
                source_language_column="CN",
                enabled_language_columns=["EN", "DE"],
                check_chinese=False,
                check_unfinished=False,
                check_spelling=False,
                check_terms=False,
                check_duplicate_translation=False,
                aggregate_empty_language_columns=aggregate,
                ui_limit=100000,
            ),
            threading.Event(),
        )
        return path, issues

    def test_numeric_equivalence_range_order_percent_and_real_change(self):
        _path, issues = self._run([
            ["level", "城堡3级", "Castle St. 3", "Schloss Niv. 3"],
            ["range", "开放1-30关", "Stages 1-30", "Stufen 1–30"],
            ["top", "地区排名前3", "Top-3 in the region", "Top-3 der Region"],
            ["order", "持续8小时提升50%", "Increase 50 percent for 8 hours", "8 Stunden, 50 Prozent"],
            ["decimal", "减少10.5%", "Reduce by 10.5%", "Um 10,5% reduzieren"],
            ["changed", "目标2", "Goal 3", "Ziel 2"],
        ], aggregate=False)
        by_id = collections.defaultdict(list)
        for item in issues:
            by_id[item.item_id].append(item)
        for text_id in ("level", "range", "top", "order", "decimal"):
            self.assertFalse(any(item.rule_name == "数值" for item in by_id[text_id]), text_id)
        self.assertIn("数值内容不一致", [item.issue_type for item in by_id["changed"]])
        self.assertFalse(any(item.rule_name == "数值" and item.language.startswith("DE") for item in by_id["changed"]))

    def test_date_cells_and_serial_text_are_not_clear_numeric_errors(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        path = Path(temp.name) / "dates.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Text"
        sheet.append(["#tid", "CN", "EN", "DE"])
        sheet.append(["date", "三月二十九日", datetime(2026, 3, 29), "46110"])
        sheet["C2"].number_format = "mmmm d"
        workbook.save(path)
        issues, _table, _columns = check_localization_file(
            str(path),
            LocalizationCheckOptions(
                table_structure="单 Sheet 多语言列",
                source_language_column="CN",
                enabled_language_columns=["EN", "DE"],
                check_chinese=False,
                check_unfinished=False,
                check_spelling=False,
                check_terms=False,
                check_duplicate_translation=False,
                aggregate_empty_language_columns=False,
            ),
            threading.Event(),
        )
        self.assertFalse(any(item.language.startswith("EN") and item.rule_name == "数值" for item in issues))
        de_numbers = [item for item in issues if item.language.startswith("DE") and item.rule_name == "数值"]
        self.assertEqual(1, len(de_numbers))
        self.assertEqual("需人工确认", de_numbers[0].result_category)
        self.assertFalse(de_numbers[0].count_in_error_stats)

    def test_context_symbols_tags_source_dedupe_and_coverage(self):
        path, issues = self._run([
            ["bracket", "活动名称", "[I Scream]", "[Eis]"],
            ["slash", "他/她", "he or she", "er oder sie"],
            ["percent", "完成50%", "Complete 50 percent", "Zu 50 Prozent abschließen"],
            ["tag_attr", "[color=#f53d39]文本[/color]", "[color=#CFBFA3]Text[/color]", "[color=#f53d39]Text[/color]"],
            ["target_bad", "[b]文本[/b]", "[color=#ffff", "[b]Text[/b]"],
            ["source_bad", "[color=#d16316]源文本", "Text", "Text"],
            ["partial_empty", "中文", "", "Deutsch"],
        ])
        by_id = collections.defaultdict(list)
        for item in issues:
            by_id[item.item_id].append(item)
        self.assertFalse(any("特殊符号" in item.issue_type or "标签" in item.issue_type for item in by_id["bracket"]))
        self.assertFalse(any("特殊符号" in item.issue_type for item in by_id["slash"]))
        self.assertFalse(any("特殊符号" in item.issue_type or item.rule_name == "数值" for item in by_id["percent"]))
        attr = [item for item in by_id["tag_attr"] if item.issue_type == "标签属性差异"]
        self.assertEqual(1, len(attr))
        self.assertEqual("需人工确认", attr[0].result_category)
        self.assertIn("目标文本标签结构异常", [item.issue_type for item in by_id["target_bad"]])
        source_problems = [item for item in by_id["source_bad"] if item.result_category == "源文本问题"]
        self.assertEqual(1, len(source_problems))
        self.assertFalse(any(item.language.startswith("EN") or item.language.startswith("DE") for item in source_problems))
        self.assertIn("空翻译", [item.issue_type for item in by_id["partial_empty"]])

        coverage = [item for item in issues if item.result_category == "语言覆盖范围"]
        self.assertEqual(2, len(coverage))
        self.assertTrue(all(not item.count_in_error_stats for item in coverage))

        report_path = path.with_name("report.xlsx")
        export_localization_issues_to_excel(issues, str(report_path))
        report = load_workbook(report_path, read_only=False, data_only=True)
        try:
            self.assertEqual(
                ["汇总说明", "明确问题", "源文本问题", "需人工确认", "语言覆盖范围", "忽略项或白名单命中", "全部明细", "规则说明"],
                report.sheetnames,
            )
            for sheet_name in report.sheetnames:
                self.assertTrue(report[sheet_name].auto_filter.ref)
        finally:
            report.close()


if __name__ == "__main__":
    unittest.main()
