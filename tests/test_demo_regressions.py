from __future__ import annotations

import collections
import os
import sys
import threading
import unittest
from pathlib import Path

from openpyxl import load_workbook


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from app.config_validator import (  # noqa: E402
    ConfigValidationOptions,
    detect_config_structure,
    validate_config_table,
)
from app.localization_checker import (  # noqa: E402
    LocalizationCheckOptions,
    check_localization_file,
)
from app.value_config_compare import (  # noqa: E402
    ValueCompareOptions,
    compare_value_reference_to_config,
)


DEMO_ROOT_TEXT = os.environ.get("QA_TOOLBOX_DEMO_ROOT", "").strip()
DEMO_ROOT = Path(DEMO_ROOT_TEXT) if DEMO_ROOT_TEXT else None


def _require_demo_root() -> Path:
    if DEMO_ROOT is None or not DEMO_ROOT.exists():
        raise unittest.SkipTest("请设置 QA_TOOLBOX_DEMO_ROOT 指向 QA测试工具盒_分享会演示素材")
    return DEMO_ROOT


def _report_rows(path: Path, sheet_index: int = 0):
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        return list(workbook.worksheets[sheet_index].iter_rows(min_row=2, values_only=True))
    finally:
        workbook.close()


class DemoRegressionTests(unittest.TestCase):
    def test_commercial_compare_matches_verified_reports(self):
        root = _require_demo_root() / "商业化配置比对范例"
        mall = root / "mall.csv"
        cases = [
            ("常规礼包.xlsx", "常规礼包差异比对结果.xlsx"),
            ("月卡周卡.xlsx", "月卡周卡差异比对结果.xlsx"),
        ]
        for source_name, report_name in cases:
            with self.subTest(source=source_name):
                diffs, _report = compare_value_reference_to_config(
                    str(mall),
                    str(root / source_name),
                    ValueCompareOptions(template_name="mall", enable_general_compare=True),
                )
                expected = {
                    (str(row[1] or ""), str(row[2] or ""), str(row[3] or ""), str(row[4] or ""), str(row[5] or ""))
                    for row in _report_rows(root / report_name)
                }
                actual = {
                    (item.diff_type, item.document_content, item.file_content, item.location, item.remark)
                    for item in diffs
                }
                self.assertEqual(expected, actual)

    def test_config_validation_matches_verified_reports(self):
        root = _require_demo_root() / "配置表校验范例"
        cases = [
            ("shop_basic.csv", "shop_basic配置表校验报告.xlsx", True),
            ("chessboard_stage.csv", "chessboard_stage配置表校验报告.xlsx", False),
        ]
        compound_rule = "field=restrict;allow_empty=true;group=none;sep=|;count=2;type=uint;allow_zero=true;forbid=0"
        for source_name, report_name, use_compound in cases:
            with self.subTest(source=source_name):
                source = root / source_name
                structure = detect_config_structure(str(source))
                options = ConfigValidationOptions(
                    header_row=structure["header_row"],
                    data_start_row=structure["data_start_row"],
                    key_column=structure["key_column"],
                    sheet_name=structure.get("sheet_name", ""),
                    description_row=structure.get("description_row", 0),
                    type_row=structure.get("type_row", 0),
                    check_type_fields=True,
                    check_compound_rules=use_compound,
                    compound_rules_text=compound_rule if use_compound else "",
                    ui_limit=100000,
                )
                issues, _table = validate_config_table(str(source), options)
                expected = {
                    tuple(str(value or "") for value in row[1:9])
                    for row in _report_rows(root / report_name)
                }
                actual = {
                    (
                        str(item.row_number or ""),
                        str(item.item_id or ""),
                        str(item.field_name or ""),
                        str(item.field_type or ""),
                        str(item.current_value or ""),
                        str(item.issue_type or ""),
                        str(item.remark or ""),
                        str(item.suggestion or ""),
                    )
                    for item in issues
                }
                self.assertEqual(expected, actual)

    def test_localization_checks_match_verified_reports(self):
        root = _require_demo_root() / "多语言检查范例"
        cases = [
            ("单Sheet语言.xlsx", "单Sheet语言文本质量检查报告.xlsx"),
            ("多Sheet语言.xlsx", "多Sheet语言文本质量检查报告.xlsx"),
        ]
        for source_name, report_name in cases:
            with self.subTest(source=source_name):
                issues, _table, _columns = check_localization_file(
                    str(root / source_name),
                    LocalizationCheckOptions(ui_limit=100000),
                    threading.Event(),
                )
                expected_rows = _report_rows(root / report_name, sheet_index=1)
                expected = {
                    (
                        str(row[1] or ""),
                        str(row[2] or ""),
                        str(row[3] or ""),
                        str(row[5] or ""),
                        str(row[8] or ""),
                        str(row[9] or ""),
                        str(row[11] or ""),
                    )
                    for row in expected_rows
                }
                actual = {
                    (
                        str(item.sheet_name or ""),
                        str(item.row_number or ""),
                        str(item.item_id or ""),
                        str(item.language or ""),
                        str(item.issue_type or ""),
                        str(item.issue_level or ""),
                        str(item.remark or ""),
                    )
                    for item in issues
                }
                self.assertEqual(expected, actual)
                self.assertEqual(
                    collections.Counter(row[8] for row in expected_rows),
                    collections.Counter(item.issue_type for item in issues),
                )


if __name__ == "__main__":
    unittest.main()
