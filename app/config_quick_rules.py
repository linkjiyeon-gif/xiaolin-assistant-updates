from __future__ import annotations

# 该模块用于承载配置表校验的快捷规则入口说明；核心校验逻辑在 app.config_validator 中执行。
# 保留独立模块，方便后续继续拆分“自然语言规则解析”和“快捷模板生成”。

from .config_rule_templates import QUICK_TEMPLATES, TEMPLATE_HELP, TYPE_TOKENS

__all__ = ["QUICK_TEMPLATES", "TEMPLATE_HELP", "TYPE_TOKENS"]
