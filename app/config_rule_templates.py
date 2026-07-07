from __future__ import annotations

TYPE_TOKENS = {"uint", "int", "string", "uint[]", "int[]", "string[]", "float", "bool"}

QUICK_TEMPLATES = [
    "必填字段",
    "允许为空或 x|y",
    "必须为 x|y",
    "禁止单值 0",
    "reward 四参数校验",
    "reward 三参数校验",
    "restrict 条件格式校验",
    "condition 条件字段校验",
    "按类型行自动校验",
]

TEMPLATE_HELP = {
    "必填字段": "字段不能为空，可选择是否把 0 / -1 当作无效值。",
    "允许为空或 x|y": "字段可空；不为空时必须为两个整数，用 | 分隔。",
    "必须为 x|y": "字段不能为空，必须为两个整数，用 | 分隔。",
    "禁止单值 0": "字段不能只填写 0。",
    "reward 四参数校验": "reward 多组用 | 分隔；每组用 * 分隔；每组必须 4 个整数参数。",
    "reward 三参数校验": "reward 多组用 | 分隔；每组用 * 分隔；每组必须 3 个整数参数。",
    "restrict 条件格式校验": "restrict 可空；不为空时必须为 x|y；禁止单值 0。",
    "condition 条件字段校验": "condition 可空；不为空时必须为 2 或 3 个整数参数。",
    "按类型行自动校验": "根据类型行自动检查 uint / int / float / bool / 数组字段。",
}
