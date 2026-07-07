from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Iterable, List


@dataclass
class NormalizeOptions:
    ignore_spaces: bool = False
    ignore_newlines: bool = False
    ignore_case: bool = False


def normalize_text(value: str, options: NormalizeOptions) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    if options.ignore_newlines:
        text = text.replace("\n", "")

    if options.ignore_spaces:
        text = re.sub(r"\s+", "", text)
    else:
        text = "\n".join(line.strip() for line in text.split("\n"))

    if options.ignore_case:
        text = text.lower()

    return text


def normalize_line(value: str, options: NormalizeOptions) -> str:
    line = "" if value is None else str(value)
    line = line.strip()
    if options.ignore_spaces:
        line = re.sub(r"\s+", "", line)
    if options.ignore_case:
        line = line.lower()
    return line


def normalize_key(value: str, options: NormalizeOptions) -> str:
    key = "" if value is None else str(value)
    key = key.strip()
    if options.ignore_spaces:
        key = re.sub(r"\s+", "", key)
    if options.ignore_case:
        key = key.lower()
    return key


def normalize_value(value: str, options: NormalizeOptions) -> str:
    return normalize_text(value, options)


def clean_ignore_fields(raw_text: str, options: NormalizeOptions) -> List[str]:
    if not raw_text:
        return []
    items = [item.strip() for item in re.split(r"[,，;；\n]", raw_text) if item.strip()]
    return [normalize_key(item, options) for item in items]


def should_ignore_key(key: str, ignore_keys: Iterable[str], options: NormalizeOptions) -> bool:
    normalized = normalize_key(key, options)
    for ignore_key in ignore_keys:
        if not ignore_key:
            continue
        if normalized == ignore_key or ignore_key in normalized:
            return True
    return False
