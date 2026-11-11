"""网络弱网目标 IP 解析与 WinDivert 过滤规则生成。

目标 IP 输入支持多个，用逗号、分号、空格或换行分隔，每一项可以是：
- 单个 IP：8.8.8.8 / 2001:db8::1
- 网段（CIDR）：10.0.0.0/24
- 范围：192.168.1.10-192.168.1.20
"""

import ipaddress
import re

_SEPARATORS = re.compile(r"[\s,，;；、]+")


def _split_items(text: str):
    return [item.strip() for item in _SEPARATORS.split(text or "") if item.strip()]


def _parse_item(item: str):
    """Return (first_ip, last_ip) for one entry, or raise ValueError."""
    if "/" in item:
        network = ipaddress.ip_network(item, strict=False)
        return network.network_address, network.broadcast_address
    if "-" in item:
        start_text, end_text = (part.strip() for part in item.split("-", 1))
        first, last = ipaddress.ip_address(start_text), ipaddress.ip_address(end_text)
        if first.version != last.version:
            raise ValueError(item)
        return (last, first) if first > last else (first, last)
    address = ipaddress.ip_address(item)
    return address, address


def parse_ip_targets(text: str):
    """Parse target text into a de-duplicated list of (first_ip, last_ip) tuples.

    Raises ValueError with the offending item when an entry is invalid.
    """
    targets = []
    seen = set()
    for item in _split_items(text):
        try:
            key = _parse_item(item)
        except ValueError:
            raise ValueError(f"目标 IP 格式不正确：{item}\n支持单个 IP、网段（如 10.0.0.0/24）或范围（如 1.1.1.1-1.1.1.20），多个之间用逗号分隔") from None
        if key not in seen:
            seen.add(key)
            targets.append(key)
    return targets


# 实测（含 !impostor 保护、同时选协议和方向）：单条规则约 126 个单 IP 或 63 个网段/范围。
# 网段/范围在规则里占两倍长度，这里按“单 IP 计 1、网段/范围计 2”估算，留一点余量。
TARGET_RULE_BUDGET = 120


def describe_ip_targets(text: str, max_chips: int = 5) -> dict:
    """Live summary for the target IP input.

    Returns a dict with:
      status  "empty" | "ok" | "warning" | "error"
      message the text for the result row
      chips   up to max_chips display names of recognised targets
      more    how many recognised targets are not shown as chips
    """
    items = _split_items(text)
    if not items:
        return {"status": "empty", "message": "留空 = 影响本机全部流量；可粘贴多行，每行一个或用逗号分隔",
                "chips": [], "more": 0}
    targets, seen, invalid, names = [], set(), [], []
    for item in items:
        try:
            key = _parse_item(item)
        except ValueError:
            invalid.append(item)
            continue
        if key not in seen:
            seen.add(key)
            targets.append(key)
            first, last = key
            if "/" in item:
                names.append(str(ipaddress.ip_network(item, strict=False)))
            else:
                names.append(str(first) if first == last else f"{first}-{last}")
    singles = sum(1 for first, last in targets if first == last)
    ranges = len(targets) - singles
    chips = names[:max_chips]
    more = len(names) - len(chips)
    if invalid:
        shown = "、".join(invalid[:3]) + (" 等" if len(invalid) > 3 else "")
        return {"status": "error", "message": f"⚠ {len(invalid)} 项格式不正确：{shown}", "chips": chips, "more": more}
    if singles + 2 * ranges > TARGET_RULE_BUDGET:
        return {"status": "warning",
                "message": f"⚠ 已识别 {len(targets)} 个目标，可能超过上限（约 120 个 IP 或 60 个网段），建议合并为网段",
                "chips": chips, "more": more}
    # 文案尽量短，给后面的 IP 标签留出位置。
    detail = f"（IP {singles} · 网段 {ranges}）" if singles and ranges else ("（网段/范围）" if ranges else "")
    return {"status": "ok", "message": f"✓ 已识别 {len(targets)} 个{detail}", "chips": chips, "more": more}


def summarize_ip_targets(text: str) -> str:
    """Short description of the targets for the run log."""
    targets = parse_ip_targets(text)
    if not targets:
        return "全部 IP"
    names = [str(first) if first == last else f"{first}-{last}" for first, last in targets]
    shown = "、".join(names[:5]) + (" 等" if len(names) > 5 else "")
    return f"{len(names)} 个（{shown}）"


def build_ip_condition(targets) -> str:
    """Build a WinDivert condition matching packets to or from any target."""
    clauses = []
    for first, last in targets:
        prefix = "ip" if first.version == 4 else "ipv6"
        if first == last:
            clauses.append(f"{prefix}.SrcAddr == {first} or {prefix}.DstAddr == {first}")
        else:
            clauses.append(
                f"({prefix}.SrcAddr >= {first} and {prefix}.SrcAddr <= {last})"
                f" or ({prefix}.DstAddr >= {first} and {prefix}.DstAddr <= {last})"
            )
    return "(" + " or ".join(clauses) + ")" if clauses else ""
