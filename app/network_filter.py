"""网络弱网目标 IP 解析与 WinDivert 过滤规则生成。

目标 IP 输入支持多个，用逗号、分号、空格或换行分隔，每一项可以是：
- 单个 IP：8.8.8.8 / 2001:db8::1
- 网段（CIDR）：10.0.0.0/24
- 范围：192.168.1.10-192.168.1.20
"""

import ipaddress
import re

_SEPARATORS = re.compile(r"[\s,，;；、]+")


def parse_ip_targets(text: str):
    """Parse target text into a de-duplicated list of (first_ip, last_ip) tuples.

    Raises ValueError with the offending item when an entry is invalid.
    """
    targets = []
    seen = set()
    for item in _SEPARATORS.split(text or ""):
        item = item.strip()
        if not item:
            continue
        try:
            if "/" in item:
                network = ipaddress.ip_network(item, strict=False)
                first, last = network.network_address, network.broadcast_address
            elif "-" in item:
                start_text, end_text = (part.strip() for part in item.split("-", 1))
                first, last = ipaddress.ip_address(start_text), ipaddress.ip_address(end_text)
                if first.version != last.version:
                    raise ValueError
                if first > last:
                    first, last = last, first
            else:
                first = last = ipaddress.ip_address(item)
        except ValueError:
            raise ValueError(f"目标 IP 格式不正确：{item}\n支持单个 IP、网段（如 10.0.0.0/24）或范围（如 1.1.1.1-1.1.1.20），多个之间用逗号分隔") from None
        key = (first, last)
        if key not in seen:
            seen.add(key)
            targets.append(key)
    return targets


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
