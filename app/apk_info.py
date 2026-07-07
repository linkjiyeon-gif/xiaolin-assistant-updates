from __future__ import annotations

import hashlib
import os
import re
import shutil
import struct
import subprocess
import zipfile
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple


@dataclass
class APKInfo:
    file_name: str = ""
    file_size: str = ""
    md5: str = ""
    sha1: str = ""
    package: str = ""
    version_name: str = ""
    version_code: str = ""
    min_sdk: str = ""
    target_sdk: str = ""
    permissions: List[str] = field(default_factory=list)
    abis: List[str] = field(default_factory=list)
    debuggable: str = ""
    signature_summary: str = ""
    parse_note: str = ""


def _format_size(size: int) -> str:
    value = float(size)
    for unit in ["B", "KB", "MB", "GB"]:
        if value < 1024 or unit == "GB":
            return f"{value:.2f} {unit}" if unit != "B" else f"{int(value)} B"
        value /= 1024
    return f"{size} B"


def _hash_file(path: str) -> Tuple[str, str]:
    md5 = hashlib.md5()
    sha1 = hashlib.sha1()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            md5.update(chunk)
            sha1.update(chunk)
    return md5.hexdigest(), sha1.hexdigest()


def _run_tool(cmd: List[str], timeout: int = 20) -> str:
    completed = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="ignore",
        timeout=timeout,
        creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
    )
    return (completed.stdout or "") + ("\n" + completed.stderr if completed.stderr else "")


def _parse_aapt_badging(path: str) -> Dict[str, object]:
    aapt = shutil.which("aapt") or shutil.which("aapt.exe")
    if not aapt:
        return {}
    try:
        text = _run_tool([aapt, "dump", "badging", path], timeout=25)
    except Exception:
        return {}
    data: Dict[str, object] = {"permissions": [], "abis": []}
    pkg = re.search(r"package: name='([^']*)' versionCode='([^']*)' versionName='([^']*)'", text)
    if pkg:
        data["package"] = pkg.group(1)
        data["version_code"] = pkg.group(2)
        data["version_name"] = pkg.group(3)
    sdk = re.search(r"sdkVersion:'([^']*)'", text)
    target = re.search(r"targetSdkVersion:'([^']*)'", text)
    if sdk:
        data["min_sdk"] = sdk.group(1)
    if target:
        data["target_sdk"] = target.group(1)
    perms = re.findall(r"uses-permission: name='([^']*)'", text)
    data["permissions"] = sorted(set(perms))
    native = re.search(r"native-code:\s*(.*)", text)
    if native:
        data["abis"] = re.findall(r"'([^']*)'", native.group(1))
    if "application-debuggable" in text:
        data["debuggable"] = "是"
    return data


class _AXMLParser:
    UTF8_FLAG = 0x00000100
    RES_STRING_POOL_TYPE = 0x0001
    RES_XML_START_ELEMENT_TYPE = 0x0102
    NO_INDEX = 0xFFFFFFFF

    TYPE_STRING = 0x03
    TYPE_INT_DEC = 0x10
    TYPE_INT_HEX = 0x11
    TYPE_INT_BOOLEAN = 0x12

    def __init__(self, data: bytes):
        self.data = data
        self.strings: List[str] = []
        self.elements: List[Tuple[str, List[Tuple[str, str]]]] = []

    def parse(self) -> Dict[str, object]:
        offset = 8  # Skip RES_XML_TYPE file header.
        while offset + 8 <= len(self.data):
            chunk_type, header_size, chunk_size = struct.unpack_from("<HHI", self.data, offset)
            if chunk_size <= 0:
                break
            if chunk_type == self.RES_STRING_POOL_TYPE:
                self._parse_string_pool(offset)
            elif chunk_type == self.RES_XML_START_ELEMENT_TYPE:
                self._parse_start_element(offset, header_size)
            offset += chunk_size
        return self._extract_manifest_info()

    def _get_string(self, idx: int) -> str:
        if idx == self.NO_INDEX or idx < 0 or idx >= len(self.strings):
            return ""
        return self.strings[idx]

    def _read_utf8_len(self, pos: int) -> Tuple[int, int]:
        first = self.data[pos]
        if first & 0x80:
            return ((first & 0x7F) << 8) | self.data[pos + 1], pos + 2
        return first, pos + 1

    def _read_utf16_len(self, pos: int) -> Tuple[int, int]:
        first = struct.unpack_from("<H", self.data, pos)[0]
        if first & 0x8000:
            second = struct.unpack_from("<H", self.data, pos + 2)[0]
            return ((first & 0x7FFF) << 16) | second, pos + 4
        return first, pos + 2

    def _parse_string_pool(self, offset: int) -> None:
        header = struct.unpack_from("<HHIIIIII", self.data, offset)
        _, header_size, chunk_size, string_count, style_count, flags, strings_start, styles_start = header
        offsets_pos = offset + header_size
        string_offsets = [struct.unpack_from("<I", self.data, offsets_pos + i * 4)[0] for i in range(string_count)]
        base = offset + strings_start
        is_utf8 = bool(flags & self.UTF8_FLAG)
        strings: List[str] = []
        for string_offset in string_offsets:
            pos = base + string_offset
            try:
                if is_utf8:
                    _, pos = self._read_utf8_len(pos)
                    byte_len, pos = self._read_utf8_len(pos)
                    raw = self.data[pos:pos + byte_len]
                    strings.append(raw.decode("utf-8", "replace"))
                else:
                    char_len, pos = self._read_utf16_len(pos)
                    raw = self.data[pos:pos + char_len * 2]
                    strings.append(raw.decode("utf-16le", "replace"))
            except Exception:
                strings.append("")
        self.strings = strings

    def _format_value(self, raw_idx: int, data_type: int, data: int) -> str:
        if raw_idx != self.NO_INDEX:
            return self._get_string(raw_idx)
        if data_type == self.TYPE_STRING:
            return self._get_string(data)
        if data_type == self.TYPE_INT_BOOLEAN:
            return "true" if data != 0 else "false"
        if data_type in (self.TYPE_INT_DEC, self.TYPE_INT_HEX):
            return str(data)
        return str(data) if data else ""

    def _parse_start_element(self, offset: int, header_size: int) -> None:
        # ResXMLTree_node is 16 bytes, then ResXMLTree_attrExt is 20 bytes.
        if offset + 36 > len(self.data):
            return
        name_idx = struct.unpack_from("<I", self.data, offset + 20)[0]
        tag_name = self._get_string(name_idx)
        attr_start, attr_size, attr_count = struct.unpack_from("<HHH", self.data, offset + 24)
        attrs_offset = offset + header_size
        attrs: List[Tuple[str, str]] = []
        for i in range(attr_count):
            pos = attrs_offset + i * attr_size
            if pos + 20 > len(self.data):
                continue
            _, attr_name_idx, raw_value_idx = struct.unpack_from("<III", self.data, pos)
            _, _, data_type, data = struct.unpack_from("<HBBI", self.data, pos + 12)
            attr_name = self._get_string(attr_name_idx)
            value = self._format_value(raw_value_idx, data_type, data)
            attrs.append((attr_name, value))
        self.elements.append((tag_name, attrs))

    def _extract_manifest_info(self) -> Dict[str, object]:
        result: Dict[str, object] = {"permissions": []}
        for tag, attrs in self.elements:
            attr_map = {name: value for name, value in attrs if name}
            if tag == "manifest":
                result["package"] = attr_map.get("package", result.get("package", ""))
                result["version_code"] = attr_map.get("versionCode", result.get("version_code", ""))
                result["version_name"] = attr_map.get("versionName", result.get("version_name", ""))
            elif tag == "uses-sdk":
                result["min_sdk"] = attr_map.get("minSdkVersion", result.get("min_sdk", ""))
                result["target_sdk"] = attr_map.get("targetSdkVersion", result.get("target_sdk", ""))
            elif tag in ("uses-permission", "uses-permission-sdk-23"):
                name = attr_map.get("name", "")
                if name:
                    result.setdefault("permissions", []).append(name)
            elif tag == "application":
                if "debuggable" in attr_map:
                    result["debuggable"] = "是" if attr_map.get("debuggable") == "true" else "否"
        if "permissions" in result:
            result["permissions"] = sorted(set(result.get("permissions", [])))
        return result


def _parse_axml_manifest(manifest_bytes: bytes) -> Dict[str, object]:
    try:
        return _AXMLParser(manifest_bytes).parse()
    except Exception:
        return {}


def analyze_apk(path: str) -> APKInfo:
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    info = APKInfo()
    info.file_name = os.path.basename(path)
    info.file_size = _format_size(os.path.getsize(path))
    info.md5, info.sha1 = _hash_file(path)

    abis = set()
    signatures = []
    manifest_data: Optional[bytes] = None
    try:
        with zipfile.ZipFile(path, "r") as apk:
            for name in apk.namelist():
                lower = name.lower()
                if lower.startswith("lib/") and lower.endswith(".so"):
                    parts = name.split("/")
                    if len(parts) >= 3 and parts[1]:
                        abis.add(parts[1])
                if lower.startswith("meta-inf/") and (lower.endswith(".rsa") or lower.endswith(".dsa") or lower.endswith(".ec") or lower.endswith(".sf")):
                    signatures.append(name)
            try:
                manifest_data = apk.read("AndroidManifest.xml")
            except KeyError:
                manifest_data = None
    except zipfile.BadZipFile as exc:
        raise ValueError("不是有效的 APK/ZIP 文件") from exc

    info.abis = sorted(abis)
    info.signature_summary = ", ".join(signatures[:8]) + (" ..." if len(signatures) > 8 else "")

    manifest_info: Dict[str, object] = {}
    if manifest_data:
        manifest_info.update(_parse_axml_manifest(manifest_data))
    aapt_info = _parse_aapt_badging(path)
    # aapt output is usually more complete; use it to fill missing values or override empty parsed values.
    manifest_info.update({k: v for k, v in aapt_info.items() if v})

    info.package = str(manifest_info.get("package", "") or "")
    info.version_name = str(manifest_info.get("version_name", "") or "")
    info.version_code = str(manifest_info.get("version_code", "") or "")
    info.min_sdk = str(manifest_info.get("min_sdk", "") or "")
    info.target_sdk = str(manifest_info.get("target_sdk", "") or "")
    info.permissions = list(manifest_info.get("permissions", []) or [])
    if manifest_info.get("abis"):
        info.abis = list(manifest_info.get("abis", []) or [])
    info.debuggable = str(manifest_info.get("debuggable", "") or "")

    if not any([info.package, info.version_name, info.version_code, info.permissions, info.min_sdk, info.target_sdk]):
        info.parse_note = "未找到 aapt，且本地 Manifest 解析未获得完整信息；可把 aapt.exe 放入 PATH 后重新解析。"
    elif not shutil.which("aapt") and not shutil.which("aapt.exe"):
        info.parse_note = "已使用本地解析；如需更完整签名/Manifest 信息，可配置 aapt.exe。"
    else:
        info.parse_note = "已结合 aapt 与本地 ZIP 信息解析。"
    return info


def format_apk_info(info: APKInfo) -> str:
    lines = [
        f"APK 文件名：{info.file_name}",
        f"文件大小：{info.file_size}",
        f"MD5：{info.md5}",
        f"SHA1：{info.sha1}",
        "",
        f"包名：{info.package or '未解析到'}",
        f"versionName：{info.version_name or '未解析到'}",
        f"versionCode：{info.version_code or '未解析到'}",
        f"minSdkVersion：{info.min_sdk or '未解析到'}",
        f"targetSdkVersion：{info.target_sdk or '未解析到'}",
        f"ABI 架构：{', '.join(info.abis) if info.abis else '未发现 lib ABI'}",
        f"是否 Debug 包：{info.debuggable or '未解析到'}",
        f"签名信息摘要：{info.signature_summary or '未发现 META-INF 签名文件'}",
        "",
        "权限列表：",
    ]
    if info.permissions:
        lines.extend([f"  - {perm}" for perm in info.permissions])
    else:
        lines.append("  未解析到")
    if info.parse_note:
        lines.extend(["", f"备注：{info.parse_note}"])
    return "\n".join(lines)
