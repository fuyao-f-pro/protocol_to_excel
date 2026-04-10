"""
parsers/engine.py — 协议智能提取引擎（完整版）
"""
import re
from typing import List, Dict, Any, Optional
import pandas as pd

import config
from models import DataPoint, ParsedDocument


class ProtocolEngine:
    """协议文档 → 标准化点位列表"""

    def __init__(self):
        self.alias = config.HEADER_ALIAS
        self.unit_map = config.UNIT_INFER
        self.dtype_map = config.DATATYPE_NORMALIZE

    # ============================================================
    # 主入口
    # ============================================================
    def process(self, doc: ParsedDocument) -> ParsedDocument:
        doc.protocol_type = self._detect_protocol(doc.raw_text)
        doc.device_name = self._detect_device(doc.raw_text)

        points = self._extract_from_tables(doc.raw_tables)

        if len(points) < 3:
            text_points = self._extract_from_text(doc.raw_text)
            existing_addrs = {p.数据地址 for p in points if p.数据地址}
            for tp in text_points:
                if tp.数据地址 not in existing_addrs:
                    points.append(tp)
                    existing_addrs.add(tp.数据地址)

        for p in points:
            self._auto_fill(p, doc)

        doc.points = points
        doc.renumber()
        return doc

    # ============================================================
    # 1. 协议类型检测
    # ============================================================
    def _detect_protocol(self, text: str) -> str:
        text_l = text.lower()
        mapping = {
            "MODBUS_TCP": ["modbus tcp", "modbus/tcp"],
            "MODBUS_RTU": ["modbus rtu", "modbus-rtu", "rtu"],
            "BACnet":     ["bacnet"],
            "DL/T645":    ["dl/t645", "dlt645", "645"],
            "OPC_UA":     ["opc ua", "opcua"],
            "IEC104":     ["iec104", "iec 104", "104规约"],
            "MQTT":       ["mqtt"],
            "KNX":        ["knx"],
        }
        for proto, keywords in mapping.items():
            for kw in keywords:
                if kw in text_l:
                    return proto
        return "MODBUS_RTU"

    def _detect_device(self, text: str) -> str:
        patterns = [
            r"(?:设备|产品|型号|model)[：:\s]*([^\n\r,，]{2,30})",
            r"(?:协议|通信|通讯)[：:\s]*([^\n\r,，]{2,30})",
        ]
        for pat in patterns:
            m = re.search(pat, text, re.IGNORECASE)
            if m:
                return m.group(1).strip()
        return ""

    # ============================================================
    # 2. 从表格提取
    # ============================================================
    def _extract_from_tables(self, tables: List[Dict]) -> List[DataPoint]:
        all_points = []
        for tbl in tables:
            df = tbl.get("data")
            if df is None or not isinstance(df, pd.DataFrame) or df.empty:
                continue
            pts = self._parse_table(df)
            all_points.extend(pts)
        return all_points

    def _parse_table(self, df: pd.DataFrame) -> List[DataPoint]:
        col_map = self._map_columns(df)
        if not col_map:
            return []

        has_addr = "数据地址" in col_map
        has_name = "点位名称" in col_map
        if not has_addr and not has_name:
            return []

        points = []
        for _, row in df.iterrows():
            d = {}
            for std_field, col_idx in col_map.items():
                val = str(row.iloc[col_idx]).strip()
                if val.lower() in ("", "nan", "none"):
                    val = ""
                d[std_field] = val

            name = d.get("点位名称", "")
            addr = d.get("数据地址", "")
            if not name and not addr:
                continue
            if self._is_header_row(d):
                continue

            if addr:
                d["数据地址"] = self._normalize_address(addr)
            if "数据类型" in d:
                d["数据类型"] = self._normalize_datatype(d["数据类型"])
            if "访问规则" in d:
                d["访问规则"] = self._normalize_rw(d["访问规则"])
            if not d.get("数据标签"):
                d["数据标签"] = name

            points.append(DataPoint.from_dict(d))
        return points

    def _map_columns(self, df: pd.DataFrame) -> Dict[str, int]:
        col_map = {}
        columns_lower = [str(c).lower().strip() for c in df.columns]

        for std_field, aliases in self.alias.items():
            best_idx = None
            for alias in aliases:
                for idx, col in enumerate(columns_lower):
                    if col == alias.lower():
                        best_idx = idx
                        break
                if best_idx is not None:
                    break
            if best_idx is None:
                for alias in aliases:
                    for idx, col in enumerate(columns_lower):
                        if alias.lower() in col or col in alias.lower():
                            best_idx = idx
                            break
                    if best_idx is not None:
                        break
            if best_idx is not None:
                col_map[std_field] = best_idx
        return col_map

    def _is_header_row(self, d: Dict[str, str]) -> bool:
        vals = [v.lower() for v in d.values() if v]
        header_words = {"地址", "名称", "address", "name", "register",
                        "类型", "type", "说明", "单位", "unit"}
        match_count = sum(1 for v in vals if v in header_words)
        return match_count >= 2

    # ============================================================
    # 3. 从文本提取
    # ============================================================
    def _extract_from_text(self, text: str) -> List[DataPoint]:
        points = []
        seen_addrs = set()

        pat_register = re.compile(
            r"(?:寄存器|地址|register|addr)[：:\s]*(\d{1,6})\s*[,，\-~\s]\s*(.{2,40})",
            re.IGNORECASE,
        )

        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue

            m2 = pat_register.search(line)
            if m2:
                addr = m2.group(1).strip()
                name = m2.group(2).strip()
                if addr not in seen_addrs and len(name) >= 2:
                    d = {"数据地址": self._normalize_address(addr), "点位名称": name}
                    points.append(DataPoint.from_dict(d))
                    seen_addrs.add(addr)
                continue

            hex_match = re.search(r"(0[xX][0-9a-fA-F]{1,6}|\d{3,6}[hH])", line)
            if hex_match:
                addr_raw = hex_match.group(1)
                after = line[hex_match.end():].strip()
                name = re.split(r"\s{2,}|\t", after)[0].strip() if after else ""
                if name and len(name) >= 2 and addr_raw not in seen_addrs:
                    d = {"数据地址": self._normalize_address(addr_raw), "点位名称": name}
                    dtype_m = re.search(
                        r"\b(UINT16|INT16|UINT32|INT32|FLOAT32|FLOAT|BOOL)\b",
                        line, re.IGNORECASE,
                    )
                    if dtype_m:
                        d["数据类型"] = self._normalize_datatype(dtype_m.group(1))
                    points.append(DataPoint.from_dict(d))
                    seen_addrs.add(addr_raw)

        return points

    # ============================================================
    # 4. 智能补全
    # ============================================================
    def _auto_fill(self, point: DataPoint, doc: ParsedDocument):
        name = point.点位名称

        if not point.单位:
            point.单位 = self._infer_unit(name)

        point.点位类型 = self._infer_point_type(name, point.数据类型)

        if not point.点位简称:
            point.点位简称 = name[:15] if len(name) > 15 else name

        if not point.数据标签:
            point.数据标签 = name

        if point.数据类型 in ("UINT32", "INT32", "FLOAT32"):
            point.寄存器数量 = max(point.寄存器数量, 2)
        elif point.数据类型 in ("FLOAT64",):
            point.寄存器数量 = max(point.寄存器数量, 4)

        if point.功能码 in ("", "nan", "None"):
            point.功能码 = "01" if point.点位类型 == "开关量" else "03"

        if point.点位类型 == "开关量" and not point.翻译枚举:
            point.翻译枚举 = "0=关,1=开"

    def _infer_unit(self, name: str) -> str:
        name_l = name.lower()
        for unit, keywords in self.unit_map.items():
            for kw in keywords:
                if kw.lower() in name_l:
                    return unit
        return ""

    def _infer_point_type(self, name: str, dtype: str) -> str:
        digital_keywords = [
            "状态", "开关", "告警", "报警", "故障", "运行",
            "启停", "通断", "on/off", "alarm", "fault",
            "status", "switch", "bool", "flag", "使能",
        ]
        name_l = name.lower()
        for kw in digital_keywords:
            if kw.lower() in name_l:
                return "开关量"
        if dtype in ("BOOL",):
            return "开关量"
        return "模拟量"

    # ============================================================
    # 5. 规范化工具
    # ============================================================
    def _normalize_address(self, raw: str) -> str:
        raw = raw.strip()
        if not raw:
            return ""
        # 去掉末尾的 H/h
        if raw.endswith(("H", "h")) and not raw.startswith("0x"):
            raw = "0x" + raw[:-1]
        # 已经是 0x 开头 → 转十进制
        if raw.lower().startswith("0x"):
            try:
                return str(int(raw, 16))
            except ValueError:
                return raw
        # 纯数字直接返回
        if re.match(r"^\d+$", raw):
            return raw
        return raw

    def _normalize_datatype(self, raw: str) -> str:
        raw_l = raw.lower().strip()
        if not raw_l:
            return "UINT16"
        for std, aliases in self.dtype_map.items():
            if raw_l == std.lower():
                return std
            for alias in aliases:
                if alias.lower() == raw_l or alias.lower() in raw_l:
                    return std
        return "UINT16"

    def _normalize_rw(self, raw: str) -> str:
        raw_l = raw.lower().strip()
        if not raw_l:
            return "R"
        if "rw" in raw_l or ("r" in raw_l and "w" in raw_l):
            return "RW"
        if "w" in raw_l or "写" in raw_l:
            return "W"
        return "R"