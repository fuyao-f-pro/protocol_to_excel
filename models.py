"""
models.py — 数据点位模型
"""
from __future__ import annotations
from dataclasses import dataclass, field
from typing import List, Dict, Any
import config
@dataclass
class DataPoint:
    """一条点位记录"""
    序号: int = 0
    点位名称: str = ""
    点位简称: str = ""
    点位类型: str = "模拟量"
    数据标签: str = ""
    数据地址: str = ""
    功能码: str = "03"
    数据类型: str = "UINT16"
    解码规则: str = "原始值"
    寄存器数量: int = 1
    Bit位: str = ""
    单位: str = ""
    变比_精度: float = 1          # 对应"变比/精度"列
    访问规则: str = "R"
    翻译枚举: str = ""
    是否保存历史: str = "是"
    状态: str = "启用"
    备注: str = ""
    # --- 内部字段 ---
    _source: str = field(default="", repr=False)
    def to_row(self) -> List[Any]:
        """按标准列顺序输出"""
        return [
            self.序号,
            self.点位名称,
            self.点位简称,
            self.点位类型,
            self.数据标签,
            self.数据地址,
            self.功能码,
            self.数据类型,
            self.解码规则,
            self.寄存器数量,
            self.Bit位,
            self.单位,
            self.变比_精度,
            self.访问规则,
            self.翻译枚举,
            self.是否保存历史,
            self.状态,
            self.备注,
        ]
    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> DataPoint:
        """从字典构建，自动补全默认值"""
        merged = {**config.DEFAULTS, **d}
        return cls(
            点位名称=str(merged.get("点位名称", "")),
            点位简称=str(merged.get("点位简称", "")),
            点位类型=str(merged.get("点位类型", "模拟量")),
            数据标签=str(merged.get("数据标签", merged.get("点位名称", ""))),
            数据地址=str(merged.get("数据地址", "")),
            功能码=str(merged.get("功能码", "03")),
            数据类型=str(merged.get("数据类型", "UINT16")),
            解码规则=str(merged.get("解码规则", "原始值")),
            寄存器数量=_safe_int(merged.get("寄存器数量", 1)),
            Bit位=str(merged.get("Bit位", "")),
            单位=str(merged.get("单位", "")),
            变比_精度=_safe_float(merged.get("变比/精度", 1)),
            访问规则=str(merged.get("访问规则", "R")),
            翻译枚举=str(merged.get("翻译枚举", "")),
            是否保存历史=str(merged.get("是否保存历史", "是")),
            状态=str(merged.get("状态", "启用")),
            备注=str(merged.get("备注", "")),
        )
@dataclass
class ParsedDocument:
    """解析后的完整文档"""
    filename: str = ""
    protocol_type: str = ""
    device_name: str = ""
    points: List[DataPoint] = field(default_factory=list)
    raw_text: str = ""
    raw_tables: list = field(default_factory=list)
    def renumber(self):
        for i, p in enumerate(self.points, 1):
            p.序号 = i
def _safe_int(v) -> int:
    try:
        return int(float(str(v).strip()))
    except (ValueError, TypeError):
        return 1
def _safe_float(v) -> float:
    try:
        return float(str(v).strip())
    except (ValueError, TypeError):
        return 1.0
