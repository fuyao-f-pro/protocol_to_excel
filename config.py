"""
config.py — 全局配置：标准列、默认值、关键词映射、样式
"""
import os
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
for _d in (INPUT_DIR, OUTPUT_DIR):
    os.makedirs(_d, exist_ok=True)
# ---------- 标准 Excel 列 ----------
STANDARD_COLUMNS = [
    "序号",
    "点位名称",
    "点位简称",
    "点位类型",
    "数据标签",
    "数据地址",
    "功能码",
    "数据类型",
    "解码规则",
    "寄存器数量",
    "Bit位",
    "单位",
    "变比/精度",
    "访问规则",
    "翻译枚举",
    "是否保存历史",
    "状态",
    "备注",
]
# ---------- 默认值 ----------
DEFAULTS = {
    "点位类型": "模拟量",
    "数据类型": "UINT16",
    "解码规则": "原始值",
    "寄存器数量": 1,
    "Bit位": "",
    "变比/精度": 1,
    "访问规则": "R",
    "翻译枚举": "",
    "是否保存历史": "是",
    "状态": "启用",
    "备注": "",
    "点位简称": "",
    "功能码": "03",
}
# ---------- 关键词 → 标准字段映射 ----------
# 表头可能出现的各种写法 → 我们的标准列名
HEADER_ALIAS = {
    "数据地址": [
        "地址", "address", "addr", "寄存器地址", "register",
        "reg", "起始地址", "start address", "寄存器",
    ],
    "点位名称": [
        "名称", "name", "参数名", "说明", "描述",
        "description", "数据项", "内容", "功能", "参数",
    ],
    "数据类型": [
        "数据类型", "type", "data type", "格式", "类型",
    ],
    "单位": [
        "单位", "unit",
    ],
    "变比/精度": [
        "变比", "倍率", "精度", "系数", "缩放", "scale",
        "ratio", "分辨率", "resolution",
    ],
    "访问规则": [
        "读写", "access", "r/w", "rw", "权限", "操作",
    ],
    "寄存器数量": [
        "长度", "length", "字节数", "寄存器数", "数量", "size",
    ],
    "功能码": [
        "功能码", "function", "func", "fc",
    ],
    "备注": [
        "备注", "remark", "note", "comment",
    ],
    "Bit位": [
        "bit", "位", "bit位",
    ],
}
# ---------- 单位推断 ----------
UNIT_INFER = {
    "°C": ["温度", "temp", "temperature", "测温"],
    "%RH": ["湿度", "humi", "humidity"],
    "V": ["电压", "voltage"],
    "A": ["电流", "current", "ampere"],
    "kW": ["功率", "power", "有功", "无功"],
    "kWh": ["电量", "电能", "能耗", "energy"],
    "Hz": ["频率", "freq"],
    "Pa": ["压力", "pressure", "气压"],
    "m³/h": ["流量", "flow"],
    "%": ["浓度", "百分比", "开度", "湿度"],
    "rpm": ["转速", "speed"],
    "m/s": ["风速", "wind"],
}
# ---------- 数据类型归一化 ----------
DATATYPE_NORMALIZE = {
    "UINT16": ["uint16", "u16", "unsigned 16", "无符号16", "16位无符号", "unsigned short"],
    "INT16":  ["int16", "i16", "signed 16", "有符号16", "16位有符号", "short"],
    "UINT32": ["uint32", "u32", "unsigned 32", "32位无符号", "unsigned int", "unsigned long"],
    "INT32":  ["int32", "i32", "signed 32", "32位有符号", "signed int", "long"],
    "FLOAT32":["float32", "float", "real", "单精度", "浮点"],
    "FLOAT64":["float64", "double", "双精度"],
    "BOOL":   ["bool", "boolean", "bit", "开关量", "线圈"],
    "STRING": ["string", "ascii", "字符串"],
}
# ---------- Excel 样式 ----------
STYLE = {
    "header_bg": "4472C4",
    "header_font_color": "FFFFFF",
    "header_font_size": 11,
    "data_font_size": 10,
    "border": "thin",
    "alt_row_colors": ["FFFFFF", "F2F7FB"],
    "freeze": "C2",
}