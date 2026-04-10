# 📄 协议文档 → 标准Excel 自动化转换工具

## 功能

将楼宇自控/工业协议的 PDF / Word 文档，自动提取点位信息并生成标准化 Excel 点位表。

### 核心能力

- **多格式解析**：PDF（含扫描版）、Word (.docx)
- **智能识别**：自动判断 Modbus RTU/TCP、BACnet、DL/T645 等协议
- **字段提取**：寄存器地址、名称、数据类型、功能码、读写属性、单位、倍率等
- **智能补全**：自动推断缺失的单位、数据类型、点位类型
- **标准输出**：17列标准点位表 + 汇总统计 Sheet

## 安装

```bash
# 克隆项目
git clone <repo_url>
cd protocol-to-excel

# 创建虚拟环境（推荐）
python -m venv venv
source venv/bin/activate   # Linux/Mac
venv\Scripts\activate      # Windows

# 安装依赖
pip install -r requirements.txt

```

# 使用方式
## 方式一：命令行（CLI）
### 处理 input/ 目录下所有文件
python main.py

### 处理单个文件
python main.py -i 协议文档.pdf

### 处理多个文件
python main.py -i doc1.pdf doc2.docx

### 指定输出路径
python main.py -i 协议.pdf -o 输出结果.xlsx

## 方式二：Web 界面（Streamlit）
streamlit run app.py
浏览器会自动打开 http://localhost:8501，拖拽上传文件即可。

# 目录结构
protocol-to-excel/
├── main.py              # CLI 入口
├── app.py               # Streamlit Web 界面
├── config.py            # 全局配置（列定义、别名、样式）
├── models.py            # 数据模型（DataPoint / ParsedDocument）
├── excel_writer.py      # Excel 生成器
├── parsers/
│   ├── __init__.py
│   ├── pdf_parser.py    # PDF 解析器
│   ├── word_parser.py   # Word 解析器
│   └── engine.py        # 协议智能提取引擎
├── input/               # 默认输入目录
├── output/              # 默认输出目录
├── requirements.txt
└── README.md

# 输出 Excel 字段说明
| 序号 | 字段名       | 说明                           |
|------|-------------|-------------------------------|
| 1    | 序号         | 自动编号                       |
| 2    | 点位名称     | 寄存器中文名称                  |
| 3    | 点位简称     | 名称缩写（≤15字符）             |
| 4    | 点位类型     | 模拟量 / 开关量                 |
| 5    | 数据地址     | 十进制寄存器地址                |
| 6    | 功能码       | Modbus功能码 01/02/03/04       |
| 7    | 数据类型     | UINT16/INT16/FLOAT32 等        |
| 8    | 寄存器数量   | 占用寄存器个数                  |
| 9    | 缩放倍率     | 原始值到工程值的缩放系数         |
| 10   | 单位         | ℃、kW、V、A 等                 |
| 11   | 访问规则     | R（只读）/ W（只写）/ RW（读写） |
| 12   | 数据标签     | 用于系统集成的标签名             |
| 13   | 上限值       | 工程量上限                      |
| 14   | 下限值       | 工程量下限                      |
| 15   | 翻译枚举     | 开关量枚举 如 0=关,1=开         |
| 16   | 备注         | 补充说明                       |
| 17   | 字节序       | Big-Endian / Little-Endian     |

# 支持的协议类型
Modbus RTU / TCP
BACnet
DL/T645
OPC UA
IEC 104
MQTT
KNX

# 常见问题
Q: PDF 扫描件能识别吗？
A: 当前版本使用 pdfplumber 提取文本层。如果 PDF 是纯扫描图片（无文本层），需要集成 OCR 引擎（如 PaddleOCR），可在 pdf_parser.py 中扩展。

Q: 表格跨页怎么办？
A: 程序会逐页提取表格并自动合并，跨页表格只要列结构一致即可正确拼接。

Q: 提取不完整怎么办？
A: 检查文档中的表格是否为标准表格格式。非结构化的纯文本描述，程序会尝试正则匹配，但准确率较低。建议将关键点位信息整理为表格形式。

# 快速启动命令

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 创建输入目录并放入文件
mkdir -p input output
cp 你的协议文档.pdf input/

# 3a. 命令行模式
python main.py

# 3b. Web模式
streamlit run app.py
```