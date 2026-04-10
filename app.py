"""
app.py — Streamlit Web 界面
运行: streamlit run app.py
"""
import os
import tempfile
import time
import streamlit as st
import pandas as pd

import config
from parsers.pdf_parser import PDFParser
from parsers.word_parser import WordParser
from parsers.engine import ProtocolEngine
from excel_writer import ExcelWriter

# ============================================================
# 页面配置
# ============================================================
st.set_page_config(
    page_title="协议文档 → 标准Excel",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# 侧边栏
# ============================================================
with st.sidebar:
    st.title("📄 协议文档转换工具")
    st.markdown("---")
    st.markdown(
        """
        **支持格式**
        - PDF（含扫描件OCR）
        - Word (.docx)

        **输出**
        - 标准化 Excel 点位表

        **功能**
        - 自动识别协议类型
        - 智能提取寄存器点位
        - 自动补全缺失字段
        - 单位/类型智能推断
        """
    )
    st.markdown("---")
    st.caption("v1.0 — 楼宇自控协议解析")

# ============================================================
# 主区域
# ============================================================
st.header("🔄 上传协议文档，一键生成标准点位表")

uploaded_files = st.file_uploader(
    "拖拽或选择文件（支持多文件）",
    type=["pdf", "docx", "doc"],
    accept_multiple_files=True,
)

if uploaded_files:
    st.info(f"已选择 {len(uploaded_files)} 个文件")

    if st.button("🚀 开始转换", type="primary", use_container_width=True):
        progress_bar = st.progress(0)
        status_text = st.empty()

        all_results = []

        for file_idx, uploaded in enumerate(uploaded_files):
            filename = uploaded.name
            status_text.text(f"处理中: {filename} ({file_idx+1}/{len(uploaded_files)})")

            # 保存临时文件
            suffix = os.path.splitext(filename)[1]
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(uploaded.read())
                tmp_path = tmp.name

            try:
                # 解析
                ext = suffix.lower()
                if ext == ".pdf":
                    parser = PDFParser(tmp_path)
                elif ext in (".docx", ".doc"):
                    parser = WordParser(tmp_path)
                else:
                    st.warning(f"不支持: {filename}")
                    continue

                doc = parser.parse()

                # 提取
                engine = ProtocolEngine()
                doc = engine.process(doc)

                # 生成Excel
                out_name = os.path.splitext(filename)[0] + "_标准点位表.xlsx"
                out_path = os.path.join(tempfile.gettempdir(), out_name)
                writer = ExcelWriter(doc)
                writer.write(out_path)

                all_results.append({
                    "filename": filename,
                    "out_name": out_name,
                    "out_path": out_path,
                    "doc": doc,
                    "ok": True,
                })

            except Exception as e:
                all_results.append({
                    "filename": filename,
                    "out_name": "",
                    "out_path": "",
                    "doc": None,
                    "ok": False,
                    "error": str(e),
                })

            finally:
                try:
                    os.unlink(tmp_path)
                except OSError:
                    pass

            progress_bar.progress((file_idx + 1) / len(uploaded_files))

        status_text.text("✅ 全部处理完成！")

        # ============================================================
        # 展示结果
        # ============================================================
        st.markdown("---")

        for res in all_results:
            if not res["ok"]:
                st.error(f"❌ {res['filename']} 处理失败: {res.get('error', '未知错误')}")
                continue

            doc = res["doc"]
            st.subheader(f"📊 {res['filename']}")

            # 指标卡片
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("协议类型", doc.protocol_type)
            col2.metric("总点位", len(doc.points))
            analog = sum(1 for p in doc.points if p.点位类型 == "模拟量")
            digital = sum(1 for p in doc.points if p.点位类型 == "开关量")
            col3.metric("模拟量", analog)
            col4.metric("开关量", digital)

            # 数据预览
            if doc.points:
                with st.expander("📋 数据预览（前30条）", expanded=True):
                    rows = [p.to_row() for p in doc.points[:30]]
                    preview_df = pd.DataFrame(rows, columns=config.STANDARD_COLUMNS)
                    st.dataframe(
                        preview_df,
                        use_container_width=True,
                        height=min(len(rows) * 35 + 50, 600),
                    )

            # 下载按钮
            with open(res["out_path"], "rb") as f:
                excel_data = f.read()

            st.download_button(
                label=f"⬇️ 下载 {res['out_name']}",
                data=excel_data,
                file_name=res["out_name"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            st.markdown("---")

else:
    # 未上传时展示示例
    st.markdown("### 📌 使用说明")
    st.markdown(
        """
        1. **上传**协议文档（PDF 或 Word 格式）
        2. 点击 **开始转换** 按钮
        3. 预览提取结果，**下载** 标准 Excel 文件

        ---

        ### 📋 输出字段说明

        | 字段 | 说明 |
        |------|------|
        | 序号 | 自动编号 |
        | 点位名称 | 从文档中提取的寄存器名称 |
        | 数据地址 | Modbus 寄存器地址（十进制） |
        | 数据类型 | UINT16 / INT16 / FLOAT32 等 |
        | 功能码 | 01/02/03/04 |
        | 访问规则 | R / W / RW |
        | 单位 | ℃、kW、V、A 等（自动推断） |
        | 点位类型 | 模拟量 / 开关量（自动推断） |
        """
    )
