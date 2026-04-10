"""
parsers/pdf_parser.py — PDF 解析
"""
import os
from typing import List, Dict, Optional
import pandas as pd
import pdfplumber
from models import ParsedDocument
class PDFParser:
    def __init__(self, filepath: str):
        self.filepath = filepath
        self.filename = os.path.basename(filepath)
    def parse(self) -> ParsedDocument:
        doc = ParsedDocument(filename=self.filename)
        tables = []
        text_parts = []
        with pdfplumber.open(self.filepath) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                # 文本
                txt = page.extract_text() or ""
                text_parts.append(txt)
                # 表格
                for raw in page.extract_tables():
                    df = self._clean(raw)
                    if df is not None:
                        tables.append({"page": page_num, "data": df})
        doc.raw_text = "\n".join(text_parts)
        doc.raw_tables = tables
        return doc
    @staticmethod
    def _clean(raw_table) -> Optional[pd.DataFrame]:
        if not raw_table or len(raw_table) < 2:
            return None
        rows = []
        for row in raw_table:
            rows.append([
                str(c).replace("\n", " ").strip() if c else ""
                for c in row
            ])
        max_cols = max(len(r) for r in rows)
        rows = [r + [""] * (max_cols - len(r)) for r in rows]
        header = rows[0]
        data = rows[1:]
        if all(h == "" for h in header):
            return None
        try:
            return pd.DataFrame(data, columns=header)
        except Exception:
            return pd.DataFrame(data)
