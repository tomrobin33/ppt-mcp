from pptx import Presentation
from typing import List, Dict, Any
from io import BytesIO
from docx import Document
import openpyxl
import tempfile
import os
from typing import Any, Dict
from PIL import Image


def extract_text_from_shape(shape) -> List[str]:
    texts = []
    if hasattr(shape, "text"):
        text = shape.text.strip()
        if text:
            texts.append(text)
    # 支持表格内容提取
    if shape.shape_type == 19:  # MSO_SHAPE_TYPE.TABLE
        for row in shape.table.rows:
            for cell in row.cells:
                cell_text = cell.text_frame.text.strip()
                if cell_text:
                    texts.append(cell_text)
    # 支持分组 shape
    if hasattr(shape, "shapes"):
        for sub_shape in shape.shapes:
            texts.extend(extract_text_from_shape(sub_shape))
    return texts

def parse_pptx(file_bytes: bytes) -> Dict[str, Any]:
    """
    解析 pptx 文件，返回结构化 JSON。
    """
    try:
        prs = Presentation(BytesIO(file_bytes))
    except Exception as e:
        raise ValueError(f"无法读取 pptx 文件: {e}")
    slides = []
    for idx, slide in enumerate(prs.slides, start=1):
        texts = []
        for shape in slide.shapes:
            try:
                texts.extend(extract_text_from_shape(shape))
            except Exception:
                continue
        slides.append({
            "slide_index": idx,
            "text": texts
        })
    return {"slides": slides}


def parse_docx(file_bytes: bytes) -> Dict[str, Any]:
    """
    解析 docx 文件，返回结构化 JSON。
    提取所有段落、表格、图片。
    """
    result = {"paragraphs": [], "tables": [], "images": []}
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    try:
        doc = Document(tmp_path)
        # 段落
        for para in doc.paragraphs:
            result["paragraphs"].append(para.text)
        # 表格
        for table in doc.tables:
            table_data = []
            for row in table.rows:
                row_data = [cell.text for cell in row.cells]
                table_data.append(row_data)
            result["tables"].append(table_data)
        # 图片
        rels = doc.part.rels
        for rel in rels:
            rel = rels[rel]
            if "image" in rel.target_ref:
                image_bytes = rel.target_part.blob
                result["images"].append({"filename": os.path.basename(rel.target_ref), "size": len(image_bytes)})
    finally:
        os.remove(tmp_path)
    return result


def parse_xlsx(file_bytes: bytes) -> Dict[str, Any]:
    """
    解析 xlsx 文件，返回结构化 JSON。
    提取所有sheet、单元格、公式。
    """
    result = {"sheets": []}
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    try:
        wb = openpyxl.load_workbook(tmp_path, data_only=False)
        for sheet in wb.worksheets:
            sheet_data = {"title": sheet.title, "cells": [], "formulas": []}
            for row in sheet.iter_rows():
                row_data = []
                for cell in row:
                    cell_info = {"value": cell.value, "coordinate": cell.coordinate}
                    if cell.data_type == 'f':
                        cell_info["formula"] = cell.value
                        sheet_data["formulas"].append({"coordinate": cell.coordinate, "formula": cell.value})
                    row_data.append(cell_info)
                sheet_data["cells"].append(row_data)
            result["sheets"].append(sheet_data)
    finally:
        os.remove(tmp_path)
    return result 