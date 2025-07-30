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
    """
    从PPT形状中提取文本内容。
    支持以下内容提取：
    1. 普通文本框中的文本
    2. 表格中的文本
    3. 分组形状中的文本
    
    Args:
        shape: PPT中的形状对象
        
    Returns:
        包含所有提取文本的列表
    """
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


def extract_table_from_shape(shape) -> Optional[List[List[str]]]:
    """
    从PPT形状中提取表格结构化数据。
    
    Args:
        shape: PPT中的形状对象
        
    Returns:
        表格数据（二维数组），如果不是表格则返回None
    """
    if shape.shape_type == 19:  # MSO_SHAPE_TYPE.TABLE
        table_data = []
        for row in shape.table.rows:
            row_data = []
            for cell in row.cells:
                cell_text = cell.text_frame.text.strip()
                row_data.append(cell_text)
            table_data.append(row_data)
        return table_data
    return None


def extract_all_content_from_shape(shape) -> Dict[str, Any]:
    """
    从PPT形状中提取所有内容（文本和表格）。
    
    Args:
        shape: PPT中的形状对象
        
    Returns:
        包含文本和表格数据的字典
    """
    content = {
        "texts": [],
        "tables": []
    }
    
    # 提取文本
    if hasattr(shape, "text"):
        text = shape.text.strip()
        if text:
            content["texts"].append(text)
    
    # 提取表格
    if shape.shape_type == 19:  # MSO_SHAPE_TYPE.TABLE
        table_data = extract_table_from_shape(shape)
        if table_data:
            content["tables"].append(table_data)
    
    # 处理分组形状
    if hasattr(shape, "shapes"):
        for sub_shape in shape.shapes:
            sub_content = extract_all_content_from_shape(sub_shape)
            content["texts"].extend(sub_content["texts"])
            content["tables"].extend(sub_content["tables"])
    
    return content

def parse_pptx(file_bytes: bytes) -> Dict[str, Any]:
    """
    解析 PPTX 文件，返回结构化 JSON。
    
    功能说明：
    1. 支持内容：
       - 文本框中的文本
       - 表格中的文本和结构化数据
       - 分组形状中的文本
       
    2. 返回格式：
       {
           "slides": [
               {
                   "slide_index": 1,
                   "text": ["文本1", "文本2", ...],
                   "tables": [
                       [["单元格1", "单元格2"], ["单元格3", "单元格4"]],
                       ...
                   ]
               },
               ...
           ]
       }
    
    Args:
        file_bytes: PPTX文件的二进制内容
        
    Returns:
        包含所有幻灯片内容的字典
        
    Raises:
        ValueError: 当文件不是有效的PPTX格式时抛出
    """
    try:
        prs = Presentation(BytesIO(file_bytes))
    except Exception as e:
        raise ValueError(f"无法读取 pptx 文件: {e}")
    slides = []
    for idx, slide in enumerate(prs.slides, start=1):
        slide_content = {
            "slide_index": idx,
            "text": [],
            "tables": []
        }
        for shape in slide.shapes:
            try:
                content = extract_all_content_from_shape(shape)
                slide_content["text"].extend(content["texts"])
                slide_content["tables"].extend(content["tables"])
            except Exception:
                continue
        slides.append(slide_content)
    return {"slides": slides}


def parse_docx(file_bytes: bytes) -> Dict[str, Any]:
    """
    解析 DOCX 文件，返回结构化 JSON。
    
    功能说明：
    1. 支持内容：
       - 文档中的所有段落文本
       - 表格内容（按行列结构保存）
       - 图片信息（文件名和大小）
       
    2. 返回格式：
       {
           "paragraphs": ["段落1", "段落2", ...],
           "tables": [
               [["单元格1", "单元格2"], ["单元格3", "单元格4"]],
               ...
           ],
           "images": [
               {"filename": "图片1.png", "size": 1024},
               ...
           ]
       }
    
    Args:
        file_bytes: DOCX文件的二进制内容
        
    Returns:
        包含文档内容的结构化字典
        
    注意：
    - 使用临时文件处理，会自动清理
    - 图片内容仅保存基本信息，不包含实际图片数据
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


def extract_tables_from_pptx(file_bytes: bytes) -> Dict[str, Any]:
    """
    专门从PPTX文件中提取所有表格数据。
    
    功能说明：
    1. 提取所有幻灯片中的表格
    2. 保持表格的行列结构
    3. 为每个表格添加来源信息
    
    2. 返回格式：
       {
           "total_tables": 5,
           "tables": [
               {
                   "table_index": 1,
                   "slide_index": 2,
                   "data": [["标题1", "标题2"], ["数据1", "数据2"]],
                   "rows": 2,
                   "columns": 2
               },
               ...
           ]
       }
    
    Args:
        file_bytes: PPTX文件的二进制内容
        
    Returns:
        包含所有表格数据的字典
        
    Raises:
        ValueError: 当文件不是有效的PPTX格式时抛出
    """
    try:
        prs = Presentation(BytesIO(file_bytes))
    except Exception as e:
        raise ValueError(f"无法读取 pptx 文件: {e}")
    
    all_tables = []
    table_index = 1
    
    for slide_idx, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            try:
                if shape.shape_type == 19:  # MSO_SHAPE_TYPE.TABLE
                    table_data = extract_table_from_shape(shape)
                    if table_data and len(table_data) > 0:
                        table_info = {
                            "table_index": table_index,
                            "slide_index": slide_idx,
                            "data": table_data,
                            "rows": len(table_data),
                            "columns": len(table_data[0]) if table_data else 0
                        }
                        all_tables.append(table_info)
                        table_index += 1
            except Exception:
                continue
    
    return {
        "total_tables": len(all_tables),
        "tables": all_tables
    }


def parse_xlsx(file_bytes: bytes) -> Dict[str, Any]:
    """
    解析 XLSX 文件，返回结构化 JSON。
    
    功能说明：
    1. 支持内容：
       - 所有工作表（sheets）的内容
       - 单元格的值和坐标信息
       - 单元格中的公式
       
    2. 返回格式：
       {
           "sheets": [
               {
                   "title": "Sheet1",
                   "cells": [
                       [
                           {"value": "A1的值", "coordinate": "A1"},
                           {"value": "B1的值", "coordinate": "B1"}
                       ],
                       ...
                   ],
                   "formulas": [
                       {"coordinate": "A1", "formula": "=SUM(B1:B10)"},
                       ...
                   ]
               },
               ...
           ]
       }
    
    Args:
        file_bytes: XLSX文件的二进制内容
        
    Returns:
        包含Excel文件内容的结构化字典
        
    注意：
    - 使用临时文件处理，会自动清理
    - data_only=False 设置可以获取公式内容
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