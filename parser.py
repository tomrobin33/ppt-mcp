from pptx import Presentation
from typing import List, Dict, Any
from io import BytesIO

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