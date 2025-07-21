from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from parser import parse_pptx, parse_docx, parse_xlsx
from fastapi import status
from fastapi.openapi.utils import get_openapi
import requests
import mimetypes

app = FastAPI(
    title="PPTX 解析微服务",
    description="上传 .pptx 文件，返回结构化 JSON 内容。",
    version="1.1.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/parse-ppt", summary="解析 PPTX 文件", response_description="结构化 JSON 内容", status_code=status.HTTP_200_OK)
async def parse_ppt(file: UploadFile = File(...)):
    """
    上传 PPTX 文件并解析为结构化 JSON。
    
    请求说明：
    1. 请求方式：POST
    2. Content-Type: multipart/form-data
    3. 参数：
       - file: PPTX文件（必需）
       
    返回格式：
    {
        "slides": [
            {
                "slide_index": 1,
                "text": ["文本1", "文本2", ...]
            },
            ...
        ]
    }
    
    错误码：
    - 400：文件格式错误，仅支持.pptx文件
    - 500：服务器解析错误
    
    使用示例：
    ```python
    import requests
    
    url = 'http://your-server/parse-ppt'
    files = {'file': open('example.pptx', 'rb')}
    response = requests.post(url, files=files)
    result = response.json()
    ```
    """
    if not file.filename or not file.filename.endswith(".pptx"):
        raise HTTPException(status_code=400, detail="只支持 .pptx 文件")
    file_bytes = await file.read()
    try:
        result = parse_pptx(file_bytes)
    except ValueError as ve:
        raise HTTPException(status_code=400, detail=str(ve))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"解析失败: {str(e)}")
    return JSONResponse(content=result)

@app.post("/parse-url", summary="通过URL解析PPT/Word/Excel文件", response_description="结构化 JSON 内容", status_code=status.HTTP_200_OK)
async def parse_url(url: str):
    """
    通过URL下载并解析PPT/Word/Excel文件，返回结构化JSON。
    
    请求说明：
    1. 请求方式：POST
    2. Content-Type: application/json
    3. 参数：
       - url: 文件的公开URL地址（必需）
       
    支持的文件类型：
    - .pptx：PowerPoint文件
    - .docx：Word文件
    - .xlsx：Excel文件
    
    返回格式：
    1. PPTX文件：
       {
           "slides": [{"slide_index": 1, "text": [...]}]
       }
       
    2. DOCX文件：
       {
           "paragraphs": [...],
           "tables": [...],
           "images": [...]
       }
       
    3. XLSX文件：
       {
           "sheets": [
               {
                   "title": "Sheet1",
                   "cells": [...],
                   "formulas": [...]
               }
           ]
       }
    
    错误码：
    - 400：URL无效或文件下载失败
    - 400：不支持的文件格式
    - 500：服务器解析错误
    
    使用示例：
    ```python
    import requests
    
    url = 'http://your-server/parse-url'
    data = {'url': 'https://example.com/document.pptx'}
    response = requests.post(url, json=data)
    result = response.json()
    ```
    """
    try:
        resp = requests.get(url)
        resp.raise_for_status()
        file_bytes = resp.content
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"文件下载失败: {e}")
    # 文件类型判断
    filename = url.split("?")[0].split("/")[-1]
    ext = filename.lower().split(".")[-1]
    if ext == "pptx":
        result = parse_pptx(file_bytes)
    elif ext == "docx":
        result = parse_docx(file_bytes)
    elif ext == "xlsx":
        result = parse_xlsx(file_bytes)
    else:
        raise HTTPException(status_code=400, detail="只支持 .pptx, .docx, .xlsx 文件")
    return JSONResponse(content=result)