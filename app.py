from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from parser import parse_pptx
from fastapi import status
from fastapi.openapi.utils import get_openapi

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
    上传 pptx 文件，返回结构化 JSON。
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