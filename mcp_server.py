import sys
import os
import base64
import logging
from typing import Optional, Dict, Any
from fastmcp import FastMCP, tool
import requests
from parser import parse_pptx

logging.basicConfig(level=logging.INFO, format='[%(asctime)s] %(levelname)s: %(message)s')

mcp = FastMCP(
    "ppt-mcp",
    version="1.0.0",
    description="PPTX结构化解析MCP Server，支持file_url和base64两种方式上传PPTX文件，返回结构化JSON。"
)

@tool(
    name="parse_pptx_handler",
    description="解析 PPTX 文件，支持 file_url 或 base64，返回结构化 JSON"
)
def parse_pptx_handler(
    file_url: Optional[str] = None,
    file_bytes_b64: Optional[str] = None
) -> Dict[str, Any]:
    """
    解析 PPTX 文件，支持 file_url 或 base64，返回结构化 JSON。
    :param file_url: PPTX文件的URL
    :param file_bytes_b64: PPTX文件的base64内容
    :return: 结构化PPT内容
    """
    try:
        if file_url:
            resp = requests.get(file_url, timeout=10)
            if resp.status_code != 200:
                return {"code": -32602, "message": f"Failed to download file from url: {file_url}"}
            file_bytes = resp.content
        elif file_bytes_b64:
            file_bytes = base64.b64decode(file_bytes_b64)
        else:
            return {"code": -32602, "message": "Missing parameter: file_url or file_bytes_b64"}
        return parse_pptx(file_bytes)
    except Exception as e:
        logging.error(f"parse_pptx_handler error: {e}")
        return {"code": -32000, "message": str(e)}

if __name__ == "__main__":
    mcp.run() 