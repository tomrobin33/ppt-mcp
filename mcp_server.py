import sys
import os
import base64
import logging
from typing import Optional, Dict, Any
from mcp.server.fastmcp import FastMCP
import requests
from parser import parse_pptx

# 配置日志
logging.basicConfig(level=logging.INFO, format='[%(asctime)s] %(levelname)s: %(message)s')
logger = logging.getLogger("ppt-mcp")

# 初始化 FastMCP 服务器
mcp = FastMCP(
    "ppt-mcp",
    version="1.0.0",
    description="PPTX结构化解析MCP Server，支持file_url和base64两种方式上传PPTX文件，返回结构化JSON。"
)

@mcp.tool()
def parse_pptx_handler(
    file_url: Optional[str] = None,
    file_bytes_b64: Optional[str] = None
) -> str:
    """
    解析 PPTX 文件，支持 file_url 或 base64，返回结构化 JSON。
    
    Args:
        file_url: PPTX文件的URL
        file_bytes_b64: PPTX文件的base64内容
        
    Returns:
        结构化PPT内容的JSON字符串
    """
    try:
        if file_url:
            logger.info(f"Downloading file from URL: {file_url}")
            resp = requests.get(file_url, timeout=10)
            if resp.status_code != 200:
                error_msg = f"Failed to download file from url: {file_url}"
                logger.error(error_msg)
                return f"Error: {error_msg}"
            file_bytes = resp.content
            logger.info(f"Successfully downloaded file, size: {len(file_bytes)} bytes")
        elif file_bytes_b64:
            logger.info("Processing base64 encoded file")
            file_bytes = base64.b64decode(file_bytes_b64)
            logger.info(f"Successfully decoded base64, size: {len(file_bytes)} bytes")
        else:
            error_msg = "Missing parameter: file_url or file_bytes_b64"
            logger.error(error_msg)
            return f"Error: {error_msg}"
        
        result = parse_pptx(file_bytes)
        logger.info(f"Successfully parsed PPTX, found {len(result.get('slides', []))} slides")
        import json
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as e:
        error_msg = f"parse_pptx_handler error: {e}"
        logger.error(error_msg)
        return f"Error: {str(e)}"

def run_stdio():
    """运行 PPT MCP 服务器在 stdio 模式"""
    try:
        logger.info("Starting PPT MCP server with stdio transport")
        mcp.run(transport="stdio")
    except KeyboardInterrupt:
        logger.info("Server stopped by user")
    except Exception as e:
        logger.error(f"Server failed: {e}")
        raise
    finally:
        logger.info("Server shutdown complete")

if __name__ == "__main__":
    run_stdio() 