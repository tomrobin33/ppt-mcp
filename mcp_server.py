import sys
import os
import base64
import logging
from typing import Optional, Dict, Any
from mcp.server.fastmcp import FastMCP
import requests
from parser import parse_pptx, parse_docx, parse_xlsx

# 配置日志
logging.basicConfig(level=logging.INFO, format='[%(asctime)s] %(levelname)s: %(message)s')
logger = logging.getLogger("ppt-mcp")

# 初始化 FastMCP 服务器
mcp = FastMCP(
    "ppt-mcp",
    version="1.0.0",
    description="文档解析MCP Server，支持PPTX、DOCX、XLSX文件的解析，返回结构化JSON。"
)

@mcp.tool()
def parse_pptx_handler(
    file_url: Optional[str] = None,
    file_bytes_b64: Optional[str] = None
) -> str:
    """
    解析 PPTX 文件，支持 file_url 或 base64，返回结构化 JSON。    注意：此工具函数仅支持解析 PPTX 格式文件，不支持 DOCX 或 XLSX。
    
    调用步骤：
    1. 文件类型检查：
       - 确保待解析的文件是 PPTX 格式
       - 可以通过文件扩展名或文件头部特征进行判断
       
    2. 文件获取方式：
       - URL方式：提供 file_url 参数，指向可下载的PPTX文件
       - Base64方式：提供 file_bytes_b64 参数，包含PPTX文件的base64编码内容
       
    3. 错误处理：
       - 如果文件不是PPTX格式，会抛出相应错误
       - 如果文件下载失败，会返回错误信息
       - 如果base64解码失败，会返回错误信息
    
    Args:
        file_url: PPTX文件的URL，与file_bytes_b64参数二选一
        file_bytes_b64: PPTX文件的base64内容，与file_url参数二选一
        
    Returns:
        结构化PPT内容的JSON字符串，包含幻灯片文本、表格等信息
        
    错误返回示例：
        - "Error: Failed to download file from url: {url}"
        - "Error: Missing parameter: file_url or file_bytes_b64"
        - "Error: Invalid file format, only PPTX files are supported"
    """
    try:
        # 获取文件内容
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
        
        # 解析PPTX文件
        result = parse_pptx(file_bytes)
        logger.info(f"Successfully parsed PPTX, found {len(result.get('slides', []))} slides")
        import json
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as e:
        error_msg = f"parse_pptx_handler error: {e}"
        logger.error(error_msg)
        return f"Error: {str(e)}"

@mcp.tool()
def parse_docx_handler(
    file_url: Optional[str] = None,
    file_bytes_b64: Optional[str] = None
) -> str:
    """
    解析 DOCX 文件，支持 file_url 或 base64，返回结构化 JSON。
    注意：此工具函数仅支持解析 DOCX 格式文件，不支持 PPTX 或 XLSX。
    
    调用步骤：
    1. 文件类型检查：
       - 确保待解析的文件是 DOCX 格式
       - 可以通过文件扩展名或文件头部特征进行判断
       
    2. 文件获取方式：
       - URL方式：提供 file_url 参数，指向可下载的DOCX文件
       - Base64方式：提供 file_bytes_b64 参数，包含DOCX文件的base64编码内容
       
    3. 错误处理：
       - 如果文件不是DOCX格式，会抛出相应错误
       - 如果文件下载失败，会返回错误信息
       - 如果base64解码失败，会返回错误信息
    
    Args:
        file_url: DOCX文件的URL，与file_bytes_b64参数二选一
        file_bytes_b64: DOCX文件的base64内容，与file_url参数二选一
        
    Returns:
        结构化Word内容的JSON字符串，包含：
        - paragraphs: 所有段落文本
        - tables: 表格内容
        - images: 图片信息
        
    错误返回示例：
        - "Error: Failed to download file from url: {url}"
        - "Error: Missing parameter: file_url or file_bytes_b64"
        - "Error: Invalid file format, only DOCX files are supported"
    """
    try:
        # 获取文件内容
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
        
        # 解析DOCX文件
        result = parse_docx(file_bytes)
        logger.info(f"Successfully parsed DOCX, found {len(result.get('paragraphs', []))} paragraphs")
        import json
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as e:
        error_msg = f"parse_docx_handler error: {e}"
        logger.error(error_msg)
        return f"Error: {str(e)}"

@mcp.tool()
def parse_xlsx_handler(
    file_url: Optional[str] = None,
    file_bytes_b64: Optional[str] = None
) -> str:
    """
    解析 XLSX 文件，支持 file_url 或 base64，返回结构化 JSON。
    注意：此工具函数仅支持解析 XLSX 格式文件，不支持 PPTX 或 DOCX。
    
    调用步骤：
    1. 文件类型检查：
       - 确保待解析的文件是 XLSX 格式
       - 可以通过文件扩展名或文件头部特征进行判断
       
    2. 文件获取方式：
       - URL方式：提供 file_url 参数，指向可下载的XLSX文件
       - Base64方式：提供 file_bytes_b64 参数，包含XLSX文件的base64编码内容
       
    3. 错误处理：
       - 如果文件不是XLSX格式，会抛出相应错误
       - 如果文件下载失败，会返回错误信息
       - 如果base64解码失败，会返回错误信息
    
    Args:
        file_url: XLSX文件的URL，与file_bytes_b64参数二选一
        file_bytes_b64: XLSX文件的base64内容，与file_url参数二选一
        
    Returns:
        结构化Excel内容的JSON字符串，包含：
        - sheets: 工作表列表，每个工作表包含：
          - title: 工作表名称
          - cells: 单元格内容和坐标
          - formulas: 公式列表
        
    错误返回示例：
        - "Error: Failed to download file from url: {url}"
        - "Error: Missing parameter: file_url or file_bytes_b64"
        - "Error: Invalid file format, only XLSX files are supported"
    """
    try:
        # 获取文件内容
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
        
        # 解析XLSX文件
        result = parse_xlsx(file_bytes)
        logger.info(f"Successfully parsed XLSX, found {len(result.get('sheets', []))} sheets")
        import json
        return json.dumps(result, ensure_ascii=False, indent=2)
    except Exception as e:
        error_msg = f"parse_xlsx_handler error: {e}"
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