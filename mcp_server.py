import sys
import os
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))
import base64
from jsonrpcserver import method, dispatch
from parser import parse_pptx
import logging
import requests

logging.basicConfig(level=logging.INFO, format='[%(asctime)s] %(levelname)s: %(message)s')

@method
def parse_pptx_handler(**kwargs):
    try:
        file_url = kwargs.get("file_url")
        file_bytes_b64 = kwargs.get("file_bytes_b64")
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

@method(name="initialize")
def initialize(**kwargs):
    protocol_version = kwargs.get("protocolVersion", "2024-11-05")
    return {
        "protocolVersion": protocol_version,
        "capabilities": {
            "parse_pptx_handler": True
        },
        "serverInfo": {
            "name": "ppt-mcp",
            "version": "1.0.0"
        }
    }

@method(name="tools/list")
def tools_list(**kwargs):
    try:
        logging.info("tools/list called, kwargs: %s", kwargs)
        result = {
            "tools": [
                {
                    "name": "parse_pptx_handler",
                    "description": "解析 PPTX 文件，支持 file_url 或 base64，返回结构化 JSON",
                    "parameters": [
                        {
                            "name": "file_url",
                            "type": "string",
                            "required": False,
                            "description": "PPTX文件的URL"
                        },
                        {
                            "name": "file_bytes_b64",
                            "type": "string",
                            "required": False,
                            "description": "PPTX文件的base64内容"
                        }
                    ],
                    "returns": {
                        "type": "object",
                        "description": "结构化PPT内容",
                        "example": {
                            "slides": [
                                {"slide_index": 1, "text": ["标题", "内容"]}
                            ]
                        }
                    }
                }
            ]
        }
        logging.info("tools/list result: %s", result)
        return result
    except Exception as e:
        logging.error(f"tools/list error: {e}")
        return {"code": -32000, "message": str(e)}

@method
def health(**kwargs):
    return {"status": "ok"}

@method
def version(**kwargs):
    return {"version": "1.0.0"}

if __name__ == "__main__":
    logging.info("MCP Server started, code版本: 2025-07-19-01-unique")
    def main():
        for line in sys.stdin:
            logging.info(f"Received: {line}")
            line = line.strip()
            if not line:
                continue
            try:
                response = dispatch(line)
                logging.info(f"Dispatch response: {response}")
                if response is not None:
                    print(response, flush=True)
            except Exception as e:
                logging.error(f"dispatch error: {e}")
                print('{"jsonrpc": "2.0", "error": {"code": -32000, "message": "Internal error", "data": "%s"}}' % str(e), flush=True)
    main() 