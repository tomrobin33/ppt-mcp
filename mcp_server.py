import sys
import os
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))
import base64
import asyncio
from jsonrpcserver import method, async_dispatch as dispatch
from parser import parse_pptx
import logging

logging.basicConfig(level=logging.INFO, format='[%(asctime)s] %(levelname)s: %(message)s')

@method
def parse_pptx_handler(**kwargs):
    try:
        file_bytes_b64 = kwargs.get("file_bytes_b64")
        if not file_bytes_b64:
            return {"code": -32602, "message": "Missing parameter: file_bytes_b64"}
        file_bytes = base64.b64decode(file_bytes_b64)
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

@method
def health(**kwargs):
    return {"status": "ok"}

@method
def version(**kwargs):
    return {"version": "1.0.0"}

if __name__ == "__main__":
    logging.info("MCP Server started, code version: 2025-07-19-01-unique")
    async def main():
        for line in sys.stdin:
            logging.info(f"Received: {line}")
            line = line.strip()
            if not line:
                continue
            try:
                response = await dispatch(line)
                logging.info(f"Dispatch response: {response}")
                if response is not None:
                    print(response, flush=True)
            except Exception as e:
                logging.error(f"dispatch error: {e}")
                print('{"jsonrpc": "2.0", "error": {"code": -32000, "message": "Internal error", "data": "%s"}}' % str(e), flush=True)
    asyncio.run(main()) 