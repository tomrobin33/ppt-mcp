import sys
import base64
import asyncio
from jsonrpcserver import method, async_dispatch as dispatch, Success, Error
from parser import parse_pptx
import logging

logging.basicConfig(level=logging.INFO, format='[%(asctime)s] %(levelname)s: %(message)s')

@method
async def parse_pptx_handler(file_bytes_b64: str):
    try:
        file_bytes = base64.b64decode(file_bytes_b64)
        return Success(parse_pptx(file_bytes))
    except Exception as e:
        logging.error(f"parse_pptx_handler error: {e}")
        return Error(str(e))

@method
async def initialize(**kwargs):
    return Success({
        "protocolVersion": "1.0",
        "capabilities": {
            "parse_pptx_handler": True
        },
        "serverInfo": {
            "name": "ppt-mcp",
            "version": "1.0.0"
        }
    })

@method
async def health():
    return Success({"status": "ok"})

@method
async def version():
    return Success({"version": "1.0.0"})

if __name__ == "__main__":
    logging.info("MCP Server started, code version: 2025-07-19-01-unique")
    async def main():
        for line in sys.stdin:
            line = line.strip()
            if not line:
                continue
            try:
                response = await dispatch(line)
                if response is not None:
                    print(response, flush=True)
            except Exception as e:
                logging.error(f"dispatch error: {e}")
                print('{"jsonrpc": "2.0", "error": {"code": -32000, "message": "Internal error", "data": "%s"}}' % str(e), flush=True)
    asyncio.run(main()) 