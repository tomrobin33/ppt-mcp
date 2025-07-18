import sys
import base64
import asyncio
from jsonrpcserver import method, async_dispatch as dispatch
from parser import parse_pptx

@method
async def parse_pptx_handler(file_bytes_b64: str):
    file_bytes = base64.b64decode(file_bytes_b64)
    return parse_pptx(file_bytes)

if __name__ == "__main__":
    async def main():
        for line in sys.stdin:
            line = line.strip()
            if not line:
                continue
            response = await dispatch(line)
            if response is not None:
                print(response, flush=True)
    asyncio.run(main()) 