import sys
import base64
from python_jsonrpc_server import serve
from parser import parse_pptx

def parse_pptx_handler(file_bytes_b64: str):
    file_bytes = base64.b64decode(file_bytes_b64)
    return parse_pptx(file_bytes)

if __name__ == "__main__":
    serve(methods={"parse_pptx": parse_pptx_handler}, input_=sys.stdin, output=sys.stdout) 