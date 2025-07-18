import base64
import json
import unittest
from jsonrpcserver import method, dispatch
from parser import parse_pptx

# Mock pptx内容
MOCK_PPTX_BYTES = b"FakePPTXContent"
MOCK_PPTX_B64 = base64.b64encode(MOCK_PPTX_BYTES).decode()

class TestMCPServer(unittest.TestCase):
    def test_initialize(self):
        req = json.dumps({"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {}})
        resp = dispatch(req)
        self.assertIn('protocolVersion', str(resp))
        self.assertIn('serverInfo', str(resp))

    def test_parse_pptx_handler(self):
        req = json.dumps({
            "jsonrpc": "2.0", "id": 2, "method": "parse_pptx_handler",
            "params": {"file_bytes_b64": MOCK_PPTX_B64}
        })
        resp = dispatch(req)
        self.assertIn('slides', str(resp))

    def test_method_not_found(self):
        req = json.dumps({"jsonrpc": "2.0", "id": 3, "method": "not_exist", "params": {}})
        resp = dispatch(req)
        self.assertIn('Method not found', str(resp))

if __name__ == "__main__":
    unittest.main() 