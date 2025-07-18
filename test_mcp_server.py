import base64
import unittest
from mcp_server import parse_pptx_handler

MOCK_PPTX_BYTES = b"FakePPTXContent"
MOCK_PPTX_B64 = base64.b64encode(MOCK_PPTX_BYTES).decode()

class TestMCPServer(unittest.TestCase):
    def test_parse_pptx_handler_with_b64(self):
        resp = parse_pptx_handler(file_bytes_b64=MOCK_PPTX_B64)
        self.assertIn("slides", resp)

    def test_parse_pptx_handler_with_url(self):
        # 这里可以用一个无效URL测试错误分支
        resp = parse_pptx_handler(file_url="http://invalid-url.com/test.pptx")
        self.assertIn("code", resp)
        self.assertNotEqual(resp["code"], 0)

    def test_parse_pptx_handler_missing_param(self):
        resp = parse_pptx_handler()
        self.assertIn("code", resp)
        self.assertNotEqual(resp["code"], 0)

if __name__ == "__main__":
    unittest.main() 