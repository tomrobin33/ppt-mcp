# PPTX 结构化解析服务（支持 HTTP & MCP 协议）

## 项目简介
本项目是一个用于解析 .pptx（PowerPoint）文件并将其内容结构化为 JSON 的服务，支持两种协议：
- **HTTP（FastAPI）**：通过 RESTful API 上传 pptx 文件，返回结构化 JSON。
- **MCP（JSON-RPC over stdio）**：通过标准输入输出（stdio）协议，适合集成到编辑器、自动化工具、AI 编排等场景。

适用于：
- 教育、办公、AI、内容审核等场景下的 PPT 内容结构化提取
- 需要批量处理 PPT 文件的自动化系统
- 需要通过 HTTP 或进程间通信（stdio）调用的多种平台

## 功能特性
- 支持解析 .pptx 文件，提取每张幻灯片的所有文本内容（包括分组、表格等）
- 输出结构化 JSON，便于后续处理
- 支持 HTTP 文件上传接口和 MCP stdio 协议
- 可容器化部署，易于分享和集成

## 目录结构
```
ppt/
├── app.py            # FastAPI HTTP 服务主程序
├── parser.py         # PPTX 解析核心逻辑
├── mcp_server.py     # MCP (JSON-RPC over stdio) 服务主程序
├── requirements.txt  # 依赖清单
├── Dockerfile        # 容器部署文件
├── __init__.py       # 包初始化
└── README.md         # 项目说明文档
```

## 安装与环境准备
建议使用 Python 3.12 及虚拟环境：
```bash
# 安装依赖
python3.12 -m venv venv
source venv/bin/activate
pip install -r ppt/requirements.txt
```

## 使用方法
### 1. HTTP (FastAPI) 服务
- 启动服务：
  ```bash
  source venv/bin/activate
  venv/bin/uvicorn ppt.app:app --reload
  ```
- 访问接口文档：http://127.0.0.1:8000/docs
- 示例调用：
  ```bash
  curl -F "file=@你的文件.pptx" http://127.0.0.1:8000/parse-ppt
  ```

### 2. MCP (JSON-RPC over stdio) 服务
- 启动服务：
  ```bash
  source venv/bin/activate
  venv/bin/python ppt/mcp_server.py
  ```
- Python 客户端调用示例：
  ```python
  import sys, json, base64
  with open('your.pptx', 'rb') as f:
      file_bytes_b64 = base64.b64encode(f.read()).decode()
  req = {
      "jsonrpc": "2.0",
      "id": 1,
      "method": "parse_pptx",
      "params": {"file_bytes_b64": file_bytes_b64}
  }
  sys.stdout.write(json.dumps(req)+'\n')
  sys.stdout.flush()
  # 读取并解析返回的 JSON
  ```

## JSON 输出格式示例
```
{
  "slides": [
    {
      "slide_index": 1,
      "text": ["标题", "副标题", "段落内容"]
    },
    ...
  ]
}
```

## 容器化部署
```bash
docker build -t pptx-parser .
docker run -p 8000:8000 pptx-parser
```

## 依赖
- fastapi
- uvicorn
- python-pptx
- python-multipart
- python-jsonrpc-server

## 许可证
MIT 