FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt ./
COPY packages ./packages
RUN pip install --no-index --find-links=packages -r requirements.txt
COPY . .
EXPOSE 8000
CMD ["python", "mcp_server.py"] 