#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
基于 fastmcp 的 MCP 服务器

提供一个工具：将 Markdown 字符串转换为 PPTX，并上传至 MinIO，返回可访问的 URL。
"""

import os
import sys
import logging
import socket
from typing import Optional

from fastmcp import FastMCP

# 兼容直接运行：将项目根目录加入 sys.path，避免找不到 core 包
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from core.services.pptx_mcp_service import get_service


# 服务器实例（FastMCP 2.x 推荐用法）
mcp = FastMCP("AI-Office-PPTX-MCP")


@mcp.tool(
    name="md_to_minio_url",
    description="将 Markdown 字符串转换为 PPTX 并上传到 MinIO，返回可访问的 URL。",
)
def md_to_minio_url(
    md_content: str,
    filename: Optional[str] = None,
    template_path: Optional[str] = None,
    enable_logging: bool = False,
) -> str:
    """将 Markdown 转 PPTX 并上传到 MinIO。

    参数:
        md_content: Markdown 内容字符串
        filename: 可选，自定义输出文件名（无需 .pptx 后缀）
        template_path: 可选，PPTX 模板路径；默认使用 config/pptx_template.pptx
        enable_logging: 是否启用详细日志
    返回:
        MinIO 上文件的可访问 URL 字符串
    """
    logger = logging.getLogger("mcp")
    logger.info("[tool] md_to_minio_url called: filename=%s, md_length=%s", filename, len(md_content) if md_content else 0)
    service = get_service(template_path=template_path, enable_logging=enable_logging)
    url = service.convert_md_to_pptx_url(md_content, filename)
    logger.info("[tool] md_to_minio_url completed: url=%s", url)
    return url


def _detect_local_ip() -> str:
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.connect(("8.8.8.8", 80))
        ip = sock.getsockname()[0]
        sock.close()
        return ip
    except Exception:
        try:
            return socket.gethostbyname(socket.gethostname())
        except Exception:
            return "127.0.0.1"


def main() -> None:
    """本地运行入口。"""
    # 日志设置
    log_level_name = os.getenv("FASTMCP_LOG_LEVEL", "INFO").upper()
    level = getattr(logging, log_level_name, logging.INFO)
    logging.basicConfig(level=level, format="%(asctime)s %(levelname)s %(name)s - %(message)s")
    logger = logging.getLogger("mcp")

    # 运行参数（通过环境变量可配置）- 按 Dify/SSE 文章示例优化
    transport = os.getenv("FASTMCP_TRANSPORT", "sse").lower()
    host = os.getenv("FASTMCP_HOST", "0.0.0.0")
    try:
        port = int(os.getenv("FASTMCP_PORT", "8099"))
    except ValueError:
        port = 8099

    local_ip = _detect_local_ip()

    logger.info("[startup] FastMCP starting… name=%s", "AI-Office-PPTX-MCP")
    logger.info("[startup] transport=%s host=%s port=%s local_ip=%s", transport, host, port, local_ip)

    # 友好提示（对照教程风格）
    http_base = f"http://{host}:{port}"
    print("启动 AI-Office-PPTX-MCP 服务…")
    print(f"服务器地址: {http_base}")
    print(f"SSE 端点: http://localhost:{port}/sse")
    print(f"Docker 中的 Dify 连接: http://host.docker.internal:{port}/sse")
    print("可用工具:")
    print("- md_to_minio_url: 将 Markdown 转为 PPTX 并上传 MinIO，返回 URL")

    # FastMCP 在不同传输下的运行方式
    try:
        if transport == "http":
            if hasattr(mcp, "run_http"):
                try:
                    mcp.run_http(host=host, port=port, stream=True)
                except TypeError:
                    logger.warning("[startup] run_http(stream=True) 不被支持，回退 run_http(host, port)")
                    mcp.run_http(host=host, port=port)
            else:
                logger.warning("[startup] 未找到 run_http，回退 run(host, port)")
                mcp.run(host=host, port=port)
        elif transport in ("sse", "ws", "websocket"):
            # 兼容 fastmcp.run 直接以 host/port 暴露 SSE/WS
            try:
                mcp.run(transport=transport, host=host, port=port)
            except TypeError:
                mcp.run(host=host, port=port)
        else:
            mcp.run()
    except TypeError:
        logger.warning("[startup] run() 不支持 host/port 形参，降级为无参运行")
        mcp.run()


if __name__ == "__main__":
    main()


