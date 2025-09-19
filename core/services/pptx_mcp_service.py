#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
PPTX MCP 服务实现（后端逻辑）

输入：Markdown字符串
输出：上传到 MinIO 的 PPTX 文件 URL
"""

import os
import tempfile
import uuid
from typing import Optional

from core.pptx_engine import PPTXBuilder
from core.minio_service import MinIOService
from core.pptx_engine.logger import get_logger


class PPTXMCPService:
    """PPTX MCP服务类（纯后端实现）"""

    def __init__(self, template_path: Optional[str] = None, enable_logging: bool = False):
        self.logger = get_logger()

        # 模板路径
        if template_path is None:
            project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
            self.template_path = os.path.join(project_root, "config", "pptx_template.pptx")
        else:
            self.template_path = template_path

        if not os.path.exists(self.template_path):
            raise FileNotFoundError(f"PPTX模板文件不存在: {self.template_path}")

        # MinIO
        self.minio_service = MinIOService()
        self.logger.log_success("MinIO服务初始化成功", "MCPService")

        # 日志控制开关（如需可在此读取并覆写配置文件，但当前保持全局配置）
        self.enable_logging = enable_logging

    def convert_md_to_pptx_url(self, md_content: str, filename: Optional[str] = None) -> str:
        """将Markdown内容转换为PPTX并上传到MinIO，返回URL"""
        # 生成文件名
        if not filename:
            filename = f"presentation_{uuid.uuid4().hex[:8]}.pptx"
        if not filename.endswith(".pptx"):
            filename += ".pptx"

        # 临时文件
        with tempfile.NamedTemporaryFile(mode='w', suffix='.md', delete=False, encoding='utf-8') as temp_md:
            temp_md.write(md_content)
            temp_md_path = temp_md.name

        with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as temp_pptx:
            temp_pptx_path = temp_pptx.name

        try:
            self.logger.log_progress(f"开始转换Markdown到PPTX: {filename}")

            builder = PPTXBuilder()
            result_path = builder.from_md(temp_md_path, self.template_path, temp_pptx_path)

            if not os.path.exists(result_path):
                raise RuntimeError("PPTX文件生成失败")

            file_size = os.path.getsize(result_path)
            self.logger.log_success(f"PPTX文件生成成功: {file_size:,} 字节", "MCPService")

            # 上传 MinIO
            self.logger.log_progress(f"开始上传文件到MinIO: {filename}")
            minio_url = self.minio_service.upload_file(result_path, object_name=filename)
            self.logger.log_success(f"文件上传成功: {minio_url}", "MCPService")
            return minio_url

        finally:
            # 清理临时文件
            try:
                if os.path.exists(temp_md_path):
                    os.unlink(temp_md_path)
                if os.path.exists(temp_pptx_path):
                    os.unlink(temp_pptx_path)
            except Exception as e:
                self.logger.log_warning(f"清理临时文件失败: {e}", "MCPService")


_global_service = None


def get_service(template_path: Optional[str] = None, enable_logging: bool = False) -> PPTXMCPService:
    global _global_service
    if _global_service is None:
        _global_service = PPTXMCPService(template_path, enable_logging)
    return _global_service


