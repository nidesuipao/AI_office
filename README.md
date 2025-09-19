# AI Office（Markdown → PPTX + MinIO 上传）

一个将 Markdown 智能渲染为 PPTX 的实用项目，内置版式与字体计算、布局管理与内容渲染，并集成 MinIO 文件存储（生成可访问 URL）。同时提供 MCP 服务端，便于在工具生态中集成调用。

## 功能特性

- Markdown → PPTX 全流程转换（标题、段落、列表、表格、图片等）
- 版式/字体/布局智能计算与渲染（`core/pptx_engine/*`）
- MinIO 上传与 URL 返回（`core/minio_service.py`）
- 模板驱动（`config/pptx_template.pptx`）
- 可配置日志（`config/pptx_log_config.yaml`）
- 提供 MCP 服务端与服务封装（`api/pptx_mcp_server.py`、`core/services/pptx_mcp_service.py`）

## 目录结构（节选）

```
AI_office-main/
├── README.md
├── api/
│   └── pptx_mcp_server.py          # MCP 服务端入口
├── config/
│   ├── pptx_template.pptx          # PPTX 模板
│   ├── pptx_log_config.yaml        # 日志配置
│   └── pptx_font_config.yaml       # 字体配置
├── core/
│   ├── pptx_engine/                # 渲染/布局/字体/构建等引擎
│   ├── services/pptx_mcp_service.py# 服务封装
│   └── minio_service.py            # MinIO 上传
├── docs/                           # 文档
├── md_input_file/                  # 示例 Markdown
├── test_md2pptx.py                 # 转换测试脚本
└── requirements.txt
```

## 环境准备

- Python 3.9+
- 安装依赖：

```bash
pip install -r requirements.txt
```

## 最快上手

1) 将 `md_input_file/pptx_test_case_compact_full.md` 转为 PPTX（本地生成）

```bash
python test_md2pptx.py
```

2) 启动 MCP 服务端（用于工具集成）

```bash
python -m api.pptx_mcp_server
```

## 示例：在代码中调用服务封装

```python
from core.services.pptx_mcp_service import PPTXMCPService

service = PPTXMCPService(enable_logging=True)
url = service.convert_md_to_pptx_url("# 标题\n\n正文", filename="demo.pptx")
print(url)
```

## MinIO 配置

默认配置位于 `core/minio_service.py`，你也可以通过环境变量或构造参数传入：

- endpoint: 例如 172.27.0.1:19100
- bucket: 例如 ai-office-test
- secure: False（HTTP）或 True（HTTPS）

确保 MinIO 已创建对应 bucket，账户有写入权限。

## 日志配置

通过 `config/pptx_log_config.yaml` 控制各组件的日志粒度。开发态可开启更详细的组件级日志；生产态建议关闭 debug 级日志，仅保留必要信息。

## Docker（可选）

如需在容器环境中运行 MinIO 或相关依赖，可参考 `docker/docker-compose.yml`。PPTX 模板与日志配置文件已纳入版本控制，便于直接使用。

## 开发与测试

- 主要代码位于 `core/pptx_engine/*`
- 运行单元/集成测试：

```bash
python test_md2pptx.py
python test_md2docx.py  # 若需要验证 docx 相关功能
```

## 文档

- `docs/README.md`：文档导航/索引
- `docs/MCP_SERVICE_README.md`：MCP 服务文档
- `docs/INTEGRATION_SUMMARY.md`：集成与架构摘要
- `docs/日志配置测试总结.md`：日志配置与测试结论

## 许可证

MIT

—— 最后更新：2025-09-19
