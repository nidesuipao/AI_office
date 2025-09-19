# PPTX MCP 服务文档

## 概述

PPTX MCP服务是一个完整的Markdown到PPTX转换服务，集成了MinIO文件存储功能。该服务将复杂的测试流程整合为一个简单的API接口。

## 功能特性

- ✅ **Markdown到PPTX转换**：支持完整的Markdown语法
- ✅ **MinIO文件上传**：自动上传生成的PPTX文件到MinIO
- ✅ **模板支持**：使用预定义的PPTX模板
- ✅ **日志配置**：可配置的日志输出级别
- ✅ **临时文件管理**：自动清理临时文件
- ✅ **错误处理**：完善的异常处理机制

## 快速开始

### 1. 基本使用

```python
from api_pptx_converter import convert_md_to_pptx

# 准备Markdown内容
md_content = """# 我的演示文稿

## 第一章：介绍

这是一个演示文稿的内容。

### 主要特点

- 功能强大
- 易于使用
- 支持多种格式

| 特性 | 状态 |
|------|------|
| 转换 | ✅ |
| 上传 | ✅ |
"""

# 转换为PPTX并获取URL
url = convert_md_to_pptx(md_content, "my_presentation.pptx")
print(f"文件URL: {url}")
```

### 2. 高级使用

```python
from mcp_pptx_service import PPTXMCPService

# 创建服务实例
service = PPTXMCPService(
    template_path="./config/pptx_template.pptx",  # 自定义模板
    enable_logging=True  # 启用详细日志
)

# 转换文件
url = service.convert_md_to_pptx_url(md_content, "custom_name.pptx")

# 获取服务信息
info = service.get_service_info()
print(info)
```

## API 参考

### `convert_md_to_pptx(md_content, filename=None)`

将Markdown内容转换为PPTX并返回MinIO URL。

**参数：**
- `md_content` (str): Markdown内容字符串
- `filename` (str, optional): 输出文件名，默认自动生成

**返回：**
- `str`: MinIO文件URL

**示例：**
```python
url = convert_md_to_pptx("# 标题\n\n内容", "test.pptx")
```

### `PPTXMCPService` 类

完整的服务类，提供更多控制选项。

**初始化参数：**
- `template_path` (str, optional): PPTX模板文件路径
- `enable_logging` (bool): 是否启用详细日志

**方法：**
- `convert_md_to_pptx_url(md_content, filename=None)`: 转换Markdown到PPTX URL
- `get_service_info()`: 获取服务信息

## 配置说明

### 日志配置

日志配置通过 `config/pptx_log_config.yaml` 文件控制：

```yaml
# 启用详细日志
log_levels:
  component_init: true
  font_calculation: true
  layout_management: true
  content_rendering: true
  slide_building: true
  file_operations: true
  debug_details: true
  performance_stats: true

output_control:
  show_progress: true
  show_content_analysis: true
  show_layout_decisions: true
  show_slide_creation: true
  show_chapter_processing: true
```

### MinIO配置

MinIO服务配置在 `core/minio_service.py` 中：

```python
# 默认配置
endpoint = "172.27.0.1:19100"
bucket_name = "ai-office-test"
secure = False
```

## 支持的Markdown语法

### 标题
```markdown
# 一级标题
## 二级标题
### 三级标题
```

### 列表
```markdown
- 无序列表项1
- 无序列表项2
  - 子列表项

1. 有序列表项1
2. 有序列表项2
```

### 表格
```markdown
| 列1 | 列2 | 列3 |
|-----|-----|-----|
| 数据1 | 数据2 | 数据3 |
| 数据4 | 数据5 | 数据6 |
```

### 图片
```markdown
![图片描述](./image.jpg)
```

## 错误处理

服务包含完善的错误处理机制：

```python
try:
    url = convert_md_to_pptx(md_content)
    print(f"转换成功: {url}")
except FileNotFoundError as e:
    print(f"模板文件不存在: {e}")
except Exception as e:
    print(f"转换失败: {e}")
```

## 性能优化

### 禁用详细日志
```python
# 生产环境建议禁用详细日志
service = PPTXMCPService(enable_logging=False)
```

### 批量处理
```python
# 批量转换多个文件
results = []
for i, md_content in enumerate(markdown_files):
    url = convert_md_to_pptx(md_content, f"presentation_{i}.pptx")
    results.append(url)
```

## 文件结构

```
AI_office-main/
├── mcp_pptx_service.py          # 主服务类
├── api_pptx_converter.py        # 简化API接口
├── config/
│   ├── pptx_template.pptx       # PPTX模板
│   └── pptx_log_config.yaml     # 日志配置
├── core/
│   ├── pptx_engine/             # PPTX引擎
│   └── minio_service.py         # MinIO服务
└── MCP_SERVICE_README.md        # 本文档
```

## 测试

运行快速测试：

```bash
python api_pptx_converter.py
```

运行完整服务测试：

```bash
python mcp_pptx_service.py
```

## 故障排除

### 常见问题

1. **模板文件不存在**
   ```
   FileNotFoundError: PPTX模板文件不存在
   ```
   解决：确保 `config/pptx_template.pptx` 文件存在

2. **MinIO连接失败**
   ```
   MinIO服务初始化失败
   ```
   解决：检查MinIO服务是否运行，网络连接是否正常

3. **字体配置文件警告**
   ```
   警告: 字体配置文件不存在
   ```
   解决：这是正常警告，不影响功能，字体计算器会使用默认配置

### 调试模式

启用详细日志进行调试：

```python
service = PPTXMCPService(enable_logging=True)
url = service.convert_md_to_pptx_url(md_content)
```

## 更新日志

### v1.0.0
- 初始版本发布
- 支持Markdown到PPTX转换
- 集成MinIO文件上传
- 可配置日志系统
- 完善的错误处理

## 许可证

本项目采用MIT许可证。
