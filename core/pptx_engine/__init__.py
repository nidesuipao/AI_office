"""
md2pptx_components - PPTX构建器组件包

包含以下组件：
- FontCalculator: 字体计算器
- ContentRenderer: 内容渲染器
- LayoutManager: 布局管理器
- SlideBuilder: 幻灯片构建器
- PPTXBuilder: 主构建器类
"""

from .font_calculator import FontCalculator
from .content_renderer import ContentRenderer
from .layout_manager import LayoutManager
from .slide_builder import SlideBuilder
from .pptx_builder import PPTXBuilder

__all__ = ['FontCalculator', 'ContentRenderer', 'LayoutManager', 'SlideBuilder', 'PPTXBuilder']
