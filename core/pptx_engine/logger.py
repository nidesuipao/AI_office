#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
PPTX日志管理器

根据配置文件控制PPTX生成过程中的日志输出
"""

import os
import yaml
from typing import Dict, Any, Optional
from datetime import datetime


class PPTXLogger:
    """PPTX日志管理器"""
    
    def __init__(self, config_path: Optional[str] = None):
        """
        初始化日志管理器
        
        Args:
            config_path: 配置文件路径，默认使用config/pptx_log_config.yaml
        """
        if config_path is None:
            # 从core/pptx_engine/logger.py 到项目根目录，然后到config目录
            project_root = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
            config_path = os.path.join(project_root, "config", "pptx_log_config.yaml")
        
        self.config_path = config_path
        self.config = self._load_config()
        
    def _load_config(self) -> Dict[str, Any]:
        """加载配置文件"""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    return yaml.safe_load(f) or {}
            else:
                print(f"⚠️  日志配置文件不存在: {self.config_path}")
                return self._get_default_config()
        except Exception as e:
            print(f"⚠️  加载日志配置失败: {e}")
            return self._get_default_config()
    
    def _get_default_config(self) -> Dict[str, Any]:
        """获取默认配置"""
        return {
            'log_levels': {
                'component_init': True,
                'font_calculation': True,
                'layout_management': True,
                'content_rendering': True,
                'slide_building': True,
                'file_operations': True,
                'debug_details': False,
                'performance_stats': True
            },
            'output_control': {
                'show_progress': True,
                'show_content_analysis': True,
                'show_layout_decisions': True,
                'show_font_calculations': True,
                'show_image_processing': True,
                'show_table_processing': True
            },
            'log_format': {
                'timestamp_format': "%Y-%m-%d %H:%M:%S",
                'include_timestamp': False,
                'include_component_name': True,
                'use_colors': True,
                'prefix': "📊"
            }
        }
    
    def _get_log_prefix(self, component_name: str = "") -> str:
        """获取日志前缀"""
        prefix = self.config.get('log_format', {}).get('prefix', '📊')
        
        if self.config.get('log_format', {}).get('include_timestamp', False):
            timestamp = datetime.now().strftime(
                self.config.get('log_format', {}).get('timestamp_format', "%Y-%m-%d %H:%M:%S")
            )
            prefix = f"[{timestamp}] {prefix}"
        
        if self.config.get('log_format', {}).get('include_component_name', False) and component_name:
            prefix = f"{prefix} [{component_name}]"
        
        return prefix
    
    def _should_log(self, log_type: str) -> bool:
        """检查是否应该输出指定类型的日志"""
        return self.config.get('log_levels', {}).get(log_type, True)
    
    def _should_show(self, show_type: str) -> bool:
        """检查是否应该显示指定类型的信息"""
        return self.config.get('output_control', {}).get(show_type, True)
    
    def log(self, message: str, log_type: str = "general", component_name: str = ""):
        """输出日志"""
        if not self._should_log(log_type):
            return
        
        prefix = self._get_log_prefix(component_name)
        print(f"{prefix} {message}")
    
    def log_component_init(self, component_name: str, message: str):
        """组件初始化日志"""
        if self._should_log("component_init"):
            self.log(f"✅ {component_name}: {message}", "component_init", component_name)
    
    def log_font_calculation(self, message: str, component_name: str = "FontCalculator"):
        """字体计算日志"""
        if self._should_log("font_calculation") and self._should_show("show_font_calculations"):
            self.log(f"🔤 {message}", "font_calculation", component_name)
    
    def log_layout_management(self, message: str, component_name: str = "LayoutManager"):
        """布局管理日志"""
        if self._should_log("layout_management") and self._should_show("show_layout_decisions"):
            self.log(f"📐 {message}", "layout_management", component_name)
    
    def log_content_rendering(self, message: str, component_name: str = "ContentRenderer"):
        """内容渲染日志"""
        if self._should_log("content_rendering"):
            self.log(f"🎨 {message}", "content_rendering", component_name)
    
    def log_slide_building(self, message: str, component_name: str = "SlideBuilder"):
        """幻灯片构建日志"""
        if self._should_log("slide_building"):
            self.log(f"🏗️  {message}", "slide_building", component_name)
    
    def log_file_operations(self, message: str, component_name: str = "FileOps"):
        """文件操作日志"""
        if self._should_log("file_operations"):
            self.log(f"📁 {message}", "file_operations", component_name)
    
    def log_performance(self, message: str, component_name: str = "Performance"):
        """性能统计日志"""
        if self._should_log("performance_stats"):
            self.log(f"⏱️  {message}", "performance_stats", component_name)
    
    def log_debug(self, message: str, component_name: str = "Debug"):
        """调试日志"""
        if self._should_log("debug_details"):
            self.log(f"🔍 {message}", "debug_details", component_name)
    
    def log_progress(self, message: str):
        """进度日志"""
        if self._should_show("show_progress"):
            self.log(f"🔄 {message}", "general")
    
    def log_content_analysis(self, message: str):
        """内容分析日志"""
        if self._should_show("show_content_analysis"):
            self.log(f"📊 {message}", "general")
    
    def log_image_processing(self, message: str):
        """图片处理日志"""
        if self._should_show("show_image_processing"):
            self.log(f"🖼️  {message}", "content_rendering")
    
    def log_table_processing(self, message: str):
        """表格处理日志"""
        if self._should_show("show_table_processing"):
            self.log(f"📋 {message}", "content_rendering")
    
    def log_success(self, message: str, component_name: str = ""):
        """成功日志"""
        self.log(f"✅ {message}", "general", component_name)
    
    def log_warning(self, message: str, component_name: str = ""):
        """警告日志"""
        self.log(f"⚠️  {message}", "general", component_name)
    
    def log_error(self, message: str, component_name: str = ""):
        """错误日志"""
        self.log(f"❌ {message}", "general", component_name)
    
    def log_info(self, message: str, component_name: str = ""):
        """信息日志"""
        self.log(f"ℹ️  {message}", "general", component_name)
    
    def log_slide_creation(self, message: str, component_name: str = "SlideBuilder"):
        """幻灯片创建日志"""
        if self._should_show("show_slide_creation"):
            self.log(f"🏗️  {message}", "slide_building", component_name)
    
    def log_chapter_processing(self, message: str, component_name: str = "ChapterProcessor"):
        """章节处理日志"""
        if self._should_show("show_chapter_processing"):
            self.log(f"📖 {message}", "slide_building", component_name)


# 全局日志实例
_global_logger = None

def get_logger() -> PPTXLogger:
    """获取全局日志实例"""
    global _global_logger
    if _global_logger is None:
        _global_logger = PPTXLogger()
    return _global_logger

def set_logger(logger: PPTXLogger):
    """设置全局日志实例"""
    global _global_logger
    _global_logger = logger
