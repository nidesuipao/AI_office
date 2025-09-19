"""
字体计算器 - 负责智能计算各种场景下的字体大小
"""

import os
import yaml
from pptx.util import Pt


class FontCalculator:
    """字体大小计算器，根据内容类型、可用空间等智能计算最佳字体大小"""
    
    def __init__(self, config_path=None):
        # 加载配置文件
        self.config = self._load_config(config_path)
        
        # 基础字体大小设定（从配置文件加载，支持动态覆盖）
        self.base_sizes = self.config.get('base_sizes', {
            'parent_title': 26,   # 父标题（章节分隔页标题）
            'title': 20,          # 子章节标题
            'text': 18,           # 正文内容（中文友好大小）
            'table_header': 18,   # 表格标题（中文友好）
            'table_data': 16      # 表格数据（中文友好）
        })
        
        # 字体大小范围限制（从配置文件加载，支持动态覆盖）
        self.size_ranges = {}
        config_ranges = self.config.get('size_ranges', {})
        for key, value in config_ranges.items():
            if isinstance(value, list) and len(value) == 2:
                self.size_ranges[key] = tuple(value)
            else:
                self.size_ranges[key] = value
        
        # 设置默认范围
        if 'default' not in self.size_ranges:
            self.size_ranges['default'] = (14, 22)
    
    def _load_config(self, config_path=None):
        """加载字体配置文件"""
        if config_path is None:
            # 默认配置文件路径
            current_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            config_path = os.path.join(current_dir, 'config', 'pptx_font_config.yaml')
        
        try:
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    return yaml.safe_load(f) or {}
            else:
                print(f"警告: 字体配置文件不存在: {config_path}")
                return {}
        except Exception as e:
            print(f"警告: 加载字体配置文件失败: {e}")
            return {}
    
    def update_base_sizes(self, **kwargs):
        """动态更新基础字体大小（优先级高于配置文件）"""
        self.base_sizes.update(kwargs)
    
    def update_size_ranges(self, **kwargs):
        """动态更新字体大小范围（优先级高于配置文件）"""
        for key, value in kwargs.items():
            if isinstance(value, (list, tuple)) and len(value) == 2:
                self.size_ranges[key] = tuple(value)
            else:
                self.size_ranges[key] = value
    
    def calculate_optimal_font_size(self, available_height, content_amount, content_type):
        """根据可用空间和内容量计算最佳字体大小"""
        base_size = self.base_sizes.get(content_type, 16)
        
        # 从配置文件获取动态调整参数
        dynamic_config = self.config.get('dynamic_adjustment', {})
        height_multipliers = dynamic_config.get('height_multipliers', {
            'small': 0.75, 'medium': 0.9, 'large': 1.0, 'extra_large': 1.2
        })
        content_multipliers = dynamic_config.get('content_multipliers', {
            'few': 1.2, 'normal': 1.0, 'many': 0.9, 'too_many': 0.8
        })
        
        # 根据可用高度调整
        # 对于caption类型，使用更宽松的高度判断
        if content_type == 'caption':
            if available_height <= 0.5:  # 标题空间很小
                size_multiplier = height_multipliers.get('small', 0.75)
            elif available_height <= 1.0:  # 标题空间中等
                size_multiplier = height_multipliers.get('medium', 0.9)
            else:  # 标题空间充足
                size_multiplier = height_multipliers.get('large', 1.0)
        else:
            if available_height <= 1.5:  # 空间很小
                size_multiplier = height_multipliers.get('small', 0.75)
            elif available_height <= 3.0:  # 空间中等
                size_multiplier = height_multipliers.get('medium', 0.9)
            elif available_height <= 5.0:  # 空间充足
                size_multiplier = height_multipliers.get('large', 1.0)
            else:  # 空间很大
                size_multiplier = height_multipliers.get('extra_large', 1.2)
        
        # 根据内容量调整
        if content_type == 'text':
            if content_amount <= 2:  # 内容很少
                content_multiplier = content_multipliers.get('few', 1.2)
            elif content_amount <= 5:  # 内容适中
                content_multiplier = content_multipliers.get('normal', 1.0)
            elif content_amount <= 10:  # 内容较多
                content_multiplier = content_multipliers.get('many', 0.9)
            else:  # 内容很多
                content_multiplier = content_multipliers.get('too_many', 0.8)
        else:
            content_multiplier = 1.0
        
        # 计算最终字体大小
        final_size = int(base_size * size_multiplier * content_multiplier)
        
        # 应用大小范围限制
        min_size, max_size = self.size_ranges.get(content_type, self.size_ranges['default'])
        final_size = max(min_size, min(final_size, max_size))
        
        return final_size
    
    def calculate_title_font_size(self, available_height):
        """计算标题字体大小"""
        return self.calculate_optimal_font_size(available_height, 1, 'title')
    
    def calculate_parent_title_font_size(self, available_height):
        """计算父标题字体大小"""
        return self.calculate_optimal_font_size(available_height, 1, 'parent_title')
    
    def calculate_table_font_size(self, table_height, rows, cols, content_type):
        """根据表格尺寸计算最优字体大小，让文字占满大部分空间（中文友好）"""
        base_size = self.base_sizes.get(content_type, 18 if content_type == 'table_header' else 16)
        
        # 从配置文件获取表格调整参数
        dynamic_config = self.config.get('dynamic_adjustment', {})
        table_config = dynamic_config.get('table_adjustment', {})
        cell_height_ratio = table_config.get('cell_height_ratio', 0.6)
        base_size_multiplier = table_config.get('base_size_multiplier', 1.5)
        col_adjustments = table_config.get('col_adjustments', {
            'normal': 1.0, 'many': 0.9, 'too_many': 0.8
        })
        
        # 计算单元格平均高度
        cell_height = table_height / rows if rows > 0 else 1.0
        
        # 根据单元格高度计算最优字体大小（字体高度约为字号的1.2倍）
        max_font_by_height = int((cell_height * 72 * cell_height_ratio))  # 英寸转点
        
        # 根据列数调整（列太多时字体要小一些）
        if cols > 5:
            col_adjustment = col_adjustments.get('too_many', 0.8)
        elif cols > 3:
            col_adjustment = col_adjustments.get('many', 0.9)
        else:
            col_adjustment = col_adjustments.get('normal', 1.0)
        
        # 计算最终字体大小
        optimal_size = int(min(base_size * base_size_multiplier, max_font_by_height) * col_adjustment)
        
        # 设置表格字体的合理范围（中文友好）
        min_size, max_size = self.size_ranges.get(content_type, self.size_ranges['default'])
        final_size = max(min_size, min(optimal_size, max_size))
        
        return final_size
    
    def get_font_name(self, font_type='default'):
        """获取字体名称"""
        base_font = self.config.get('base_font', {})
        return base_font.get('name', '微软雅黑')
    
    def get_font_color(self, color_type='default'):
        """获取字体颜色"""
        colors = self.config.get('colors', {})
        return colors.get(color_type, '#000000')
    
    def get_font_style(self, style_type='default'):
        """获取字体样式"""
        styles = self.config.get('styles', {})
        return styles.get(style_type, {'bold': False, 'italic': False})
    
    def get_text_estimation_config(self):
        """获取文本估算配置"""
        return self.config.get('text_estimation', {
            'line_height_ratio': 1.2,
            'gap_list': 6,
            'gap_paragraph': 8,
            'min_chars_per_line': 8
        })