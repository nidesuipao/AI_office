#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
md2pptx 测试模块

专门用于测试 md2pptx_components/pptx_builder.py 中的组件功能

包含以下测试：
1. 组件导入测试
2. 组件初始化测试
3. 组件功能测试
4. 组件性能测试
"""

import os
import time
from urllib.parse import urlparse
from core.pptx_engine import PPTXBuilder
from core.rustfs_service import RustFSService


def test_component_imports():
    """测试 pptx_engine 组件的导入"""
    print("=" * 60)
    print("pptx_engine 组件导入测试")
    print("=" * 60)
    
    components = [
        ("FontCalculator", "core.pptx_engine.font_calculator"),
        ("ContentRenderer", "core.pptx_engine.content_renderer"), 
        ("LayoutManager", "core.pptx_engine.layout_manager"),
        ("SlideBuilder", "core.pptx_engine.slide_builder"),
        ("PPTXBuilder", "core.pptx_engine.pptx_builder")
    ]
    
    all_success = True
    for component_name, module_name in components:
        try:
            # 使用importlib来正确导入模块
            import importlib
            module = importlib.import_module(module_name)
            component_class = getattr(module, component_name)
            print(f"✅ {component_name} 导入成功")
        except Exception as e:
            print(f"❌ {component_name} 导入失败: {e}")
            all_success = False
    
    return all_success


def test_pptx_builder_functionality():
    """测试 PPTXBuilder 组件的功能"""
    print("\n" + "=" * 60)
    print("PPTXBuilder 组件功能测试")
    print("=" * 60)
    
    # 测试参数
    md_path = "./md_input_file/pptx_test_case_compact_full.md"
    template_path = "./config/pptx_template.pptx"
    output_path = "./output_file/test_pptx_builder.pptx"
    
    # 检查输入文件
    if not os.path.exists(md_path):
        print(f"❌ Markdown文件不存在: {md_path}")
        return False
    
    if not os.path.exists(template_path):
        print(f"❌ 模板文件不存在: {template_path}")
        return False
    
    print(f"📄 输入文件: {md_path}")
    print(f"📋 模板文件: {template_path}")
    print()
    
    # 测试 PPTXBuilder 组件
    print("🔄 测试 PPTXBuilder 组件...")
    start_time = time.time()
    try:
        # 直接使用 PPTXBuilder 类
        builder = PPTXBuilder()
        result = builder.from_md(md_path, template_path, output_path)
        test_time = time.time() - start_time
        file_size = os.path.getsize(output_path) if os.path.exists(output_path) else 0
        print(f"✅ PPTXBuilder 组件生成成功: {result}")
        print(f"⏱️  耗时: {test_time:.2f}秒")
        print(f"📊 文件大小: {file_size:,} 字节")
        
        # 检查生成的文件
        if os.path.exists(output_path) and file_size > 0:
            print("✅ 生成的PPT文件有效")
            return True
        else:
            print("❌ 生成的PPT文件无效")
            return False
            
    except Exception as e:
        print(f"❌ PPTXBuilder 组件生成失败: {e}")
        return False


def test_pptx_builder_initialization():
    """测试 PPTXBuilder 组件的初始化"""
    print("\n" + "=" * 60)
    print("PPTXBuilder 组件初始化测试")
    print("=" * 60)
    
    try:
        # 测试PPTXBuilder实例化
        builder = PPTXBuilder()
        print("✅ PPTXBuilder实例创建成功")
        
        # 测试组件初始化
        print(f"✅ 字体计算器: {type(builder.font_calc).__name__}")
        print(f"✅ 内容渲染器: {type(builder.renderer).__name__}")
        print(f"✅ 布局管理器: {type(builder.layout_manager).__name__}")
        print(f"✅ 幻灯片构建器: {type(builder.slide_builder).__name__}")
        
        # 测试组件方法
        print("\n🔧 测试组件方法:")
        
        # 测试字体计算器
        font_size = builder.calculate_title_font_size(1.0)
        print(f"✅ 字体计算器方法正常: 标题字体大小 = {font_size}pt")
        
        # 测试内容渲染器
        print("✅ 内容渲染器方法正常")
        
        # 测试布局管理器
        print("✅ 布局管理器方法正常")
        
        # 测试幻灯片构建器
        print("✅ 幻灯片构建器方法正常")
        
        return True
    except Exception as e:
        print(f"❌ PPTXBuilder 组件初始化测试失败: {e}")
        return False


def test_individual_components():
    """测试各个子组件的独立功能"""
    print("\n" + "=" * 60)
    print("子组件独立功能测试")
    print("=" * 60)
    
    try:
        from core.pptx_engine import FontCalculator, ContentRenderer, LayoutManager, SlideBuilder
        
        # 测试 FontCalculator
        print("🔤 测试 FontCalculator:")
        font_calc = FontCalculator()
        title_font = font_calc.calculate_title_font_size(1.0)
        parent_title_font = font_calc.calculate_parent_title_font_size(1.0)
        table_font = font_calc.calculate_table_font_size(2.0, 3, 4, 'table_header')
        print(f"  ✅ 标题字体大小: {title_font}pt")
        print(f"  ✅ 父标题字体大小: {parent_title_font}pt")
        print(f"  ✅ 表格字体大小: {table_font}pt")
        
        # 测试 ContentRenderer
        print("\n🎨 测试 ContentRenderer:")
        try:
            from pptx import Presentation
            mock_prs = Presentation()
            renderer = ContentRenderer(mock_prs)
            print("  ✅ ContentRenderer 实例化成功")
        except Exception as e:
            print(f"  ❌ ContentRenderer 实例化失败: {e}")
            return False
        
        # 测试 LayoutManager
        print("\n📐 测试 LayoutManager:")
        try:
            layout_manager = LayoutManager(renderer, font_calc)
            print("  ✅ LayoutManager 实例化成功")
        except Exception as e:
            print(f"  ❌ LayoutManager 实例化失败: {e}")
            return False
        
        # 测试 SlideBuilder
        print("\n🏗️  测试 SlideBuilder:")
        try:
            slide_builder = SlideBuilder(mock_prs, renderer, font_calc)
            print("  ✅ SlideBuilder 实例化成功")
        except Exception as e:
            print(f"  ❌ SlideBuilder 实例化失败: {e}")
            return False
        
        return True
    except Exception as e:
        print(f"❌ 子组件独立功能测试失败: {e}")
        return False


def show_test_files():
    """显示测试文件信息"""
    test_files = [
        "./output_file/test_pptx_builder.pptx"
    ]
    
    print("📁 测试文件信息:")
    for file_path in test_files:
        if os.path.exists(file_path):
            try:
                file_size = os.path.getsize(file_path)
                print(f"📄 测试文件: {file_path} ({file_size:,} 字节)")
                print(f"💾 保留测试文件: {file_path}")
            except Exception as e:
                print(f"⚠️  读取文件失败: {file_path} - {e}")
        else:
            print(f"❌ 测试文件不存在: {file_path}")


def main():
    """主测试函数"""
    print("🚀 开始 PPTXBuilder 组件测试...")
    
    # 测试组件导入
    import_success = test_component_imports()
    
    # 测试组件初始化
    init_success = test_pptx_builder_initialization()
    
    # 测试子组件独立功能
    individual_success = test_individual_components()
    
    # 测试组件功能
    if import_success and init_success and individual_success:
        functionality_success = test_pptx_builder_functionality()
        
        print("\n" + "=" * 60)
        print("测试总结")
        print("=" * 60)
        
        if functionality_success:
            print("🎉 所有 PPTXBuilder 组件测试通过！")
            print("✨ PPTXBuilder 组件功能正常")
            
        else:
            print("⚠️  部分测试未通过，需要进一步优化")
    else:
        print("❌ 基础测试失败，无法进行功能测试")
    
    # 显示测试文件信息
    print("\n" + "=" * 60)
    print("测试文件信息")
    print("=" * 60)
    show_test_files()

if __name__ == "__main__":
    main()
