#!/usr/bin/env python3
"""
简单的Markdown到DOCX转换器
专门用于转换 /home/yzy/document/project/AI_office-main/md_input_file/input.md 文件
"""

from md2docx import MarkdownToDocxConverter
import os

def convert_input_md():
    """转换input.md文件为DOCX"""
    input_file = './md_input_file/input.md'
    output_file = './output_file/output.docx'
    
    print("🔄 开始转换 input.md 文件...")
    print(f"📄 输入文件: {input_file}")
    print(f"📄 输出文件: {output_file}")
    
    # 检查输入文件是否存在
    if not os.path.exists(input_file):
        print(f"❌ 输入文件不存在: {input_file}")
        return False
    
    try:
        # 使用 md2docx 的转换器，遵循“从 h2 开始遍历”的实现
        converter = MarkdownToDocxConverter()
        converter.convert(input_file, output_file)
        
        print("✅ 转换成功！")
        # 检查输出文件
        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            print(f"📊 输出文件大小: {file_size:,} 字节")
            print(f"📄 输出文件: {output_file}")
            return True
        else:
            print("❌ 输出文件未生成")
            return False
            
    except Exception as e:
        print(f"❌ 转换过程中出错: {e}")
        return False

if __name__ == "__main__":
    convert_input_md()
