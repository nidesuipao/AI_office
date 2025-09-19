"""
布局管理器 - 负责各种布局策略的选择和执行
"""

from .content_renderer import ContentRenderer
from .font_calculator import FontCalculator


class LayoutManager:
    """布局管理器，负责根据内容类型选择和执行最佳布局策略"""
    
    def __init__(self, content_renderer: ContentRenderer, font_calculator: FontCalculator):
        self.renderer = content_renderer
        self.font_calc = font_calculator
    
    def add_content_auto_layout(self, slide, content_blocks, content_top, content_height):
        """智能布局自动匹配：根据内容类型选择最佳布局"""
        if not content_blocks:
            return
        
        # 分析内容类型
        text_blocks = [block for block in content_blocks if block['type'] in ['list', 'paragraph']]
        table_blocks = [block for block in content_blocks if block['type'] == 'table']
        image_blocks = [block for block in content_blocks if block['type'] == 'image']
        
        # 计算内容量
        text_content_amount = sum(len(block.get('items', [])) if block['type'] == 'list' 
                                 else 1 if block['type'] == 'paragraph' else 0 
                                 for block in text_blocks)
        table_count = len(table_blocks)
        image_count = len(image_blocks)
        
        print(f"    内容分析: {text_content_amount}条文字, {table_count}个表格, {image_count}张图片")
        
        # 根据内容类型选择布局策略
        # 仅图片
        if image_count > 0 and table_count == 0 and text_content_amount == 0:
            self.layout_images_only(slide, image_blocks, content_top, content_height)
            return

        # 文字 + 图片（常见）
        if image_count > 0 and table_count == 0 and text_content_amount > 0:
            self.layout_text_and_images(slide, text_blocks, image_blocks, content_top, content_height)
            return

        # 表格 + 图片
        if image_count > 0 and table_count == 1 and text_content_amount == 0:
            self.layout_table_and_images(slide, table_blocks[0], image_blocks, content_top, content_height)
            return

        # 文字 + 表格 + 图片
        if image_count > 0 and table_count >= 1 and text_content_amount > 0:
            self.layout_text_table_images(slide, text_blocks, table_blocks[0], image_blocks, content_top, content_height)
            return

        # 只有表格
        if table_count == 1 and text_content_amount == 0:
            self.layout_table_only(slide, table_blocks[0], content_top, content_height)
            return

        # 文字 + 表格
        if table_count == 1 and text_content_amount > 0:
            self.layout_text_and_table(slide, text_blocks, table_blocks[0], content_top, content_height)
            return

        # 多表格或复杂内容（无图片）
        self.layout_complex_content(slide, text_blocks, table_blocks, content_top, content_height)
    
    def layout_images_only(self, slide, image_blocks, content_top, content_height):
        """布局1: 纯图片内容"""
        print(f"    应用布局: 纯图片布局")
        # 图片区域占用大部分空间
        images_top = content_top + 0.1
        images_height = min(content_height * 0.85, content_height)
        
        # 直接使用insert_image插入图片，确保高度不超过可用区域
        try:
            max_images = 3
            imgs = (image_blocks or [])[:max_images]
            count = len(imgs)
            if count == 0:
                return
            
            total_width = 11.0
            gap = 0.3
            idx = self.renderer._get_slide_index(slide)
            
            if count == 1:
                # 单张图片：高度优先，假设16:9比例
                assumed_aspect = 16.0 / 9.0
                width_each = min(total_width, images_height * assumed_aspect, 10.0)
                left0 = (13.33 - width_each) / 2
                # 垂直居中到可用区域
                est_height = width_each / assumed_aspect
                top0 = images_top + max(0.0, (images_height - est_height) / 2)
                path = self.renderer._resolve_image_path(imgs[0]['src'])
                # 使用计算出的高度，确保不超过可用区域
                pic_height = min(est_height, images_height)
                # 获取图片标题
                image_caption = imgs[0].get('caption', f"图片 1")
                self.renderer.insert_image(idx, path, left0, top0, width_each, pic_height, image_caption)
                return
            
            # 多张图片：等宽排列
            assumed_aspect = 4.0 / 3.0
            width_by_row = (total_width - gap * (count - 1)) / count
            width_by_height = images_height * assumed_aspect
            width_each = min(width_by_row, width_by_height, 5.0)
            est_height = width_each / assumed_aspect
            top0 = images_top + max(0.0, (images_height - est_height) / 2)
            # 左边距统一为1.0，与文字区域对齐
            start_left = 1.0
            for i, block in enumerate(imgs):
                path = self.renderer._resolve_image_path(block['src'])
                left = start_left + i * (width_each + gap)
                # 使用计算出的高度，确保不超过可用区域
                pic_height = min(est_height, images_height)
                # 获取图片标题
                image_caption = block.get('caption', f"图片 {i+1}")
                self.renderer.insert_image(idx, path, left, top0, width_each, pic_height, image_caption)
            print(f"已添加图片 {count} 张，等宽布局")
        except Exception as e:
            print(f"添加图片时出错: {e}")
    
    def layout_table_only(self, slide, table_block, content_top, content_height):
        """布局2: 纯表格内容 - 表格居中，垂直居中"""
        print(f"    应用布局: 纯表格布局")
        # 计算垂直居中位置
        # 动态估计表格高度：根据可用高度与行列数自适应
        try:
            lines = [ln for ln in table_block.get('text', '').split('\n') if ln.strip()]
            header = lines[0] if lines else ''
            cols = max(1, len([c for c in header.split('|') if c.strip()]))
            rows = max(1, 1 + max(0, len(lines) - 1))
        except Exception:
            cols, rows = 3, 4
        base_row_h = max(0.4, min(0.7, 0.6 - 0.02 * max(0, rows - 4)))
        required_h = base_row_h * min(rows, 6) + 0.2
        table_height = max(0.8, min(content_height * 0.9, required_h))
        vertical_margin = max(0.0, (content_height - table_height) / 2)
        table_top_centered = content_top + vertical_margin
        
        # 使用insert_table替代_add_table_grid
        idx = self.renderer._get_slide_index(slide)
        # 获取表格标题
        table_caption = table_block.get('caption', f"表格 1")
        # 计算表格居中位置
        slide_width = 13.33
        table_width = min(10.0, cols * 2.0)
        table_left = (slide_width - table_width) / 2
        self.renderer.insert_table(
            slide_index=idx,
            rows=rows,
            cols=cols,
            data=None,
            left=table_left,
            top=table_top_centered,
            width=table_width,
            height=table_height,
            caption=table_caption
        )
    
    def layout_text_and_table(self, slide, text_blocks, table_block, content_top, content_height):
        """布局3: 文字+表格 - 文字顶部靠左，表格中下方居中"""
        print(f"    应用布局: 文字+表格布局")
        
        # 智能分配空间比例
        text_ratio = 0.3 if len(text_blocks) <= 2 else 0.4  # 文字少时占30%，多时占40%
        table_ratio = 0.5  # 表格固定占50%
        gap_ratio = 0.1    # 间隙占10%
        
        # 文字区域（顶部靠左）
        if text_blocks:
            text_height = content_height * text_ratio
            self.renderer.add_text_content_left_aligned(slide, text_blocks, content_top, text_height, self.font_calc)
        
        # 表格区域（中下方居中）
        table_top = content_top + content_height * (text_ratio + gap_ratio)
        # 动态估计高度
        try:
            lines = [ln for ln in table_block.get('text', '').split('\n') if ln.strip()]
            header = lines[0] if lines else ''
            cols = max(1, len([c for c in header.split('|') if c.strip()]))
            rows = max(1, 1 + max(0, len(lines) - 1))
        except Exception:
            cols, rows = 3, 4
        base_row_h = max(0.4, min(0.7, 0.6 - 0.02 * max(0, rows - 4)))
        required_h = base_row_h * min(rows, 6) + 0.2
        table_height = max(0.8, min(content_height * table_ratio, required_h))
        # 使用insert_table替代_add_table_grid
        idx = self.renderer._get_slide_index(slide)
        # 获取表格标题
        table_caption = table_block.get('caption', f"表格 1")
        # 计算表格居中位置
        slide_width = 13.33
        table_width = min(10.0, cols * 2.0)
        table_left = (slide_width - table_width) / 2
        self.renderer.insert_table(
            slide_index=idx,
            rows=rows,
            cols=cols,
            data=None,
            left=table_left,
            top=table_top,
            width=table_width,
            height=table_height,
            caption=table_caption
        )
    
    def layout_text_and_images(self, slide, text_blocks, image_blocks, content_top, content_height):
        """布局4: 文字+图片 - 文字顶部，图片下方居中"""
        print(f"    应用布局: 文字+图片布局")
        text_ratio = 0.35 if len(text_blocks) > 2 else 0.3
        gap_ratio = 0.08
        images_ratio = 1.0 - text_ratio - gap_ratio
        # 文字区域
        text_height = content_height * text_ratio
        self.renderer.add_text_content_left_aligned(slide, text_blocks, content_top, text_height, self.font_calc)
        # 图片区域：显示在文字下方和页面底部的中间
        text_bottom = content_top + text_height
        page_bottom = content_top + content_height
        images_height = content_height * images_ratio
        # 计算图片区域的顶部位置，使其在文字下方和页面底部之间居中
        available_height = page_bottom - text_bottom
        images_top = text_bottom + (available_height - images_height) / 2
        
        # 直接使用insert_image插入图片，确保高度不超过可用区域
        try:
            max_images = 3
            imgs = (image_blocks or [])[:max_images]
            count = len(imgs)
            if count == 0:
                return
            
            total_width = 11.0
            gap = 0.3
            idx = self.renderer._get_slide_index(slide)
            
            if count == 1:
                # 单张图片：高度优先，假设16:9比例
                assumed_aspect = 16.0 / 9.0
                width_each = min(total_width, images_height * assumed_aspect, 10.0)
                left0 = (13.33 - width_each) / 2
                # 垂直居中到可用区域
                est_height = width_each / assumed_aspect
                top0 = images_top + max(0.0, (images_height - est_height) / 2)
                path = self.renderer._resolve_image_path(imgs[0]['src'])
                # 使用计算出的高度，确保不超过可用区域
                pic_height = min(est_height, images_height)
                # 获取图片标题
                image_caption = imgs[0].get('caption', f"图片 1")
                self.renderer.insert_image(idx, path, left0, top0, width_each, pic_height, image_caption)
                return
            
            # 多张图片：等宽排列
            assumed_aspect = 4.0 / 3.0
            width_by_row = (total_width - gap * (count - 1)) / count
            width_by_height = images_height * assumed_aspect
            width_each = min(width_by_row, width_by_height, 5.0)
            est_height = width_each / assumed_aspect
            top0 = images_top + max(0.0, (images_height - est_height) / 2)
            # 左边距统一为1.0，与文字区域对齐
            start_left = 1.0
            for i, block in enumerate(imgs):
                path = self.renderer._resolve_image_path(block['src'])
                left = start_left + i * (width_each + gap)
                # 使用计算出的高度，确保不超过可用区域
                pic_height = min(est_height, images_height)
                # 获取图片标题
                image_caption = block.get('caption', f"图片 {i+1}")
                self.renderer.insert_image(idx, path, left, top0, width_each, pic_height, image_caption)
            print(f"已添加图片 {count} 张，等宽布局")
        except Exception as e:
            print(f"添加图片时出错: {e}")
    
    def layout_table_and_images(self, slide, table_block, image_blocks, content_top, content_height):
        """布局5: 表格+图片 - 底部横排对齐"""
        print(f"    应用布局: 表格+图片布局（底部横排对齐）")
        # 在内容区域底部创建一个横向分区，将表格与图片并排放置，并整体居中
        band_margin_h = 0.0
        band_left = 1.0
        band_width = 11.0
        # 底部带高度尽量大一些，但不超过可用高度
        band_height = max(1.8, min(content_height, 4.0))
        band_top = content_top + content_height - band_height
        # 两列布局
        gap = 0.3
        col_width = (band_width - gap) / 2.0
        left_col_left = band_left
        right_col_left = band_left + col_width + gap
        # 表格区域：左列，直接调用 insert_table 顶部对齐
        idx = self.renderer._get_slide_index(slide)
        try:
            lines = table_block.get('lines', []) if table_block else []
            header_line = lines[0].strip() if len(lines) > 0 else ''
            data_lines = lines[2:] if len(lines) > 2 else []
            header_cells = [cell.strip() for cell in header_line.split('|') if cell.strip()] or ['']
            cols = len(header_cells)
            rows = 1 + len(data_lines)
            # 限制规模，保证适配区域
            max_cols = min(cols, 5)
            max_rows = min(rows, 8)
            # 组装数据
            data = []
            data.append(header_cells[:max_cols] + [''] * (max_cols - len(header_cells[:max_cols])))
            for line in data_lines[:max_rows-1]:
                cells = [c.strip() for c in line.split('|') if c.strip()]
                row = cells[:max_cols] + [''] * (max_cols - len(cells[:max_cols]))
                data.append(row)
            # 获取表格标题
            table_caption = table_block.get('caption', f"表格 1")
            self.renderer.insert_table(
                slide_index=idx,
                rows=max_rows,
                cols=max_cols,
                data=data,
                left=left_col_left,
                top=band_top,
                width=col_width,
                height=band_height,
                caption=table_caption
            )
        except Exception as e:
            print(f"直接插入表格失败: {e}")

        # 图片区域：右列，直接调用 insert_image 顶部对齐，等宽布局
        try:
            max_images = 3
            imgs = (image_blocks or [])[:max_images]
            count = len(imgs)
            if count > 0:
                gap = 0.3
                if count == 1:
                    assumed_aspect = 16.0 / 9.0
                    width_each = min(col_width * 0.9, band_height * assumed_aspect, 8.0)
                    left_offset = (col_width - width_each) / 2
                    path = self.renderer._resolve_image_path(imgs[0]['src'])
                    # 直接使用区域高度，与表格保持一致
                    pic_height = band_height
                    # 获取图片标题
                    image_caption = imgs[0].get('caption', f"图片 1")
                    self.renderer.insert_image(idx, path, right_col_left + left_offset, band_top, width_each, pic_height, image_caption)
                else:
                    assumed_aspect = 4.0 / 3.0
                    width_by_row = (col_width - gap * (count - 1)) / count
                    width_by_height = band_height * assumed_aspect
                    width_each = min(width_by_row, width_by_height, 4.5)
                    total_group_width = width_each * count + gap * (count - 1)
                    start_offset = (col_width - total_group_width) / 2
                    for i, block in enumerate(imgs):
                        path = self.renderer._resolve_image_path(block['src'])
                        left = right_col_left + start_offset + i * (width_each + gap)
                        # 直接使用区域高度，与表格保持一致
                        pic_height = band_height
                        # 获取图片标题
                        image_caption = block.get('caption', f"图片 {i+1}")
                        self.renderer.insert_image(idx, path, left, band_top, width_each, pic_height, image_caption)
        except Exception as e:
            print(f"直接插入图片失败: {e}")
    
    def layout_text_table_images(self, slide, text_blocks, table_block, image_blocks, content_top, content_height):
        """布局6: 文字+表格+图片 - 上文字、左表格、右图片"""
        print(f"    应用布局: 文字+表格+图片布局（上文字、左表格、右图片）")
        
        # 1) 顶部文字：根据内容动态估算高度，设定最小/最大边界
        max_text_ratio = 0.35
        min_text_in = 0.6
        max_text_in = 2.0
        text_height = 0.0
        if text_blocks:
            estimated = self.estimate_text_block_height(text_blocks, container_width_in=11.0, available_height_in=content_height * max_text_ratio)
            text_height = max(min_text_in, min(max_text_in, estimated))
            self.renderer.add_text_content_left_aligned(slide, text_blocks, content_top, text_height, self.font_calc)
        
        remaining_top = content_top + (text_height if text_blocks else 0)
        remaining_height = max(0.0, content_height - (text_height if text_blocks else 0))
        if remaining_height <= 0.4:
            if table_block:
                # 高度已由区域限制，此处直接使用
                # 使用insert_table替代_add_table_grid
                idx = self.renderer._get_slide_index(slide)
                # 获取表格标题
                table_caption = table_block.get('caption', f"表格 1")
                # 计算表格居中位置
                slide_width = 13.33
                table_width = min(10.0, 5 * 2.0)  # 假设最多5列
                table_left = (slide_width - table_width) / 2
                self.renderer.insert_table(
                    slide_index=idx,
                    rows=4,  # 假设4行
                    cols=3,  # 假设3列
                    data=None,
                    left=table_left,
                    top=remaining_top,
                    width=table_width,
                    height=remaining_height,
                    caption=table_caption
                )
            if image_blocks:
                # 直接使用insert_image插入图片，确保高度不超过可用区域
                try:
                    max_images = 3
                    imgs = (image_blocks or [])[:max_images]
                    count = len(imgs)
                    if count > 0:
                        total_width = 11.0
                        gap = 0.3
                        idx = self.renderer._get_slide_index(slide)
                        
                        if count == 1:
                            # 单张图片：高度优先，假设16:9比例
                            assumed_aspect = 16.0 / 9.0
                            width_each = min(total_width, remaining_height * assumed_aspect, 10.0)
                            left0 = (13.33 - width_each) / 2
                            # 垂直居中到可用区域
                            est_height = width_each / assumed_aspect
                            top0 = remaining_top + max(0.0, (remaining_height - est_height) / 2)
                            path = self.renderer._resolve_image_path(imgs[0]['src'])
                            # 使用计算出的高度，确保不超过可用区域
                            pic_height = min(est_height, remaining_height)
                            # 获取图片标题
                            image_caption = imgs[0].get('caption', f"图片 1")
                            self.renderer.insert_image(idx, path, left0, top0, width_each, pic_height, image_caption)
                        else:
                            # 多张图片：等宽排列
                            assumed_aspect = 4.0 / 3.0
                            width_by_row = (total_width - gap * (count - 1)) / count
                            width_by_height = remaining_height * assumed_aspect
                            width_each = min(width_by_row, width_by_height, 5.0)
                            est_height = width_each / assumed_aspect
                            top0 = remaining_top + max(0.0, (remaining_height - est_height) / 2)
                            # 左边距统一为1.0，与文字区域对齐
                            start_left = 1.0
                            for i, block in enumerate(imgs):
                                path = self.renderer._resolve_image_path(block['src'])
                                left = start_left + i * (width_each + gap)
                                # 使用计算出的高度，确保不超过可用区域
                                pic_height = min(est_height, remaining_height)
                                # 获取图片标题
                                image_caption = block.get('caption', f"图片 {i+1}")
                                self.renderer.insert_image(idx, path, left, top0, width_each, pic_height, image_caption)
                except Exception as e:
                    print(f"添加图片时出错: {e}")
            return
        
        # 2) 剩余区域左右分栏：左表格、右图片
        left_margin = 0.8
        right_margin = 0.8
        middle_gap = 0.4
        slide_width_in = 13.33
        available_width = slide_width_in - left_margin - right_margin
        left_region_width = (available_width - middle_gap) / 2
        right_region_width = (available_width - middle_gap) / 2
        
        table_left = left_margin
        images_left = left_margin + left_region_width + middle_gap
        
        region_top = remaining_top + 0.1
        region_height = max(0.5, remaining_height - 0.1)
        
        # 表格：在左侧区域内直接插入表格
        idx = self.renderer._get_slide_index(slide)
        try:
            lines = table_block.get('lines', []) if table_block else []
            header_line = lines[0].strip() if len(lines) > 0 else ''
            data_lines = lines[2:] if len(lines) > 2 else []
            header_cells = [cell.strip() for cell in header_line.split('|') if cell.strip()] or ['']
            cols = len(header_cells)
            rows = 1 + len(data_lines)
            max_cols = min(cols, 5)
            max_rows = min(rows, 8)
            data = []
            data.append(header_cells[:max_cols] + [''] * (max_cols - len(header_cells[:max_cols])))
            for line in data_lines[:max_rows-1]:
                cells = [c.strip() for c in line.split('|') if c.strip()]
                row = cells[:max_cols] + [''] * (max_cols - len(cells[:max_cols]))
                data.append(row)
            # 获取表格标题
            table_caption = table_block.get('caption', f"表格 1")
            self.renderer.insert_table(
                slide_index=idx,
                rows=max_rows,
                cols=max_cols,
                data=data,
                left=table_left,
                top=region_top,
                width=left_region_width,
                height=region_height,
                caption=table_caption
            )
        except Exception as e:
            print(f"直接插入表格失败: {e}")

        # 图片：在右侧区域内直接插入图片，顶部对齐
        try:
            max_images = 3
            imgs = (image_blocks or [])[:max_images]
            count = len(imgs)
            if count > 0:
                gap = 0.3
                if count == 1:
                    assumed_aspect = 16.0 / 9.0
                    width_each = min(right_region_width * 0.9, region_height * assumed_aspect, 8.0)
                    left_offset = (right_region_width - width_each) / 2
                    path = self.renderer._resolve_image_path(imgs[0]['src'])
                    # 计算图片高度，使其与表格高度一致
                    pic_height = region_height
                    # 获取图片标题
                    image_caption = imgs[0].get('caption', f"图片 1")
                    self.renderer.insert_image(idx, path, images_left + left_offset, region_top, width_each, pic_height, image_caption)
                else:
                    assumed_aspect = 4.0 / 3.0
                    width_by_row = (right_region_width - gap * (count - 1)) / count
                    width_by_height = region_height * assumed_aspect
                    width_each = min(width_by_row, width_by_height, 4.5)
                    total_group_width = width_each * count + gap * (count - 1)
                    start_offset = (right_region_width - total_group_width) / 2
                    top_offset = 0.0  # 顶部对齐
                    for i, block in enumerate(imgs):
                        path = self.renderer._resolve_image_path(block['src'])
                        left = images_left + start_offset + i * (width_each + gap)
                        # 计算图片高度，使其与表格高度一致
                        pic_height = region_height
                        # 获取图片标题
                        image_caption = block.get('caption', f"图片 {i+1}")
                        self.renderer.insert_image(idx, path, left, region_top, width_each, pic_height, image_caption)
        except Exception as e:
            print(f"直接插入图片失败: {e}")
    
    def layout_complex_content(self, slide, text_blocks, table_blocks, content_top, content_height):
        """布局7: 复杂内容 - 紧凑布局"""
        print(f"    应用布局: 复杂内容布局")
        
        current_top = content_top
        remaining_height = content_height
        
        # 文字内容在顶部
        if text_blocks:
            text_height = min(remaining_height * 0.3, 1.5)
            self.renderer.add_text_content_left_aligned(slide, text_blocks, current_top, text_height, self.font_calc)
            current_top += text_height + 0.1
            remaining_height -= (text_height + 0.1)
        
        # 多个表格垂直排列
        if table_blocks and remaining_height > 0.5:
            table_height = remaining_height / len(table_blocks)
            for i, table_block in enumerate(table_blocks):
                if remaining_height > 0.3:
                    actual_table_height = min(table_height, remaining_height)
                    # 使用insert_table替代_add_table_grid
                    idx = self.renderer._get_slide_index(slide)
                    # 获取表格标题
                    table_caption = table_block.get('caption', f"表格 {i+1}")
                    # 计算表格居中位置
                    slide_width = 13.33
                    table_width = min(10.0, 5 * 2.0)  # 假设最多5列
                    table_left = (slide_width - table_width) / 2
                    self.renderer.insert_table(
                        slide_index=idx,
                        rows=4,  # 假设4行
                        cols=3,  # 假设3列
                        data=None,
                        left=table_left,
                        top=current_top,
                        width=table_width,
                        height=actual_table_height,
                        caption=table_caption
                    )
                    current_top += actual_table_height
                    remaining_height -= actual_table_height
    
    def estimate_text_block_height(self, text_blocks, container_width_in, available_height_in) -> float:
        """更精确估算文本高度（英寸）"""
        try:
            if not text_blocks:
                return 0.0

            # 从FontCalculator获取文本估算配置
            text_config = self.font_calc.get_text_estimation_config()
            font_size_pt = self.font_calc.base_sizes.get('text', 18)  # 使用配置的正文字号
            line_height_in = (font_size_pt * text_config.get('line_height_ratio', 1.2)) / 72.0
            gap_list_in = text_config.get('gap_list', 6) / 72.0
            gap_para_in = text_config.get('gap_paragraph', 8) / 72.0

            # 每行可容纳字符数（CJK 近似方宽：字符宽≈字号pt/72 英寸）
            def chars_per_line_for(width_in: float) -> int:
                min_chars = text_config.get('min_chars_per_line', 8)
                cpp = int(max(min_chars, (width_in * 72.0) / max(8.0, float(font_size_pt))))
                return cpp

            total_height_in = 0.0
            block_count_for_gaps = 0

            for block in text_blocks:
                if block.get('type') == 'list':
                    items = block.get('items', []) or []
                    if not items:
                        continue
                    # 列表有效宽度：考虑项目符号与缩进
                    effective_width_in = max(1.0, container_width_in - 0.3)
                    cpp = chars_per_line_for(effective_width_in)
                    for item in items:
                        text = str(item or '')
                        char_len = len(text)
                        lines = max(1, (char_len + cpp - 1) // cpp)
                        total_height_in += lines * line_height_in
                        # 段后距：每个列表项后添加
                        total_height_in += gap_list_in
                        block_count_for_gaps += 1
                elif block.get('type') == 'paragraph':
                    text = str(block.get('text', '') or '')
                    if text.strip() == '':
                        continue
                    effective_width_in = container_width_in
                    cpp = chars_per_line_for(effective_width_in)
                    char_len = len(text)
                    lines = max(1, (char_len + cpp - 1) // cpp)
                    total_height_in += lines * line_height_in
                    total_height_in += gap_para_in
                    block_count_for_gaps += 1

            # 去除最后一个段落/列表项多加的段后距
            if block_count_for_gaps > 0:
                total_height_in -= min(gap_para_in, gap_list_in)

            # 约束到可用高度
            return max(0.0, min(total_height_in, max(0.0, available_height_in)))
        except Exception:
            # 回退到安全估计
            fallback = min(1.2, max(0.0, available_height_in))
            return fallback