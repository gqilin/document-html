# -*- coding: utf-8 -*-
"""
改进的PDF到HTML转换器
针对双层可复制文字PDF优化
支持图片提取和本地保存
"""

import fitz  # PyMuPDF
import re
import os
import io
from html import escape
from pathlib import Path
from typing import List, Dict, Any, Tuple, Optional
from PIL import Image as PILImage
import hashlib

# 导入文本标准化器
try:
    from .text_normalizer import normalize_pdf_text
    TEXT_NORMALIZER_AVAILABLE = True
except ImportError:
    try:
        from src.text_normalizer import normalize_pdf_text
        TEXT_NORMALIZER_AVAILABLE = True
    except ImportError:
        TEXT_NORMALIZER_AVAILABLE = False

# 导入图片下载器
try:
    from .image_downloader import ImageDownloader
    IMAGE_DOWNLOADER_AVAILABLE = True
except ImportError:
    try:
        from src.image_downloader import ImageDownloader
        IMAGE_DOWNLOADER_AVAILABLE = True
    except ImportError:
        IMAGE_DOWNLOADER_AVAILABLE = False


class ImprovedPDFConverter:
    """
    改进的PDF转换器
    
    针对双层PDF（图像层+文字层）优化：
    1. 使用PyMuPDF提取带位置信息的文本
    2. 合并同一行上的分散文本块
    3. 智能标点位置校正
    4. 去除重复文字层
    """
    
    def __init__(self, images_dir: Optional[str] = None):
        self.y_tolerance = 3.0  # Y轴容差，用于判断同一行
        self.x_tolerance = 5.0  # X轴容差，用于判断相邻文本
        self.min_text_length = 2  # 最小文本长度
        
        # 图片相关设置
        self.extract_images = True  # 是否提取图片
        self.image_min_size = 100  # 图片最小尺寸（像素）
        self.image_min_bytes = 1024  # 图片最小字节数
        
        # 装饰性图片过滤设置
        self.filter_decorative_images = True  # 是否过滤装饰性图片
        self.decorative_max_size = 50  # 装饰性图片最大尺寸（像素）
        self.decorative_max_bytes = 20480  # 装饰性图片最大字节数（20KB）
        self.min_image_width = 80  # 正文图片最小宽度
        self.min_image_height = 80  # 正文图片最小高度
        self.min_aspect_ratio = 0.2  # 最小宽高比
        self.max_aspect_ratio = 5.0  # 最大宽高比
        self.seen_image_hashes = set()  # 用于去重
        
        # 初始化图片下载器
        if images_dir:
            self.images_dir = Path(images_dir)
        else:
            # 默认使用项目根目录下的uploads/images
            project_root = Path(__file__).parent.parent
            self.images_dir = project_root / 'uploads' / 'images'
        
        self.images_dir.mkdir(parents=True, exist_ok=True)
        
        if IMAGE_DOWNLOADER_AVAILABLE:
            self.image_downloader = ImageDownloader(str(self.images_dir))
        else:
            self.image_downloader = None
        
        # 记录提取的图片
        self.extracted_images = []
        self.skipped_images = []  # 记录被过滤的图片
        
        # 重置图片hash集合
        self.reset_image_filter()
        
    def convert_to_html(self, pdf_path: str, output_path: Optional[str] = None) -> str:
        """
        将PDF转换为HTML
        
        Args:
            pdf_path: PDF文件路径
            output_path: 输出HTML路径，如果为None则自动生成
            
        Returns:
            HTML文件路径
        """
        pdf_path = Path(pdf_path)
        if not pdf_path.exists():
            raise FileNotFoundError(f"PDF文件不存在: {pdf_path}")
        
        if output_path is None:
            output_path = pdf_path.with_suffix('.html')
        else:
            output_path = Path(output_path)
        
        # 打开PDF
        doc = fitz.open(str(pdf_path))
        
        # 重置图片过滤状态
        self.reset_image_filter()
        self.extracted_images = []
        self.skipped_images = []
        
        # 构建HTML
        html_parts = []
        html_parts.append('<!DOCTYPE html>')
        html_parts.append('<html lang="zh-CN">')
        html_parts.append('<head>')
        html_parts.append('    <meta charset="UTF-8">')
        html_parts.append(f'    <title>{escape(pdf_path.stem)}</title>')
        html_parts.append('    <style>')
        html_parts.append(self._get_css_styles())
        html_parts.append('    </style>')
        html_parts.append('</head>')
        html_parts.append('<body>')
        
        # 处理每一页，直接将内容添加到body
        for page_num in range(len(doc)):
            page = doc[page_num]
            page_content = self._convert_page_content(page, page_num)
            html_parts.extend(page_content)
        
        html_parts.append('</body>')
        html_parts.append('</html>')
        
        doc.close()
        
        # 写入文件
        html_content = '\n'.join(html_parts)
        output_path.write_text(html_content, encoding='utf-8')
        
        return str(output_path)
    
    def _get_css_styles(self) -> str:
        """获取CSS样式"""
        return '''
        body {
            font-family: "SimSun", "宋体", "Noto Serif CJK SC", serif;
            line-height: 1.8;
            margin: 40px;
            color: #333;
            max-width: 800px;
            margin-left: auto;
            margin-right: auto;
        }
        p {
            margin: 0.8em 0;
            text-indent: 2em;
            text-align: justify;
        }
        p.no-indent {
            text-indent: 0;
        }
        h1, h2, h3, h4, h5, h6 {
            margin: 1.5em 0 0.8em 0;
            font-weight: bold;
        }
        h1 { font-size: 1.8em; }
        h2 { font-size: 1.5em; }
        h3 { font-size: 1.3em; }
        .center { text-align: center; }
        .right { text-align: right; }
        
        /* PDF图片样式 */
        p.pdf-image {
            text-indent: 0;
            margin: 1em 0;
        }
        p.pdf-image img {
            max-width: 100%;
            height: auto;
            display: block;
            margin: 0 auto;
        }
        '''
    
    def _convert_page_content(self, page: fitz.Page, page_num: int) -> List[str]:
        """转换单个页面的内容为HTML元素列表"""
        # 获取页面文本块
        blocks = self._extract_text_blocks(page)
        
        # 合并同一行的文本
        lines = self._merge_blocks_to_lines(blocks)
        
        # 将行组合成段落
        paragraphs = self._lines_to_paragraphs(lines)
        
        # 提取页面中的图片
        page_images = []
        if self.extract_images:
            page_images = self._extract_page_images(page, page_num)
        
        # 生成HTML元素列表
        html_parts = []
        
        # 插入图片到合适的位置
        for para in paragraphs:
            # 检查是否有图片应该插入在这个段落附近
            para_y = para.get('y0', 0)
            para_html = self._format_paragraph(para)
            
            # 查找应该在这个段落之前插入的图片
            images_to_insert = []
            for img in page_images[:]:
                img_y = img.get('y0', 0)
                # 如果图片在这个段落上方附近，先插入图片
                if img_y < para_y and abs(img_y - para_y) < 100:
                    images_to_insert.append(img)
                    page_images.remove(img)
            
            # 插入图片
            for img in images_to_insert:
                img_html = self._format_image(img)
                if img_html:
                    html_parts.append(img_html)
            
            if para_html:
                html_parts.append(para_html)
        
        # 添加剩余的图片（在页面底部）
        for img in page_images:
            img_html = self._format_image(img)
            if img_html:
                html_parts.append(img_html)
        
        return html_parts
    
    def _extract_text_blocks(self, page: fitz.Page) -> List[Dict[str, Any]]:
        """
        提取页面中的文本块
        
        返回按位置排序的文本块列表
        """
        blocks = []
        
        # 使用dict模式获取详细文本信息
        page_dict = page.get_text("dict")
        
        for block in page_dict.get("blocks", []):
            if block.get("type") != 0:  # 跳过非文本块
                continue
            
            block_text = []
            block_fonts = []
            block_sizes = []
            
            for line in block.get("lines", []):
                line_text = []
                for span in line.get("spans", []):
                    text = span.get("text", "")
                    if text.strip():
                        line_text.append(text)
                        block_fonts.append(span.get("font", ""))
                        block_sizes.append(span.get("size", 0))
                
                if line_text:
                    block_text.append("".join(line_text))
            
            if block_text:
                full_text = "".join(block_text)
                # 过滤掉过短的文本（可能是页眉页脚或噪点）
                if len(full_text.strip()) >= self.min_text_length:
                    bbox = block.get("bbox", [0, 0, 0, 0])
                    blocks.append({
                        "text": full_text,
                        "x0": bbox[0],
                        "y0": bbox[1],
                        "x1": bbox[2],
                        "y1": bbox[3],
                        "fonts": list(set(block_fonts)),
                        "sizes": block_sizes,
                        "avg_size": sum(block_sizes) / len(block_sizes) if block_sizes else 0
                    })
        
        # 按Y坐标排序，然后按X坐标排序
        blocks.sort(key=lambda b: (round(b["y0"] / self.y_tolerance), b["x0"]))
        
        return blocks
    
    def _merge_blocks_to_lines(self, blocks: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        将文本块合并成行
        
        同一水平线上的文本块会被合并
        """
        if not blocks:
            return []
        
        lines = []
        current_line = [blocks[0]]
        current_y = blocks[0]["y0"]
        
        for block in blocks[1:]:
            # 检查是否在同一个水平线上
            if abs(block["y0"] - current_y) <= self.y_tolerance:
                current_line.append(block)
            else:
                # 处理当前行
                line = self._merge_line_blocks(current_line)
                if line:
                    lines.append(line)
                
                # 开始新行
                current_line = [block]
                current_y = block["y0"]
        
        # 处理最后一行
        if current_line:
            line = self._merge_line_blocks(current_line)
            if line:
                lines.append(line)
        
        return lines
    
    def _merge_line_blocks(self, blocks: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
        """合并同一行上的多个文本块"""
        if not blocks:
            return None
        
        # 按X坐标排序
        blocks.sort(key=lambda b: b["x0"])
        
        # 合并文本
        merged_text = ""
        last_x1 = None
        
        for block in blocks:
            # 检查是否需要添加空格
            if last_x1 is not None:
                gap = block["x0"] - last_x1
                if gap > self.x_tolerance:
                    merged_text += " "
            
            merged_text += block["text"]
            last_x1 = block["x1"]
        
        # 计算平均字体大小
        all_sizes = []
        for b in blocks:
            all_sizes.extend(b.get("sizes", []))
        avg_size = sum(all_sizes) / len(all_sizes) if all_sizes else 12
        
        return {
            "text": merged_text,
            "y0": blocks[0]["y0"],
            "y1": blocks[0]["y1"],
            "x0": blocks[0]["x0"],
            "x1": blocks[-1]["x1"],
            "size": avg_size,
            "fonts": list(set(f for b in blocks for f in b.get("fonts", [])))
        }
    
    def _lines_to_paragraphs(self, lines: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        将行组合成段落
        
        根据行间距和缩进判断段落边界
        """
        if not lines:
            return []
        
        paragraphs = []
        current_para_lines = [lines[0]]
        
        for i in range(1, len(lines)):
            current_line = lines[i]
            prev_line = lines[i - 1]
            
            # 判断是否是新段落
            is_new_para = self._is_new_paragraph(prev_line, current_line)
            
            if is_new_para:
                # 保存当前段落
                para = self._create_paragraph(current_para_lines)
                if para:
                    paragraphs.append(para)
                current_para_lines = [current_line]
            else:
                current_para_lines.append(current_line)
        
        # 处理最后一个段落
        if current_para_lines:
            para = self._create_paragraph(current_para_lines)
            if para:
                paragraphs.append(para)
        
        return paragraphs
    
    def _is_new_paragraph(self, prev_line: Dict, current_line: Dict) -> bool:
        """判断当前行是否开始新段落"""
        # 检查行间距
        line_gap = current_line["y0"] - prev_line["y1"]
        normal_line_height = prev_line["y1"] - prev_line["y0"]
        
        # 如果行间距大于1.5倍行高，认为是新段落
        if line_gap > normal_line_height * 1.5:
            return True
        
        # 检查缩进
        if current_line["x0"] - prev_line["x0"] > 20:
            return True
        
        # 检查上一行是否以标点结尾
        prev_text = prev_line.get("text", "").strip()
        if prev_text and prev_text[-1] in '。！？.!?':
            return True
        
        return False
    
    def _create_paragraph(self, lines: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
        """从行列表创建段落"""
        if not lines:
            return None
        
        # 合并文本
        text = "".join(line["text"] for line in lines)
        text = self._clean_text(text)
        
        if not text.strip():
            return None
        
        # 计算平均字体大小
        avg_size = sum(line.get("size", 12) for line in lines) / len(lines)
        
        # 判断段落类型
        para_type = self._detect_paragraph_type(text, avg_size, lines[0]["x0"])
        
        return {
            "text": text,
            "type": para_type,
            "size": avg_size,
            "y0": lines[0]["y0"],
            "is_center": abs(lines[0]["x0"] - 200) < 50  # 粗略判断是否居中
        }
    
    def _detect_paragraph_type(self, text: str, font_size: float, x0: float) -> str:
        """检测段落类型（标题、正文等）"""
        text = text.strip()
        
        # 根据字体大小判断
        if font_size > 16:
            return "h2"
        if font_size > 14:
            return "h3"
        
        # 根据内容判断
        if re.match(r'^第[一二三四五六七八九十\d]+章', text):
            return "h1"
        if re.match(r'^\d+[\.、]', text):
            return "h3"
        
        return "p"
    
    def _clean_text(self, text: str) -> str:
        """清理文本"""
        # 使用文本标准化器处理全角字符和特殊符号
        if TEXT_NORMALIZER_AVAILABLE:
            text = normalize_pdf_text(text)
        else:
            # 降级处理：基本的标点修复
            text = self._fix_punctuation(text)
        
        # 去除控制字符
        text = re.sub(r'[\x00-\x08\x0b-\x0c\x0e-\x1f]', '', text)
        
        return text.strip()
    
    def _fix_punctuation(self, text: str) -> str:
        """
        修复标点位置问题
        
        双层PDF中，标点符号可能位置不正确
        """
        # 修复空格+标点的问题（如"文字 ，" -> "文字，"）
        text = re.sub(r'\s+([，。、；：""''（）【】《》？！])', r'\1', text)
        
        # 修复标点后缺少空格的问题（英文标点后的空格）
        text = re.sub(r'([,;:!?])([^ \n])', r'\1 \2', text)
        
        # 修复连续标点
        text = re.sub(r'([，。、；：])​+', r'\1', text)  # 零宽空格
        
        # 修复引号配对问题
        text = self._fix_quotes(text)
        
        return text
    
    def _fix_quotes(self, text: str) -> str:
        """修复引号问题"""
        # 统一引号 - 使用ASCII兼容的方式
        text = text.replace('"', '"').replace('"', '"')
        text = text.replace("'", "'").replace("'", "'")
        
        return text
    
    def _format_paragraph(self, para: Dict[str, Any]) -> str:
        """格式化段落为HTML"""
        text = para.get("text", "")
        if not text:
            return ""
        
        para_type = para.get("type", "p")
        is_center = para.get("is_center", False)
        
        # 转义HTML特殊字符
        text = escape(text)
        
        # 根据类型生成HTML
        if para_type == "h1":
            class_attr = ' class="center"' if is_center else ''
            return f'<h1{class_attr}>{text}</h1>'
        elif para_type == "h2":
            class_attr = ' class="center"' if is_center else ''
            return f'<h2{class_attr}>{text}</h2>'
        elif para_type == "h3":
            return f'<h3>{text}</h3>'
        else:
            class_attr = ' class="no-indent"' if is_center else ''
            return f'<p{class_attr}>{text}</p>'
    
    def _extract_page_images(self, page: fitz.Page, page_num: int) -> List[Dict[str, Any]]:
        """
        提取页面中的图片，并过滤装饰性图片
        
        Args:
            page: PDF页面
            page_num: 页码
            
        Returns:
            图片信息列表
        """
        images = []
        
        try:
            # 获取页面中的图片列表
            image_list = page.get_images(full=True)
            
            for img_index, img in enumerate(image_list):
                xref = img[0]
                base_image = page.parent.extract_image(xref)
                
                if not base_image:
                    continue
                
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]
                
                # 过滤太小的图片（字节数）
                if len(image_bytes) < self.image_min_bytes:
                    continue
                
                # 获取图片在页面上的位置
                img_rect = self._get_image_rect(page, xref)
                width = img_rect[2] - img_rect[0] if img_rect else 0
                height = img_rect[3] - img_rect[1] if img_rect else 0
                
                # 过滤装饰性图片
                if self.filter_decorative_images:
                    is_decorative, reason = self._is_decorative_image(
                        image_bytes, width, height
                    )
                    if is_decorative:
                        print(f"  跳过装饰性图片: page{page_num + 1}_img{img_index + 1} - {reason}")
                        continue
                
                # 生成文件名
                filename = f"pdf_page{page_num + 1}_img{img_index + 1}.{image_ext}"
                file_path = self.images_dir / filename
                
                # 保存图片
                with open(file_path, 'wb') as f:
                    f.write(image_bytes)
                
                image_info = {
                    'filename': filename,
                    'path': str(file_path),
                    'relative_path': f"../uploads/images/{filename}",
                    'page': page_num + 1,
                    'index': img_index + 1,
                    'size': len(image_bytes),
                    'ext': image_ext,
                    'y0': img_rect[1] if img_rect else 0,
                    'x0': img_rect[0] if img_rect else 0,
                    'width': width,
                    'height': height
                }
                
                images.append(image_info)
                self.extracted_images.append(image_info)
                
                print(f"  提取图片: {filename} ({len(image_bytes)} bytes, {int(width)}x{int(height)})")
                
        except Exception as e:
            print(f"  提取图片时出错: {e}")
        
        # 按Y坐标排序（从上到下）
        images.sort(key=lambda x: x['y0'])
        
        return images
    
    def _is_decorative_image(self, image_bytes: bytes, width: float, height: float) -> Tuple[bool, str]:
        """
        判断图片是否为装饰性图片
        
        装饰性图片特征：
        1. 尺寸过小（如页眉页脚的分隔线、小图标）
        2. 极端的宽高比（如细长的分隔线）
        3. 文件过小（简单的图形）
        4. 纯色或简单图案（通过颜色数量判断）
        
        Args:
            image_bytes: 图片字节数据
            width: 图片宽度
            height: 图片高度
            
        Returns:
            (是否为装饰性图片, 原因)
        """
        # 1. 检查尺寸 - 太小的图片可能是装饰性的
        if width < self.min_image_width or height < self.min_image_height:
            return True, f"尺寸过小 ({int(width)}x{int(height)})"
        
        # 2. 检查文件大小 - 太小的文件可能是简单图形
        if len(image_bytes) < self.decorative_max_bytes:
            # 小文件需要进一步检查宽高比
            pass
        
        # 3. 检查宽高比 - 极端比例可能是分隔线
        if width > 0 and height > 0:
            aspect_ratio = width / height
            if aspect_ratio < self.min_aspect_ratio:
                return True, f"宽高比过小 ({aspect_ratio:.2f})"
            if aspect_ratio > self.max_aspect_ratio:
                return True, f"宽高比过大 ({aspect_ratio:.2f})"
        
        # 4. 检查图片内容复杂度（使用PIL）
        try:
            from PIL import Image as PILImage
            import io
            
            img = PILImage.open(io.BytesIO(image_bytes))
            
            # 获取图片实际尺寸
            img_width, img_height = img.size
            
            # 如果实际像素尺寸很小
            if img_width < 50 or img_height < 50:
                return True, f"像素尺寸过小 ({img_width}x{img_height})"
            
            # 检查颜色数量 - 装饰性图片通常颜色很少
            if img.mode in ('RGB', 'RGBA', 'L'):
                # 转换为RGB模式统计颜色
                rgb_img = img.convert('RGB')
                # 缩小图片以加快统计
                small_img = rgb_img.resize((100, 100))
                colors = small_img.getcolors(maxcolors=256)
                
                if colors and len(colors) < 10:
                    return True, f"颜色数量过少 ({len(colors)}种颜色)"
            
            # 检查是否为纯色或接近纯色
            if img.mode == 'L' or img.mode == '1':
                # 灰度图或二值图可能是线条
                extrema = img.getextrema()
                if isinstance(extrema, tuple):
                    min_val, max_val = extrema
                    if max_val - min_val < 20:  # 颜色变化很小
                        return True, "接近纯色"
                
        except Exception as e:
            # PIL检查失败，继续保留图片
            pass
        
        # 5. 检查图片hash去重
        try:
            image_hash = hashlib.md5(image_bytes).hexdigest()
            if image_hash in self.seen_image_hashes:
                return True, "重复图片"
            self.seen_image_hashes.add(image_hash)
        except Exception:
            pass
        
        return False, ""
    
    def reset_image_filter(self):
        """重置图片过滤状态（用于新的PDF转换）"""
        self.seen_image_hashes.clear()
    
    def _get_image_rect(self, page: fitz.Page, xref: int) -> Optional[Tuple[float, float, float, float]]:
        """
        获取图片在页面上的位置
        
        Args:
            page: PDF页面
            xref: 图片xref
            
        Returns:
            图片位置 (x0, y0, x1, y1) 或 None
        """
        try:
            # 遍历页面内容查找图片位置
            for img in page.get_images():
                if img[0] == xref:
                    # 尝试从页面字典中获取位置信息
                    page_dict = page.get_text("dict")
                    for block in page_dict.get("blocks", []):
                        if block.get("type") == 1:  # 图片块
                            return block.get("bbox")
            
            # 如果找不到精确位置，返回页面中心位置
            rect = page.rect
            return (rect.x0, rect.y0, rect.x1, rect.y1)
            
        except Exception:
            return None
    
    def _format_image(self, img_info: Dict[str, Any]) -> str:
        """
        格式化图片为HTML
        
        Args:
            img_info: 图片信息字典
            
        Returns:
            HTML img标签
        """
        relative_path = img_info.get('relative_path', '')
        width = img_info.get('width', 0)
        height = img_info.get('height', 0)
        
        # 构建img标签
        img_attrs = [f'src="{relative_path}"', 'alt="PDF图片"']
        
        # 添加尺寸属性（如果有效）
        if width > 0 and height > 0:
            # 限制最大宽度
            max_width = 600
            if width > max_width:
                ratio = max_width / width
                width = max_width
                height = int(height * ratio)
            
            img_attrs.append(f'width="{int(width)}"')
            img_attrs.append(f'height="{int(height)}"')
        
        img_tag = f'<img {" ".join(img_attrs)} />'
        
        # 包装在段落中，居中显示
        return f'<p class="pdf-image" style="text-align: center;">{img_tag}</p>'


# 兼容性函数
def convert_pdf_to_html(pdf_path: str, output_path: Optional[str] = None) -> str:
    """
    将PDF转换为HTML的便捷函数
    
    Args:
        pdf_path: PDF文件路径
        output_path: 输出HTML路径
        
    Returns:
        HTML文件路径
    """
    converter = ImprovedPDFConverter()
    return converter.convert_to_html(pdf_path, output_path)
