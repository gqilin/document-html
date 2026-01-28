# -*- coding: utf-8 -*-
"""
改进的PDF到HTML转换器
针对双层可复制文字PDF优化
"""

import fitz  # PyMuPDF
import re
from html import escape
from pathlib import Path
from typing import List, Dict, Any, Tuple, Optional

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


class ImprovedPDFConverter:
    """
    改进的PDF转换器
    
    针对双层PDF（图像层+文字层）优化：
    1. 使用PyMuPDF提取带位置信息的文本
    2. 合并同一行上的分散文本块
    3. 智能标点位置校正
    4. 去除重复文字层
    """
    
    def __init__(self):
        self.y_tolerance = 3.0  # Y轴容差，用于判断同一行
        self.x_tolerance = 5.0  # X轴容差，用于判断相邻文本
        self.min_text_length = 2  # 最小文本长度
        
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
        html_parts.append('    <div class="document">')
        
        # 处理每一页
        for page_num in range(len(doc)):
            page = doc[page_num]
            page_html = self._convert_page(page, page_num)
            html_parts.append(page_html)
        
        html_parts.append('    </div>')
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
        }
        .document {
            max-width: 800px;
            margin: 0 auto;
        }
        .page {
            margin-bottom: 40px;
            padding-bottom: 20px;
            border-bottom: 1px solid #ddd;
        }
        .page-number {
            text-align: center;
            color: #666;
            font-size: 0.9em;
            margin-top: 20px;
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
        '''
    
    def _convert_page(self, page: fitz.Page, page_num: int) -> str:
        """转换单个页面"""
        # 获取页面文本块
        blocks = self._extract_text_blocks(page)
        
        # 合并同一行的文本
        lines = self._merge_blocks_to_lines(blocks)
        
        # 将行组合成段落
        paragraphs = self._lines_to_paragraphs(lines)
        
        # 生成HTML
        html_parts = []
        html_parts.append(f'        <div class="page" id="page-{page_num + 1}">')
        
        for para in paragraphs:
            para_html = self._format_paragraph(para)
            if para_html:
                html_parts.append(f'            {para_html}')
        
        html_parts.append(f'            <div class="page-number">- {page_num + 1} -</div>')
        html_parts.append('        </div>')
        
        return '\n'.join(html_parts)
    
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
