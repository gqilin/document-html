import os
import subprocess
import tempfile
from pathlib import Path
from typing import Optional, Union
import re
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from src.config import ConversionConfig, DEFAULT_CONFIG
from src.image_downloader import ImageDownloader, process_html_images
try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False

try:
    from ebooklib import epub
    from bs4 import BeautifulSoup
    EPUBLIB_AVAILABLE = True
except ImportError:
    EPUBLIB_AVAILABLE = False

try:
    import re
    CSS_PARSER_AVAILABLE = True
except ImportError:
    CSS_PARSER_AVAILABLE = False


class DocumentConverter:
    """基于pandoc的文档转换器"""
    
    def __init__(self, config: Optional[ConversionConfig] = None):
        self.pandoc_available = self._check_pandoc()
        self.config = config if config is not None else DEFAULT_CONFIG
        # 创建图片下载器，使用相对于HTML文件的images目录
        self.image_downloader = ImageDownloader()
    
    def _check_pandoc(self) -> bool:
        """检查pandoc是否可用"""
        try:
            subprocess.run(['pandoc', '--version'], 
                         capture_output=True, check=True)
            return True
        except (subprocess.CalledProcessError, FileNotFoundError):
            return False
    
    def convert_to_html(self, input_file: str, output_file: Optional[str] = None) -> Optional[str]:
        """
        转换文档到HTML
        
        Args:
            input_file: 输入文件路径
            output_file: 输出文件路径，如果为None则自动生成
            
        Returns:
            转换后的HTML文件路径，失败返回None
        """
        if not self.pandoc_available:
            raise RuntimeError("pandoc is not available")
        
        input_path = Path(input_file)
        if not input_path.exists():
            raise FileNotFoundError(f"Input file not found: {input_file}")
        
        if output_file is None:
            output_file = str(input_path.with_suffix('.html'))
        
# 创建图片下载器，使用项目根目录的uploads/images目录
        project_root = Path(__file__).parent.parent
        images_dir = project_root / 'uploads' / 'images'
        image_downloader = ImageDownloader(str(images_dir))
        
        # 如果是docx文件且docx库可用，先提取对齐信息和图片
        alignment_info = []
        if input_path.suffix.lower() == '.docx' and DOCX_AVAILABLE:
            try:
                doc = Document(input_file)
                
                # 提取段落对齐信息
                for para in doc.paragraphs:
                    if para.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                        alignment_info.append('center')
                    elif para.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                        alignment_info.append('right')
                    elif para.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
                        alignment_info.append('justify')
                    elif para.alignment == WD_ALIGN_PARAGRAPH.DISTRIBUTE:
                        alignment_info.append('justify')
                    else:
                        alignment_info.append('left')
                
                # 提取图片
                try:
                    for rel in doc.part.rels.values():
                        if "image" in rel.target_ref:
                            try:
                                image_data = rel.target_part.blob
                                image_name = rel.target_ref.split('/')[-1]
                                local_path = image_downloader.save_from_bytes(image_data, image_name)
                                # 这里可以将图片路径存储起来，后续在HTML中使用
                            except Exception:
                                pass
                except Exception:
                    pass
                    
            except Exception:
                alignment_info = []
        
        input_path = Path(input_file)
        
        # 检测实际文件类型（不仅仅依赖扩展名）
        file_type = None
        try:
            with open(input_file, 'rb') as f:
                header = f.read(200)
                if header.startswith(b'PK\x03\x04') and b'mimetypeapplication/epub+zip' in header:
                    file_type = 'epub'
                elif header.startswith(b'%PDF'):
                    file_type = 'pdf'
        except Exception:
            pass
        
        # 如果是PDF文件，使用pdfplumber处理
        if (file_type == 'pdf' or input_path.suffix.lower() == '.pdf') and PDFPLUMBER_AVAILABLE:
            return self._convert_pdf_to_html(input_file, output_file)
        
        # 如果是EPUB文件，使用专门的epub处理
        if file_type == 'epub' or input_path.suffix.lower() == '.epub':
            return self._convert_epub_to_html(input_file, output_file)
        
        # 非PDF文件，使用原有的转换逻辑
        try:
            # 根据文件扩展名和内容确定输入格式
            input_path = Path(input_file)
            ext = input_path.suffix.lower()
            
            # 检查文件内容类型
            detected_format = None
            try:
                with open(input_file, 'rb') as f:
                    header = f.read(200)
                    
                    # 检测文件类型
                    if header.startswith(b'PK\x03\x04'):
                        # ZIP格式，可能是epub、docx等
                        if b'mimetypeapplication/epub+zip' in header:
                            detected_format = 'epub'
                        elif b'word/' in header or b'docx' in header:
                            detected_format = 'docx'
                    elif header.startswith(b'%PDF'):
                        detected_format = 'pdf'
                    elif header.startswith(b'{\\rtf'):
                        detected_format = 'rtf'
            except Exception:
                pass
            
            format_map = {
                '.docx': 'docx',
                '.doc': 'doc', 
                '.epub': 'epub',
                '.txt': 'plain',
                '.rtf': 'rtf',
                '.md': 'markdown',
                '.html': 'html',
                '.htm': 'html',
                '.odt': 'odt',
                '.pdf': 'pdf'
            }
            
            # 优先使用文件扩展名，如果没有则使用检测到的格式
            if ext and ext in format_map:
                input_format = format_map[ext]
            elif detected_format:
                input_format = detected_format
            else:
                input_format = 'markdown'
            
            # 使用明确的格式参数转换
            cmd = ['pandoc', '-f', input_format, '-t', 'html', 
                   input_file, '-o', output_file, 
                   '--standalone', '--embed-resources',
                   '--wrap=none']
            
            subprocess.run(cmd, check=True, capture_output=True)
            
            # 读取生成的HTML
            with open(output_file, 'r', encoding='utf-8') as f:
                html_content = f.read()
            
            # 如果有对齐信息，应用到HTML段落（根据配置过滤）
            if alignment_info:
                p_pattern = r'<p[^>]*>(.*?)</p>'
                paragraphs = re.findall(p_pattern, html_content, re.DOTALL)
                
                for i, (paragraph, alignment) in enumerate(zip(paragraphs, alignment_info)):
                    if i < len(paragraphs):
                        # 根据配置检查对齐样式是否被允许
                        allowed_styles = self.config.get_allowed_styles('doc')
                        style_map = {}
                        
                        if 'text-align' in allowed_styles:
                            if alignment == 'center':
                                style_map['text-align'] = 'center'
                            elif alignment == 'right':
                                style_map['text-align'] = 'right'
                            elif alignment == 'justify':
                                style_map['text-align'] = 'justify'
                            else:
                                style_map['text-align'] = 'left'
                        
                        # 构建样式字符串
                        style_str = '; '.join([f'{k}: {v}' for k, v in style_map.items()])
                        
                        old_p = f'<p>{paragraph}</p>'
                        new_p = f'<p style="{style_str}">{paragraph}</p>' if style_str else f'<p>{paragraph}</p>'
                        html_content = html_content.replace(old_p, new_p, 1)
            
            # 根据配置添加CSS样式
            allowed_styles = self.config.get_allowed_styles('doc')
            allowed_classes = self.config.get_allowed_classes('doc')
            
            css_style = """
            <style>
            body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; }
            p { margin-bottom: 1em; line-height: 1.6; }
            """
            
            # 根据允许的类名添加CSS规则
            if 'center' in allowed_classes:
                css_style += ".center { text-align: center !important; }\n"
            if 'right' in allowed_classes:
                css_style += ".right { text-align: right !important; }\n"
            if 'justify' in allowed_classes:
                css_style += ".justify { text-align: justify !important; }\n"
            
            css_style += "</style>\n"
            
            # 在</head>前插入CSS样式
            if '</head>' in html_content:
                html_content = html_content.replace('</head>', css_style + '</head>', 1)
            elif '<head>' in html_content:
                html_content = html_content.replace('<head>', '<head>' + css_style, 1)
            else:
                html_content = css_style + html_content
            
# 使用图片处理器处理HTML中的图片
            html_content = process_html_images(html_content, image_downloader)
            
            # 写回文件
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            return output_file
            
        except subprocess.CalledProcessError as e:
            raise RuntimeError(f"Conversion failed: {e.stderr.decode()}")
    
    def _convert_pdf_to_html(self, input_file: str, output_file: str) -> str:
        """使用pdfplumber将PDF转换为HTML，优化双层PDF处理"""
        if not PDFPLUMBER_AVAILABLE:
            raise RuntimeError("pdfplumber is not available for PDF conversion")
        
# 创建图片下载器，使用项目根目录的uploads/images目录
        project_root = Path(__file__).parent.parent
        images_dir = project_root / 'uploads' / 'images'
        image_downloader = ImageDownloader(str(images_dir))
        
        # 根据配置生成CSS样式
        allowed_styles = self.config.get_allowed_styles('pdf')
        allowed_classes = self.config.get_allowed_classes('pdf')
        
        css_style = """
        <style>
        body { 
            font-family: Arial, sans-serif; 
            max-width: 800px; 
            margin: 0 auto; 
            padding: 20px; 
            line-height: 1.6;
        }
        .page { 
            margin-bottom: 30px; 
            border-bottom: 1px solid #eee; 
            padding-bottom: 20px; 
        }
        .page-header { 
            font-weight: bold; 
            margin-bottom: 15px; 
            color: #666; 
        }
        p { 
            margin-bottom: 0.8em; 
            text-indent: 2em;
            white-space: pre-wrap;
        }
        h1, h2, h3, h4, h5, h6 { 
            margin: 1.5em 0 0.8em 0; 
            text-indent: 0;
        }
        """
        
        # 根据允许的类名添加CSS规则
        if 'center' in allowed_classes:
            css_style += ".center { text-align: center !important; }\n"
        if 'right' in allowed_classes:
            css_style += ".right { text-align: right !important; }\n"
        if 'bold' in allowed_classes:
            css_style += ".bold { font-weight: bold; }\n"
        if 'italic' in allowed_classes:
            css_style += ".italic { font-style: italic; }\n"
        if 'underline' in allowed_classes:
            css_style += ".underline { text-decoration: underline; }\n"
        
        css_style += "</style>\n"
        
        html_content = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>PDF Conversion</title>
    {css_style}
</head>
<body>
"""
        
        try:
            with pdfplumber.open(input_file) as pdf:
                for page_num, page in enumerate(pdf.pages, 1):
                    html_content += f"<div class='page' id='page-{page_num}'>\n"
                    html_content += f"<div class='page-header'>第 {page_num} 页</div>\n"
                    
                    # 提取图片
                    try:
                        if hasattr(page, 'images') and page.images:
                            for img_info in page.images:
                                try:
                                    # 提取图片数据
                                    image = page.within_bbox((img_info['x0'], img_info['top'], img_info['x1'], img_info['bottom'])).to_image()
                                    # 注意：pdfplumber的图片提取功能有限，这里使用PIL/Pillow方法
                                    # 由于pdfplumber的图片提取较复杂，这里提供基础框架
                                    pass
                                except Exception:
                                    pass
                    except Exception:
                        pass
                    
                    # 优先使用文本提取方法（适合双层PDF）
                    text = page.extract_text()
                    
                    if text and text.strip():
                        # 处理提取的文本
                        paragraphs = self._extract_text_paragraphs(text)
                        for paragraph in paragraphs:
                            if paragraph.strip():
                                html_content += f"<p>{self._escape_html(paragraph.strip())}</p>\n"
                    else:
                        # 回退到字符级处理（扫描版PDF）
                        chars = page.chars
                        if chars:
                            lines = self._group_chars_into_lines(chars)
                            for line_chars in lines:
                                line_chars.sort(key=lambda x: x['x0'])
                                line_text = ''.join(char['text'] for char in line_chars)
                                if line_text.strip():
                                    html_content += f"<p>{self._escape_html(line_text.strip())}</p>\n"
                    
                    html_content += "</div>\n"
            
            html_content += "</body>\n</html>"
            
            # 写入文件
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            return output_file
            
        except Exception as e:
            raise RuntimeError(f"PDF conversion failed: {str(e)}")
    
    def _extract_text_paragraphs(self, text: str) -> list:
        """从提取的文本中识别段落"""
        if not text:
            return []
        
        # 按换行符分割，过滤空行
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        paragraphs = []
        current_paragraph = ""
        
        for line in lines:
            # 如果行很短（可能是标题）或以特殊字符结尾（可能是段落结束）
            if len(line) < 20 and current_paragraph:
                if current_paragraph.strip():
                    paragraphs.append(current_paragraph.strip())
                current_paragraph = line
            elif line.endswith(('。', '！', '？', '.', '!', '?', ';', '：')):
                current_paragraph += line + " "
                paragraphs.append(current_paragraph.strip())
                current_paragraph = ""
            else:
                current_paragraph += line + " "
        
        # 添加最后一个段落
        if current_paragraph.strip():
            paragraphs.append(current_paragraph.strip())
        
        return [p for p in paragraphs if len(p) > 5]  # 过滤太短的段落
    
    def _group_chars_into_lines(self, chars):
        """改进的字符行分组算法"""
        if not chars:
            return []
        
        # 按Y坐标聚类，使用更精确的算法
        lines = []
        used_chars = set()
        
        # 按Y坐标排序
        sorted_chars = sorted(chars, key=lambda x: x['top'])
        
        for i, char in enumerate(sorted_chars):
            if i in used_chars:
                continue
                
            # 找到同一行的字符
            line_chars = [char]
            char_y = char['top']
            char_height = char.get('height', 10)
            
            # 使用字体高度的30%作为行内字符的Y坐标容差
            y_tolerance = char_height * 0.3
            
            for j, other_char in enumerate(sorted_chars):
                if j != i and j not in used_chars:
                    other_y = other_char['top']
                    if abs(other_y - char_y) <= y_tolerance:
                        line_chars.append(other_char)
                        used_chars.add(j)
            
            used_chars.add(i)
            lines.append(line_chars)
        
        return lines
    
    def _escape_html(self, text: str) -> str:
        """转义HTML特殊字符"""
        return (text.replace('&', '&amp;')
                    .replace('<', '&lt;')
                    .replace('>', '&gt;')
                    .replace('"', '&quot;')
                    .replace("'", '&#39;'))
    
    def _convert_epub_to_html(self, input_file: str, output_file: str) -> str:
        """直接解析EPUB并保留对齐等格式"""
        if not EPUBLIB_AVAILABLE:
            # 回退到pandoc方法
            return self._convert_epub_fallback(input_file, output_file)
        
        try:
            # 读取epub文件
            book = epub.read_epub(input_file)
            
            # 创建图片下载器，使用项目根目录的uploads/images目录
            project_root = Path(__file__).parent.parent
            images_dir = project_root / 'uploads' / 'images'
            image_downloader = ImageDownloader(str(images_dir))
            
            # 检查CSS文件，提取对齐样式
            css_content = ""
            images_map = {}  # 存储图片本地路径映射
            
            for item in book.get_items():
                # 处理CSS文件
                if hasattr(item, 'get_name') and item.get_name().endswith('.css'):
                    try:
                        content = item.get_content()
                        if isinstance(content, bytes):
                            css_content += content.decode('utf-8', errors='ignore') + "\n"
                        else:
                            css_content += str(content) + "\n"
                    except Exception:
                        pass
                
                # 处理图片文件
                elif hasattr(item, 'media_type') and item.media_type.startswith('image/'):
                    try:
                        image_data = item.get_content()
                        image_name = item.get_name().split('/')[-1]  # 获取文件名
                        
                        # 保存图片到本地
                        local_path = image_downloader.save_from_bytes(image_data, image_name)
                        if local_path:
                            # 生成相对路径
                            relative_path = image_downloader.get_relative_path(local_path)
                            images_map[image_name] = relative_path
                        
                    except Exception:
                        pass
            
            print(f"CSS content found: {len(css_content)} characters")
            print(f"CSS sample: {css_content[:500]}...")
            
            # 开始构建HTML
            html_content = """<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>""" + self._escape_html(book.get_metadata('DC', 'title')[0][0] if book.get_metadata('DC', 'title') else 'EPUB Conversion') + """</title>
    <style>
        body { 
            font-family: serif; 
            max-width: 800px; 
            margin: 0 auto; 
            padding: 20px; 
            line-height: 1.6;
        }
        p { margin-bottom: 1em; }
        h1, h2, h3, h4, h5, h6 { margin-top: 1.5em; margin-bottom: 0.8em; }
        /* 所有样式已转为行内样式，无需CSS class规则 */
    </style>
</head>
<body>
"""
            
            # 获取所有章节（按spine顺序）
            for spine_item in book.spine:
                # spine_item可能是字符串ID或元组
                if isinstance(spine_item, tuple):
                    spine_item = spine_item[0]
                
                # 获取实际的HTML内容项
                html_item = book.get_item_with_id(spine_item)
                if html_item and hasattr(html_item, 'media_type') and html_item.media_type == 'application/xhtml+xml':
                    try:
                        content = html_item.get_content()
                        if isinstance(content, bytes):
                            content = content.decode('utf-8', errors='ignore')
                        soup = BeautifulSoup(content, 'html.parser')
                        
                        # 应用CSS样式为行内样式
                        if css_content:
                            css_rules = self._parse_css_rules(css_content)
                            self._apply_css_inline(soup, css_rules, 'epub')
                        
# 替换图片src为本地路径
                        for img in soup.find_all('img'):
                            src = img.get('src', '')
                            if src and images_map:
                                # 尝试匹配图片文件名
                                for img_name, local_path in images_map.items():
                                    if img_name in src or src.endswith(img_name):
                                        img['src'] = local_path
                                        break
                        
                        # 直接使用已经处理过的HTML内容（包含行内样式和嵌套标签）
                        for element in soup.find_all(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'div', 'img']):
                            # 检查元素是否有实际内容（图片除外）
                            if element.name != 'img' and not element.get_text().strip():
                                continue
                            
                            # 获取处理后的属性，但跳过class（因为样式已转为行内）
                            attrs_str = ''
                            for attr_name, attr_value in element.attrs.items():
                                if attr_name == 'class':
                                    continue  # 跳过class属性
                                if attr_value:  # 只添加非空属性
                                    if isinstance(attr_value, list):
                                        attr_value = ' '.join(attr_value)
                                    attrs_str += f' {attr_name}="{attr_value}"'
                            
                            # 获取元素的完整HTML内容（包括嵌套标签）
                            element_html = str(element)
                            
                            # 移除外层标签，因为我们重新构建
                            inner_html = re.sub(rf'^<{element.name}[^>]*>(.*)</{element.name}>$', r'\1', element_html, flags=re.DOTALL)
                            
                            html_content += f'<{element.name}{attrs_str}>{inner_html}</{element.name}>\n'
                        
                        # 标题已在上面的循环中处理，无需重复
                    
                    except Exception:
                        continue
            
            html_content += """
</body>
</html>"""
            
            # 写入文件
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            return output_file
            
        except Exception:
            return self._convert_epub_fallback(input_file, output_file)
    
    def _parse_css_rules(self, css_content):
        """解析CSS规则，返回选择器和样式的映射"""
        rules = {}
        
        # 简单的CSS解析器
        # 移除注释
        css_content = re.sub(r'/\*.*?\*/', '', css_content, flags=re.DOTALL)
        
        # 匹配CSS规则
        pattern = r'([^{]+)\{([^}]+)\}'
        matches = re.findall(pattern, css_content)
        
        for selector, declarations in matches:
            selector = selector.strip()
            declarations = declarations.strip()
            
            # 清理声明
            style_dict = {}
            for decl in declarations.split(';'):
                decl = decl.strip()
                if ':' in decl:
                    prop, value = decl.split(':', 1)
                    style_dict[prop.strip()] = value.strip()
            
            # 处理多个选择器（逗号分隔）
            selectors = [s.strip() for s in selector.split(',')]
            for sel in selectors:
                rules[sel] = style_dict
        
        return rules
    
    def _apply_css_inline(self, soup, css_rules, format_type: str = 'epub'):
        """将CSS规则应用为行内样式，根据配置过滤样式"""
        
        for selector, styles in css_rules.items():
            # 处理class选择器 (.classname)
            if selector.startswith('.'):
                class_name = selector[1:]
                # 检查类名是否被允许
                if self.config.is_class_allowed(format_type, class_name):
                    elements = soup.find_all(class_=lambda c: c and class_name in str(c).split())
                    for element in elements:
                        self._add_inline_style(element, styles, format_type)
                        # 递归应用到子元素
                        for child in element.find_all(True):
                            self._add_inline_style(child, styles, format_type)
            
            # 处理标签选择器 (p, h1, span, em, strong, etc.)
            elif selector.isalpha() or selector in ['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'div', 'span', 'em', 'strong', 'b', 'i', 'u']:
                elements = soup.find_all(selector)
                for element in elements:
                    self._add_inline_style(element, styles, format_type)
            
            # 处理ID选择器 (#id)
            elif selector.startswith('#'):
                id_name = selector[1:]
                element = soup.find(id=id_name)
                if element:
                    self._add_inline_style(element, styles, format_type)
                    # 递归应用到子元素
                    for child in element.find_all(True):
                        self._add_inline_style(child, styles, format_type)
            
            # 处理组合选择器 (element.class)
            elif '.' in selector and not selector.startswith('.'):
                parts = selector.split('.')
                if len(parts) == 2:
                    tag, class_name = parts
                    # 检查类名是否被允许
                    if self.config.is_class_allowed(format_type, class_name):
                        elements = soup.find_all(tag, class_=lambda c: c and class_name in str(c).split())
                        for element in elements:
                            self._add_inline_style(element, styles, format_type)
                            # 递归应用到子元素
                            for child in element.find_all(True):
                                self._add_inline_style(child, styles, format_type)
    
    def _add_inline_style(self, element, styles, format_type: str = 'epub'):
        """为元素添加行内样式，根据配置过滤样式"""
        # 根据格式类型过滤样式
        filtered_styles = self.config.filter_styles(format_type, styles)
        
        existing_style = element.get('style', '')
        
        # 将过滤后的样式添加到现有样式
        for prop, value in filtered_styles.items():
            # 移除已存在的同名属性
            existing_style = re.sub(rf'{re.escape(prop)}\s*:\s*[^;]*;?', '', existing_style, flags=re.IGNORECASE)
            # 添加新样式
            if existing_style and not existing_style.endswith(';'):
                existing_style += ';'
            existing_style += f'{prop}: {value};'
        
        if existing_style:
            element['style'] = existing_style
    
    def _extract_alignment(self, element):
        """从元素中提取对齐信息"""
        # 检查class
        classes = element.get('class', [])
        if isinstance(classes, str):
            classes = [classes]
        
        for cls in classes:
            if 'center' in cls.lower():
                return 'center'
            elif 'right' in cls.lower():
                return 'right'
            elif 'justify' in cls.lower():
                return 'justify'
        
        # 检查style属性
        style = element.get('style', '').lower()
        if 'text-align:center' in style:
            return 'center'
        elif 'text-align:right' in style:
            return 'right'
        elif 'text-align:justify' in style:
            return 'justify'
        
        # 检查align属性
        align = element.get('align', '').lower()
        if align in ['center', 'right', 'justify']:
            return align
        
        return None
    
    def _convert_epub_fallback(self, input_file: str, output_file: str) -> str:
        """使用pandoc的回退方法"""
        try:
            cmd = ['pandoc', '-f', 'epub', '-t', 'html5',
                   input_file, '-o', output_file,
                   '--standalone', '--embed-resources',
                   '--wrap=none']
            
            subprocess.run(cmd, check=True, capture_output=True)
            return output_file
            
        except subprocess.CalledProcessError as e:
            raise RuntimeError(f"EPUB conversion failed: {e.stderr.decode()}")
    
    def get_supported_formats(self) -> list:
        """获取支持的输入格式"""
        formats = ['.doc', '.docx', '.epub', '.txt', '.rtf', '.md', '.html', '.htm', '.odt']
        if PDFPLUMBER_AVAILABLE:
            formats.append('.pdf')
        return formats
    
    def is_supported(self, file_path: str) -> bool:
        """检查文件格式是否支持"""
        return Path(file_path).suffix.lower() in self.get_supported_formats()
    
    def _process_paragraph_alignment(self, doc_json):
        """处理段落对齐信息"""
        if isinstance(doc_json, dict) and 'blocks' in doc_json:
            for block in doc_json['blocks']:
                if block[0] == 'Para':
                    # 检查段落是否有对齐信息
                    para_content = block[1]
                    # 如果没有对齐类，根据内容推断或添加默认类
                    if not any(isinstance(item, list) and len(item) > 1 and isinstance(item[1], dict) and 'classes' in item[1] 
                              for item in para_content if isinstance(item, list)):
                        # 添加默认左对齐类
                        for item in para_content:
                            if isinstance(item, list) and len(item) > 1 and isinstance(item[1], dict):
                                if 'classes' not in item[1]:
                                    item[1]['classes'] = []
                                if 'left-align' not in item[1]['classes']:
                                    item[1]['classes'].append('left-align')
                                break