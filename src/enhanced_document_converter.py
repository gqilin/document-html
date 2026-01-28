"""
增强版文档转换器 - 集成paraId映射功能
替换当前的文本匹配逻辑为精确的paraId映射
"""

import os
import subprocess
import tempfile
from pathlib import Path
from typing import Optional, Dict, Any
import sys

# 添加项目路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 导入现有的模块
from src.config import ConversionConfig, DEFAULT_CONFIG
from src.image_downloader import ImageDownloader, process_html_images

# 导入新的paraId模块
from src.paragraph_id_extractor import ParagraphIdExtractor
from src.para_id_pandoc_converter import ParaIdPandocConverter
from src.para_id_style_applicator import ParaIdStyleApplicator

# 保留原有的导入以支持其他格式
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
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False

try:
    from ebooklib import epub
    from bs4 import BeautifulSoup
    EPUBLIB_AVAILABLE = True
except ImportError:
    EPUBLIB_AVAILABLE = False


class EnhancedDocumentConverter:
    """增强版文档转换器 - 使用paraId精确映射样式"""
    
    def __init__(self, config: Optional[ConversionConfig] = None):
        # 原有配置
        self.config = config if config is not None else DEFAULT_CONFIG
        
        # 创建图片目录并初始化图片下载器
        project_root = Path(__file__).parent.parent
        self.images_dir = project_root / 'uploads' / 'images'
        self.images_dir.mkdir(parents=True, exist_ok=True)
        self.image_downloader = ImageDownloader(str(self.images_dir))
        
        # 新的paraId组件
        self.para_id_extractor = ParagraphIdExtractor()
        self.pandoc_converter = ParaIdPandocConverter()
        self.style_applicator = ParaIdStyleApplicator(self.config)
        
        # 检查可用性
        self.pandoc_available = self.pandoc_converter.pandoc_available
        
        # 导入改进的PDF转换器
        if PYMUPDF_AVAILABLE:
            from src.pdf_converter import ImprovedPDFConverter
            self.pdf_converter = ImprovedPDFConverter()
    
    def convert_to_html(self, input_file: str, output_file: Optional[str] = None) -> Optional[str]:
        """
        增强版文档转换 - 支持paraId精确映射
        
        Args:
            input_file: 输入文件路径
            output_file: 输出文件路径，如果为None则自动生成
            
        Returns:
            转换后的HTML文件路径，失败返回None
        """
        input_path = Path(input_file)
        if not input_path.exists():
            raise FileNotFoundError(f"Input file not found: {input_file}")
        
        if output_file is None:
            output_file = str(input_path.with_suffix('.html'))
        
        # 检测文件类型并选择合适的转换方法
        file_type = self._detect_file_type(input_file)
        
        if file_type == 'docx':
            return self._convert_docx_with_para_id(input_file, output_file)
        elif file_type == 'pdf':
            return self._convert_pdf_to_html(input_file, output_file)
        elif file_type == 'epub':
            return self._convert_epub_to_html(input_file, output_file)
        else:
            # 对于其他格式，使用原有的pandoc转换
            return self._convert_with_pandoc(input_file, output_file)
    
    def _detect_file_type(self, file_path: str) -> str:
        """检测文件类型"""
        path = Path(file_path)
        ext = path.suffix.lower()
        
        # 首先基于扩展名
        ext_map = {
            '.docx': 'docx',
            '.doc': 'doc',
            '.pdf': 'pdf',
            '.epub': 'epub',
            '.txt': 'plain',
            '.rtf': 'rtf',
            '.md': 'markdown',
            '.html': 'html',
            '.htm': 'html',
            '.odt': 'odt'
        }
        
        if ext in ext_map:
            return ext_map[ext]
        
        # 如果扩展名不明确，基于文件头检测
        try:
            with open(file_path, 'rb') as f:
                header = f.read(200)
                
                if header.startswith(b'%PDF'):
                    return 'pdf'
                elif header.startswith(b'PK\\x03\\x04'):
                    # ZIP格式，可能是epub或docx
                    if b'mimetypeapplication/epub+zip' in header:
                        return 'epub'
                    elif b'word/' in header or b'docx' in header:
                        return 'docx'
                elif header.startswith(b'{\\\\rtf'):
                    return 'rtf'
        except Exception:
            pass
        
        return 'plain'  # 默认为纯文本
    
    def _convert_docx_with_para_id(self, docx_file: str, output_file: str) -> str:
        """
        使用paraId精确映射转换docx文件
        
        Args:
            docx_file: docx文件路径
            output_file: 输出HTML文件路径
            
        Returns:
            HTML文件路径
        """
        try:
            # 第一步：提取段落样式数据（包含paraId）
            print(f"提取paraId和样式信息...")
            paragraph_data = self.para_id_extractor.extract_with_text_fallback(docx_file)
            print(f"提取到 {len(paragraph_data)} 个段落的样式数据")
            
            # 第二步：提取图片
            print(f"提取文档中的图片...")
            self._extract_images_from_docx(docx_file)
            
            # 第三步：使用pandoc转换并保留paraId
            print(f"使用Pandoc转换并保留paraId...")
            html_with_ids = self.pandoc_converter.convert_with_para_id(docx_file, paragraph_data)
            print("Pandoc转换完成")
            
            # 第四步：根据paraId应用样式
            print(f"应用样式到HTML段落...")
            final_html = self.style_applicator.apply_styles_by_para_id(html_with_ids, paragraph_data)
            
            # 第五步：添加CSS样式和图片处理
            enhanced_html = self._enhance_html_with_resources(final_html, paragraph_data)
            
            # 第六步：保存结果
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(enhanced_html)
            
            print(f"转换完成: {output_file}")
            return output_file
            
        except Exception as e:
            print(f"paraId转换失败，回退到原有方法: {str(e)}")
            # 如果paraId方法失败，回退到原有方法
            return self._convert_with_pandoc(docx_file, output_file)
    
    def _extract_images_from_docx(self, docx_file: str):
        """
        从docx文件中提取图片到本地
        
        Args:
            docx_file: docx文件路径
        """
        import tempfile
        import shutil
        import zipfile
        from pathlib import Path
        
        try:
            # 方法1：直接从docx zip中提取
            print("  尝试从docx文件直接提取图片...")
            
            # 创建目标目录
            target_images_dir = Path('uploads/images')
            target_images_dir.mkdir(parents=True, exist_ok=True)
            
            extracted_count = 0
            with zipfile.ZipFile(docx_file, 'r') as zf:
                for file_info in zf.infolist():
                    if file_info.filename.startswith('word/media/') and file_info.filename.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp')):
                        # 提取文件名
                        filename = Path(file_info.filename).name
                        
                        # 提取图片数据
                        with zf.open(file_info) as source:
                            target_path = target_images_dir / filename
                            with open(target_path, 'wb') as target:
                                shutil.copyfileobj(source, target)
                            extracted_count += 1
                            print(f"    提取图片: {filename}")
            
            if extracted_count > 0:
                print(f"  成功提取 {extracted_count} 个图片")
                return
            
            # 方法2：使用pandoc --extract-media
            print("  尝试使用Pandoc提取图片...")
            
            with tempfile.TemporaryDirectory() as temp_dir:
                # 使用pandoc提取media
                cmd = [
                    'pandoc', '-f', 'docx', '-t', 'html',
                    '--extract-media=temp_media',
                    '--standalone',
                    docx_file
                ]
                
                # 在临时目录中运行
                original_cwd = os.getcwd()
                os.chdir(temp_dir)
                
                try:
                    result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
                    
                    # 检查是否创建了media目录
                    media_dir = Path(temp_dir) / 'temp_media'
                    if media_dir.exists():
                        copied_count = 0
                        for img_file in media_dir.rglob('*'):
                            if img_file.is_file() and img_file.suffix.lower() in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp']:
                                target_path = target_images_dir / img_file.name
                                shutil.copy2(img_file, target_path)
                                copied_count += 1
                                print(f"    复制图片: {img_file.name}")
                        
                        if copied_count > 0:
                            print(f"  成功复制 {copied_count} 个图片")
                    
                finally:
                    os.chdir(original_cwd)
                    
        except Exception as e:
            print(f"  图片提取失败: {e}")
            # 不阻止转换流程，只记录警告
    
    def _wrap_images_with_paragraphs(self, html_content: str, paragraph_data: Dict[str, Any]) -> str:
        """
        为独立的图片标签添加包装段落，应用对齐样式
        
        Args:
            html_content: HTML内容
            paragraph_data: 段落样式数据
            
        Returns:
            处理后的HTML内容
        """
        import re
        
        # 找到图片段落（空文本但有对齐设置）
        image_paragraphs = []
        for para_id, data in paragraph_data.items():
            text = data.get('text', '')
            alignment = data.get('alignment')
            
            if text.strip() == '' and alignment is not None:
                image_paragraphs.append((para_id, alignment))
        
        if not image_paragraphs:
            return html_content
        
        print(f"  为 {len(image_paragraphs)} 个图片段落添加包装...")
        
        # 创建段落ID到对齐的映射
        alignment_map = {para_id: alignment for para_id, alignment in image_paragraphs}
        
        # 查找独立的图片标签并包装
        def wrap_images(match):
            img_tag = match.group(0)
            
            # 尝试查找最近的未使用的图片段落样式
            for para_id, alignment in image_paragraphs:
                if para_id in alignment_map:
                    # 生成CSS样式
                    from src.para_id_style_applicator import ParaIdStyleApplicator
                    applicator = ParaIdStyleApplicator()
                    
                    # 转换对齐方式
                    alignment_name = applicator.alignment_map.get(alignment, 'left')
                    
                    # 添加样式和ID
                    wrapped_img = f'<p id="{para_id}" class="word-style-image-paragraph" style="text-align: {alignment_name}">{img_tag}</p>'
                    
                    # 从映射中移除已使用的段落
                    del alignment_map[para_id]
                    
                    print(f"    为图片添加包装: {para_id} -> {alignment_name}")
                    return wrapped_img
            
            # 如果没有找到匹配的段落，返回原标签
            return img_tag
        
        # 查找独立的图片标签（不在p标签内的）
        # 简化方法：先查找所有img标签，然后检查是否在p标签内
        img_matches = list(re.finditer(r'<img[^>]+>', html_content))
        processed_html = html_content
        
        # 查找只包含图片的段落（<p><img /></p>格式）
        simple_img_pattern = r'<p>\s*<img[^>]+>\s*</p>'
        img_para_matches = list(re.finditer(simple_img_pattern, processed_html))
        
        # 从后往前处理，避免位置偏移
        for match in reversed(img_para_matches):
            full_match = match.group()
            
            # 提取其中的img标签
            img_match = re.search(r'<img[^>]+>', full_match)
            if not img_match:
                continue
                
            img_tag = img_match.group()
            
            # 这是纯图片段落，需要添加对齐样式
            for para_id, alignment in image_paragraphs:
                if para_id in alignment_map:
                    # 转换对齐方式
                    alignment_name = self.style_applicator.alignment_map.get(alignment, 'left')
                    
                    # 添加样式和ID
                    wrapped_img = f'<p id="{para_id}" class="word-style-image-paragraph" style="text-align: {alignment_name}">{img_tag}</p>'
                    
                    # 替换整个匹配的段落
                    start_pos = match.start()
                    end_pos = match.end()
                    processed_html = processed_html[:start_pos] + wrapped_img + processed_html[end_pos:]
                    
                    # 从映射中移除已使用的段落
                    del alignment_map[para_id]
                    
                    print(f"    为图片段落添加样式: {para_id} -> {alignment_name}")
                    break
        
        return processed_html
    
    def _enhance_html_with_resources(self, html_content: str, paragraph_data: Dict[str, Any]) -> str:
        """
        增强HTML：添加CSS样式和图片处理
        
        Args:
            html_content: 原始HTML内容
            paragraph_data: 段落样式数据
            
        Returns:
            增强后的HTML
        """
        # 处理图片路径
        html_content = process_html_images(html_content, self.image_downloader)
        
        # 为独立图片添加包装段落和样式
        html_content = self._wrap_images_with_paragraphs(html_content, paragraph_data)
        
        # 生成CSS样式
        css_rules = self.style_applicator.generate_css_rules(paragraph_data)
        
        # 将CSS插入HTML
        if '</head>' in html_content:
            css_block = f"<style>\\n{css_rules}\\n</style>\\n"
            html_content = html_content.replace('</head>', css_block + '</head>')
        elif '<head>' in html_content:
            css_block = f"<style>\\n{css_rules}\\n</style>\\n"
            html_content = html_content.replace('<head>', '<head>' + css_block)
        else:
            # 如果没有head标签，添加到开头
            css_block = f"<head><style>\\n{css_rules}\\n</style></head>\\n"
            html_content = css_block + html_content
        
        return html_content
    
    def _convert_with_pandoc(self, input_file: str, output_file: str) -> str:
        """
        使用原有pandoc方法转换文档
        
        Args:
            input_file: 输入文件路径
            output_file: 输出文件路径
            
        Returns:
            HTML文件路径
        """
        if not self.pandoc_available:
            raise RuntimeError("pandoc is not available")
        
        try:
            # 确定输入格式
            file_type = self._detect_file_type(input_file)
            format_map = {
                'docx': 'docx',
                'doc': 'doc',
                'epub': 'epub',
                'plain': 'plain',
                'rtf': 'rtf',
                'markdown': 'markdown',
                'html': 'html',
                'odt': 'odt'
            }
            
            input_format = format_map.get(file_type, 'plain')
            
            # 确保epub文件使用正确的格式
            if file_type == 'epub' and input_format == 'plain':
                input_format = 'epub'  # 强制使用epub格式
                print(f"强制将epub文件格式设置为 'epub'")
            
            # 对于非docx格式，使用pandoc直接转换（不使用paraId相关逻辑）
            if file_type != 'docx':
                cmd = [
                    'pandoc', '-f', input_format, '-t', 'html',
                    input_file, '-o', output_file,
                    '--standalone', '--embed-resources',
                    '--wrap=none'
                ]
                
                subprocess.run(cmd, check=True, capture_output=True)
                return output_file
            
            # docx格式：使用paraId相关逻辑
            # 第一步：提取段落样式数据（包含paraId）
            print(f"提取paraId和样式信息...")
            paragraph_data = self.para_id_extractor.extract_with_text_fallback(input_file)
            print(f"提取到 {len(paragraph_data)} 个段落的样式数据")
            
            # 第二步：提取图片
            print(f"提取文档中的图片...")
            self._extract_images_from_docx(input_file)
            
            # 第三步：使用pandoc转换并保留paraId
            print(f"使用Pandoc转换并保留paraId...")
            html_with_ids = self.pandoc_converter.convert_with_para_id(input_file, paragraph_data)
            print("Pandoc转换完成")
            
            # 第四步：根据paraId应用样式
            print(f"应用样式到HTML段落...")
            final_html = self.style_applicator.apply_styles_by_para_id(html_with_ids, paragraph_data)
            
            # 第五步：添加CSS样式和图片处理
            enhanced_html = self._enhance_html_with_resources(final_html, paragraph_data)
            
            # 第六步：保存结果
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(enhanced_html)
            
            print(f"转换完成: {output_file}")
            return output_file
            
        except subprocess.CalledProcessError as e:
            raise RuntimeError(f"Pandoc conversion failed: {e.stderr.decode()}")
    
    def _convert_pdf_to_html(self, input_file: str, output_file: str) -> str:
        """PDF转换（使用改进的转换器）"""
        if PYMUPDF_AVAILABLE:
            # 使用改进的PyMuPDF转换器
            print("使用改进的PyMuPDF转换器...")
            return self.pdf_converter.convert_to_html(input_file, output_file)
        elif PDFPLUMBER_AVAILABLE:
            # 回退到旧的pdfplumber方案
            print("PyMuPDF不可用，使用pdfplumber转换器...")
            return self._convert_pdf_with_pdfplumber(input_file, output_file)
        else:
            raise RuntimeError("没有可用的PDF转换库（需要PyMuPDF或pdfplumber）")
    
    def _convert_pdf_with_pdfplumber(self, input_file: str, output_file: str) -> str:
        """PDF转换（使用pdfplumber提取文本 - 旧方案）"""
        import pdfplumber
        from html import escape
        
        try:
            html_parts = []
            html_parts.append('<!DOCTYPE html>')
            html_parts.append('<html>')
            html_parts.append('<head>')
            html_parts.append('<meta charset="UTF-8">')
            html_parts.append('<title>PDF Conversion</title>')
            html_parts.append('<style>')
            html_parts.append('body { font-family: Arial, sans-serif; line-height: 1.6; margin: 40px; }')
            html_parts.append('p { margin: 0.5em 0; }')
            html_parts.append('</style>')
            html_parts.append('</head>')
            html_parts.append('<body>')
            
            with pdfplumber.open(input_file) as pdf:
                for page_num, page in enumerate(pdf.pages, 1):
                    # 提取页面文本
                    text = page.extract_text()
                    if text:
                        # 将文本按段落分割
                        paragraphs = text.split('\n')
                        for para in paragraphs:
                            para = para.strip()
                            if para:
                                # 转义HTML特殊字符
                                escaped_para = escape(para)
                                html_parts.append(f'<p>{escaped_para}</p>')
                    
                    # 添加分页标记（可选）
                    if page_num < len(pdf.pages):
                        html_parts.append('<hr style="page-break-after: always;" />')
            
            html_parts.append('</body>')
            html_parts.append('</html>')
            
            # 写入输出文件
            html_content = '\n'.join(html_parts)
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            return output_file
            
        except Exception as e:
            raise RuntimeError(f"PDF conversion failed: {str(e)}")
    
    def _convert_epub_to_html(self, input_file: str, output_file: str) -> str:
        """EPUB转换（使用pandoc直接指定epub格式）"""
        if not self.pandoc_available:
            raise RuntimeError("pandoc is not available")
        
        # 直接使用pandoc转换，强制指定epub格式
        cmd = [
            'pandoc', '-f', 'epub', '-t', 'html',
            input_file, '-o', output_file,
            '--standalone', '--embed-resources',
            '--wrap=none'
        ]
        
        try:
            subprocess.run(cmd, check=True, capture_output=True)
            return output_file
        except subprocess.CalledProcessError as e:
            raise RuntimeError(f"Pandoc conversion failed: {e.stderr.decode()}")
    
    def get_supported_formats(self) -> list:
        """获取支持的输入格式"""
        formats = ['.doc', '.docx', '.epub', '.txt', '.rtf', '.md', '.html', '.htm', '.odt']
        if PDFPLUMBER_AVAILABLE:
            formats.append('.pdf')
        return formats
    
    def is_supported(self, file_path: str) -> bool:
        """检查文件格式是否支持"""
        return Path(file_path).suffix.lower() in self.get_supported_formats()
    
    def test_conversion(self, docx_file: str) -> Dict[str, Any]:
        """
        测试转换功能，返回详细的转换信息
        
        Args:
            docx_file: 测试docx文件路径
            
        Returns:
            转换结果信息
        """
        if not os.path.exists(docx_file):
            return {"error": "测试文件不存在"}
        
        result = {
            "file": docx_file,
            "para_id_extraction": False,
            "pandoc_conversion": False,
            "style_application": False,
            "final_result": False,
            "error": None
        }
        
        try:
            # 测试paraId提取
            paragraph_data = self.para_id_extractor.extract_with_text_fallback(docx_file)
            result["para_id_extraction"] = True
            result["paragraphs_extracted"] = len(paragraph_data)
            
            # 测试pandoc转换
            html_with_ids = self.pandoc_converter.convert_with_para_id(docx_file, paragraph_data)
            result["pandoc_conversion"] = True
            
            # 测试样式应用
            final_html = self.style_applicator.apply_styles_by_para_id(html_with_ids, paragraph_data)
            result["style_application"] = True
            
            result["final_result"] = True
            result["html_length"] = len(final_html)
            
        except Exception as e:
            result["error"] = str(e)
        
        return result
    
    def _extract_images_from_docx(self, docx_file: str):
        """
        从docx文件中提取图片到本地
        
        Args:
            docx_file: docx文件路径
        """
        import tempfile
        import shutil
        import zipfile
        from pathlib import Path
        
        try:
            # 方法1：直接从docx zip中提取
            print("  尝试从docx文件直接提取图片...")
            
            # 创建目标目录
            target_images_dir = Path('uploads/images')
            target_images_dir.mkdir(parents=True, exist_ok=True)
            
            extracted_count = 0
            with zipfile.ZipFile(docx_file, 'r') as zf:
                for file_info in zf.infolist():
                    if file_info.filename.startswith('word/media/') and file_info.filename.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp')):
                        # 提取文件名
                        filename = Path(file_info.filename).name
                        
                        # 提取图片数据
                        with zf.open(file_info) as source:
                            target_path = target_images_dir / filename
                            with open(target_path, 'wb') as target:
                                shutil.copyfileobj(source, target)
                            extracted_count += 1
                            print(f"    提取图片: {filename}")
            
            if extracted_count > 0:
                print(f"  成功提取 {extracted_count} 个图片")
                return
            
            # 方法2：使用pandoc --extract-media
            print("  尝试使用Pandoc提取图片...")
            
            with tempfile.TemporaryDirectory() as temp_dir:
                # 使用pandoc提取media
                cmd = [
                    'pandoc', '-f', 'docx', '-t', 'html',
                    '--extract-media=temp_media',
                    '--standalone',
                    docx_file
                ]
                
                # 在临时目录中运行
                original_cwd = os.getcwd()
                os.chdir(temp_dir)
                
                try:
                    result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
                    
                    # 检查是否创建了media目录
                    media_dir = Path(temp_dir) / 'temp_media'
                    if media_dir.exists():
                        copied_count = 0
                        for img_file in media_dir.rglob('*'):
                            if img_file.is_file() and img_file.suffix.lower() in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp']:
                                target_path = target_images_dir / img_file.name
                                shutil.copy2(img_file, target_path)
                                copied_count += 1
                                print(f"    复制图片: {img_file.name}")
                        
                        if copied_count > 0:
                            print(f"  成功复制 {copied_count} 个图片")
                    
                finally:
                    os.chdir(original_cwd)
                    
        except Exception as e:
            print(f"  图片提取失败: {e}")
            # 不阻止转换流程，只记录警告


# 测试函数
def test_enhanced_converter():
    """测试增强版转换器"""
    converter = EnhancedDocumentConverter()
    
    print("EnhancedDocumentConverter 已实现")
    print("功能:")
    print("- paraId精确映射样式转换")
    print("- 文件类型自动检测")
    print("- 回退机制保证兼容性")
    print("- 多格式支持")
    print("- 转换测试和诊断")
    
    # 测试支持格式
    print(f"\\n支持的格式: {converter.get_supported_formats()}")
    print(f"Pandoc可用: {converter.pandoc_available}")


if __name__ == "__main__":
    test_enhanced_converter()