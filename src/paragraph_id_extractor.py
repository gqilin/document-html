"""
Word文档段落ID提取器
利用Word内置的paraId实现精确的样式映射
"""

import xml.etree.ElementTree as ET
import re
from typing import Dict, Any, Optional
try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False


class ParagraphIdExtractor:
    """提取Word文档段落的唯一ID和样式信息"""
    
    def __init__(self):
        self.namespace = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        }
    
    def extract_paragraph_data(self, docx_file: str) -> Dict[str, Any]:
        """
        提取段落的paraId和样式信息
        
        Args:
            docx_file: Word文档路径
            
        Returns:
            {paraId: {样式信息}} 的字典
        """
        if not DOCX_AVAILABLE:
            raise ImportError("python-docx is required for paraId extraction")
        
        doc = Document(docx_file)
        paragraph_data = {}
        
        # 遍历所有段落
        for para in doc.paragraphs:
            para_id = self._get_paragraph_id(para)
            
            # 如果没有paraId，但有特殊内容需要处理，强制生成ID
            if not para_id:
                # 1. 包含图片的段落
                if self._paragraph_contains_image(para):
                    para_id = f"hash_img_{hash(str(para._element.xml)) % 1000000:06d}"
                # 2. 有对齐设置的段落（如图片标题）
                elif para.alignment is not None:
                    para_id = f"hash_align_{hash(str(para._element.xml)) % 1000000:06d}"
                # 3. 有特殊样式名的段落
                elif para.style and para.style.name:
                    para_id = f"hash_style_{hash(str(para._element.xml)) % 1000000:06d}"
            
            if para_id:
                paragraph_data[para_id] = {
                    'text': para.text,
                    'alignment': para.alignment,
                    'style_name': para.style.name if para.style else None,
                    'runs': self._extract_runs_data(para),
                    'paragraph_format': self._extract_paragraph_format(para)
                }
        
        return paragraph_data
    
    def _get_paragraph_id(self, para) -> Optional[str]:
        """
        从段落的XML中获取paraId
        
        Args:
            para: python-docx段落对象
            
        Returns:
            paraId字符串或None
        """
        try:
            # 访问底层XML
            para_xml = para._p.xml
            root = ET.fromstring(para_xml)
            
            # 查找paraId属性
            para_id = root.get(f"{{{self.namespace['w']}}}paraId")
            return para_id
            
        except Exception as e:
            # 如果获取失败，生成基于内容的hash
            text = para.text.strip()
            if text:
                return f"hash_{hash(text) % 1000000:06d}"
            return None
    
    def _extract_runs_data(self, para) -> list:
        """提取段落中文本片段的样式信息"""
        runs_data = []
        
        try:
            for i, run in enumerate(para.runs):
                run_data = {
                    'id': f"{self._get_paragraph_id(para)}_run_{i}",
                    'text': run.text,
                    'bold': run.bold,
                    'italic': run.italic,
                    'underline': run.underline,
                    'font_name': run.font.name,
                    'font_size': run.font.size.pt if run.font.size else None,
                    'font_color': self._get_color_hex(run.font.color) if run.font.color else None,
                    'highlight_color': self._get_color_hex(run.font.highlight) if run.font.highlight else None
                }
                runs_data.append(run_data)
                
        except Exception:
            pass
        
        return runs_data
    
    def _paragraph_contains_image(self, para) -> bool:
        """
        检查段落是否包含图片
        
        Args:
            para: python-docx段落对象
            
        Returns:
            如果包含图片返回True
        """
        try:
            xml = para._element.xml
            return any(keyword in xml.lower() for keyword in ['graphic', 'blip', 'a:blip', 'pic:pic'])
        except Exception:
            return False
        """提取段落中文本片段的样式信息"""
        runs_data = []
        
        try:
            for i, run in enumerate(para.runs):
                run_data = {
                    'id': f"{self._get_paragraph_id(para)}_run_{i}",
                    'text': run.text,
                    'bold': run.bold,
                    'italic': run.italic,
                    'underline': run.underline,
                    'font_name': run.font.name,
                    'font_size': run.font.size.pt if run.font.size else None,
                    'font_color': self._get_color_hex(run.font.color) if run.font.color else None,
                    'highlight_color': self._get_color_hex(run.font.highlight) if run.font.highlight else None
                }
                runs_data.append(run_data)
                
        except Exception:
            pass  # 如果提取失败，跳过runs
            
        return runs_data
    
    def _extract_paragraph_format(self, para) -> Dict[str, Any]:
        """提取段落格式信息"""
        try:
            # 访问段落的格式属性
            if hasattr(para, 'paragraph_format'):
                fmt = para.paragraph_format
                return {
                    'left_indent': fmt.left_indent.pt if fmt.left_indent else None,
                    'right_indent': fmt.right_indent.pt if fmt.right_indent else None,
                    'first_line_indent': fmt.first_line_indent.pt if fmt.first_line_indent else None,
                    'space_before': fmt.space_before.pt if fmt.space_before else None,
                    'space_after': fmt.space_after.pt if fmt.space_after else None,
                    'line_spacing': fmt.line_spacing
                }
        except Exception:
            pass
            
        return {}
    
    def _get_color_hex(self, color_obj) -> Optional[str]:
        """
        获取颜色的十六进制值
        
        Args:
            color_obj: docx颜色对象
            
        Returns:
            十六进制颜色字符串
        """
        try:
            if hasattr(color_obj, 'rgb'):
                # RGB颜色
                rgb = color_obj.rgb
                return f"#{rgb:06X}" if rgb else None
            elif hasattr(color_obj, 'theme_color'):
                # 主题颜色
                return f"theme-{color_obj.theme_color}"
        except Exception:
            pass
            
        return None
    
    def get_alignment_name(self, alignment) -> str:
        """将对齐枚举转换为字符串"""
        alignment_map = {
            WD_ALIGN_PARAGRAPH.LEFT: 'left',
            WD_ALIGN_PARAGRAPH.CENTER: 'center',
            WD_ALIGN_PARAGRAPH.RIGHT: 'right',
            WD_ALIGN_PARAGRAPH.JUSTIFY: 'justify',
            WD_ALIGN_PARAGRAPH.DISTRIBUTE: 'justify'
        }
        
        return alignment_map.get(alignment, 'left')
    
    def extract_with_text_fallback(self, docx_file: str) -> Dict[str, Any]:
        """
        如果paraId提取失败，使用文本哈希作为备用方案
        
        Args:
            docx_file: Word文档路径
            
        Returns:
            {paraId: {样式信息}} 的字典
        """
        paragraph_data = self.extract_paragraph_data(docx_file)
        
        # 检查是否有空的paraId
        has_empty_ids = any(not pid for pid in paragraph_data.keys())
        
        if has_empty_ids or len(paragraph_data) == 0:
            # 使用文本哈希作为备用方案
            doc = Document(docx_file)
            paragraph_data = {}
            
            for i, para in enumerate(doc.paragraphs):
                if para.text.strip():  # 只处理非空段落
                    # 基于文本内容生成稳定的ID
                    text_hash = f"text_{hash(para.text) % 1000000:06d}"
                    paragraph_data[text_hash] = {
                        'text': para.text,
                        'alignment': para.alignment,
                        'style_name': para.style.name if para.style else None,
                        'runs': self._extract_runs_data(para),
                        'paragraph_format': self._extract_paragraph_format(para),
                        'fallback': True  # 标记这是备用ID
                    }
        
        return paragraph_data


# 测试函数
def test_paragraph_id_extraction():
    """测试paraId提取功能"""
    if not DOCX_AVAILABLE:
        print("python-docx not available for testing")
        return
    
    extractor = ParagraphIdExtractor()
    
    # 这里需要实际的docx文件进行测试
    # docx_file = "test.docx"
    # data = extractor.extract_paragraph_data(docx_file)
    # print(f"提取到 {len(data)} 个段落的数据")
    
    print("ParagraphIdExtractor 类已实现，可以提取:")
    print("- Word原生paraId")
    print("- 段落对齐、样式信息")
    print("- 文本片段runs的字体颜色、大小等")
    print("- 段落格式信息")


if __name__ == "__main__":
    test_paragraph_id_extraction()