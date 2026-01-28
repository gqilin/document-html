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


# Word高亮颜色到CSS十六进制的映射
# 基于 WD_COLOR_INDEX 枚举值
HIGHLIGHT_COLOR_MAP = {
    1: '#000000',   # BLACK
    2: '#0000FF',   # BLUE
    3: '#00FFFF',   # TURQUOISE
    4: '#00FF00',   # BRIGHT_GREEN
    5: '#FF00FF',   # PINK
    6: '#FF0000',   # RED
    7: '#FFFF00',   # YELLOW
    8: '#FFFFFF',   # WHITE
    9: '#000080',   # DARK_BLUE
    10: '#008080',  # TEAL
    11: '#008000',  # GREEN
    12: '#800080',  # VIOLET
    13: '#800000',  # DARK_RED
    14: '#808000',  # DARK_YELLOW
    15: '#808080',  # GRAY_50
    16: '#C0C0C0',  # GRAY_25
}


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
            
            # 如果没有paraId，但有文本内容或包含图片，强制生成ID
            has_image = self._paragraph_contains_image(para)
            if not para_id and (para.text.strip() or has_image):
                para_id = f"hash_text_{hash(str(para._element.xml)) % 1000000:06d}"
            
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
                try:
                    # 安全获取属性
                    font_color_hex = None
                    if run.font.color:
                        font_color_hex = self._get_color_hex(run.font.color)

                    # 字体大小处理
                    font_size = None
                    if run.font.size:
                        try:
                            font_size = run.font.size.pt
                        except:
                            # 如果不是Length对象，尝试直接使用
                            font_size = float(run.font.size)

                    # 提取高亮颜色
                    highlight_color_hex = None
                    try:
                        highlight = run.font.highlight_color
                        if highlight:
                            # highlight 是 WD_COLOR_INDEX 枚举，获取其数值
                            highlight_value = highlight.value if hasattr(highlight, 'value') else int(highlight)
                            highlight_color_hex = HIGHLIGHT_COLOR_MAP.get(highlight_value)
                    except ValueError:
                        # 忽略 'none' 值等无效高亮颜色
                        pass
                    except Exception:
                        # 其他错误也忽略，不影响整体流程
                        pass

                    run_data = {
                        'id': f"run_{i}",  # 简化ID生成
                        'text': run.text if run.text else '',
                        'bold': run.bold,
                        'italic': run.italic,
                        'underline': run.underline,
                        'font_name': run.font.name,
                        'font_size': font_size,
                        'font_color': font_color_hex,
                        'highlight_color': highlight_color_hex  # 现在正确提取高亮颜色
                    }
                    runs_data.append(run_data)
                except Exception as e:
                    # 单个run失败不影响其他run
                    print(f"Run {i} 提取失败: {e}")
                    # 添加基本的run数据
                    runs_data.append({
                        'id': f"run_{i}",
                        'text': run.text if run.text else '',
                        'bold': False,
                        'italic': False,
                        'underline': False,
                        'font_name': None,
                        'font_size': None,
                        'font_color': None,
                        'highlight_color': None
                    })
                
        except Exception as e:
            print(f"提取runs数据失败: {e}")
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
                    'space_after': fmt.space_after.pt if fmt.space_before else None,
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
            if color_obj is None:
                return None
                
            # 检查是否有RGB颜色
            if hasattr(color_obj, 'rgb') and color_obj.rgb:
                # RGB颜色 - 可能是RGBColor对象或整数
                rgb = color_obj.rgb
                if hasattr(rgb, '__iter__'):  # 如果是可迭代的(R, G, B)
                    r, g, b = rgb
                    return f"#{r:02X}{g:02X}{b:02X}"
                elif isinstance(rgb, int):  # 如果是整数
                    return f"#{rgb:06X}"
                else:
                    # 尝试转换为整数
                    rgb_int = int(rgb)
                    return f"#{rgb_int:06X}"
            
            # 检查主题颜色
            if hasattr(color_obj, 'theme_color') and color_obj.theme_color is not None:
                # 尝试获取主题颜色的实际值
                theme_color = color_obj.theme_color
                return f"theme-{theme_color}"
                
            # 检查其他颜色属性
            if hasattr(color_obj, 'color') and color_obj.color:
                color_val = color_obj.color
                if isinstance(color_val, int):
                    return f"#{color_val:06X}"
                elif hasattr(color_val, '__iter__'):
                    r, g, b = color_val
                    return f"#{r:02X}{g:02X}{b:02X}"
                
        except Exception as e:
            # 静默处理颜色提取错误，不影响转换流程
            pass
            
        return None
    
    def get_alignment_name(self, alignment) -> str:
        """将对齐枚举转换为字符串"""
        if not DOCX_AVAILABLE:
            return 'left'
            
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