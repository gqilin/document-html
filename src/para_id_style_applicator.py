"""
基于paraId的HTML样式应用器
精确匹配HTML段落并应用Word样式
"""

import re
import logging
from typing import Dict, Any, Optional, List

try:
    from bs4 import BeautifulSoup
    BS4_AVAILABLE = True
except ImportError:
    BS4_AVAILABLE = False
    BeautifulSoup = None

try:
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

# 配置日志
logger = logging.getLogger(__name__)


class ParaIdStyleApplicator:
    """基于paraId的HTML样式应用器"""
    
    def __init__(self, config=None):
        self.bs4_available = BS4_AVAILABLE
        self.config = config
        
        # 安全的alignment映射，避免导入错误
        self.alignment_map = {
            None: 'left',
            'left': 'left',
            'center': 'center', 
            'right': 'right',
            'justify': 'justify',
            'distribute': 'justify'
        }
        
        # 如果docx库可用，添加枚举映射
        if DOCX_AVAILABLE:
            try:
                from docx.enum.text import WD_ALIGN_PARAGRAPH
                self.alignment_map.update({
                    WD_ALIGN_PARAGRAPH.LEFT: 'left',
                    WD_ALIGN_PARAGRAPH.CENTER: 'center',
                    WD_ALIGN_PARAGRAPH.RIGHT: 'right',
                    WD_ALIGN_PARAGRAPH.JUSTIFY: 'justify',
                    WD_ALIGN_PARAGRAPH.DISTRIBUTE: 'justify'
                })
                logger.info("成功加载WD_ALIGN_PARAGRAPH映射")
            except Exception as e:
                logger.warning(f"无法初始化WD_ALIGN_PARAGRAPH映射: {e}")
        else:
            logger.info("docx库不可用，使用字符串映射")
    
    def apply_styles_by_para_id(self, html_content: str, paragraph_data: Dict[str, Any]) -> str:
        """
        根据paraId应用样式到HTML
        
        Args:
            html_content: 原始HTML内容
            paragraph_data: 段落样式数据 {paraId: 样式信息}
            
        Returns:
            应用样式后的HTML
        """
        logger.info(f"开始应用样式，HTML长度: {len(html_content)}, 段落数: {len(paragraph_data)}")
        
        if not self.bs4_available:
            logger.warning("BeautifulSoup不可用，使用正则表达式方式")
            # 如果BeautifulSoup不可用，使用正则表达式方式
            return self._apply_styles_with_regex(html_content, paragraph_data)
        
        try:
            if not BS4_AVAILABLE or BeautifulSoup is None:
                raise ImportError("BeautifulSoup not available")
            soup = BeautifulSoup(html_content, 'html.parser')
            
            logger.info(f"BeautifulSoup解析成功，找到段落元素: {len(soup.find_all(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6']))} 个")
            
            elements_with_id = 0
            elements_styled = 0
            
            # 处理段落元素
            for element in soup.find_all(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6']):
                para_id = element.get('id')
                if para_id:
                    elements_with_id += 1
                    logger.debug(f"找到带ID元素: {element.name}#{para_id}")
                    
                if para_id and para_id in paragraph_data:
                    elements_styled += 1
                    style_info = paragraph_data[para_id]
                    logger.debug(f"应用样式到 {element.name}#{para_id}: {style_info}")
                    
                    self._apply_paragraph_style(element, style_info)
                    
                    # 处理内联元素（runs）
                    if 'runs' in style_info:
                        self._apply_inline_styles(element, style_info['runs'])
            
            logger.info(f"样式应用完成: {elements_with_id} 个带ID元素, {elements_styled} 个成功应用样式")
            
            return str(soup)
            
        except Exception as e:
            # 如果BeautifulSoup处理失败，回退到正则表达式
            logger.error(f"BeautifulSoup处理失败: {e}")
            logger.info("回退到正则表达式方式")
            return self._apply_styles_with_regex(html_content, paragraph_data)
    
    def _apply_paragraph_style(self, element, style_info: Dict[str, Any]):
        """
        为段落元素应用样式
        
        Args:
            element: BeautifulSoup元素
            style_info: 样式信息字典
        """
        css_styles = []
        
        logger.debug(f"应用段落样式: {element.name}#{element.get('id', 'no-id')} -> {style_info}")
        
        # 1. 对齐方式
        alignment = style_info.get('alignment')
        if alignment:
            alignment_name = self.alignment_map.get(alignment, 'left')
            if alignment_name != 'left':  # 只添加非默认的对齐方式
                css_styles.append(f"text-align: {alignment_name}")
                logger.debug(f"  对齐方式: {alignment} -> {alignment_name}")
        
        # 2. 段落格式
        paragraph_format = style_info.get('paragraph_format', {})
        
        # 缩进
        if paragraph_format.get('left_indent'):
            css_styles.append(f"margin-left: {paragraph_format['left_indent']}pt")
            logger.debug(f"  左缩进: {paragraph_format['left_indent']}pt")
        
        if paragraph_format.get('right_indent'):
            css_styles.append(f"margin-right: {paragraph_format['right_indent']}pt")
            logger.debug(f"  右缩进: {paragraph_format['right_indent']}pt")
        
        if paragraph_format.get('first_line_indent'):
            css_styles.append(f"text-indent: {paragraph_format['first_line_indent']}pt")
            logger.debug(f"  首行缩进: {paragraph_format['first_line_indent']}pt")
        
        # 段落间距
        if paragraph_format.get('space_before'):
            css_styles.append(f"margin-top: {paragraph_format['space_before']}pt")
            logger.debug(f"  段前间距: {paragraph_format['space_before']}pt")
        
        if paragraph_format.get('space_after'):
            css_styles.append(f"margin-bottom: {paragraph_format['space_after']}pt")
            logger.debug(f"  段后间距: {paragraph_format['space_after']}pt")
        
        # 3. 行间距
        line_spacing = paragraph_format.get('line_spacing')
        if line_spacing:
            if isinstance(line_spacing, (int, float)):
                css_styles.append(f"line-height: {line_spacing}")
                logger.debug(f"  行间距: {line_spacing}")
        
        # 应用样式到元素
        if css_styles:
            self._add_css_styles(element, css_styles)
            logger.debug(f"  最终CSS样式: {'; '.join(css_styles)}")
        
        # 添加样式类名
        style_name = style_info.get('style_name')
        if style_name:
            # 将样式名转换为CSS类名
            class_name = self._style_name_to_class(style_name)
            existing_classes = element.get('class', [])
            if isinstance(existing_classes, str):
                existing_classes = [existing_classes]
            
            if class_name not in existing_classes:
                existing_classes.append(class_name)
                element['class'] = existing_classes
                logger.debug(f"  添加CSS类: {class_name}")
    
    def _apply_inline_styles(self, paragraph_element, runs_data: List[Dict[str, Any]]):
        """
        应用内联样式到段落中的文本片段
        智能判断：如果所有run的样式一致，则将样式应用到段落级别
        如果样式不一致，则只对不同样式的文本应用span标签
        
        Args:
            paragraph_element: 段落元素
            runs_data: runs样式数据列表
        """
        if not runs_data:
            return
        
        # 获取段落纯文本内容
        para_text = paragraph_element.get_text()
        if not para_text.strip():
            return
        
        # 过滤掉空文本的runs
        valid_runs = [r for r in runs_data if r.get('text', '').strip()]
        if not valid_runs:
            return
        
        # 检查是否所有run的字体样式都一致
        all_same_style = self._check_all_runs_same_style(valid_runs)
        
        if all_same_style and len(valid_runs) > 0:
            # 所有run样式一致，将样式应用到段落级别
            logger.debug(f"  检测到段落级统一样式，应用到p标签")
            self._apply_uniform_style_to_paragraph(paragraph_element, valid_runs[0])
            return
        
        # 样式不一致，需要对不同样式的文本应用span标签
        logger.debug(f"  检测到混合样式，应用内联span标签")
        self._apply_mixed_inline_styles(paragraph_element, valid_runs)
    
    def _check_all_runs_same_style(self, runs_data: List[Dict[str, Any]]) -> bool:
        """
        检查所有run的字体样式是否完全一致
        
        Args:
            runs_data: runs样式数据列表
            
        Returns:
            如果所有run样式一致返回True
        """
        if len(runs_data) <= 1:
            return True
        
        # 获取第一个run的样式作为基准
        first_run = runs_data[0]
        style_keys = ['bold', 'italic', 'underline', 'font_name', 'font_size', 'font_color', 'highlight_color']
        
        for run in runs_data[1:]:
            for key in style_keys:
                if first_run.get(key) != run.get(key):
                    return False
        
        return True
    
    def _apply_uniform_style_to_paragraph(self, paragraph_element, run_data: Dict[str, Any]):
        """
        将统一样式应用到段落级别（p标签）
        
        Args:
            paragraph_element: 段落元素
            run_data: run样式数据
        """
        css_styles = []
        
        # 字体名称
        font_name = run_data.get('font_name')
        if font_name:
            css_styles.append(f"font-family: '{font_name}'")
        
        # 字体大小
        font_size = run_data.get('font_size')
        if font_size:
            try:
                size_value = float(font_size)
                css_styles.append(f"font-size: {size_value}pt")
            except (ValueError, TypeError):
                pass
        
        # 字体颜色
        font_color = run_data.get('font_color')
        if font_color and font_color.startswith('#'):
            css_styles.append(f"color: {font_color}")
        
        # 高亮颜色
        highlight_color = run_data.get('highlight_color')
        if highlight_color and highlight_color.startswith('#'):
            css_styles.append(f"background-color: {highlight_color}")
        
        # 应用样式到段落元素
        if css_styles:
            self._add_css_styles(paragraph_element, css_styles)
            logger.debug(f"  段落级样式: {'; '.join(css_styles)}")
        
        # 处理加粗、斜体、下划线 - 如果整个段落都是这些样式，可以考虑使用strong/em标签
        # 但这里我们保持段落内容不变，因为这些是语义标签
        # 如果需要，可以添加font-weight/font-style到CSS
        if run_data.get('bold'):
            existing_style = paragraph_element.get('style', '')
            if 'font-weight' not in existing_style:
                self._add_css_styles(paragraph_element, ['font-weight: bold'])
        
        if run_data.get('italic'):
            existing_style = paragraph_element.get('style', '')
            if 'font-style' not in existing_style:
                self._add_css_styles(paragraph_element, ['font-style: italic'])
        
        if run_data.get('underline'):
            existing_style = paragraph_element.get('style', '')
            if 'text-decoration' not in existing_style:
                self._add_css_styles(paragraph_element, ['text-decoration: underline'])
    
    def _apply_mixed_inline_styles(self, paragraph_element, runs_data: List[Dict[str, Any]]):
        """
        对混合样式的段落应用内联span标签
        只对与段落默认样式不同的run应用span标签
        
        Args:
            paragraph_element: 段落元素
            runs_data: runs样式数据列表
        """
        # 获取段落HTML内容
        para_html = str(paragraph_element)
        if not para_html:
            return
        
        # 找到最常见的样式作为段落默认样式
        default_style = self._get_most_common_style(runs_data)
        logger.debug(f"  段落默认样式: {default_style}")
        
        processed_runs = set()
        
        # 遍历每个run并应用样式
        for i, run_data in enumerate(runs_data):
            run_text = run_data.get('text', '')
            if not run_text or i in processed_runs:
                continue
            
            # 检查这个run的样式是否与默认样式不同
            if self._is_same_style(run_data, default_style):
                # 样式与默认相同，不需要添加span
                logger.debug(f"  跳过默认样式run: {run_text[:20]}...")
                processed_runs.add(i)
                continue
            
            # 创建带样式的run文本
            styled_run = self._create_styled_run(run_text, run_data)
            
            # 尝试替换HTML中的文本
            import html
            import re
            
            run_html = html.escape(run_text)
            run_html_escaped = run_text.replace('<', '&lt;').replace('>', '&gt;').replace('&', '&amp;')
            
            # 尝试多种替换方式
            new_html = None
            if run_text in para_html:
                new_html = para_html.replace(run_text, styled_run, 1)
            elif run_html in para_html:
                new_html = para_html.replace(run_html, styled_run, 1)
            elif run_html_escaped in para_html:
                new_html = para_html.replace(run_html_escaped, styled_run, 1)
            else:
                # 如果都不匹配，使用正则表达式
                escaped_text = re.escape(run_text)
                pattern = re.compile(escaped_text)
                match = pattern.search(para_html)
                if match:
                    new_html = para_html[:match.start()] + styled_run + para_html[match.end():]
            
            if new_html:
                para_html = new_html
                processed_runs.add(i)
                logger.debug(f"  应用内联样式: {run_text[:20]}... -> {styled_run[:50]}...")
        
        # 如果有修改，更新段落内容
        if processed_runs:
            try:
                soup = BeautifulSoup(para_html, 'html.parser')
                # 保留原始标签属性
                for attr_name, attr_value in paragraph_element.attrs.items():
                    for child in soup.find_all():
                        if child.name == paragraph_element.name:
                            child[attr_name] = attr_value
                            break
                
                # 清空原元素并添加新内容
                paragraph_element.clear()
                for content in soup.contents:
                    paragraph_element.append(content)
                    
            except Exception as e:
                logger.error(f"    内联样式应用失败: {e}")
                # 回退方案：直接设置HTML
                paragraph_element.clear()
                paragraph_element.append(para_html)
    
    def _get_most_common_style(self, runs_data: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        获取最常见的样式作为段落默认样式
        
        Args:
            runs_data: runs样式数据列表
            
        Returns:
            最常见的样式字典
        """
        if not runs_data:
            return {}
        
        # 统计每种样式的出现次数
        style_counts = {}
        for run in runs_data:
            # 创建样式的哈希键
            style_key = tuple(sorted([
                (k, str(v)) for k, v in run.items() 
                if k in ['bold', 'italic', 'underline', 'font_name', 'font_size', 'font_color', 'highlight_color'] and v is not None
            ]))
            style_counts[style_key] = style_counts.get(style_key, 0) + 1
        
        # 找到最常见的样式
        if style_counts:
            most_common_key = max(style_counts.keys(), key=lambda k: style_counts[k])
            # 找到对应的run数据
            for run in runs_data:
                style_key = tuple(sorted([
                    (k, str(v)) for k, v in run.items() 
                    if k in ['bold', 'italic', 'underline', 'font_name', 'font_size', 'font_color', 'highlight_color'] and v is not None
                ]))
                if style_key == most_common_key:
                    return run
        
        return runs_data[0] if runs_data else {}
    
    def _is_same_style(self, run_data: Dict[str, Any], default_style: Dict[str, Any]) -> bool:
        """
        检查run的样式是否与默认样式相同
        
        Args:
            run_data: run样式数据
            default_style: 默认样式数据
            
        Returns:
            如果样式相同返回True
        """
        style_keys = ['bold', 'italic', 'underline', 'font_name', 'font_size', 'font_color', 'highlight_color']
        
        for key in style_keys:
            if run_data.get(key) != default_style.get(key):
                return False
        
        return True
    
    def _create_styled_run(self, text: str, run_data: Dict[str, Any]) -> str:
        """
        创建带样式的内联文本
        
        Args:
            text: 文本内容
            run_data: run样式数据
            
        Returns:
            带样式的HTML字符串
        """
        css_styles = []
        html_tags = []
        
        # 加粗
        if run_data.get('bold'):
            html_tags.append('strong')
        
        # 斜体
        if run_data.get('italic'):
            html_tags.append('em')
        
        # 下划线
        if run_data.get('underline'):
            css_styles.append("text-decoration: underline")
        
        # 字体名称
        font_name = run_data.get('font_name')
        if font_name:
            css_styles.append(f"font-family: '{font_name}'")
        
        # 字体大小
        font_size = run_data.get('font_size')
        if font_size:
            # 确保字体大小是数字
            try:
                size_value = float(font_size)
                css_styles.append(f"font-size: {size_value}pt")
            except (ValueError, TypeError):
                pass
        
        # 字体颜色
        font_color = run_data.get('font_color')
        if font_color and font_color.startswith('#'):
            css_styles.append(f"color: {font_color}")
        
        # 高亮颜色
        highlight_color = run_data.get('highlight_color')
        if highlight_color and highlight_color.startswith('#'):
            css_styles.append(f"background-color: {highlight_color}")
        
        # 构建HTML
        result = text
        css_style = '; '.join(css_styles)
        
        # 先应用字体样式（使用span）
        if css_style:
            result = f'<span style="{css_style}">{result}</span>'
        
        # 然后应用语义标签
        for tag in reversed(html_tags):
            result = f'<{tag}>{result}</{tag}>'
        
        return result
    
    def _add_css_styles(self, element, css_styles: List[str]):
        """
        为元素添加CSS样式
        
        Args:
            element: BeautifulSoup元素
            css_styles: CSS样式列表
        """
        if not css_styles:
            return
        
        existing_style = element.get('style', '')
        if existing_style and not existing_style.endswith(';'):
            existing_style += ';'
        
        new_styles = '; '.join(css_styles)
        combined_style = f"{existing_style} {new_styles}".strip()
        
        element['style'] = combined_style
    
    def _style_name_to_class(self, style_name: Optional[str]) -> str:
        """
        将Word样式名转换为CSS类名
        
        Args:
            style_name: Word样式名
            
        Returns:
            CSS类名
        """
        if not style_name:
            return ""
        
        # 移除空格和特殊字符，转换为小写
        class_name = re.sub(r'[^\w]', '-', style_name.lower())
        return f"word-style-{class_name}"
    
    def _apply_styles_with_regex(self, html_content: str, paragraph_data: Dict[str, Any]) -> str:
        """
        使用正则表达式应用样式（BeautifulSoup不可用时的备用方案）
        
        Args:
            html_content: 原始HTML内容
            paragraph_data: 段落样式数据
            
        Returns:
            应用样式后的HTML
        """
        logger.info("使用正则表达式方式应用样式")
        
        # 简单的正则表达式实现
        result = html_content
        
        # 查找所有段落标签
        para_pattern = r'<(p|h[1-6])([^>]*)(?:\sid="([^"]+)")?([^>]*)>(.*?)</\1>'
        
        matches_count = 0
        styled_count = 0
        
        def replace_para(match):
            nonlocal matches_count, styled_count
            matches_count += 1
            
            tag = match.group(1)
            attrs_before = match.group(2) or ''
            para_id = match.group(3)
            attrs_after = match.group(4) or ''
            content = match.group(5)
            
            logger.debug(f"正则匹配: {tag}#{para_id}")
            
            if para_id and para_id in paragraph_data:
                styled_count += 1
                style_info = paragraph_data[para_id]
                logger.debug(f"  应用样式: {style_info}")
                
                # 构建样式
                css_styles = []
                
                # 对齐方式
                alignment = style_info.get('alignment')
                if alignment:
                    alignment_name = self.alignment_map.get(alignment, 'left')
                    css_styles.append(f"text-align: {alignment_name}")
                
                # 组合样式字符串
                if css_styles:
                    style_attr = f' style="{"; ".join(css_styles)}"'
                    logger.debug(f"  生成的样式属性: {style_attr}")
                else:
                    style_attr = ''
                
                # 重新构建标签
                return f'<{tag}{attrs_before}{style_attr}{attrs_after}>{content}</{tag}>'
            
            return match.group(0)
        
        # 应用替换
        result = re.sub(para_pattern, replace_para, result, flags=re.DOTALL)
        
        logger.info(f"正则方式处理完成: {styled_count}/{matches_count} 个段落应用样式")
        
        return result
    
    def generate_css_rules(self, paragraph_data: Dict[str, Any]) -> str:
        """
        生成CSS规则文件
        
        Args:
            paragraph_data: 段落样式数据
            
        Returns:
            CSS规则字符串
        """
        css_rules = []
        
        # 收集所有样式名称
        style_classes = set()
        for style_info in paragraph_data.values():
            style_name = style_info.get('style_name')
            if style_name:
                style_classes.add(self._style_name_to_class(style_name))
        
        # 生成CSS规则
        for class_name in sorted(style_classes):
            css_rules.append(f".{class_name} {{ margin: 0.5em 0; }}")
        
        # 对齐类
        alignment_classes = [
            ".align-center { text-align: center; }",
            ".align-right { text-align: right; }", 
            ".align-justify { text-align: justify; }",
            ".align-left { text-align: left; }"
        ]
        
        css_rules.extend(alignment_classes)
        
        return "\n".join(css_rules)


# 测试函数
def test_style_applicator():
    """测试样式应用器"""
    applicator = ParaIdStyleApplicator()
    
    if applicator.bs4_available:
        print("ParaIdStyleApplicator 已实现")
        print("功能:")
        print("- 基于paraId的精确样式匹配")
        print("- 段落对齐、缩进、间距")
        print("- 内联文本样式（加粗、斜体、颜色等）")
        print("- CSS类名生成和样式应用")
        print("- 正则表达式备用方案")
    else:
        print("BeautifulSoup不可用，功能受限")


if __name__ == "__main__":
    test_style_applicator()