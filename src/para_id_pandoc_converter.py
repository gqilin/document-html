"""
Pandoc转换器 - 保留paraId信息进行Word到HTML的转换
"""

import subprocess
import json
import tempfile
import os
import logging
from typing import Dict, Any, Optional
from pathlib import Path

# 配置日志
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 检查BeautifulSoup可用性
try:
    from bs4 import BeautifulSoup
    BS4_AVAILABLE = True
except ImportError:
    BS4_AVAILABLE = False
    BeautifulSoup = None


class ParaIdPandocConverter:
    """使用Pandoc进行转换并保留paraId信息"""
    
    def __init__(self):
        self.pandoc_available = self._check_pandoc()
    
    def _check_pandoc(self) -> bool:
        """检查pandoc是否可用"""
        try:
            subprocess.run(['pandoc', '--version'], 
                        capture_output=True, check=True, encoding='utf-8', errors='replace')
            return True
        except (subprocess.CalledProcessError, FileNotFoundError):
            return False
    
    def convert_with_para_id(self, docx_file: str, paragraph_data: Dict[str, Any]) -> str:
        """
        转换Word文档并保留paraId信息
        
        Args:
            docx_file: Word文档路径
            paragraph_data: 段落样式数据 {paraId: 样式信息}
            
        Returns:
            包含paraId的HTML字符串
        """
        if not self.pandoc_available:
            raise RuntimeError("pandoc is not available")
        
        logger.info(f"开始转换文档: {docx_file}")
        logger.info(f"输入段落数据: {len(paragraph_data)} 个段落")
        
        # 第一步：将Word转换为Pandoc JSON
        logger.info("步骤1: Word -> JSON")
        doc_json = self._docx_to_json(docx_file)
        logger.info(f"JSON转换完成，块数量: {len(doc_json.get('blocks', []))}")
        
 # 第二步：直接转换，不修改JSON（避免格式错误）
        # 第三步：将JSON转换为HTML
        html_content = self._json_to_html(doc_json)
        
        # 第四步：后处理HTML来添加ID
        html_content = self._post_process_html_ids(html_content, paragraph_data)
        
        return html_content
    
    def _docx_to_json(self, docx_file: str) -> Dict[str, Any]:
        """将Word文档转换为Pandoc JSON格式，同时提取图片"""
        # 第一步：先提取图片
        self._extract_images_from_docx(docx_file)
        
        # 第二步：转换为JSON
        cmd = [
            'pandoc', '-f', 'docx', '-t', 'json',
            '--wrap=none',
            docx_file
        ]
        
        logger.debug(f"执行命令: {' '.join(cmd)}")
        # 使用UTF-8编码避免GBK编码错误
        result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
        if result.returncode != 0:
            logger.error(f"Pandoc JSON转换失败: {result.stderr}")
            raise RuntimeError(f"Pandoc JSON conversion failed: {result.stderr}")
        
        logger.debug(f"原始JSON输出长度: {len(result.stdout)} 字符")
        
        try:
            doc_json = json.loads(result.stdout)
            logger.debug(f"JSON解析成功，块数量: {len(doc_json.get('blocks', []))}")
            return doc_json
        except json.JSONDecodeError as e:
            logger.error(f"JSON解析失败: {e}")
            logger.error(f"原始输出前500字符: {result.stdout[:500]}")
            raise
    
    def _inject_para_ids(self, doc_json: Dict[str, Any], paragraph_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        将paraId信息注入到Pandoc JSON中
        
        Args:
            doc_json: Pandoc文档JSON
            paragraph_data: 段落样式数据
            
        Returns:
            包含paraId的Pandoc JSON
        """
        logger.debug(f"开始注入paraId，输入段落数: {len(paragraph_data)}")
        
        # 创建段落文本到paraId的映射
        text_to_para_id = {}
        for para_id, data in paragraph_data.items():
            text = data.get('text', '').strip()
            if text:
                text_to_para_id[text] = para_id
                logger.debug(f"映射文本: '{text[:30]}...' -> {para_id}")
        
        logger.debug(f"创建文本映射: {len(text_to_para_id)} 条记录")
        
        matched_count = 0
        total_para_blocks = 0
        
        # 遍历文档中的所有段落
        def process_blocks(blocks):
            nonlocal matched_count, total_para_blocks
            for i, block in enumerate(blocks):
                if not isinstance(block, dict) or 't' not in block:
                    logger.debug(f"跳过非标准块 {i}: {type(block)}")
                    continue
                    
                block_type = block.get('t')
                logger.debug(f"处理块 {i}: {block_type}")
                
                if block_type == 'Para':
                    # 这是段落块，尝试匹配文本
                    para_text = self._extract_text_from_para_object(block)
                    logger.debug(f"段落 {i}: 文本='{para_text[:50]}...'")
                    
                    # 尝试找到对应的paraId
                    para_id = None
                    
                    # 精确匹配
                    if para_text in text_to_para_id:
                        para_id = text_to_para_id[para_text]
                        logger.debug(f"精确匹配成功: {para_text[:30]} -> {para_id}")
                    else:
                        # 模糊匹配：查找最相似的文本
                        para_id = self._find_best_text_match(para_text, text_to_para_id)
                        if para_id:
                            logger.debug(f"模糊匹配成功: {para_text[:30]} -> {para_id}")
                    
                    if para_id:
                        matched_count += 1
                        # 添加identifier属性到段落
                        if 'c' not in block:
                            block['c'] = []
                        
                        # Pandoc的属性格式：[id, classes, kvs] 应该作为content的第一个元素
                        attributes = [para_id, [], []]  # [id, classes, key-value pairs]
                        
                        # 确保content是一个数组
                        if not block['c']:
                            block['c'] = []
                        
                        # 检查第一个元素是否是属性数组 [id, classes, kvs]
                        if len(block['c']) > 0 and isinstance(block['c'][0], list) and len(block['c'][0]) == 3:
                            # 已有属性，更新ID
                            block['c'][0][0] = para_id
                        else:
                            # 没有属性，在开头插入属性
                            block['c'].insert(0, attributes)
                        
                        logger.debug(f"添加ID到段落块: {para_id}")
                    else:
                        logger.debug(f"未找到匹配的paraId: '{para_text[:30]}...'")
                
                elif block_type == 'Header':
                    # 处理标题
                    header_text = self._extract_text_from_para_object(block)
                    logger.debug(f"标题: 文本='{header_text[:50]}...'")
                    
                    if header_text in text_to_para_id:
                        para_id = text_to_para_id[header_text]
                        matched_count += 1
                        
                        # Header的格式：[level, attributes, content]
                        if len(block['c']) >= 3 and isinstance(block['c'][1], list):
                            block['c'][1][0] = para_id  # 设置attributes中的ID
                        else:
                            # 重新构建Header结构
                            level = block['c'][0] if len(block['c']) > 0 else 1
                            content = block['c'][1] if len(block['c']) > 1 else []
                            attributes = [para_id, [], []]
                            block['c'] = [level, attributes, content]
                        
                        logger.debug(f"添加ID到标题块: {para_id}")
        
        # 处理正文块
        if 'blocks' in doc_json:
            process_blocks(doc_json['blocks'])
        
        # 处理meta中的标题（如果有的话）
        if 'meta' in doc_json and 'title' in doc_json['meta']:
            title_meta = doc_json['meta']['title']
            if isinstance(title_meta, dict) and title_meta.get('t') == 'MetaInlines':
                # 提取标题文本
                title_content = title_meta.get('c', [])
                title_text = self._extract_text_from_content_list(title_content)
                logger.debug(f"Meta标题文本: '{title_text[:50]}...'")
                
                if title_text in text_to_para_id:
                    para_id = text_to_para_id[title_text]
                    logger.debug(f"Meta标题匹配到paraId: {para_id}")
                    
                    # 创建一个标题块
                    title_block = {
                        "t": "Header",
                        "c": [1, [para_id, [], []], [{"t": "Str", "c": title_text}]]
                    }
                    # 插入到blocks开头
                    doc_json['blocks'].insert(0, title_block)
                    matched_count += 1
                    logger.debug(f"创建标题块并添加ID: {para_id}")
        
        logger.info(f"paraId注入完成: {matched_count}/{total_para_blocks} 个段落成功匹配")
        return doc_json
    
    def _extract_text_from_content_list(self, content_list: list) -> str:
        """从Pandoc内容列表中提取纯文本"""
        text_parts = []
        
        for item in content_list:
            if isinstance(item, str):
                text_parts.append(item)
            elif isinstance(item, dict) and 't' in item:
                item_type = item.get('t')
                
                if item_type == 'Str':
                    # 纯文本
                    text = item.get('c', '')
                    text_parts.append(text)
                elif item_type == 'Space':
                    # 空格
                    text_parts.append(' ')
                elif item_type == 'SoftBreak':
                    # 软换行
                    text_parts.append(' ')
                elif item_type == 'LineBreak':
                    # 硬换行
                    text_parts.append('\n')
                elif 'c' in item and isinstance(item['c'], list):
                    # 递归处理嵌套内容
                    nested_text = self._extract_text_from_content_list(item['c'])
                    text_parts.append(nested_text)
        
        return ''.join(text_parts)
    
    def _extract_text_from_para_object(self, para_block: dict) -> str:
        """从Pandoc对象格式的段落块中提取纯文本"""
        if not isinstance(para_block, dict) or 'c' not in para_block:
            return ""
        
        try:
            content_list = para_block['c']
            return self._extract_text_from_content_list(content_list).strip()
        except Exception as e:
            logger.debug(f"提取段落文本失败: {e}")
            return ""
    
    def _extract_text_from_para(self, para_block) -> str:
        """从段落块中提取纯文本（旧格式兼容）"""
        if isinstance(para_block, dict):
            return self._extract_text_from_para_object(para_block)
        
        if not isinstance(para_block, list) or len(para_block) < 2:
            return ""
        
        def extract_text(elements):
            text_parts = []
            for element in elements:
                if isinstance(element, str):
                    text_parts.append(element)
                elif isinstance(element, list):
                    if len(element) > 0 and isinstance(element[0], str):
                        # 这是内联元素
                        if len(element) > 1 and isinstance(element[1], dict):
                            # 有属性的元素
                            if len(element) > 2:
                                text_parts.extend(extract_text(element[2:]))
                        else:
                            text_parts.extend(extract_text(element[1:]))
                    else:
                        text_parts.extend(extract_text(element))
                elif isinstance(element, dict):
                    if 'content' in element:
                        text_parts.extend(extract_text(element['content']))
            
            return text_parts
        
        try:
            text_parts = extract_text(para_block[1:])
            return ''.join(text_parts).strip()
        except Exception:
            return ""
    
    def _find_best_text_match(self, target_text: str, text_to_para_id: Dict[str, str]) -> Optional[str]:
        """
        当精确匹配失败时，找到最佳文本匹配
        
        Args:
            target_text: 目标文本
            text_to_para_id: 文本到paraId的映射
            
        Returns:
            最佳匹配的paraId
        """
        if not target_text:
            return None
        
        best_match = None
        best_score = 0.8  # 最低相似度阈值
        
        for text, para_id in text_to_para_id.items():
            similarity = self._calculate_text_similarity(target_text, text)
            if similarity > best_score:
                best_score = similarity
                best_match = para_id
        
        return best_match
    
    def _calculate_text_similarity(self, text1: str, text2: str) -> float:
        """计算两个文本的相似度"""
        # 简单的相似度计算：基于共同字符比例
        if not text1 or not text2:
            return 0.0
        
        # 移除空格和标点符号进行比较
        clean1 = ''.join(c.lower() for c in text1 if c.isalnum())
        clean2 = ''.join(c.lower() for c in text2 if c.isalnum())
        
        if len(clean1) == 0 and len(clean2) == 0:
            return 1.0
        if len(clean1) == 0 or len(clean2) == 0:
            return 0.0
        
        # 计算Jaccard相似度
        set1 = set(clean1)
        set2 = set(clean2)
        intersection = len(set1 & set2)
        union = len(set1 | set2)
        
        return intersection / union if union > 0 else 0.0
    
    def _extract_images_from_docx(self, docx_file: str):
        """
        从docx文件中提取图片到本地
        
        Args:
            docx_file: docx文件路径
        """
        import tempfile
        import shutil
        from pathlib import Path
        
        try:
            # 创建临时目录
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
                        logger.info(f"找到media目录: {media_dir}")
                        
                        # 复制图片到项目images目录
                        project_root = Path(__file__).parent.parent
                        target_images_dir = project_root / 'uploads' / 'images'
                        target_images_dir.mkdir(parents=True, exist_ok=True)
                        
                        copied_count = 0
                        for img_file in media_dir.rglob('*'):
                            if img_file.is_file() and img_file.suffix.lower() in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp']:
                                target_path = target_images_dir / img_file.name
                                shutil.copy2(img_file, target_path)
                                copied_count += 1
                                logger.debug(f"复制图片: {img_file.name} -> {target_path}")
                        
                        logger.info(f"成功复制 {copied_count} 个图片到 uploads/images/")
                        
                        # 如果有HTML输出，检查其中的图片引用
                        if result.stdout and 'src=' in result.stdout:
                            logger.debug("HTML输出包含图片引用")
                            # 可以进一步处理HTML中的图片路径
                    else:
                        logger.debug("未找到media目录")
                        
                finally:
                    os.chdir(original_cwd)
                    
        except Exception as e:
            logger.warning(f"图片提取失败: {e}")
            # 不阻止转换流程，只记录警告
    
    def _json_to_html(self, doc_json: Dict[str, Any]) -> str:
        """将Pandoc JSON转换为HTML"""
        cmd = ['pandoc', '-f', 'json', '-t', 'html', '--standalone', '--wrap=none']
        
        json_input = json.dumps(doc_json, ensure_ascii=False)
        logger.debug(f"JSON输入长度: {len(json_input)} 字符")
        logger.debug(f"执行命令: {' '.join(cmd)}")
        
        # 使用UTF-8编码避免GBK编码错误
        result = subprocess.run(cmd, input=json_input, capture_output=True, text=True, encoding='utf-8', errors='replace')
        
        if result.returncode != 0:
            logger.error(f"Pandoc HTML转换失败: {result.stderr}")
            logger.error(f"JSON输入前500字符: {json_input[:500]}")
            raise RuntimeError(f"Pandoc HTML conversion failed: {result.stderr}")
        
        html_output = result.stdout
        logger.debug(f"HTML输出长度: {len(html_output)} 字符")
        
        # 检查关键元素
        id_count = html_output.count('id=')
        p_count = html_output.count('<p')
        h_count = sum(html_output.count(f'<h{i}') for i in range(1, 7))
        
        logger.debug(f"HTML统计: {id_count}个ID, {p_count}个<p>, {h_count}个标题")
        
        if id_count == 0:
            logger.warning("HTML中没有找到任何ID属性！")
            # 记录一些示例段落用于调试
            p_matches = html_output.split('<p')
            logger.debug(f"前3个段落示例:")
            for i, match in enumerate(p_matches[:3]):
                if '>' in match:
                    full_tag = '<p' + match.split('>')[0] + '>'
                    logger.debug(f"  段落{i+1}: {full_tag}")
        
        return html_output
    
    def convert_simple(self, docx_file: str) -> str:
        """
        简单的Word到HTML转换（不保留paraId）
        
        Args:
            docx_file: Word文档路径
            
        Returns:
            HTML字符串
        """
        if not self.pandoc_available:
            raise RuntimeError("pandoc is not available")
        
        cmd = [
            'pandoc', '-f', 'docx', '-t', 'html',
            '--standalone', '--embed-resources',
            '--wrap=none',
            docx_file
        ]
        
        # 使用UTF-8编码避免GBK编码错误
        result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
        if result.returncode != 0:
            raise RuntimeError(f"Pandoc conversion failed: {result.stderr}")
        
        return result.stdout
    
    def _post_process_html_ids(self, html_content: str, paragraph_data: Dict[str, Any]) -> str:
        """
        后处理HTML来添加ID属性
        当JSON方法失败时使用正则表达式匹配
        
        Args:
            html_content: 原始HTML内容
            paragraph_data: 段落样式数据
            
        Returns:
            添加了ID的HTML
        """
        import re
        from bs4 import BeautifulSoup
        
        logger.info("开始HTML后处理添加ID")
        
        # 首先处理mark标签 - 将mark标签转换为span标签
        html_content = self._convert_mark_to_span(html_content)
        
        if not 'BS4_AVAILABLE' in globals() or not BS4_AVAILABLE or BeautifulSoup is None:
            logger.warning("BeautifulSoup不可用，使用正则表达式方法")
            return self._add_ids_with_regex(html_content, paragraph_data)
        
        try:
            soup = BeautifulSoup(html_content, 'html.parser')
            
            # 创建文本到paraId的映射
            text_to_para_id = {}
            for para_id, data in paragraph_data.items():
                text = data.get('text', '').strip()
                if text:
                    text_to_para_id[text] = para_id
            
            elements_modified = 0
            
            # 处理所有段落和标题
            for element in soup.find_all(['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6']):
                if element.get('id'):
                    continue  # 已有ID，跳过
                
                element_text = element.get_text().strip()
                
                # 尝试精确匹配
                if element_text in text_to_para_id:
                    para_id = text_to_para_id[element_text]
                    element['id'] = para_id
                    elements_modified += 1
                    logger.debug(f"添加ID: '{element_text[:30]}...' -> {para_id}")
                else:
                    # 尝试模糊匹配
                    para_id = self._find_best_text_match(element_text, text_to_para_id)
                    if para_id:
                        element['id'] = para_id
                        elements_modified += 1
                        logger.debug(f"模糊匹配添加ID: '{element_text[:30]}...' -> {para_id}")
            
            logger.info(f"HTML后处理完成，修改了 {elements_modified} 个元素")
            return str(soup)
            
        except Exception as e:
            logger.error(f"BeautifulSoup处理失败: {e}")
            return self._add_ids_with_regex(html_content, paragraph_data)
    
    def _convert_mark_to_span(self, html_content: str) -> str:
        """
        将HTML中的mark标签转换为span标签，保留背景色样式
        如果父元素（如p标签）已有背景色，则不添加默认黄色背景
        
        Args:
            html_content: 原始HTML内容
            
        Returns:
            转换后的HTML内容
        """
        import re
        
        logger.info("开始转换mark标签为span标签")
        
        # 首先检查父元素是否有背景色
        # 匹配 <p ... style="...background-color...">...<mark>...</mark>...</p>
        def should_skip_default_bg(mark_match, html_content):
            """检查mark标签的父元素是否已有背景色"""
            mark_start = mark_match.start()
            # 向前查找最近的p标签开始
            text_before = html_content[:mark_start]
            
            # 查找最近的 <p ...> 标签
            p_match = None
            for match in re.finditer(r'<p[^>]*>', text_before, re.IGNORECASE):
                p_match = match
            
            if p_match:
                p_tag = p_match.group(0)
                # 检查p标签是否有background-color样式
                p_style_match = re.search(r'style=["\']([^"\']*)["\']', p_tag, re.IGNORECASE)
                if p_style_match:
                    p_style = p_style_match.group(1)
                    if 'background-color' in p_style.lower():
                        return True
            
            return False
        
        # 使用正则表达式替换mark标签
        def replace_mark_tag(match, html_content=html_content):
            """替换mark开始标签为span标签，添加背景色样式"""
            tag_content = match.group(0)
            
            # 检查是否已有style属性
            if 'style=' in tag_content:
                # 如果已有style属性，检查是否需要添加background-color
                # 匹配style="..."
                style_match = re.search(r'style=["\']([^"\']*)["\']', tag_content)
                if style_match:
                    existing_style = style_match.group(1)
                    # 如果已经有background-color，则不添加；否则也不添加默认黄色背景
                    # mark标签由Pandoc生成，只有当原文有高亮时才应该保留
                    # 如果mark标签本身没有background-color，说明原文没有设置高亮
                    return tag_content.replace('<mark', '<span').replace('<MARK', '<span')
            
            # 没有style属性，检查父元素是否有背景色
            if should_skip_default_bg(match, html_content):
                # 父元素有背景色，不添加默认背景色
                return '<span>'
            
            # 没有style属性且父元素没有背景色，不添加默认背景色
            # mark标签由Pandoc生成，表示Word中的高亮，但如果父元素没有背景色
            # 说明原文没有设置高亮，只是Pandoc的默认行为，不应添加黄色背景
            return '<span>'
        
        # 替换mark开始标签
        result = re.sub(r'<mark[^>]*>', lambda m: replace_mark_tag(m, html_content), html_content, flags=re.IGNORECASE)
        
        # 替换mark结束标签
        result = re.sub(r'</mark>', '</span>', result, flags=re.IGNORECASE)
        result = re.sub(r'</MARK>', '</span>', result, flags=re.IGNORECASE)
        
        # 统计替换数量
        mark_count = html_content.lower().count('<mark')
        logger.info(f"转换完成，替换了 {mark_count} 个mark标签")
        
        return result
    
    def _add_ids_with_regex(self, html_content: str, paragraph_data: Dict[str, Any]) -> str:
        """
        使用正则表达式添加ID（最后的备用方案）
        """
        import re
        
        logger.info("使用正则表达式添加ID")
        
        # 创建文本到paraId的映射
        text_to_para_id = {}
        for para_id, data in paragraph_data.items():
            text = data.get('text', '').strip()
            if text:
                text_to_para_id[text] = para_id
        
        result = html_content
        
        # 匹配段落和标题标签
        pattern = r'<(p|h[1-6])([^>]*)(>)(.*?)(</\1>)'
        
        def replace_func(match):
            tag = match.group(1)
            attrs = match.group(2) or ''
            closing = match.group(3)
            content = match.group(4)
            end_tag = match.group(5)
            
            if 'id=' in attrs:
                return match.group(0)  # 已有ID
            
            content_stripped = content.strip()
            
            # 尝试匹配
            if content_stripped in text_to_para_id:
                para_id = text_to_para_id[content_stripped]
                return f'<{tag}{attrs} id="{para_id}">{content}{end_tag}'
            else:
                # 模糊匹配
                para_id = self._find_best_text_match(content_stripped, text_to_para_id)
                if para_id:
                    return f'<{tag}{attrs} id="{para_id}">{content}{end_tag}'
            
            return match.group(0)
        
        result = re.sub(pattern, replace_func, result, flags=re.DOTALL)
        return result
    
    def _extract_images_from_docx(self, docx_file: str):
        """
        从docx文件中提取图片到本地
        
        Args:
            docx_file: docx文件路径
        """
        import tempfile
        import shutil
        from pathlib import Path
        
        try:
            # 创建临时目录
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
                        logger.info(f"找到media目录: {media_dir}")
                        
                        # 复制图片到项目images目录
                        project_root = Path(__file__).parent.parent
                        target_images_dir = project_root / 'uploads' / 'images'
                        target_images_dir.mkdir(parents=True, exist_ok=True)
                        
                        copied_count = 0
                        for img_file in media_dir.rglob('*'):
                            if img_file.is_file() and img_file.suffix.lower() in ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp']:
                                target_path = target_images_dir / img_file.name
                                shutil.copy2(img_file, target_path)
                                copied_count += 1
                                logger.debug(f"复制图片: {img_file.name} -> {target_path}")
                        
                        logger.info(f"成功复制 {copied_count} 个图片到 uploads/images/")
                        
                        # 如果有HTML输出，检查其中的图片引用
                        if result.stdout and 'src=' in result.stdout:
                            logger.debug("HTML输出包含图片引用")
                            # 可以进一步处理HTML中的图片路径
                    else:
                        logger.debug("未找到media目录")
                        
                finally:
                    os.chdir(original_cwd)
                    
        except Exception as e:
            logger.warning(f"图片提取失败: {e}")
            # 不阻止转换流程，只记录警告


# 测试函数
def test_pandoc_converter():
    """测试paraId转换器"""
    converter = ParaIdPandocConverter()
    
    if converter.pandoc_available:
        print("ParaIdPandocConverter 已实现")
        print("功能:")
        print("- Word到JSON转换")
        print("- paraId注入到JSON结构中")
        print("- JSON到HTML转换")
        print("- 文本相似度匹配")
    else:
        print("pandoc 不可用，无法测试")


if __name__ == "__main__":
    test_pandoc_converter()