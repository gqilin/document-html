import os
import hashlib
import requests
import urllib.parse
import shutil
from pathlib import Path
from typing import Optional, Dict, Tuple
import base64
import re


class ImageDownloader:
    """图片下载和本地化管理器"""
    
    def __init__(self, images_dir: str = "images"):
        """
        初始化图片下载器
        
        Args:
            images_dir: 图片存储目录
        """
        self.images_dir = Path(images_dir)
        self.images_dir.mkdir(exist_ok=True)
        self.downloaded_images = {}  # 缓存已下载的图片
    
    def _generate_filename(self, url: str, content: bytes = None) -> str:
        """
        生成唯一的文件名
        
        Args:
            url: 图片URL或base64数据
            content: 图片内容（用于生成hash）
            
        Returns:
            生成的文件名
        """
        if content:
            # 使用内容hash生成文件名
            file_hash = hashlib.md5(content).hexdigest()[:16]
        else:
            # 使用URL生成hash
            file_hash = hashlib.md5(url.encode()).hexdigest()[:16]
        
        # 尝试从URL中获取文件扩展名
        ext = ""
        if not url.startswith('data:'):
            parsed = urllib.parse.urlparse(url)
            path = parsed.path.lower()
            if path.endswith(('.jpg', '.jpeg')):
                ext = '.jpg'
            elif path.endswith('.png'):
                ext = '.png'
            elif path.endswith('.gif'):
                ext = '.gif'
            elif path.endswith('.webp'):
                ext = '.webp'
            elif path.endswith('.bmp'):
                ext = '.bmp'
        else:
            # 从base64数据中提取MIME类型
            if ';' in url:
                mime_type = url.split(';')[0].split(':')[1]
                ext_map = {
                    'image/jpeg': '.jpg',
                    'image/png': '.png',
                    'image/gif': '.gif',
                    'image/webp': '.webp',
                    'image/bmp': '.bmp'
                }
                ext = ext_map.get(mime_type, '.jpg')
        
        return f"{file_hash}{ext}" if ext else f"{file_hash}.jpg"
    
    def download_from_url(self, url: str) -> Optional[str]:
        """
        从URL下载图片
        
        Args:
            url: 图片URL
            
        Returns:
            本地文件路径，失败返回None
        """
        if url in self.downloaded_images:
            return self.downloaded_images[url]
        
        try:
            # 跳过无效URL
            if not url or url.startswith('#') or url.startswith('mailto:'):
                return None
            
            # 处理相对URL（这里假设是绝对URL）
            if not url.startswith(('http://', 'https://')):
                return None
            
            # 下载图片
            response = requests.get(url, timeout=30, stream=True)
            response.raise_for_status()
            
            content = response.content
            if not content or len(content) < 100:  # 太小的文件可能不是有效图片
                return None
            
            # 生成文件名并保存
            filename = self._generate_filename(url, content)
            file_path = self.images_dir / filename
            
            with open(file_path, 'wb') as f:
                f.write(content)
            
            # 缓存结果
            local_path = str(file_path)
            self.downloaded_images[url] = local_path
            
            return local_path
            
        except Exception:
            return None
    
    def save_from_base64(self, data_url: str, original_name: str = None) -> Optional[str]:
        """
        保存base64编码的图片
        
        Args:
            data_url: data:URL格式的base64数据
            original_name: 原始文件名（可选）
            
        Returns:
            本地文件路径，失败返回None
        """
        if data_url in self.downloaded_images:
            return self.downloaded_images[data_url]
        
        try:
            if not data_url.startswith('data:image/'):
                return None
            
            # 提取base64数据
            if ',' in data_url:
                header, base64_data = data_url.split(',', 1)
            else:
                return None
            
            # 解码base64
            content = base64.b64decode(base64_data)
            if not content:
                return None
            
            # 生成文件名
            if original_name:
                filename = self._generate_filename(original_name, content)
            else:
                filename = self._generate_filename(data_url, content)
            
            file_path = self.images_dir / filename
            
            with open(file_path, 'wb') as f:
                f.write(content)
            
            # 缓存结果
            local_path = str(file_path)
            self.downloaded_images[data_url] = local_path
            
            return local_path
            
        except Exception:
            return None
    
    def save_from_bytes(self, content: bytes, filename: str = None) -> Optional[str]:
        """
        保存字节数据图片
        
        Args:
            content: 图片字节数据
            filename: 建议的文件名（可选）
            
        Returns:
            本地文件路径，失败返回None
        """
        try:
            if not content:
                return None
            
            # 生成文件名
            if filename:
                final_filename = self._generate_filename(filename, content)
            else:
                final_filename = self._generate_filename("bytes", content)
            
            file_path = self.images_dir / final_filename
            
            with open(file_path, 'wb') as f:
                f.write(content)
            
            return str(file_path)
            
        except Exception:
            return None
    
    def handle_pandoc_image(self, img_path: str) -> Optional[str]:
        """
        处理Pandoc生成的本地图片路径
        
        Args:
            img_path: Pandoc生成的图片路径（如 media/image1.png）
            
        Returns:
            本地文件路径，如果图片不存在返回None
        """
        try:
            # 如果图片已经被提取到uploads/images目录，直接返回
            if img_path.startswith('media/'):
                filename = img_path.replace('media/', '')
                target_path = self.images_dir / filename
                
                if target_path.exists():
                    print(f"找到已提取的图片: {target_path}")
                    return str(target_path)
                else:
                    # 尝试在当前目录下查找
                    current_dir = Path('.')
                    potential_path = current_dir / img_path
                    if potential_path.exists():
                        # 复制到images目录
                        shutil.copy2(potential_path, target_path)
                        print(f"复制图片到images目录: {potential_path} -> {target_path}")
                        return str(target_path)
            
            # 其他情况返回None
            return None
            
        except Exception as e:
            print(f"处理Pandoc图片时出错: {e}")
            return None
    
    def get_relative_path(self, local_path: str, base_dir: str = None) -> str:
        """
        获取相对路径，使用../uploads/images前缀
        
        Args:
            local_path: 本地绝对路径
            base_dir: 基础目录，默认为images_dir
            
        Returns:
            相对路径，以../uploads/images开头
        """
        if not local_path:
            return ""
        
        try:
            # 获取文件名
            path = Path(local_path)
            filename = path.name
            
            # 返回以../uploads/images开头的路径
            return f"../uploads/images/{filename}"
            
        except Exception:
            return local_path
    
    def clean_old_files(self, max_files: int = 1000):
        """
        清理旧文件，保持文件数量在限制范围内
        
        Args:
            max_files: 最大文件数量
        """
        try:
            files = []
            for file_path in self.images_dir.iterdir():
                if file_path.is_file():
                    files.append((file_path.stat().st_mtime, file_path))
            
            # 按修改时间排序
            files.sort()
            
            # 删除最旧的文件
            if len(files) > max_files:
                for _, file_path in files[:-max_files]:
                    try:
                        file_path.unlink()
                    except Exception:
                        pass
                        
        except Exception:
            pass
    
    def get_stats(self) -> Dict:
        """
        获取统计信息
        
        Returns:
            统计信息字典
        """
        try:
            files = list(self.images_dir.iterdir())
            total_size = sum(f.stat().st_size for f in files if f.is_file())
            
            return {
                'files_count': len(files),
                'total_size_bytes': total_size,
                'total_size_mb': round(total_size / (1024 * 1024), 2),
                'cache_entries': len(self.downloaded_images)
            }
        except Exception:
            return {
                'files_count': 0,
                'total_size_bytes': 0,
                'total_size_mb': 0,
                'cache_entries': 0
            }


def process_html_images(html_content: str, downloader: ImageDownloader, base_path: str = None) -> str:
    """
    处理HTML中的图片标签，将网络图片下载到本地并更新路径
    现在也处理Pandoc生成的本地图片路径
    
    Args:
        html_content: HTML内容
        downloader: 图片下载器实例
        base_path: 基础路径，用于生成相对路径
        
    Returns:
        处理后的HTML内容
    """
    # 匹配img标签
    img_pattern = r'<img[^>]+src=["\']([^"\']+)["\'][^>]*>'
    images = re.findall(img_pattern, html_content, re.IGNORECASE)
    
    processed_html = html_content
    
    for img_src in images:
        local_path = None
        
        # 处理网络图片
        if img_src.startswith(('http://', 'https://')):
            local_path = downloader.download_from_url(img_src)
        
        # 处理base64图片
        elif img_src.startswith('data:image/'):
            local_path = downloader.save_from_base64(img_src)
        
        # 处理Pandoc生成的本地图片路径 (如 media/image1.png)
        elif img_src.startswith('media/') or (not img_src.startswith(('http://', 'https://', 'data:', '#', 'mailto:'))):
            local_path = downloader.handle_pandoc_image(img_src)
        
        # 更新HTML中的图片路径
        if local_path:
            # 生成相对路径，使用../uploads/images前缀
            relative_path = downloader.get_relative_path(local_path)
            
            # 替换src属性
            old_img_tag = re.search(f'<img[^>]+src=["\']{re.escape(img_src)}["\'][^>]*>', processed_html, re.IGNORECASE)
            if old_img_tag:
                new_img_tag = old_img_tag.group(0).replace(img_src, relative_path)
                processed_html = processed_html.replace(old_img_tag.group(0), new_img_tag)
    
    return processed_html