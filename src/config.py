"""
文档转换配置文件
定义不同格式转换时要保留的样式配置
"""

from typing import Dict, List, Set


class ConversionConfig:
    """文档转换配置类"""
    
    # 默认样式保留配置
    DEFAULT_STYLE_CONFIG = {
        'doc': {
            'allowed_styles': [
                'font-size', 'color', 'font-weight', 'font-style', 
                'text-decoration', 'text-align', 'font-family'
            ],
            'allowed_classes': [
                'center', 'right', 'left', 'justify', 'bold', 'italic', 'underline'
            ]
        },
        'epub': {
            'allowed_styles': [
                'font-size', 'color', 'font-weight', 'font-style', 
                'text-decoration', 'text-align', 'font-family', 
                'line-height', 'margin', 'padding'
            ],
            'allowed_classes': [
                'center', 'right', 'left', 'justify', 'bold', 'italic', 'underline',
                'chapter', 'title', 'subtitle', 'paragraph', 'quote'
            ]
        },
        'pdf': {
            'allowed_styles': [
                'font-size', 'color', 'font-weight', 'font-style', 
                'text-decoration', 'text-align', 'font-family'
            ],
            'allowed_classes': [
                'center', 'right', 'left', 'justify', 'bold', 'italic', 'underline'
            ]
        }
    }
    
    def __init__(self, custom_config: Dict = None):
        """
        初始化配置
        
        Args:
            custom_config: 自定义配置，会与默认配置合并
        """
        self.config = self.DEFAULT_STYLE_CONFIG.copy()
        
        if custom_config:
            self._merge_config(custom_config)
    
    def _merge_config(self, custom_config: Dict):
        """合并自定义配置"""
        for format_type, format_config in custom_config.items():
            if format_type in self.config:
                if 'allowed_styles' in format_config:
                    self.config[format_type]['allowed_styles'].extend(
                        [s for s in format_config['allowed_styles'] 
                         if s not in self.config[format_type]['allowed_styles']]
                    )
                if 'allowed_classes' in format_config:
                    self.config[format_type]['allowed_classes'].extend(
                        [c for c in format_config['allowed_classes'] 
                         if c not in self.config[format_type]['allowed_classes']]
                    )
            else:
                self.config[format_type] = format_config
    
    def get_allowed_styles(self, format_type: str) -> Set[str]:
        """获取指定格式允许的样式属性"""
        format_config = self.config.get(format_type, {})
        return set(format_config.get('allowed_styles', []))
    
    def get_allowed_classes(self, format_type: str) -> Set[str]:
        """获取指定格式允许的CSS类名"""
        format_config = self.config.get(format_type, {})
        return set(format_config.get('allowed_classes', []))
    
    def filter_styles(self, format_type: str, styles: Dict[str, str]) -> Dict[str, str]:
        """
        过滤样式属性，只保留允许的样式
        
        Args:
            format_type: 文档格式 (doc, epub, pdf)
            styles: 原始样式字典
            
        Returns:
            过滤后的样式字典
        """
        allowed_styles = self.get_allowed_styles(format_type)
        return {k: v for k, v in styles.items() if k in allowed_styles}
    
    def filter_classes(self, format_type: str, classes: List[str]) -> List[str]:
        """
        过滤CSS类名，只保留允许的类
        
        Args:
            format_type: 文档格式 (doc, epub, pdf)
            classes: 原始类名列表
            
        Returns:
            过滤后的类名列表
        """
        allowed_classes = self.get_allowed_classes(format_type)
        return [c for c in classes if c in allowed_classes]
    
    def is_style_allowed(self, format_type: str, style_name: str) -> bool:
        """检查样式属性是否被允许"""
        return style_name in self.get_allowed_styles(format_type)
    
    def is_class_allowed(self, format_type: str, class_name: str) -> bool:
        """检查CSS类名是否被允许"""
        return class_name in self.get_allowed_classes(format_type)


# 全局默认配置实例
DEFAULT_CONFIG = ConversionConfig()


def create_config_from_request(request_data: Dict) -> ConversionConfig:
    """
    从请求数据创建配置实例
    
    Args:
        request_data: 包含配置信息的请求数据
        
    Returns:
        配置实例
    """
    custom_config = {}
    
    # 解析样式配置
    if 'style_config' in request_data:
        style_config = request_data['style_config']
        
        for format_type in ['doc', 'epub', 'pdf']:
            if format_type in style_config:
                format_config = {}
                
                if 'allowed_styles' in style_config[format_type]:
                    format_config['allowed_styles'] = style_config[format_type]['allowed_styles']
                
                if 'allowed_classes' in style_config[format_type]:
                    format_config['allowed_classes'] = style_config[format_type]['allowed_classes']
                
                if format_config:
                    custom_config[format_type] = format_config
    
    return ConversionConfig(custom_config)