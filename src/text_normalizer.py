# -*- coding: utf-8 -*-
"""
PDF文本字符标准化模块
结合Python内置NFKC和自定义映射，处理PDF提取后的全角字符和特殊符号
"""

import unicodedata
import re
from typing import Dict, List, Tuple, Optional


class TextNormalizer:
    """
    文本标准化器
    
    处理流程：
    1. Unicode NFKC标准化（转换大部分全角字符）
    2. 自定义字符映射（处理NFKC无法转换的特殊字符）
    3. 标点符号规范化
    """
    
    # 自定义字符映射表 - PDF中常见的特殊符号
    CUSTOM_CHAR_MAPPINGS: Dict[str, str] = {
        # 特殊点符号（PDF中常用来替代小数点）
        '\U00010236': '.',      # 𐆶 -> .
        '\u2024': '.',          # ․ -> . (One Dot Leader)
        '\u2025': '..',         # ‥ -> .. (Two Dot Leader)
        
        # 特殊连字符和破折号
        '\u2043': '-',          # ⁃ -> - (Hyphen Bullet)
        '\u2010': '-',          # ‐ -> - (Hyphen)
        '\u2011': '-',          # ‑ -> - (Non-Breaking Hyphen)
        '\u2012': '-',          # ‒ -> - (Figure Dash)
        '\u2013': '-',          # – -> - (En Dash)
        '\u2014': '--',         # — -> -- (Em Dash)
        '\u2015': '--',         # ― -> -- (Horizontal Bar)
        
        # 特殊省略号
        '\u2026': '...',        # … -> ... (Horizontal Ellipsis)
        '\u22ef': '...',        # ⋯ -> ... (Midline Horizontal Ellipsis)
        
        # 特殊空格
        '\u00a0': ' ',          #   ->   (Non-Breaking Space)
        '\u2000': ' ',          #   ->   (En Quad)
        '\u2001': ' ',          #   ->   (Em Quad)
        '\u2002': ' ',          #   ->   (En Space)
        '\u2003': ' ',          #   ->   (Em Space)
        '\u2004': ' ',          #   ->   (Three-Per-Em Space)
        '\u2005': ' ',          #   ->   (Four-Per-Em Space)
        '\u2006': ' ',          #   ->   (Six-Per-Em Space)
        '\u2007': ' ',          #   ->   (Figure Space)
        '\u2008': ' ',          #   ->   (Punctuation Space)
        '\u2009': ' ',          #   ->   (Thin Space)
        '\u200a': ' ',          #   ->   (Hair Space)
        '\u202f': ' ',          #   ->   (Narrow No-Break Space)
        '\u205f': ' ',          #   ->   (Medium Mathematical Space)
        '\u3000': ' ',          #   ->   (Ideographic Space)
        
        # 特殊引号
        '\u2018': "'",          # ' -> ' (Left Single Quotation Mark)
        '\u2019': "'",          # ' -> ' (Right Single Quotation Mark)
        '\u201a': "'",          # ‚ -> ' (Single Low-9 Quotation Mark)
        '\u201b': "'",          # ‛ -> ' (Single High-Reversed-9 Quotation Mark)
        '\u201c': '"',          # " -> " (Left Double Quotation Mark)
        '\u201d': '"',          # " -> " (Right Double Quotation Mark)
        '\u201e': '"',          # „ -> " (Double Low-9 Quotation Mark)
        '\u201f': '"',          # ‟ -> " (Double High-Reversed-9 Quotation Mark)
        '\u2032': "'",          # ′ -> ' (Prime)
        '\u2033': '"',          # ″ -> " (Double Prime)
        '\u2035': "'",          # ‵ -> ' (Reversed Prime)
        '\u2036': '"',          # ‶ -> " (Reversed Double Prime)
        
        # 特殊撇号
        '\u02b9': "'",          # ʹ -> ' (Modifier Letter Prime)
        '\u02ba': '"',          # ʺ -> " (Modifier Letter Double Prime)
        '\u02bc': "'",          # ʼ -> ' (Modifier Letter Apostrophe)
        '\u02c8': "'",          # ˈ -> ' (Modifier Letter Vertical Line)
        
        # 特殊括号
        '\u3008': '<',          # 〈 -> < (Left Angle Bracket)
        '\u3009': '>',          # 〉 -> > (Right Angle Bracket)
        '\u300a': '<<',         # 《 -> << (Left Double Angle Bracket)
        '\u300b': '>>',         # 》 -> >> (Right Double Angle Bracket)
        '\u300c': '"',          # 「 -> " (Left Corner Bracket)
        '\u300d': '"',          # 」 -> " (Right Corner Bracket)
        '\u300e': '"',          # 『 -> " (Left White Corner Bracket)
        '\u300f': '"',          # 』 -> " (Right White Corner Bracket)
        '\u3010': '[',          # 【 -> [ (Left Black Lenticular Bracket)
        '\u3011': ']',          # 】 -> ] (Right Black Lenticular Bracket)
        '\u3014': '[',          # 〔 -> [ (Left Tortoise Shell Bracket)
        '\u3015': ']',          # 〕 -> ] (Right Tortoise Shell Bracket)
        '\u3016': '[[',         # 〖 -> [[ (Left White Lenticular Bracket)
        '\u3017': ']]',         # 〗 -> ]] (Right White Lenticular Bracket)
        
        # 特殊波浪号和连接号
        '\uff5e': '~',          # ～ -> ~ (Fullwidth Tilde)
        '\u301c': '~',          # 〜 -> ~ (Wave Dash)
        '\u3030': '~',          # 〰 -> ~ (Wavy Dash)
        
        # 特殊分隔符
        '\u2022': '*',          # • -> * (Bullet)
        '\u2023': '>',          # ‣ -> > (Triangular Bullet)
        '\u25b6': '>',          # ▶ -> > (Black Right-Pointing Triangle)
        '\u25b8': '>',          # ▸ -> > (Black Right-Pointing Small Triangle)
        '\u25cf': '*',          # ● -> * (Black Circle)
        '\u25cb': 'o',          # ○ -> o (White Circle)
        
        # 特殊数学符号
        '\u00d7': '*',          # × -> * (Multiplication Sign)
        '\u00f7': '/',          # ÷ -> / (Division Sign)
        '\u2212': '-',          # − -> - (Minus Sign)
        '\u2215': '/',          # ∕ -> / (Division Slash)
        '\u2217': '*',          # ∗ -> * (Asterisk Operator)
        '\u2223': '|',          # ∣ -> | (Divides)
        '\u2236': ':',          # ∶ -> : (Ratio)
        '\u223c': '~',          # ∼ -> ~ (Tilde Operator)
        '\u2248': '~',          # ≈ -> ~ (Almost Equal To)
        '\u2260': '!=',         # ≠ -> != (Not Equal To)
        '\u2264': '<=',         # ≤ -> <= (Less-Than or Equal To)
        '\u2265': '>=',         # ≥ -> >= (Greater-Than or Equal To)
        
        # 特殊货币符号
        '\u00a2': 'c',          # ¢ -> c (Cent Sign)
        '\u00a3': 'GBP',        # £ -> GBP (Pound Sign)
        '\u00a5': 'CNY',        # ¥ -> CNY (Yen Sign)
        '\u20ac': 'EUR',        # € -> EUR (Euro Sign)
        
        # 特殊版权和商标符号
        '\u00a9': '(C)',        # © -> (C) (Copyright Sign)
        '\u00ae': '(R)',        # ® -> (R) (Registered Sign)
        '\u2122': '(TM)',       # ™ -> (TM) (Trade Mark Sign)
        
        # 特殊度数符号
        '\u00b0': 'deg',        # ° -> deg (Degree Sign)
        '\u2103': 'degC',       # ℃ -> degC (Degree Celsius)
        '\u2109': 'degF',       # ℉ -> degF (Degree Fahrenheit)
        
        # 特殊上标和下标数字（NFKC应该能处理，但以防万一）
        '\u2070': '0',          # ⁰ -> 0
        '\u00b9': '1',          # ¹ -> 1
        '\u00b2': '2',          # ² -> 2
        '\u00b3': '3',          # ³ -> 3
        '\u2074': '4',          # ⁴ -> 4
        '\u2075': '5',          # ⁵ -> 5
        '\u2076': '6',          # ⁶ -> 6
        '\u2077': '7',          # ⁷ -> 7
        '\u2078': '8',          # ⁸ -> 8
        '\u2079': '9',          # ⁹ -> 9
        '\u2080': '0',          # ₀ -> 0
        '\u2081': '1',          # ₁ -> 1
        '\u2082': '2',          # ₂ -> 2
        '\u2083': '3',          # ₃ -> 3
        '\u2084': '4',          # ₄ -> 4
        '\u2085': '5',          # ₅ -> 5
        '\u2086': '6',          # ₆ -> 6
        '\u2087': '7',          # ₇ -> 7
        '\u2088': '8',          # ₈ -> 8
        '\u2089': '9',          # ₉ -> 9
        
        # 特殊罗马数字（NFKC应该能处理）
        '\u2160': 'I',          # Ⅰ -> I
        '\u2161': 'II',         # Ⅱ -> II
        '\u2162': 'III',        # Ⅲ -> III
        '\u2163': 'IV',         # Ⅳ -> IV
        '\u2164': 'V',          # Ⅴ -> V
        '\u2165': 'VI',         # Ⅵ -> VI
        '\u2166': 'VII',        # Ⅶ -> VII
        '\u2167': 'VIII',       # Ⅷ -> VIII
        '\u2168': 'IX',         # Ⅸ -> IX
        '\u2169': 'X',          # Ⅹ -> X
        '\u2170': 'i',          # ⅰ -> i
        '\u2171': 'ii',         # ⅱ -> ii
        '\u2172': 'iii',        # ⅲ -> iii
        '\u2173': 'iv',         # ⅳ -> iv
        '\u2174': 'v',          # ⅴ -> v
        '\u2175': 'vi',         # ⅵ -> vi
        '\u2176': 'vii',        # ⅶ -> vii
        '\u2177': 'viii',       # ⅷ -> viii
        '\u2178': 'ix',         # ⅸ -> ix
        '\u2179': 'x',          # ⅹ -> x
        
        # 特殊圆圈数字
        '\u2460': '(1)',        # ① -> (1)
        '\u2461': '(2)',        # ② -> (2)
        '\u2462': '(3)',        # ③ -> (3)
        '\u2463': '(4)',        # ④ -> (4)
        '\u2464': '(5)',        # ⑤ -> (5)
        '\u2465': '(6)',        # ⑥ -> (6)
        '\u2466': '(7)',        # ⑦ -> (7)
        '\u2467': '(8)',        # ⑧ -> (8)
        '\u2468': '(9)',        # ⑨ -> (9)
        '\u2469': '(10)',       # ⑩ -> (10)
        '\u246a': '(11)',       # ⑪ -> (11)
        '\u246b': '(12)',       # ⑫ -> (12)
        '\u246c': '(13)',       # ⑬ -> (13)
        '\u246d': '(14)',       # ⑭ -> (14)
        '\u246e': '(15)',       # ⑮ -> (15)
        '\u246f': '(16)',       # ⑯ -> (16)
        '\u2470': '(17)',       # ⑰ -> (17)
        '\u2471': '(18)',       # ⑱ -> (18)
        '\u2472': '(19)',       # ⑲ -> (19)
        '\u2473': '(20)',       # ⑳ -> (20)
        
        # 特殊带括号数字
        '\u2474': '(1)',        # ⑴ -> (1)
        '\u2475': '(2)',        # ⑵ -> (2)
        '\u2476': '(3)',        # ⑶ -> (3)
        '\u2477': '(4)',        # ⑷ -> (4)
        '\u2478': '(5)',        # ⑸ -> (5)
        '\u2479': '(6)',        # ⑹ -> (6)
        '\u247a': '(7)',        # ⑺ -> (7)
        '\u247b': '(8)',        # ⑻ -> (8)
        '\u247c': '(9)',        # ⑼ -> (9)
        '\u247d': '(10)',       # ⑽ -> (10)
        
        # 特殊带圈字母
        '\u24b6': '(A)',        # Ⓐ -> (A)
        '\u24b7': '(B)',        # Ⓑ -> (B)
        '\u24b8': '(C)',        # Ⓒ -> (C)
        '\u24b9': '(D)',        # Ⓓ -> (D)
        '\u24ba': '(E)',        # Ⓔ -> (E)
        '\u24bb': '(F)',        # Ⓕ -> (F)
        '\u24bc': '(G)',        # Ⓖ -> (G)
        '\u24bd': '(H)',        # Ⓗ -> (H)
        '\u24be': '(I)',        # Ⓘ -> (I)
        '\u24bf': '(J)',        # Ⓙ -> (J)
        '\u24c0': '(K)',        # Ⓚ -> (K)
        '\u24c1': '(L)',        # Ⓛ -> (L)
        '\u24c2': '(M)',        # Ⓜ -> (M)
        '\u24c3': '(N)',        # Ⓝ -> (N)
        '\u24c4': '(O)',        # Ⓞ -> (O)
        '\u24c5': '(P)',        # Ⓟ -> (P)
        '\u24c6': '(Q)',        # Ⓠ -> (Q)
        '\u24c7': '(R)',        # Ⓡ -> (R)
        '\u24c8': '(S)',        # Ⓢ -> (S)
        '\u24c9': '(T)',        # Ⓣ -> (T)
        '\u24ca': '(U)',        # Ⓤ -> (U)
        '\u24cb': '(V)',        # Ⓥ -> (V)
        '\u24cc': '(W)',        # Ⓦ -> (W)
        '\u24cd': '(X)',        # Ⓧ -> (X)
        '\u24ce': '(Y)',        # Ⓨ -> (Y)
        '\u24cf': '(Z)',        # Ⓩ -> (Z)
        
        # 特殊带圈小写字母
        '\u24d0': '(a)',        # ⓐ -> (a)
        '\u24d1': '(b)',        # ⓑ -> (b)
        '\u24d2': '(c)',        # ⓒ -> (c)
        '\u24d3': '(d)',        # ⓓ -> (d)
        '\u24d4': '(e)',        # ⓔ -> (e)
        '\u24d5': '(f)',        # ⓕ -> (f)
        '\u24d6': '(g)',        # ⓖ -> (g)
        '\u24d7': '(h)',        # ⓗ -> (h)
        '\u24d8': '(i)',        # ⓘ -> (i)
        '\u24d9': '(j)',        # ⓙ -> (j)
        '\u24da': '(k)',        # ⓚ -> (k)
        '\u24db': '(l)',        # ⓛ -> (l)
        '\u24dc': '(m)',        # ⓜ -> (m)
        '\u24dd': '(n)',        # ⓝ -> (n)
        '\u24de': '(o)',        # ⓞ -> (o)
        '\u24df': '(p)',        # ⓟ -> (p)
        '\u24e0': '(q)',        # ⓠ -> (q)
        '\u24e1': '(r)',        # ⓡ -> (r)
        '\u24e2': '(s)',        # ⓢ -> (s)
        '\u24e3': '(t)',        # ⓣ -> (t)
        '\u24e4': '(u)',        # ⓤ -> (u)
        '\u24e5': '(v)',        # ⓥ -> (v)
        '\u24e6': '(w)',        # ⓦ -> (w)
        '\u24e7': '(x)',        # ⓧ -> (x)
        '\u24e8': '(y)',        # ⓨ -> (y)
        '\u24e9': '(z)',        # ⓩ -> (z)
    }
    
    # 标点规范化规则 - 用于清理标点周围的空格
    PUNCTUATION_RULES: List[Tuple[str, str]] = [
        # 移除零宽空格（先处理，避免影响其他规则）
        (r'[\u200b\u200c\u200d\ufeff]', ''),
        # 移除中文标点前空格
        (r'\s+([，。、；：""''（）【】《》？！])', r'\1'),
        # 移除英文标点前空格
        (r'\s+([,;:.!?])', r'\1'),
        # 在英文标点后添加空格（如果不是紧跟数字或字母）
        (r'([,;:.!?])([^ \n\d])', r'\1 \2'),
        # 移除连续的中文标点间的空格
        (r'([，。、；：])\s+([，。、；：])', r'\1\2'),
        # 移除多个连续空格
        (r' +', ' '),
    ]
    
    def __init__(self, 
                 use_nfkc: bool = True,
                 use_custom_mappings: bool = True,
                 use_punctuation_rules: bool = True,
                 additional_mappings: Optional[Dict[str, str]] = None):
        """
        初始化文本标准化器
        
        Args:
            use_nfkc: 是否使用NFKC标准化
            use_custom_mappings: 是否使用自定义字符映射
            use_punctuation_rules: 是否使用标点规范化规则
            additional_mappings: 额外的自定义字符映射
        """
        self.use_nfkc = use_nfkc
        self.use_custom_mappings = use_custom_mappings
        self.use_punctuation_rules = use_punctuation_rules
        
        # 合并额外映射
        self.mappings = self.CUSTOM_CHAR_MAPPINGS.copy()
        if additional_mappings:
            self.mappings.update(additional_mappings)
    
    def normalize(self, text: str) -> str:
        """
        标准化文本
        
        Args:
            text: 输入文本
            
        Returns:
            标准化后的文本
        """
        if not text:
            return text
        
        # 第一步：NFKC标准化
        if self.use_nfkc:
            text = unicodedata.normalize('NFKC', text)
        
        # 第二步：自定义字符映射
        if self.use_custom_mappings:
            text = self._apply_custom_mappings(text)
        
        # 第三步：标点规范化
        if self.use_punctuation_rules:
            text = self._apply_punctuation_rules(text)
        
        return text.strip()
    
    def _apply_custom_mappings(self, text: str) -> str:
        """应用自定义字符映射"""
        for old_char, new_char in self.mappings.items():
            text = text.replace(old_char, new_char)
        return text
    
    def _apply_punctuation_rules(self, text: str) -> str:
        """应用标点规范化规则"""
        for pattern, replacement in self.PUNCTUATION_RULES:
            text = re.sub(pattern, replacement, text)
        return text
    
    def normalize_paragraphs(self, paragraphs: List[str]) -> List[str]:
        """
        批量标准化段落列表
        
        Args:
            paragraphs: 段落文本列表
            
        Returns:
            标准化后的段落列表
        """
        return [self.normalize(p) for p in paragraphs if p.strip()]
    
    def get_character_info(self, char: str) -> Dict[str, str]:
        """
        获取字符的Unicode信息
        
        Args:
            char: 单个字符
            
        Returns:
            包含字符信息的字典
        """
        if len(char) != 1:
            return {'error': 'Input must be a single character'}
        
        code = ord(char)
        try:
            name = unicodedata.name(char)
        except ValueError:
            name = 'UNKNOWN'
        
        nfkc = unicodedata.normalize('NFKC', char)
        
        return {
            'char': char,
            'code': f'U+{code:04X}',
            'decimal': str(code),
            'name': name,
            'category': unicodedata.category(char),
            'nfkc': nfkc,
            'in_custom_mappings': char in self.mappings
        }
    
    def analyze_text(self, text: str) -> List[Dict[str, str]]:
        """
        分析文本中的所有字符
        
        Args:
            text: 输入文本
            
        Returns:
            字符信息列表
        """
        unique_chars = set(text)
        return [self.get_character_info(c) for c in sorted(unique_chars)]


# 便捷函数
def normalize_text(text: str, **kwargs) -> str:
    """
    快速标准化文本
    
    Args:
        text: 输入文本
        **kwargs: 传递给TextNormalizer的参数
        
    Returns:
        标准化后的文本
    """
    normalizer = TextNormalizer(**kwargs)
    return normalizer.normalize(text)


def normalize_pdf_text(text: str) -> str:
    """
    专为PDF文本优化的标准化
    
    启用所有优化选项，适合处理PDF提取的文本
    
    Args:
        text: PDF提取的文本
        
    Returns:
        标准化后的文本
    """
    return normalize_text(
        text,
        use_nfkc=True,
        use_custom_mappings=True,
        use_punctuation_rules=True
    )
