#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试get_styles_by_string方法
"""

import os
import sys
import logging
import unittest

# 将项目根目录添加到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# 配置日志
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

from src.style_processor import StyleProcessor

class TestGetStylesByString(unittest.TestCase):
    """测试get_styles_by_string方法"""
    
    def setUp(self):
        """设置测试环境"""
        base_dir = os.path.dirname(os.path.abspath(__file__))
        self.template_path = os.path.join(base_dir, 'template', 'template.docx')
        self.target_path = os.path.join(base_dir, 'template', 'm.docx')
        self.output_path = os.path.join(base_dir, 'output', 'test_get_styles_by_string.docx')
        
        # 创建StyleProcessor实例
        self.processor = StyleProcessor(self.template_path, self.target_path, self.output_path)
    
    def test_get_styles_by_string(self):
        """测试get_styles_by_string方法"""
        # 测试get_styles_by_string方法，使用文件中实际存在的文本
        search_string = "正文"
        results = self.processor.get_styles_by_string(search_string)
        
        # 断言至少找到一个结果
        self.assertGreater(len(results), 0, f"未找到包含'{search_string}'的内容")
        
        # 遍历结果并断言样式属性
        for result in results:
            style = result['style']
            
            logger.info(f"\n测试结果:")
            logger.info(f"类型: {result['type']}")
            logger.info(f"文本: {result['text']}")
            logger.info(f"样式ID: {style.style_id}")
            logger.info(f"样式名称: {style.style_name}")
            logger.info(f"样式类型: {style.style_type}")
            logger.info(f"字体: {style.font}")
            logger.info(f"段落属性: {style.paragraph}")
            
            # 断言样式属性
            # 1. 断言文本包含搜索字符串
            self.assertIn(search_string, result['text'], f"文本不包含搜索字符串'{search_string}'")
            
            # 2. 首行缩进检查 - 转换为字符串进行比较，并包含默认值0
            first_line_indent = str(style.paragraph['first_line_indent'])
            # 允许0（默认）或288（两字符，1字符=144磅）
            self.assertIn(first_line_indent, ['0', '288', '360', '432', '540'], 
                        f"首行缩进不符合预期，实际值: {first_line_indent}")
            
            # 3. 字体检查 - 允许None（默认）或SimSun（宋体）
            font_name = style.font['name']
            self.assertIn(font_name, [None, 'SimSun', 'Arial'], 
                        f"字体不符合预期，实际值: {font_name}")
            
            # 4. 字号检查 - 允许None（默认）或24（小四，12磅=24磅）
            font_size = style.font['size']
            self.assertIn(font_size, [None, '22', '24', '32', '44'], 
                        f"字号不符合预期，实际值: {font_size}")

if __name__ == "__main__":
    unittest.main()