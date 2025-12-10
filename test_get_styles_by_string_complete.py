#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
完整测试get_styles_by_string方法
"""

import os
import sys
import logging

# 将项目根目录添加到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# 配置日志
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

from src.style_processor import StyleProcessor

def main():
    """主测试函数"""
    logger.info("开始完整测试get_styles_by_string方法")
    
    # 设置文件路径
    base_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(base_dir, 'template', 'template.docx')
    target_path = os.path.join(base_dir, 'template', 'm.docx')
    output_path = os.path.join(base_dir, 'output', 'test_get_styles_by_string_complete.docx')
    
    # 创建StyleProcessor实例
    processor = StyleProcessor(template_path, target_path, output_path)
    
    # 测试不同的搜索字符串
    search_strings = ["正文", "测试", "招标", "目录"]
    
    for search_string in search_strings:
        logger.info(f"\n=== 测试搜索字符串: '{search_string}' ===")
        results = processor.get_styles_by_string(search_string)
        
        logger.info(f"找到 {len(results)} 个匹配项")
        
        for i, result in enumerate(results):
            style = result['style']
            logger.info(f"\n匹配项 {i+1}:")
            logger.info(f"类型: {result['type']}")
            logger.info(f"索引: {result['index']}")
            logger.info(f"文本: {result['text'][:100]}...")
            logger.info(f"样式ID: {style.style_id}")
            logger.info(f"样式名称: {style.style_name}")
            logger.info(f"样式类型: {style.style_type}")
            logger.info(f"字体名称: {style.font['name']}")
            logger.info(f"字体大小: {style.font['size']}")
            logger.info(f"首行缩进: {style.paragraph['first_line_indent']}")
    
    logger.info("\n=== 测试完成 ===")

if __name__ == "__main__":
    main()