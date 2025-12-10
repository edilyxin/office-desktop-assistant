#!/usr/bin/env python3
"""
测试脚本：手动设置模板文件和目标文件，使用output文件夹作为输出目录
"""

import os
import sys
import logging
import random
import string

# 将项目根目录添加到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# 配置日志
logging.basicConfig(level=logging.DEBUG, 
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

from src.style_processor import StyleProcessor

def main():
    """主测试函数"""
    logger.info("测试消息：开始手动样式处理测试")
    
    # 手动设置文件路径
    base_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(base_dir, 'template', 'template.docx')
    target_path = os.path.join(base_dir, 'template', 'm.docx')
    
    # 生成随机四位字符
    random_suffix = ''.join(random.choices(string.ascii_letters + string.digits, k=4))
    output_path = os.path.join(base_dir, 'output', f'processed_test_{random_suffix}.docx')
    
    # 检查模板文件是否存在
    if not os.path.exists(template_path):
        logger.error(f"测试消息：模板文件不存在: {template_path}")
        return 1
    
    # 检查目标文件是否存在
    if not os.path.exists(target_path):
        logger.error(f"测试消息：目标文件不存在: {target_path}")
        return 1
    
    logger.info(f"测试消息：模板文件: {template_path}")
    logger.info(f"测试消息：目标文件: {target_path}")
    logger.info(f"测试消息：输出文件: {output_path}")
    
    try:
        # 创建StyleProcessor实例
        processor = StyleProcessor(template_path, target_path, output_path)
        
        # 1. 获取模板样式并打印
        logger.info("\n=== 测试1: 获取模板样式 ===")
        template_styles = processor.get_template_styles()
        processor.print_styles(template_styles)
        
        # 2. 获取目标样式并打印
        logger.info("\n=== 测试2: 获取目标样式 ===")
        target_styles = processor.get_target_styles()
        processor.print_styles(target_styles)
        
        # 3. 根据字符串查询样式
        logger.info("\n=== 测试3: 根据字符串查询样式 ===")
        search_string = "测试"
        styles_by_string = processor.get_styles_by_string(search_string)
        
        for result in styles_by_string:
            logger.info(f"\n找到类型: {result['type']}, 索引: {result['index']}")
            logger.info(f"文本内容: {result['text'][:100]}...")
            logger.info(f"样式ID: {result['style'].style_id}")
            logger.info(f"样式名称: {result['style'].style_name}")
            logger.info(f"样式类型: {result['style'].style_type}")
            logger.info(f"字体: {result['style'].font}")
            logger.info(f"段落属性: {result['style'].paragraph}")
            if result['style'].style_type == 'table':
                logger.info(f"表格属性: {result['style'].table}")
            elif result['style'].style_type == 'list':
                logger.info(f"列表属性: {result['style'].list}")
        
        # 4. 执行完整的样式处理流程
        logger.info("\n=== 测试4: 完整样式处理流程 ===")
        success = processor.process()
        
        if success:
            logger.info(f"测试消息：样式处理成功！输出文件已保存到: {output_path}")
            return 0
        else:
            logger.error("测试消息：样式处理失败！")
            return 1
            
    except Exception as e:
        logger.error(f"测试消息：测试过程中发生错误: {str(e)}", exc_info=True)  
        return 1

if __name__ == "__main__":
    sys.exit(main())