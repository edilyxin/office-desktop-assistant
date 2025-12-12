#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF文件处理工具类
用于将PDF文件转换为图片，支持多页PDF
"""

import logging
import os
import tempfile
from PIL import Image
import fitz  # PyMuPDF

# 配置日志记录
logger = logging.getLogger('PaddleOCRVL')


class PDFUtils:
    """
    PDF文件处理工具类
    """
    
    @staticmethod
    def pdf_to_images(pdf_path, dpi=300, max_width=800):
        """
        将PDF文件转换为图片列表
        
        Args:
            pdf_path: PDF文件路径
            dpi: 图片分辨率，默认300
            max_width: 图片最大宽度，默认800
            
        Returns:
            list: 图片字节流列表，每个元素是一页的图片字节流
        """
        logger.info(f"开始将PDF转换为图片: {pdf_path}")
        
        try:
            # 打开PDF文件
            pdf_document = fitz.open(pdf_path)
            logger.info(f"成功打开PDF文件，共 {pdf_document.page_count} 页")
            
            images = []
            
            # 遍历所有页面
            for page_num in range(pdf_document.page_count):
                logger.debug(f"处理PDF页面 {page_num + 1}/{pdf_document.page_count}")
                
                # 获取页面
                page = pdf_document[page_num]
                
                # 设置分辨率
                zoom = dpi / 72  # 72是PDF的默认DPI
                matrix = fitz.Matrix(zoom, zoom)
                
                # 渲染页面为图片
                pix = page.get_pixmap(matrix=matrix, alpha=False)
                
                # 将图片转换为PIL Image对象
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                
                # 调整图片大小，保持比例
                if img.width > max_width:
                    ratio = max_width / img.width
                    new_height = int(img.height * ratio)
                    img = img.resize((max_width, new_height), Image.LANCZOS)
                
                # 将图片转换为字节流
                img_bytes = tempfile.SpooledTemporaryFile()
                img.save(img_bytes, format="PNG")
                img_bytes.seek(0)
                images.append(img_bytes.read())
                img_bytes.close()
            
            # 关闭PDF文件
            pdf_document.close()
            
            logger.info(f"成功将PDF转换为 {len(images)} 张图片")
            return images
            
        except Exception as e:
            logger.error(f"将PDF转换为图片失败: {str(e)}", exc_info=True)
            return []
    
    @staticmethod
    def is_pdf_file(file_path):
        """
        检查文件是否为PDF文件
        
        Args:
            file_path: 文件路径
            
        Returns:
            bool: 是否为PDF文件
        """
        return file_path.lower().endswith('.pdf')
    
    @staticmethod
    def get_pdf_page_count(pdf_path):
        """
        获取PDF文件的页数
        
        Args:
            pdf_path: PDF文件路径
            
        Returns:
            int: PDF页数，失败返回0
        """
        try:
            pdf_document = fitz.open(pdf_path)
            page_count = pdf_document.page_count
            pdf_document.close()
            return page_count
        except Exception as e:
            logger.error(f"获取PDF页数失败: {str(e)}")
            return 0
