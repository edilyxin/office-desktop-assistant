#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
查看目标文件内容
"""

from zipfile import ZipFile
from lxml import etree
import os

base_dir = os.path.dirname(os.path.abspath(__file__))
target_path = os.path.join(base_dir, 'template', 'm.docx')

with ZipFile(target_path, 'r') as target_zip:
    document_xml_content = target_zip.read('word/document.xml')
    root = etree.fromstring(document_xml_content)
    
    # 搜索所有文本元素
    namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    text_elements = root.xpath('//w:t', namespaces=namespaces)
    
    print("目标文件内容:")
    print("=" * 50)
    for i, elem in enumerate(text_elements):
        if elem.text:
            print(f"文本{i+1}: {elem.text}")
    print("=" * 50)
    print(f"共找到 {len(text_elements)} 个文本元素")