#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word文件样式处理模块
基于OpenXML底层实现样式提取、映射和应用
"""

import logging
import os
from lxml import etree
from zipfile import ZipFile
import tempfile

# 配置日志记录
logger = logging.getLogger('PaddleOCRVL')

class Style:
    """
    样式类，封装样式数据
    借鉴OpenXML样式定义设计
    """
    
    def __init__(self, style_id, style_name, style_type='paragraph'):
        """
        初始化样式对象
        
        Args:
            style_id (str): 样式ID
            style_name (str): 样式名称
            style_type (str): 样式类型 (paragraph, character, table, list, page)
        """
        self.style_id = style_id
        self.style_name = style_name
        self.style_type = style_type
        
        # 基本属性
        self.font = {
            'name': None,
            'size': None,
            'bold': False,
            'italic': False,
            'underline': False,
            'color': None,
            'highlight': None,
            'superscript': False,
            'subscript': False,
            'strike': False,
            'double_strike': False,
            'all_caps': False,
            'small_caps': False,
            'shadow': False,
            'outline': False,
            'emboss': False,
            'engrave': False
        }
        
        # 段落属性
        self.paragraph = {
            'alignment': None,
            'line_spacing': 1.0,
            'space_before': 0,
            'space_after': 0,
            'indent_left': 0,
            'indent_right': 0,
            'first_line_indent': 0,
            'keep_together': False,
            'keep_with_next': False,
            'page_break_before': False,
            'widow_control': True
        }
        
        # 表格属性
        self.table = {
            'border_top': None,
            'border_bottom': None,
            'border_left': None,
            'border_right': None,
            'cell_margins': None,
            'shading': None
        }
        
        # 列表属性
        self.list = {
            'format': None,
            'level': 0,
            'start': 1,
            'indent': 0,
            'followed_by': None
        }
        
        # 页面属性
        self.page = {
            'size': None,
            'margins': None,
            'orientation': None,
            'header_distance': None,
            'footer_distance': None
        }
    
    def __str__(self):
        """字符串表示"""
        return f"Style(id='{self.style_id}', name='{self.style_name}', type='{self.style_type}')"
    
    def to_dict(self):
        """转换为字典"""
        return {
            'id': self.style_id,
            'name': self.style_name,
            'type': self.style_type,
            'font': self.font,
            'paragraph': self.paragraph,
            'table': self.table,
            'list': self.list,
            'page': self.page
        }
    
    def to_json(self):
        """转换为JSON字符串"""
        import json
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=2)

class StyleProcessor:
    """
    基于OpenXML底层的Word样式处理器
    实现模板样式的提取、映射和应用
    """
    
    def __init__(self, template_path, target_path, output_path):
        """
        初始化样式处理器
        
        Args:
            template_path: 模板文件路径
            target_path: 目标文件路径
            output_path: 输出文件路径
        """
        self.template_path = template_path
        self.target_path = target_path
        self.output_path = output_path
        
        # 样式映射表
        self.style_mapping = {}
        
        # 命名空间
        self.NS = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006'
        }
    
    def extract_template_styles(self):
        """
        从模板文件中提取完整的styles.xml内容
        
        Returns:
            tuple: (styles_xml_root, style_id_mapping)
        """
        logger.info(f"开始从模板文件提取样式: {self.template_path}")
        
        try:
            # 打开模板文件
            with ZipFile(self.template_path, 'r') as template_zip:
                # 读取styles.xml内容
                styles_xml_content = template_zip.read('word/styles.xml')
                styles_xml_root = etree.fromstring(styles_xml_content)
                
                # 提取所有样式ID和名称的映射
                style_id_mapping = {}
                for style_elem in styles_xml_root.xpath('//w:style', namespaces=self.NS):
                    style_id = style_elem.get(f'{{{self.NS["w"]}}}styleId')
                    style_name_elem = style_elem.find('.//w:name', namespaces=self.NS)
                    if style_name_elem is not None:
                        style_name = style_name_elem.get(f'{{{self.NS["w"]}}}val')
                        style_id_mapping[style_id] = style_name
                        logger.debug(f"模板样式: ID={style_id}, 名称={style_name}")
                
                logger.info(f"成功提取模板样式，共 {len(style_id_mapping)} 个样式")
                logger.debug(f"模板样式列表: {style_id_mapping}")
                return styles_xml_root, style_id_mapping
                
        except Exception as e:
            logger.error(f"提取模板样式失败: {str(e)}", exc_info=True)
            return None, None
    
    def analyze_target_document(self):
        """
        分析目标文档，提取原有样式和内容引用
        
        Returns:
            tuple: (target_styles_xml_root, used_style_ids)
        """
        logger.info(f"开始分析目标文档: {self.target_path}")
        
        try:
            # 打开目标文件
            with ZipFile(self.target_path, 'r') as target_zip:
                # 读取styles.xml内容
                target_styles_xml_content = target_zip.read('word/styles.xml')
                target_styles_xml_root = etree.fromstring(target_styles_xml_content)
                
                # 提取目标文档的样式ID和名称映射
                target_style_id_mapping = {}
                for style_elem in target_styles_xml_root.xpath('//w:style', namespaces=self.NS):
                    style_id = style_elem.get(f'{{{self.NS["w"]}}}styleId')
                    style_name_elem = style_elem.find('.//w:name', namespaces=self.NS)
                    if style_name_elem is not None:
                        style_name = style_name_elem.get(f'{{{self.NS["w"]}}}val')
                        target_style_id_mapping[style_id] = style_name
                        logger.debug(f"目标文档样式: ID={style_id}, 名称={style_name}")
                
                # 读取document.xml内容，分析使用的样式ID
                document_xml_content = target_zip.read('word/document.xml')
                document_xml_root = etree.fromstring(document_xml_content)
                
                # 提取所有使用的样式ID
                used_style_ids = set()
                
                # 提取段落样式引用
                logger.debug("开始分析目标文档正文中的样式引用...")
                for pstyle_elem in document_xml_root.xpath('//w:pStyle', namespaces=self.NS):
                    style_id = pstyle_elem.get(f'{{{self.NS["w"]}}}val')
                    if style_id:
                        used_style_ids.add(style_id)
                        logger.debug(f"段落使用样式ID: {style_id}")
                
                # 提取字符样式引用
                for rstyle_elem in document_xml_root.xpath('//w:rStyle', namespaces=self.NS):
                    style_id = rstyle_elem.get(f'{{{self.NS["w"]}}}val')
                    if style_id:
                        used_style_ids.add(style_id)
                        logger.debug(f"字符使用样式ID: {style_id}")
                
                # 提取表格样式引用
                for tblstyle_elem in document_xml_root.xpath('//w:tblStyle', namespaces=self.NS):
                    style_id = tblstyle_elem.get(f'{{{self.NS["w"]}}}val')
                    if style_id:
                        used_style_ids.add(style_id)
                        logger.debug(f"表格使用样式ID: {style_id}")
                
                logger.info(f"成功分析目标文档，共使用 {len(used_style_ids)} 个样式")
                logger.debug(f"目标文档使用的样式ID列表: {used_style_ids}")
                return target_styles_xml_root, used_style_ids
                
        except Exception as e:
            logger.error(f"分析目标文档失败: {str(e)}", exc_info=True)
            return None, None
    
    def build_style_mapping(self, template_style_id_mapping, target_styles_xml_root):
        """
        建立新旧样式ID映射表
        
        Args:
            template_style_id_mapping: 模板样式ID到名称的映射
            target_styles_xml_root: 目标文档样式XML根节点
            
        Returns:
            dict: 样式映射表 {旧样式ID: 新样式ID}
        """
        logger.info("开始建立样式映射关系")
        
        # 提取目标文档的样式ID和名称映射
        target_style_id_mapping = {}
        for style_elem in target_styles_xml_root.xpath('//w:style', namespaces=self.NS):
            style_id = style_elem.get(f'{{{self.NS["w"]}}}styleId')
            style_name_elem = style_elem.find('.//w:name', namespaces=self.NS)
            if style_name_elem is not None:
                style_name = style_name_elem.get(f'{{{self.NS["w"]}}}val')
                target_style_id_mapping[style_id] = style_name
                logger.debug(f"目标文档样式: ID={style_id}, 名称={style_name}")
        
        # 打印模板样式映射
        logger.debug(f"模板样式映射: {template_style_id_mapping}")
        logger.debug(f"目标样式映射: {target_style_id_mapping}")
        
        # 建立映射表：基于样式名称匹配
        mapping = {}
        
        # 1. 首先进行精确匹配
        for old_style_id, old_style_name in target_style_id_mapping.items():
            # 查找模板中同名的样式
            for new_style_id, new_style_name in template_style_id_mapping.items():
                if old_style_name == new_style_name:
                    mapping[old_style_id] = new_style_id
                    logger.debug(f"建立样式映射: {old_style_id} -> {new_style_id} (精确匹配: {old_style_name})")
                    break
        
        # 2. 然后进行前缀匹配，处理标题样式等情况
        for old_style_id, old_style_name in target_style_id_mapping.items():
            if old_style_id not in mapping:
                # 检查是否为标题样式
                if old_style_name.startswith('heading'):
                    # 在模板中查找标题样式
                    for new_style_id, new_style_name in template_style_id_mapping.items():
                        if new_style_name.startswith('heading'):
                            mapping[old_style_id] = new_style_id
                            logger.debug(f"建立样式映射: {old_style_id} -> {new_style_id} (前缀匹配: {old_style_name} -> {new_style_name})")
                            break
                # 可以添加更多的匹配规则
        
        # 3. 最后，为所有未匹配的样式添加默认映射
        for old_style_id in target_style_id_mapping:
            if old_style_id not in mapping:
                # 使用模板中的Normal样式作为默认样式
                normal_style_id = None
                for new_style_id, new_style_name in template_style_id_mapping.items():
                    if new_style_name == 'Normal':
                        normal_style_id = new_style_id
                        break
                if normal_style_id:
                    mapping[old_style_id] = normal_style_id
                    logger.debug(f"建立样式映射: {old_style_id} -> {normal_style_id} (默认匹配)")
        
        logger.info(f"成功建立 {len(mapping)} 个样式映射关系")
        logger.debug(f"样式映射表: {mapping}")
        return mapping
    
    def update_styles_definition(self, template_styles_xml_root, target_zip_path):
        """
        更新目标文档的样式定义库
        
        Args:
            template_styles_xml_root: 模板样式XML根节点
            target_zip_path: 目标文件路径
            
        Returns:
            bool: 是否成功
        """
        logger.info("开始更新目标文档样式定义库")
        
        try:
            # 1. 创建临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                logger.debug(f"创建临时目录: {temp_dir}")
                
                # 2. 解压目标文件到临时目录
                logger.debug(f"解压目标文件到临时目录: {target_zip_path}")
                with ZipFile(target_zip_path, 'r') as target_zip:
                    target_zip.extractall(temp_dir)
                
                # 3. 读取目标文档的styles.xml
                target_styles_path = os.path.join(temp_dir, 'word', 'styles.xml')
                logger.debug(f"读取目标文档的styles.xml: {target_styles_path}")
                
                with open(target_styles_path, 'rb') as f:
                    target_styles_xml_content = f.read()
                
                # 4. 解析目标文档的styles.xml
                target_styles_xml_root = etree.fromstring(target_styles_xml_content)
                logger.debug(f"成功解析目标文档的styles.xml")
                
                # 5. 找到或创建目标文档的styles元素，我看到了问题所在。
                # 在更新样式定义库的过程中，代码尝试在目标文档中查找styles元素，但失败了，日志显示"ERROR - 在目标文档中未找到styles元素"。
                # 这是因为在目标文档的styles.xml中，styles元素可能位于不同的位置，或者有不同的结构。
                # 修改代码，处理这种情况，确保即使styles元素不存在，也能正确创建它。
                styles_elem = target_styles_xml_root.find('.//w:styles', namespaces=self.NS)
                if styles_elem is None:
                    # 如果不存在，创建styles元素作为根元素的子元素
                    logger.debug("在目标文档中未找到styles元素，创建新的styles元素")
                    styles_elem = etree.SubElement(target_styles_xml_root, f'{{{self.NS["w"]}}}styles')
                
                # 6. 移除目标文档中的所有样式元素
                existing_styles = styles_elem.xpath('./w:style', namespaces=self.NS)
                logger.debug(f"移除目标文档中的所有样式元素，共 {len(existing_styles)} 个样式")
                for style_elem in existing_styles:
                    styles_elem.remove(style_elem)
                
                # 7. 移除目标文档中的现有默认样式设置
                existing_doc_defaults = styles_elem.find('./w:docDefaults', namespaces=self.NS)
                if existing_doc_defaults is not None:
                    styles_elem.remove(existing_doc_defaults)
                    logger.debug("移除目标文档中的现有默认样式设置")
                
                # 8. 从模板中获取所有样式元素
                template_styles = template_styles_xml_root.xpath('//w:style', namespaces=self.NS)
                logger.debug(f"从模板中获取到 {len(template_styles)} 个样式元素")
                
                # 9. 添加模板中的所有样式到目标文档
                added_count = 0
                for template_style_elem in template_styles:
                    # 创建样式元素的深拷贝
                    new_style_elem = etree.ElementTree(template_style_elem).getroot()
                    # 添加到目标文档的styles.xml中
                    styles_elem.append(new_style_elem)
                    added_count += 1
                
                logger.debug(f"已将 {added_count} 个样式元素添加到目标文档")
                
                # 10. 添加模板中的默认样式设置
                template_doc_defaults = template_styles_xml_root.find('.//w:docDefaults', namespaces=self.NS)
                if template_doc_defaults is not None:
                    # 创建默认样式设置的深拷贝
                    new_doc_defaults = etree.ElementTree(template_doc_defaults).getroot()
                    # 添加到目标文档的styles.xml中
                    styles_elem.append(new_doc_defaults)
                    logger.debug("已将模板中的默认样式设置添加到目标文档")
                else:
                    logger.debug("模板中没有默认样式设置")
                
                # 11. 保存更新后的styles.xml
                updated_content = etree.tostring(target_styles_xml_root, pretty_print=True, encoding='utf-8', xml_declaration=True)
                logger.debug(f"更新后的styles.xml大小: {len(updated_content)} 字节")
                
                with open(target_styles_path, 'wb') as f:
                    f.write(updated_content)
                
                logger.debug(f"已保存更新后的styles.xml: {target_styles_path}")
                
                # 12. 重新打包为新的Word文件
                logger.debug(f"开始重新打包文件到: {self.output_path}")
                
                with ZipFile(self.output_path, 'w') as output_zip:
                    # 遍历临时目录中的所有文件，添加到输出zip
                    for root, dirs, files in os.walk(temp_dir):
                        for file in files:
                            file_path = os.path.join(root, file)
                            # 计算在zip中的相对路径
                            zip_path = os.path.relpath(file_path, temp_dir)
                            output_zip.write(file_path, zip_path)
                
                logger.info(f"成功更新目标文档样式定义库，输出文件: {self.output_path}")
                return True
                
        except Exception as e:
            logger.error(f"更新目标文档样式定义库失败: {str(e)}", exc_info=True)
            return False
    
    def update_content_references(self, style_mapping):
        """
        更新文档内容中的样式引用
        
        Args:
            style_mapping: 样式映射表
            
        Returns:
            bool: 是否成功
        """
        logger.info("开始更新文档内容样式引用")
        
        try:
            # 创建临时目录
            with tempfile.TemporaryDirectory() as temp_dir:
                # 解压输出文件到临时目录
                with ZipFile(self.output_path, 'r') as output_zip:
                    output_zip.extractall(temp_dir)
                
                # 读取document.xml
                document_path = os.path.join(temp_dir, 'word', 'document.xml')
                with open(document_path, 'rb') as f:
                    document_xml_content = f.read()
                
                document_xml_root = etree.fromstring(document_xml_content)
                
                # 更新所有样式引用
                updated_count = 0
                
                # 更新段落样式引用
                for pstyle_elem in document_xml_root.xpath('//w:pStyle', namespaces=self.NS):
                    old_style_id = pstyle_elem.get(f'{{{self.NS["w"]}}}val')
                    if old_style_id in style_mapping:
                        new_style_id = style_mapping[old_style_id]
                        pstyle_elem.set(f'{{{self.NS["w"]}}}val', new_style_id)
                        updated_count += 1
                        logger.debug(f"更新段落样式引用: {old_style_id} -> {new_style_id}")
                
                # 更新字符样式引用
                for rstyle_elem in document_xml_root.xpath('//w:rStyle', namespaces=self.NS):
                    old_style_id = rstyle_elem.get(f'{{{self.NS["w"]}}}val')
                    if old_style_id in style_mapping:
                        new_style_id = style_mapping[old_style_id]
                        rstyle_elem.set(f'{{{self.NS["w"]}}}val', new_style_id)
                        updated_count += 1
                        logger.debug(f"更新字符样式引用: {old_style_id} -> {new_style_id}")
                
                # 更新表格样式引用
                for tblstyle_elem in document_xml_root.xpath('//w:tblStyle', namespaces=self.NS):
                    old_style_id = tblstyle_elem.get(f'{{{self.NS["w"]}}}val')
                    if old_style_id in style_mapping:
                        new_style_id = style_mapping[old_style_id]
                        tblstyle_elem.set(f'{{{self.NS["w"]}}}val', new_style_id)
                        updated_count += 1
                        logger.debug(f"更新表格样式引用: {old_style_id} -> {new_style_id}")
                
                # 保存更新后的document.xml
                with open(document_path, 'wb') as f:
                    f.write(etree.tostring(document_xml_root, pretty_print=True, encoding='utf-8', xml_declaration=True))
                
                # 重新打包为新的Word文件
                with ZipFile(self.output_path, 'w') as output_zip:
                    # 遍历临时目录中的所有文件，添加到输出zip
                    for root, dirs, files in os.walk(temp_dir):
                        for file in files:
                            file_path = os.path.join(root, file)
                            # 计算在zip中的相对路径
                            zip_path = os.path.relpath(file_path, temp_dir)
                            output_zip.write(file_path, zip_path)
                
                logger.info(f"成功更新 {updated_count} 个样式引用")
                return True
                
        except Exception as e:
            logger.error(f"更新文档内容样式引用失败: {str(e)}", exc_info=True)
            return False
    
    def get_template_styles(self):
        """
        获取模板文件的所有样式
        
        Returns:
            list: Style对象列表
        """
        logger.info(f"开始获取模板文件样式: {self.template_path}")
        
        try:
            # 提取模板样式XML
            template_styles_xml_root, template_style_id_mapping = self.extract_template_styles()
            if template_styles_xml_root is None:
                return []
            
            styles = []
            
            # 处理每个样式元素
            for style_elem in template_styles_xml_root.xpath('//w:style', namespaces=self.NS):
                style_id = style_elem.get(f'{{{self.NS["w"]}}}styleId')
                style_name_elem = style_elem.find('.//w:name', namespaces=self.NS)
                
                if style_name_elem is not None:
                    style_name = style_name_elem.get(f'{{{self.NS["w"]}}}val')
                    
                    # 确定样式类型
                    style_type = style_elem.get(f'{{{self.NS["w"]}}}type', 'paragraph')
                    if style_type == 'paragraph':
                        style_type = 'paragraph'
                    elif style_type == 'character':
                        style_type = 'character'
                    elif style_type == 'table':
                        style_type = 'table'
                    else:
                        style_type = 'paragraph'  # 默认
                    
                    # 创建Style对象
                    style = Style(style_id, style_name, style_type)
                    
                    # 解析样式属性
                    self._parse_style_attributes(style, style_elem)
                    
                    styles.append(style)
            
            logger.info(f"成功获取模板文件样式，共 {len(styles)} 个")
            return styles
            
        except Exception as e:
            logger.error(f"获取模板文件样式失败: {str(e)}", exc_info=True)
            return []
    
    def get_target_styles(self):
        """
        获取目标文件的所有样式
        
        Returns:
            list: Style对象列表
        """
        logger.info(f"开始获取目标文件样式: {self.target_path}")
        
        try:
            # 分析目标文档
            target_styles_xml_root, used_style_ids = self.analyze_target_document()
            if target_styles_xml_root is None:
                return []
            
            styles = []
            
            # 处理每个样式元素
            for style_elem in target_styles_xml_root.xpath('//w:style', namespaces=self.NS):
                style_id = style_elem.get(f'{{{self.NS["w"]}}}styleId')
                style_name_elem = style_elem.find('.//w:name', namespaces=self.NS)
                
                if style_name_elem is not None:
                    style_name = style_name_elem.get(f'{{{self.NS["w"]}}}val')
                    
                    # 确定样式类型
                    style_type = style_elem.get(f'{{{self.NS["w"]}}}type', 'paragraph')
                    if style_type == 'paragraph':
                        style_type = 'paragraph'
                    elif style_type == 'character':
                        style_type = 'character'
                    elif style_type == 'table':
                        style_type = 'table'
                    else:
                        style_type = 'paragraph'  # 默认
                    
                    # 创建Style对象
                    style = Style(style_id, style_name, style_type)
                    
                    # 解析样式属性
                    self._parse_style_attributes(style, style_elem)
                    
                    styles.append(style)
            
            logger.info(f"成功获取目标文件样式，共 {len(styles)} 个")
            return styles
            
        except Exception as e:
            logger.error(f"获取目标文件样式失败: {str(e)}", exc_info=True)
            return []
    
    def get_styles_by_string(self, search_string):
        """
        根据字符串查询所在段落、标题、表格或列表的当前样式
        
        Args:
            search_string (str): 要查询的字符串
            
        Returns:
            list: 包含Style对象和位置信息的字典列表
        """
        logger.info(f"开始根据字符串查询样式: '{search_string}'")
        
        try:
            # 打开目标文件
            with ZipFile(self.target_path, 'r') as target_zip:
                # 读取document.xml
                document_xml_content = target_zip.read('word/document.xml')
                document_xml_root = etree.fromstring(document_xml_content)
                
                results = []
                
                # 1. 搜索段落
                logger.debug("搜索段落...")
                paragraphs = document_xml_root.xpath('//w:body//w:p', namespaces=self.NS)
                
                for para_idx, paragraph in enumerate(paragraphs):
                    # 获取段落完整文本
                    para_text = ""
                    for text_elem in paragraph.xpath('.//w:t', namespaces=self.NS):
                        if text_elem.text:
                            para_text += text_elem.text
                            logger.debug(f"段落 {para_idx} 文本: {para_text}")
                    
                    if search_string in para_text:
                        # 获取段落样式
                        pstyle_elem = paragraph.find('.//w:pStyle', namespaces=self.NS)
                        if pstyle_elem is not None:
                            style_id = pstyle_elem.get(f'{{{self.NS["w"]}}}val')
                            
                            # 获取样式名称
                            target_styles_xml_content = target_zip.read('word/styles.xml')
                            target_styles_xml_root = etree.fromstring(target_styles_xml_content)
                            style_elem = target_styles_xml_root.xpath(f'//w:style[@w:styleId="{style_id}"]', namespaces=self.NS)
                            
                            if style_elem:
                                style_name_elem = style_elem[0].find('.//w:name', namespaces=self.NS)
                                style_name = style_name_elem.get(f'{{{self.NS["w"]}}}val') if style_name_elem is not None else style_id
                                
                                # 创建Style对象
                                style = Style(style_id, style_name, 'paragraph')
                                self._parse_style_attributes(style, style_elem[0])
                            else:
                                # 没有找到对应的样式元素，创建一个默认样式
                                style = Style(style_id, f"未命名样式_{style_id}", 'paragraph')
                        else:
                            # 没有样式，创建一个默认样式
                            style = Style('default', '默认样式', 'paragraph')
                        
                        # 添加到结果中
                        results.append({
                            'type': 'paragraph',
                            'index': para_idx,
                            'text': para_text,
                            'style': style
                        })
                
                # 2. 搜索表格
                logger.debug("搜索表格...")
                tables = document_xml_root.xpath('//w:body//w:tbl', namespaces=self.NS)
                
                for table_idx, table in enumerate(tables):
                    # 检查表格是否包含搜索字符串
                    table_text = ""
                    for text_elem in table.xpath('.//w:t', namespaces=self.NS):
                        if text_elem.text:
                            table_text += text_elem.text
                    
                    if search_string in table_text:
                        # 获取表格样式
                        tblstyle_elem = table.find('.//w:tblStyle', namespaces=self.NS)
                        if tblstyle_elem is not None:
                            style_id = tblstyle_elem.get(f'{{{self.NS["w"]}}}val')
                            
                            # 获取样式名称
                            target_styles_xml_content = target_zip.read('word/styles.xml')
                            target_styles_xml_root = etree.fromstring(target_styles_xml_content)
                            style_elem = target_styles_xml_root.xpath(f'//w:style[@w:styleId="{style_id}"]', namespaces=self.NS)
                            
                            if style_elem:
                                style_name_elem = style_elem[0].find('.//w:name', namespaces=self.NS)
                                style_name = style_name_elem.get(f'{{{self.NS["w"]}}}val') if style_name_elem is not None else style_id
                                
                                # 创建Style对象
                                style = Style(style_id, style_name, 'table')
                                self._parse_style_attributes(style, style_elem[0])
                            else:
                                # 没有找到对应的样式元素，创建一个默认样式
                                style = Style(style_id, f"未命名表格样式_{style_id}", 'table')
                        else:
                            # 没有样式，创建一个默认样式
                            style = Style('default_table', '默认表格样式', 'table')
                        
                        results.append({
                            'type': 'table',
                            'index': table_idx,
                            'text': table_text,
                            'style': style
                        })
                
                # 3. 检查列表项
                logger.debug("检查列表项...")
                for para_idx, paragraph in enumerate(paragraphs):
                    # 检查是否为列表项
                    num_pr = paragraph.find('.//w:pPr//w:numPr', namespaces=self.NS)
                    if num_pr is not None:
                        # 获取段落文本
                        text = ""
                        for text_elem in paragraph.xpath('.//w:t', namespaces=self.NS):
                            if text_elem.text:
                                text += text_elem.text
                        
                        if search_string in text:
                            # 获取段落样式
                            pstyle_elem = paragraph.find('.//w:pStyle', namespaces=self.NS)
                            if pstyle_elem is not None:
                                style_id = pstyle_elem.get(f'{{{self.NS["w"]}}}val')
                                
                                # 获取样式名称和详细属性
                                target_styles_xml_content = target_zip.read('word/styles.xml')
                                target_styles_xml_root = etree.fromstring(target_styles_xml_content)
                                style_elem = target_styles_xml_root.xpath(f'//w:style[@w:styleId="{style_id}"]', namespaces=self.NS)
                                
                                if style_elem:
                                    style_name_elem = style_elem[0].find('.//w:name', namespaces=self.NS)
                                    style_name = style_name_elem.get(f'{{{self.NS["w"]}}}val') if style_name_elem is not None else style_id
                                    
                                    # 创建Style对象并解析属性
                                    style = Style(style_id, style_name, 'list')
                                    self._parse_style_attributes(style, style_elem[0])
                                else:
                                    # 没有找到对应的样式元素，创建一个默认样式
                                    style = Style(style_id, f"未命名列表样式_{style_id}", 'list')
                            else:
                                # 没有样式，创建一个默认样式
                                style = Style('default_list', '默认列表样式', 'list')
                            
                            results.append({
                                'type': 'list',
                                'index': para_idx,
                                'text': text,
                                'style': style
                            })
                
                logger.info(f"成功找到 {len(results)} 个包含字符串的元素")
                return results
                
        except Exception as e:
            logger.error(f"根据字符串查询样式失败: {str(e)}", exc_info=True)
            return []
    
    def _parse_style_attributes(self, style, style_elem):
        """
        解析样式元素的属性并填充到Style对象
        
        Args:
            style (Style): Style对象
            style_elem (Element): 样式XML元素
        """
        # 解析字体属性
        font_elem = style_elem.find('.//w:rPr', namespaces=self.NS)
        if font_elem is not None:
            # 字体名称
            font_name = font_elem.find('.//w:rFonts', namespaces=self.NS)
            if font_name is not None:
                style.font['name'] = font_name.get(f'{{{self.NS["w"]}}}ascii')
            
            # 字体大小
            font_size = font_elem.find('.//w:sz', namespaces=self.NS)
            if font_size is not None:
                style.font['size'] = font_size.get(f'{{{self.NS["w"]}}}val')
            
            # 粗体
            bold = font_elem.find('.//w:b', namespaces=self.NS)
            style.font['bold'] = bold is not None
            
            # 斜体
            italic = font_elem.find('.//w:i', namespaces=self.NS)
            style.font['italic'] = italic is not None
            
            # 下划线
            underline = font_elem.find('.//w:u', namespaces=self.NS)
            style.font['underline'] = underline is not None
            
            # 颜色
            color = font_elem.find('.//w:color', namespaces=self.NS)
            if color is not None:
                style.font['color'] = color.get(f'{{{self.NS["w"]}}}val')
        
        # 解析段落属性
        para_elem = style_elem.find('.//w:pPr', namespaces=self.NS)
        if para_elem is not None:
            # 对齐方式
            alignment = para_elem.find('.//w:jc', namespaces=self.NS)
            if alignment is not None:
                style.paragraph['alignment'] = alignment.get(f'{{{self.NS["w"]}}}val')
            
            # 行间距
            line_spacing = para_elem.find('.//w:spacing', namespaces=self.NS)
            if line_spacing is not None:
                style.paragraph['line_spacing'] = line_spacing.get(f'{{{self.NS["w"]}}}line')
                style.paragraph['space_before'] = line_spacing.get(f'{{{self.NS["w"]}}}before')
                style.paragraph['space_after'] = line_spacing.get(f'{{{self.NS["w"]}}}after')
            
            # 缩进
            indent = para_elem.find('.//w:ind', namespaces=self.NS)
            if indent is not None:
                style.paragraph['indent_left'] = indent.get(f'{{{self.NS["w"]}}}left')
                style.paragraph['indent_right'] = indent.get(f'{{{self.NS["w"]}}}right')
                style.paragraph['first_line_indent'] = indent.get(f'{{{self.NS["w"]}}}firstLine')
        
        # 解析表格属性
        table_elem = style_elem.find('.//w:tblPr', namespaces=self.NS)
        if table_elem is not None:
            # 边框
            borders = table_elem.find('.//w:tblBorders', namespaces=self.NS)
            if borders is not None:
                style.table['border_top'] = borders.find('.//w:top', namespaces=self.NS)
                style.table['border_bottom'] = borders.find('.//w:bottom', namespaces=self.NS)
                style.table['border_left'] = borders.find('.//w:left', namespaces=self.NS)
                style.table['border_right'] = borders.find('.//w:right', namespaces=self.NS)
        
        # 解析列表属性
        list_elem = style_elem.find('.//w:pPr//w:numPr', namespaces=self.NS)
        if list_elem is not None:
            style.style_type = 'list'
            # 列表格式
            num_id = list_elem.find('.//w:numId', namespaces=self.NS)
            if num_id is not None:
                style.list['format'] = num_id.get(f'{{{self.NS["w"]}}}val')
            
            # 列表级别
            ilvl = list_elem.find('.//w:ilvl', namespaces=self.NS)
            if ilvl is not None:
                style.list['level'] = ilvl.get(f'{{{self.NS["w"]}}}val')
    
    def print_styles(self, styles, style_type=None):
        """
        打印样式信息
        
        Args:
            styles (list): Style对象列表
            style_type (str): 样式类型过滤
        """
        if style_type:
            filtered_styles = [s for s in styles if s.style_type == style_type]
            logger.info(f"打印{style_type}样式，共 {len(filtered_styles)} 个")
        else:
            filtered_styles = styles
            logger.info(f"打印所有样式，共 {len(filtered_styles)} 个")
        
        for style in filtered_styles:
            logger.info(f"\n样式ID: {style.style_id}")
            logger.info(f"样式名称: {style.style_name}")
            logger.info(f"样式类型: {style.style_type}")
            logger.info(f"字体: {style.font}")
            logger.info(f"段落属性: {style.paragraph}")
            if style.style_type == 'table':
                logger.info(f"表格属性: {style.table}")
            elif style.style_type == 'list':
                logger.info(f"列表属性: {style.list}")
    
    def process(self):
        """
        完整的样式处理流程
        
        Returns:
            bool: 是否处理成功
        """
        logger.info("开始Word文件样式处理流程")
        
        try:
            # 1. 提取模板样式
            template_styles_xml_root, template_style_id_mapping = self.extract_template_styles()
            if template_styles_xml_root is None:
                return False
            
            # 2. 分析目标文档
            target_styles_xml_root, used_style_ids = self.analyze_target_document()
            if target_styles_xml_root is None:
                return False
            
            # 3. 建立样式映射
            style_mapping = self.build_style_mapping(template_style_id_mapping, target_styles_xml_root)
            if not style_mapping:
                logger.warning("未建立任何样式映射，可能需要手动配置")
            
            # 4. 更新样式定义库
            if not self.update_styles_definition(template_styles_xml_root, self.target_path):
                return False
            
            # 5. 更新内容引用
            if not self.update_content_references(style_mapping):
                return False
            
            logger.info("Word文件样式处理流程完成")
            return True
            
        except Exception as e:
            logger.error(f"样式处理流程失败: {str(e)}", exc_info=True)
            return False
