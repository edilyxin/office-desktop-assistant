import requests
import base64
from .config_manager import load_config

class PaddleOCRVL:
    def __init__(self):
        config = load_config()
        self.api_url = config['api_url']
        self.token = config['token']
    
    def _encode_file(self, file_path):
        """将文件编码为base64格式"""
        with open(file_path, 'rb') as file:
            file_bytes = file.read()
            file_data = base64.b64encode(file_bytes).decode('ascii')
        return file_data
    
    def _get_file_type(self, file_path):
        """根据文件扩展名获取文件类型"""
        import os
        ext = os.path.splitext(file_path)[1].lower()
        if ext in ['.pdf']:
            return 0  # PDF文件
        else:
            return 1  # 图片文件
    
    def recognize(self, file_path, use_doc_orientation_classify=False, 
                  use_doc_unwarping=False, use_chart_recognition=False,
                  prettifyMarkdown=True, visualize=True):
        """调用PaddleOCR-vl API进行内容识别"""
        # 编码文件
        file_data = self._encode_file(file_path)
        file_type = self._get_file_type(file_path)
        
        # 构建请求头
        headers = {
            "Authorization": f"token {self.token}",
            "Content-Type": "application/json"
        }
        
        # 构建请求体
        payload = {
            "file": file_data,
            "fileType": file_type,
            "useDocOrientationClassify": use_doc_orientation_classify,
            "useDocUnwarping": use_doc_unwarping,
            "useChartRecognition": use_chart_recognition,
            "prettifyMarkdown": prettifyMarkdown,
            "visualize": visualize
        }
        
        # 发送请求
        response = requests.post(self.api_url, json=payload, headers=headers)
        response.raise_for_status()  # 抛出HTTP错误
        
        # 返回结果
        return response.json()['result']
    
    def recognize_image_from_bytes(self, image_bytes, use_doc_orientation_classify=False, 
                                  use_doc_unwarping=False, use_chart_recognition=False,
                                  prettifyMarkdown=True, visualize=True):
        """从字节流识别图片内容"""
        # 编码图片字节
        file_data = base64.b64encode(image_bytes).decode('ascii')
        
        # 构建请求头
        headers = {
            "Authorization": f"token {self.token}",
            "Content-Type": "application/json"
        }
        
        # 构建请求体
        payload = {
            "file": file_data,
            "fileType": 1,  # 图片文件
            "useDocOrientationClassify": use_doc_orientation_classify,
            "useDocUnwarping": use_doc_unwarping,
            "useChartRecognition": use_chart_recognition,
            "prettifyMarkdown": prettifyMarkdown,
            "visualize": visualize
        }
        
        # 发送请求
        response = requests.post(self.api_url, json=payload, headers=headers)
        response.raise_for_status()  # 抛出HTTP错误
        
        # 返回结果
        return response.json()['result']
