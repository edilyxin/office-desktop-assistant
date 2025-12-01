import os
import mimetypes

class FileUtils:
    @staticmethod
    def is_supported_file(file_path):
        """检查文件是否为支持的格式"""
        supported_formats = [
            '.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff', '.pdf'
        ]
        ext = os.path.splitext(file_path)[1].lower()
        return ext in supported_formats
    
    @staticmethod
    def get_file_size(file_path):
        """获取文件大小，单位为字节"""
        return os.path.getsize(file_path)
    
    @staticmethod
    def get_file_mime_type(file_path):
        """获取文件的MIME类型"""
        mime_type, _ = mimetypes.guess_type(file_path)
        return mime_type
    
    @staticmethod
    def read_file_bytes(file_path):
        """读取文件内容为字节流"""
        with open(file_path, 'rb') as f:
            return f.read()
    
    @staticmethod
    def get_file_name(file_path):
        """获取文件名（不含路径）"""
        return os.path.basename(file_path)
    
    @staticmethod
    def get_file_name_without_ext(file_path):
        """获取不含扩展名的文件名"""
        return os.path.splitext(os.path.basename(file_path))[0]

if __name__ == "__main__":
    # 测试文件工具功能
    file_utils = FileUtils()
    test_file = "test_screenshot.png"
    if os.path.exists(test_file):
        print(f"文件支持: {file_utils.is_supported_file(test_file)}")
        print(f"文件大小: {file_utils.get_file_size(test_file)} 字节")
        print(f"MIME类型: {file_utils.get_file_mime_type(test_file)}")
        print(f"文件名: {file_utils.get_file_name(test_file)}")
        print(f"无扩展名文件名: {file_utils.get_file_name_without_ext(test_file)}")
