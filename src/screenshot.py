import mss
import mss.tools
import os
from PIL import Image
import io


class Screenshot:
    def __init__(self):
        self.sct = mss.mss()
    
    def capture_fullscreen(self):
        """捕获全屏"""
        # 获取主显示器的尺寸
        monitor = self.sct.monitors[1]  # 1表示主显示器
        
        # 捕获屏幕
        screenshot = self.sct.grab(monitor)
        
        # 转换为PIL Image对象
        img = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")
        
        # 转换为字节流
        img_byte_arr = io.BytesIO()
        img.save(img_byte_arr, format='PNG')
        img_byte_arr = img_byte_arr.getvalue()
        
        return img_byte_arr
    
    def capture_region(self, region):
        """捕获指定区域
        region格式：(left, top, width, height)
        """
        left, top, width, height = region
        
        # 构建monitor字典
        monitor = {
            "left": left,
            "top": top,
            "width": width,
            "height": height
        }
        
        # 捕获屏幕
        screenshot = self.sct.grab(monitor)
        
        # 转换为PIL Image对象
        img = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")
        
        # 转换为字节流
        img_byte_arr = io.BytesIO()
        img.save(img_byte_arr, format='PNG')
        img_byte_arr = img_byte_arr.getvalue()
        
        return img_byte_arr
    
    def save_screenshot(self, img_bytes, file_path):
        """保存截图到文件"""
        img = Image.open(io.BytesIO(img_bytes))
        img.save(file_path)

if __name__ == "__main__":
    # 测试截图功能
    screenshot = Screenshot()
    img_bytes = screenshot.capture_fullscreen()
    screenshot.save_screenshot(img_bytes, "test_screenshot.png")
    print("截图已保存到test_screenshot.png")
