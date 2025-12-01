import os
import requests

class MarkdownManager:
    @staticmethod
    def save_markdown(ocr_result, output_dir="output"):
        """保存OCR识别结果为markdown文件"""
        # 创建输出目录
        os.makedirs(output_dir, exist_ok=True)
        
        saved_files = []
        
        # 处理每个页面的结果
        for i, res in enumerate(ocr_result["layoutParsingResults"]):
            # 保存markdown文本
            md_filename = os.path.join(output_dir, f"doc_{i}.md")
            with open(md_filename, "w", encoding="utf-8") as md_file:
                md_file.write(res["markdown"]["text"])
            saved_files.append(md_filename)
            
            # 处理图片
            for img_path, img_url in res["markdown"]["images"].items():
                full_img_path = os.path.join(output_dir, img_path)
                os.makedirs(os.path.dirname(full_img_path), exist_ok=True)
                
                # 下载图片
                img_response = requests.get(img_url)
                if img_response.status_code == 200:
                    with open(full_img_path, "wb") as img_file:
                        img_file.write(img_response.content)
            
            # 处理输出图片
            for img_name, img_url in res["outputImages"].items():
                img_response = requests.get(img_url)
                if img_response.status_code == 200:
                    img_filename = os.path.join(output_dir, f"{img_name}_{i}.jpg")
                    with open(img_filename, "wb") as f:
                        f.write(img_response.content)
                    saved_files.append(img_filename)
        
        return saved_files
    
    @staticmethod
    def extract_markdown_text(ocr_result):
        """从OCR结果中提取markdown文本"""
        markdown_texts = []
        for res in ocr_result["layoutParsingResults"]:
            markdown_texts.append(res["markdown"]["text"])
        return "\n\n".join(markdown_texts)

if __name__ == "__main__":
    # 示例使用
    sample_result = {
        "layoutParsingResults": [
            {
                "markdown": {
                    "text": "# 测试文档\n\n这是一个测试文档。",
                    "images": {}
                },
                "outputImages": {}
            }
        ]
    }
    
    markdown_manager = MarkdownManager()
    saved_files = markdown_manager.save_markdown(sample_result)
    print(f"保存的文件: {saved_files}")
    
    markdown_text = markdown_manager.extract_markdown_text(sample_result)
    print(f"提取的markdown文本: {markdown_text}")
