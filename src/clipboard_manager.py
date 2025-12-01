import pyperclip

class ClipboardManager:
    @staticmethod
    def copy_to_clipboard(text):
        """将文本复制到剪贴板"""
        pyperclip.copy(text)
    
    @staticmethod
    def paste_from_clipboard():
        """从剪贴板获取文本"""
        return pyperclip.paste()

if __name__ == "__main__":
    # 测试剪贴板功能
    clipboard = ClipboardManager()
    test_text = "测试剪贴板功能"
    clipboard.copy_to_clipboard(test_text)
    pasted_text = clipboard.paste_from_clipboard()
    print(f"复制的文本: {test_text}")
    print(f"粘贴的文本: {pasted_text}")
    print(f"测试结果: {'成功' if test_text == pasted_text else '失败'}")
