#!/usr/bin/env python3
"""PaddleOCR-vl桌面助手入口文件"""

import sys
import os
from src.main import main

# 将项目根目录添加到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
if __name__ == "__main__":
    main()
