# 办公桌面助手 / Office Desktop Assistant

[![GitHub license](https://img.shields.io/github/license/edilyxin/office-desktop-assistant)](https://github.com/edilyxin/office-desktop-assistant/blob/main/LICENSE)
[![GitHub stars](https://img.shields.io/github/stars/edilyxin/office-desktop-assistant)](https://github.com/edilyxin/office-desktop-assistant/stargazers)
[![GitHub forks](https://img.shields.io/github/forks/edilyxin/office-desktop-assistant)](https://github.com/edilyxin/office-desktop-assistant/network)
[![GitHub issues](https://img.shields.io/github/issues/edilyxin/office-desktop-assistant)](https://github.com/edilyxin/office-desktop-assistant/issues)

## 语言切换 / Language Switch

- [中文版本 (Chinese Version)](#中文版本)
- [English Version](#english-version)

---

## 中文版本

### 项目概述
办公桌面助手是一个基于PaddleOCR-vl API开发的桌面OCR工具，支持截图识别和文件上传识别，识别结果可复制到剪贴板或保存为Markdown文件。

### 功能特性

#### 📸 截图识别
- 支持全屏截图和鼠标选择区域截图
- 截图后实时显示在界面右侧
- 自动识别图片中的文字和区域

#### 📁 文件上传
- 支持多种文件格式：png、jpg、jpeg、bmp、gif、tiff、pdf
- 上传后先显示图片，再进行识别
- 支持大文件上传

#### 📝 识别结果处理
- 识别结果以Markdown格式渲染显示
- 支持一键复制到剪贴板
- 支持保存为Markdown文件
- 识别结果中的图片自动保存到本地imgs文件夹

#### 📊 图片显示
- 右侧显示识别的图片和识别到的区域
- 支持图片等比例缩放
- 显示识别区域的边界框

#### 📄 Word文件样式调整 (v0.6开发中)
- 支持选择Word样式模板文件
- 支持上传目标Word文件
- 基于模板自动调整目标文件样式
- 支持多级标题、正文、行间距、段落、表格等样式调整

#### 📋 详细日志
- 关键动作记录到日志文件
- Debug信息输出到控制台
- 模板分析和文件处理的详细日志

### 技术栈

| 技术 | 版本 | 用途 |
|------|------|------|
| Python | 3.8+ | 开发语言 |
| PyQt5 | 5.15+ | GUI框架 |
| PaddleOCR-vl API | - | OCR识别 |
| requests | 2.31+ | HTTP请求 |
| pillow | 10.0+ | 图像处理 |
| pyperclip | 1.8+ | 剪贴板操作 |
| mss | 9.0+ | 屏幕截图 |
| pyyaml | 6.0+ | 配置文件管理 |
| pyqtwebengine | 5.15+ | Markdown渲染 |
| python-docx | 0.8+ | Word文档处理 |

### 安装说明

#### 1. 克隆仓库
```bash
git clone https://github.com/edilyxin/office-desktop-assistant.git
cd office-desktop-assistant
```

#### 2. 创建虚拟环境
```bash
# Windows
python -m venv venv
venv\Scripts\activate

# macOS/Linux
python3 -m venv venv
source venv/bin/activate
```

#### 3. 安装依赖
```bash
pip install -r requirements.txt
```

#### 4. 配置API密钥
1. 复制配置文件模板：
   ```bash
   cp config/config.yaml.example config/config.yaml
   ```
2. 编辑`config/config.yaml`文件，填入你的PaddleOCR-vl API密钥

### 使用方法

#### 1. 启动程序
```bash
python run.py
```

#### 2. 截图识别
1. 点击"截图识别"按钮
2. 用鼠标拖动选择截图区域
3. 松开鼠标后自动开始识别
4. 在右侧查看识别结果

#### 3. 文件上传识别
1. 点击"上传文件"按钮
2. 选择要识别的图片或PDF文件
3. 等待文件上传和识别完成
4. 在右侧查看识别结果

#### 4. 处理识别结果
- 点击"复制到剪贴板"按钮，将识别结果复制到剪贴板
- 点击"保存为Markdown"按钮，将识别结果保存为Markdown文件

#### 5. Word文件样式调整
1. 切换到"文件样式"标签页
2. 点击"选择模板"按钮，选择Word样式模板文件
3. 点击"上传文件"按钮，选择目标Word文件
4. 点击"开始处理"按钮，等待处理完成
5. 处理后的文件将保存到output目录

### 项目结构

```
office-desktop-assistant/
├── src/                  # 源代码目录
│   ├── main.py           # 主程序入口
│   ├── ocr_api.py        # OCR API调用模块
│   ├── screenshot.py     # 截图功能模块
│   ├── config_manager.py # 配置管理模块
│   ├── file_utils.py     # 文件处理模块
│   ├── clipboard_manager.py # 剪贴板管理模块
│   ├── markdown_manager.py # Markdown处理模块
│   └── log_manager.py    # 日志管理模块
├── config/               # 配置文件目录
│   ├── config.yaml.example # 配置文件模板
│   └── config.yaml       # 配置文件（需自行创建）
├── logs/                 # 日志文件目录
├── imgs/                 # 识别结果图片保存目录
├── output/               # 处理后的文件输出目录
├── template/             # Word模板文件目录
├── venv/                 # 虚拟环境目录
├── requirements.txt      # 依赖包列表
├── run.py                # 程序启动脚本
├── FEATURES.md           # 功能开发记录
├── LICENSE               # 许可证文件
└── README.md             # 项目说明文档
```

### 配置说明

配置文件`config/config.yaml`包含以下参数：

```yaml
api_key: your_api_key_here
api_url: https://paddleocr-vl-api.example.com
use_doc_orientation_classify: true
use_doc_unwarping: true
use_chart_recognition: true
prettifyMarkdown: true
visualize: true
```

### 日志说明

- 日志文件保存在`logs/`目录下
- 日志级别：DEBUG、INFO、WARNING、ERROR、CRITICAL
- 控制台输出DEBUG级别以上的日志
- 日志文件记录INFO级别以上的日志

### 版本历史

| 版本 | 日期 | 更新内容 |
|------|------|----------|
| v0.5 | 2025-12-01 | 初始版本，包含所有基本功能 |
| v0.6 | 开发中 | 添加Word文件格式自动调整功能 |

### 后续计划

| 功能名称 | 描述 | 优先级 |
|---------|------|--------|
| 区域截图优化 | 优化区域截图的用户体验 | 中 |
| 快捷键支持 | 添加常用操作的快捷键 | 中 |
| 多语言支持 | 支持多种语言的识别和界面 | 低 |
| 批量处理 | 支持批量文件处理 | 高 |
| 结果预览 | 添加识别结果的预览功能 | 中 |

### 贡献指南

欢迎对项目进行贡献，包括：
- 修复bug
- 添加新功能
- 优化代码
- 改进文档

### 联系方式

如有问题或建议，请联系项目维护者：
- GitHub Issues：https://github.com/edilyxin/office-desktop-assistant/issues

---

[切换到英文版本](#english-version)

---

## English Version

### Project Overview
Office Desktop Assistant is a desktop OCR tool developed based on PaddleOCR-vl API, supporting screenshot recognition and file upload recognition. Recognition results can be copied to clipboard or saved as Markdown files.

### Features

#### 📸 Screenshot Recognition
- Support full-screen screenshot and mouse-selectable area screenshot
- Real-time display of screenshots on the right side of the interface
- Automatic recognition of text and regions in images

#### 📁 File Upload
- Support multiple file formats: png, jpg, jpeg, bmp, gif, tiff, pdf
- Display images first after upload, then perform recognition
- Support large file uploads

#### 📝 Recognition Result Processing
- Recognition results rendered and displayed in Markdown format
- One-click copy to clipboard
- Support saving as Markdown files
- Images in recognition results are automatically saved to local imgs folder

#### 📊 Image Display
- Display recognized images and recognized regions on the right side
- Support image proportional scaling
- Display bounding boxes of recognized regions

#### 📄 Word File Style Adjustment (v0.6 in development)
- Support selecting Word style template files
- Support uploading target Word files
- Automatically adjust target file styles based on templates
- Support style adjustment for multi-level headings, body text, line spacing, paragraphs, tables, etc.

#### 📋 Detailed Logging
- Key actions recorded to log files
- Debug information output to console
- Detailed logs for template analysis and file processing

### Technology Stack

| Technology | Version | Purpose |
|------------|---------|---------|
| Python | 3.8+ | Development language |
| PyQt5 | 5.15+ | GUI framework |
| PaddleOCR-vl API | - | OCR recognition |
| requests | 2.31+ | HTTP requests |
| pillow | 10.0+ | Image processing |
| pyperclip | 1.8+ | Clipboard operations |
| mss | 9.0+ | Screen capture |
| pyyaml | 6.0+ | Configuration file management |
| pyqtwebengine | 5.15+ | Markdown rendering |
| python-docx | 0.8+ | Word document processing |

### Installation Instructions

#### 1. Clone the Repository
```bash
git clone https://github.com/edilyxin/office-desktop-assistant.git
cd office-desktop-assistant
```

#### 2. Create Virtual Environment
```bash
# Windows
python -m venv venv
venv\Scripts\activate

# macOS/Linux
python3 -m venv venv
source venv/bin/activate
```

#### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

#### 4. Configure API Key
1. Copy the configuration file template:
   ```bash
   cp config/config.yaml.example config/config.yaml
   ```
2. Edit `config/config.yaml` file and fill in your PaddleOCR-vl API key

### Usage

#### 1. Start the Program
```bash
python run.py
```

#### 2. Screenshot Recognition
1. Click the "Screenshot Recognition" button
2. Drag with the mouse to select the screenshot area
3. Automatic recognition starts after releasing the mouse
4. View recognition results on the right side

#### 3. File Upload Recognition
1. Click the "Upload File" button
2. Select the image or PDF file to be recognized
3. Wait for file upload and recognition to complete
4. View recognition results on the right side

#### 4. Process Recognition Results
- Click "Copy to Clipboard" button to copy recognition results to clipboard
- Click "Save as Markdown" button to save recognition results as Markdown files

#### 5. Word File Style Adjustment
1. Switch to the "File Style" tab
2. Click "Select Template" button to select a Word style template file
3. Click "Upload File" button to select the target Word file
4. Click "Start Processing" button and wait for processing to complete
5. Processed files will be saved to the output directory

### Project Structure

```
office-desktop-assistant/
├── src/                  # Source code directory
│   ├── main.py           # Main program entry
│   ├── ocr_api.py        # OCR API call module
│   ├── screenshot.py     # Screenshot function module
│   ├── config_manager.py # Configuration management module
│   ├── file_utils.py     # File processing module
│   ├── clipboard_manager.py # Clipboard management module
│   ├── markdown_manager.py # Markdown processing module
│   └── log_manager.py    # Log management module
├── config/               # Configuration file directory
│   ├── config.yaml.example # Configuration file template
│   └── config.yaml       # Configuration file (need to create by yourself)
├── logs/                 # Log file directory
├── imgs/                 # Recognition result image save directory
├── output/               # Processed file output directory
├── template/             # Word template file directory
├── venv/                 # Virtual environment directory
├── requirements.txt      # Dependency package list
├── run.py                # Program startup script
├── FEATURES.md           # Feature development record
├── LICENSE               # License file
└── README.md             # Project description document
```

### Configuration Description

The configuration file `config/config.yaml` contains the following parameters:

```yaml
api_key: your_api_key_here
api_url: https://paddleocr-vl-api.example.com
use_doc_orientation_classify: true
use_doc_unwarping: true
use_chart_recognition: true
prettifyMarkdown: true
visualize: true
```

### Log Description

- Log files are saved in the `logs/` directory
- Log levels: DEBUG, INFO, WARNING, ERROR, CRITICAL
- Console outputs logs above DEBUG level
- Log files record logs above INFO level

### Version History

| Version | Date | Update Content |
|---------|------|----------------|
| v0.5 | 2025-12-01 | Initial version with all basic features |
| v0.6 | In development | Add Word file format automatic adjustment feature |

### Future Plans

| Feature Name | Description | Priority |
|--------------|-------------|----------|
| Area Screenshot Optimization | Optimize user experience of area screenshot | Medium |
| Shortcut Support | Add shortcuts for common operations | Medium |
| Multi-language Support | Support recognition and interface in multiple languages | Low |
| Batch Processing | Support batch file processing | High |
| Result Preview | Add preview function for recognition results | Medium |

### Contribution Guidelines

Welcome to contribute to the project, including:
- Fixing bugs
- Adding new features
- Optimizing code
- Improving documentation

### Contact Information

For questions or suggestions, please contact the project maintainer:
- GitHub Issues: https://github.com/edilyxin/office-desktop-assistant/issues

---

[Switch to Chinese Version](#中文版本)