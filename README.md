# Word文档内容提取系统

## 项目概述

这是一个基于Django框架开发的Web应用系统，用于上传ZIP文件并自动提取其中Word文档的关键内容，包括单号和各类费用信息。系统还会对提取的费用进行自动计算和验证。

## 功能特性

- **多格式文件上传**：支持直接上传ZIP压缩文件或Word文档(.doc, .docx)，可通过拖放方式上传
- **自动提取**：自动解压ZIP文件并提取Word文档中的关键信息
- **费用计算**：自动计算各类费用并进行验证
- **结果展示**：以友好的界面展示提取结果和验证信息
- **响应式设计**：支持在不同设备上使用

## 技术栈

- **后端框架**：Django 5.x
- **文档处理**：python-docx, docx2txt
- **前端技术**：HTML5, CSS3, JavaScript
- **数据库**：SQLite

## 安装说明

### 1. 克隆项目

```bash
git clone git@github.com:minkinoe/Yncmcc_doc_extractor.git
cd wordextractor
```

### 2. 创建并激活虚拟环境

```bash
# Windows
python -m venv venv
venv\Scripts\activate

# macOS/Linux
python3 -m venv venv
source venv/bin/activate
```

### 3. 安装依赖包

```bash
pip install -r requirements.txt
```

### 4. 运行数据库迁移

```bash
python manage.py migrate
```

### 5. 启动开发服务器

```bash
python manage.py runserver
```

## 使用方法

1. 访问 `http://127.0.0.1:8000/`
2. 点击上传区域或直接拖放ZIP文件或Word文档(.doc, .docx)到页面上
3. 系统将自动处理文件并显示提取结果
4. 查看提取的单号、各类费用信息和验证结果
5. 点击"返回上传"按钮可以上传新文件

## 支持的文件格式

- ZIP压缩文件，包含Word文档(.doc, .docx)
- 直接上传的Word文档(.doc, .docx)

## 项目结构

```
wordextractor/
├── wordextractor/      # 项目配置目录
├── uploader/           # 主应用目录
│   ├── templates/      # HTML模板
│   ├── static/         # 静态文件
│   ├── utils.py        # 工具函数
│   ├── views.py        # 视图函数
│   └── urls.py         # URL路由
├── media/              # 媒体文件存储
├── manage.py           # Django管理脚本
└── README.md           # 项目说明文档
```

## 注意事项

1. 请确保上传的ZIP文件中包含有效的Word文档，或上传有效的Word文档(.doc, .docx)
2. 系统会在后台自动清理临时文件，确保磁盘空间合理使用
3. 对于大型文件，处理时间可能会有所延长

## 许可证

本项目仅供内部使用

## 联系方式

如有问题或建议，请联系项目维护人员。