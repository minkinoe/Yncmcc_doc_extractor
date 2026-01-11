# Word文档内容提取系统

## 项目概述

这是一个基于 Django 的 Web 应用，用于上传 ZIP（内含 Word 文档）并自动提取关键信息（单号、费用、光缆信息等），同时对部分费用进行自动计算与验算。

## 功能特性

- **ZIP 上传**：仅支持上传 `.zip`（可点击或拖拽上传）
- **自动提取**：自动解压 ZIP 并处理其中 `.doc/.docx`
- **费用计算与验算**：提取维护费、服务费、终端费等并计算汇总
- **统一面板**：左侧上传与历史记录，右侧展示提取结果
- **便捷复制**：结果页的单号与关键数值支持点击复制

## 技术栈

- **后端框架**：Django 5.x
- **文档处理**：python-docx, docx2txt
- **前端技术**：HTML5, CSS3, JavaScript
- **数据库**：SQLite

## 上传规范（重要）

系统仅接收 ZIP 文件，且 ZIP 文件名通常形如：

`A+B+C.zip`

- `A`：单号（用于结果筛选、展示）
- `B`：集团名称（上传时写入数据库）
- `C`：地址（上传时写入数据库；若地址本身含 `+`，会将后续段落合并为地址）

ZIP 内部需包含 Word 文档（`.doc` 或 `.docx`）。

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

> Windows 环境如需处理 `.doc`，通常需要本机安装 Microsoft Word，并安装 pywin32（项目中使用 win32com 调用 Word 读取文档）。

### 4. 运行数据库迁移

```bash
python manage.py migrate
```

### 5. 启动开发服务器

```bash
python manage.py runserver
```

如需允许局域网其他设备访问（开发环境）：

```bash
python manage.py runserver 0.0.0.0:8000
```

并确保 `ALLOWED_HOSTS` 已放行（开发环境可用 `['*']`），配置见 [settings.py](file:///d:/Dev/django_project/wordextractor/wordextractor/settings.py#L26-L30)。

## 使用方法

1. 访问 `http://127.0.0.1:8000/`
2. 在左侧上传区域点击或拖拽 `.zip` 文件
3. 处理完成后会自动跳转到该次上传详情
4. 点击蓝色单号/数值可复制到剪贴板

## 支持的文件格式

- ZIP 压缩文件（包含 Word 文档 `.doc/.docx`）

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

## 开发要点

### 关键入口

- 主页面与上传处理：`dashboard` 视图  
  - [views.py](file:///d:/Dev/django_project/wordextractor/uploader/views.py#L17-L185)
- ZIP 解压与 Word 提取：  
  - [extract_info_from_zip](file:///d:/Dev/django_project/wordextractor/uploader/utils.py#L495-L671)
- 统一页面模板：  
  - [dashboard.html](file:///d:/Dev/django_project/wordextractor/uploader/templates/uploader/dashboard.html)

### 数据模型

- 上传记录：`UploadedFile`（保存原文件名、文件路径、集团名称、地址等）  
  - [models.py](file:///d:/Dev/django_project/wordextractor/uploader/models.py#L13-L37)
- 提取结果：`ExtractedInfo`（按文档维度存储单号、费用、文本等）  
  - [models.py](file:///d:/Dev/django_project/wordextractor/uploader/models.py#L39-L82)

### 数据库字典

数据库为 SQLite（开发默认文件：`db.sqlite3`）。以下字典以 Django 模型为准，字段类型以 Django Field 表达，并补充关键约束/默认值。

#### 表：`uploader_uploadedfile`（UploadedFile，上传记录）

| 字段 | 类型 | 允许空 | 默认值 | 说明 |
|---|---|---:|---|---|
| id | BigAutoField | 否 | 自增 | 主键 |
| file | FileField | 是 | NULL | 上传文件路径（存于 `media/`） |
| original_filename | CharField(255) | 否 |  | 原始文件名（含扩展名） |
| file_size | IntegerField | 否 | 0 | 文件大小（字节） |
| file_type | CharField(50) | 否 |  | 文件类型（当前业务为 `zip`） |
| group_name | CharField(255) | 是 | NULL | 集团名称（来自 ZIP 文件名 B 段） |
| address | CharField(255) | 是 | NULL | 地址（来自 ZIP 文件名 C 段） |
| uploaded_at | DateTimeField | 否 | auto_now_add | 上传时间 |
| processed_at | DateTimeField | 是 | NULL | 处理完成时间 |
| is_processed | BooleanField | 否 | False | 是否处理完成 |
| processing_error | TextField | 是 | NULL | 处理错误信息（如有） |
| document_count | IntegerField | 否 | 0 | ZIP 内提取到的文档数量 |

关系：

- `UploadedFile (1) -> ExtractedInfo (N)`，反向访问名：`uploaded_file.extracted_infos`

#### 表：`uploader_extractedinfo`（ExtractedInfo，提取结果）

| 字段 | 类型 | 允许空 | 默认值 | 说明 |
|---|---|---:|---|---|
| id | BigAutoField | 否 | 自增 | 主键 |
| uploaded_file_id | ForeignKey | 否 |  | 关联上传记录（级联删除） |
| order_code | CharField(100) | 是 | NULL | 单号（db_index=True） |
| document_name | CharField(255) | 否 | "" | 文档文件名 |
| document_content | TextField | 是 | NULL | 从 Word 提取的完整文本 |
| extraction_status | CharField(20) | 否 | "待处理" | 提取状态 |
| extraction_error | TextField | 是 | NULL | 提取错误信息（如有） |
| maintenance_fee | DecimalField(10,2) | 否 | 0.00 | 宽带维护费 |
| service_fee | DecimalField(10,2) | 否 | 0.00 | 宽带服务费 |
| terminal_fee | DecimalField(10,2) | 否 | 0.00 | 终端费 |
| other_fees | DecimalField(10,2) | 否 | 0.00 | 其他费用 |
| total_fees | DecimalField(10,2) | 否 | 0.00 | 费用合计（计算值） |
| doc_maintenance_total | DecimalField(10,2) | 是 | NULL | 文档中的“维护费合计” |
| overall_total_price | DecimalField(10,2) | 是 | NULL | 文档中的“总体估算价格” |
| total_price | DecimalField(10,2) | 是 | NULL | 文档中的“总估算价格” |
| fiber_info | JSONField | 是 | NULL | 光缆信息列表 |
| equipment_items | JSONField | 是 | NULL | 设备清单 |
| verification_passed | BooleanField | 否 | False | 验算是否通过 |
| verification_message | TextField | 是 | NULL | 验算说明 |
| extracted_at | DateTimeField | 否 | auto_now_add | 记录创建时间 |

索引：

- `order_code`（模型 Meta.indexes + 字段 `db_index=True`）
- `uploaded_file`（模型 Meta.indexes）

### 常用命令

```bash
# 数据库迁移
python manage.py makemigrations
python manage.py migrate

# 配置自检
python manage.py check
```

## 注意事项

1. ZIP 内部 Word 读取依赖运行环境：Windows 下处理 `.doc` 往往需要 Word + pywin32。
2. ZIP 处理使用临时目录解压，处理完成后会清理临时目录（见 [utils.py](file:///d:/Dev/django_project/wordextractor/uploader/utils.py#L666-L670)）。
3. `media/` 存储上传文件，`db.sqlite3` 为开发数据库文件。

## 许可证

本项目仅供内部使用

## 联系方式

如有问题或建议，请联系项目维护人员。
