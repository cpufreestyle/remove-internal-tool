# 全盘文档批量去除首行'内部'工具 v3.0

递归扫描全盘（所有磁盘）下的文档和图片文件，将每个文件首行中的"内部"两个字删除。

## 支持格式

### 文档格式
- Word (.docx)
- Excel (.xlsx)
- PDF (.pdf)
- PowerPoint (.ppt/.pptx)
- 文本文件 (.txt)
- CSV表格 (.csv)
- Markdown (.md)
- RTF富文本 (.rtf)

### 图片格式（需OCR）
- JPEG (.jpg/.jpeg)
- PNG (.png)
- BMP (.bmp)
- TIFF (.tiff/.tif)
- WebP (.webp)

## 使用方法

### 方式一：直接运行EXE
1. 下载 `RemoveInternal.exe`
2. 右键以管理员身份运行
3. 等待扫描完成
4. 结果保存在 `changed_files.csv`

### 方式二：Python源码运行
```bash
pip install -r requirements.txt
python main.py
```

### 指定目录
```bash
python main.py --dir C:/Users/michael/Documents
```

## 注意事项
- 需以管理员权限运行
- 建议先备份重要文件
- 自动跳过系统目录和网盘目录
- 图片处理需要安装 PaddleOCR

## 输出
- `changed_files.csv` - 被修改的文件列表
