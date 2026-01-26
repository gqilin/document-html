# Pandoc文档转换器

基于Flask和Pandoc的文档转换服务，支持多种格式转换为HTML。

## 功能特性

- 支持DOC、DOCX、PDF、EPUB、TXT、RTF、Markdown等格式
- 美观的Web界面，支持拖拽上传
- RESTful API接口
- 实时转换进度显示
- 文件下载功能

## 环境要求

- Python 3.8+
- Pandoc (必须安装)
- Flask

## 安装步骤

1. 安装Python依赖：
```bash
pip install -r requirements.txt
```

2. 安装Pandoc：
访问 https://pandoc.org/installing.html 下载并安装Pandoc

3. 运行服务：
```bash
python src/app.py
```

4. 打开浏览器访问：
http://localhost:5000

## API接口

### POST /api/convert
上传并转换文档

**参数：**
- file: 要转换的文档文件

**响应：**
```json
{
  "success": true,
  "download_url": "/download/filename.html",
  "filename": "filename.html"
}
```

### GET /api/formats
获取支持的文件格式

### GET /api/health
健康检查

## 项目结构

```
pandoc-converter/
├── src/
│   └── app.py                 # Flask应用主文件
├── converters/
│   └── document_converter.py  # 文档转换器
├── templates/
│   └── index.html            # Web界面
├── uploads/                  # 临时上传目录
├── downloads/                # 转换后文件目录
├── requirements.txt          # Python依赖
└── README.md                # 项目说明
```

## 使用说明

1. 启动服务后，在浏览器中打开Web界面
2. 点击选择文件或直接拖拽文件到上传区域
3. 等待转换完成
4. 点击下载按钮获取转换后的HTML文件

## 注意事项

- 确保已正确安装Pandoc
- 文件大小限制为50MB
- 转换后的文件会自动清理
- 支持的文件格式取决于Pandoc的安装配置