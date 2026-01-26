import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from flask import Flask, request, jsonify, render_template, send_file
from werkzeug.utils import secure_filename
import tempfile
from pathlib import Path
from converters.document_converter import DocumentConverter

# 配置模板目录为项目根目录下的templates文件夹
template_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'templates')
app = Flask(__name__, template_folder=template_dir)
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200MB max file size

# 确保必要的目录存在（使用项目根目录）
project_root = os.path.dirname(os.path.dirname(__file__))
uploads_dir = os.path.join(project_root, 'uploads')
downloads_dir = os.path.join(project_root, 'downloads')
os.makedirs(uploads_dir, exist_ok=True)
os.makedirs(downloads_dir, exist_ok=True)

converter = DocumentConverter()

@app.route('/')
def index():
    """主页"""
    return render_template('index.html')

@app.route('/api/convert', methods=['POST'])
def convert_document():
    """文档转换API"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file provided'}), 400
        
        file = request.files['file']
        if file.filename is None or file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        # 检查文件类型
        if not converter.is_supported(file.filename):
            return jsonify({
                'error': f'Unsupported file type. Supported formats: {", ".join(converter.get_supported_formats())}'
            }), 400
        
        # 保存上传的文件
        filename = secure_filename(file.filename)
        upload_path = os.path.join(uploads_dir, filename)
        file.save(upload_path)
        
        # 转换文档
        try:
            output_path = converter.convert_to_html(upload_path)
            
            if output_path is None:
                raise RuntimeError("Conversion returned None")
            
            # 生成下载文件名
            base_name = Path(filename).stem
            download_name = f"{base_name}.html"
            download_path = os.path.join(downloads_dir, download_name)
            
            # 移动转换后的文件到下载目录
            if os.path.exists(output_path):
                # 如果目标文件已存在，先删除再移动
                if os.path.exists(download_path):
                    os.remove(download_path)
                os.rename(output_path, download_path)
            
            # 清理上传文件
            os.unlink(upload_path)
            
            return jsonify({
                'success': True,
                'download_url': f'/download/{download_name}',
                'filename': download_name
            })
            
        except Exception as e:
            # 清理上传文件
            if os.path.exists(upload_path):
                os.unlink(upload_path)
            raise e
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>')
def download_file(filename):
    """下载转换后的文件"""
    try:
        file_path = os.path.join(downloads_dir, secure_filename(filename))
        if not os.path.exists(file_path):
            return jsonify({'error': 'File not found'}), 404
        
        return send_file(file_path, as_attachment=True, download_name=filename)
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/formats')
def get_supported_formats():
    """获取支持的文件格式"""
    return jsonify({
        'formats': converter.get_supported_formats(),
        'pandoc_available': converter.pandoc_available,
        'note': 'PDF files are not supported for conversion. Pandoc can export to PDF but not convert from PDF.'
    })

@app.route('/api/health')
def health_check():
    """健康检查"""
    return jsonify({
        'status': 'healthy',
        'pandoc_available': converter.pandoc_available
    })

if __name__ == '__main__':
    # 检查pandoc是否可用
    if not converter.pandoc_available:
        print("警告: pandoc未安装或不可用，某些转换功能可能无法使用")
        print("请访问 https://pandoc.org/installing.html 安装pandoc")
    
    app.run(debug=True, host='0.0.0.0', port=5000)