# -*- coding: utf-8 -*-
"""
Flask Web Server for Gongwen Formatter
Provides REST API for document formatting
"""

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import tempfile
from gongwen_formatter import GongwenFormatter, create_demo_document, format_existing_document, FORMAT_MODE_STANDARD, FORMAT_MODE_GOVERNMENT

app = Flask(__name__, static_folder='.', static_url_path='')
CORS(app)

# Serve the main page
@app.route('/')
def index():
    return send_file('index.html')

# API endpoint to format a document
@app.route('/api/format', methods=['POST'])
def format_document():
    try:
        # Check if file was uploaded
        if 'file' not in request.files:
            return jsonify({'error': '未上传文件'}), 400

        file = request.files['file']

        if file.filename == '':
            return jsonify({'error': '文件名为空'}), 400

        if not file.filename.endswith('.docx'):
            return jsonify({'error': '仅支持 .docx 格式'}), 400

        # 获取格式模式参数（默认为标准模式）
        format_mode = request.form.get('format_mode', FORMAT_MODE_STANDARD)
        if format_mode not in [FORMAT_MODE_STANDARD, FORMAT_MODE_GOVERNMENT]:
            format_mode = FORMAT_MODE_STANDARD

        # Save uploaded file to temp directory
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_input:
            file.save(tmp_input.name)
            input_path = tmp_input.name

        # Create output file path
        output_path = input_path.replace('.docx', '_formatted.docx')

        # Format the document with selected mode
        format_existing_document(input_path, output_path, format_mode=format_mode)
        
        # Send the formatted file
        response = send_file(
            output_path,
            as_attachment=True,
            download_name=file.filename.replace('.docx', '_formatted.docx'),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        
        # Clean up temp files after sending
        @response.call_on_close
        def cleanup():
            try:
                os.unlink(input_path)
                os.unlink(output_path)
            except:
                pass
        
        return response
        
    except Exception as e:
        print(f"Error formatting document: {e}")
        return jsonify({'error': f'格式化失败: {str(e)}'}), 500

# API endpoint to create demo document
@app.route('/api/create-demo', methods=['POST'])
def create_demo():
    try:
        # 获取格式模式参数
        data = request.get_json() or {}
        format_mode = data.get('format_mode', FORMAT_MODE_STANDARD)
        if format_mode not in [FORMAT_MODE_STANDARD, FORMAT_MODE_GOVERNMENT]:
            format_mode = FORMAT_MODE_STANDARD

        # Create temp file for demo
        with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_output:
            output_path = tmp_output.name

        # Create demo document with selected mode
        create_demo_document(output_path, format_mode=format_mode)
        
        # Send the demo file
        response = send_file(
            output_path,
            as_attachment=True,
            download_name='demo_gongwen.docx',
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        
        # Clean up temp file after sending
        @response.call_on_close
        def cleanup():
            try:
                os.unlink(output_path)
            except:
                pass
        
        return response
        
    except Exception as e:
        print(f"Error creating demo: {e}")
        return jsonify({'error': f'创建示例文档失败: {str(e)}'}), 500

# Health check endpoint
@app.route('/api/health', methods=['GET'])
def health_check():
    return jsonify({'status': 'ok', 'message': '公文格式化服务运行中'})

if __name__ == '__main__':
    print("=" * 60)
    print("公文格式化工具 Web 服务器")
    print("=" * 60)
    print("\n服务器启动中...")
    print(f"访问地址: http://localhost:5000")
    print(f"API 文档: http://localhost:5000/api/health")
    print("\n按 Ctrl+C 停止服务器\n")
    
    app.run(host='0.0.0.0', port=5000, debug=True)
