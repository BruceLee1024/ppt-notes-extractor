#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from flask import Flask, request, render_template, send_file, jsonify
from werkzeug.utils import secure_filename
import os
from extract_notes import extract_notes

app = Flask(__name__)

# 配置上传文件存储目录
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': '没有上传文件'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '没有选择文件'}), 400
    
    if not file.filename.endswith('.pptx'):
        return jsonify({'error': '请上传PPTX格式的文件'}), 400
    
    try:
        # 保存上传的文件
        filename = secure_filename(file.filename)
        pptx_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(pptx_path)
        
        # 生成输出文件路径
        output_filename = os.path.splitext(filename)[0] + '_notes.md'
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
        
        # 提取备注
        success = extract_notes(pptx_path, output_path)
        
        if success:
            # 读取生成的Markdown文件内容
            with open(output_path, 'r', encoding='utf-8') as f:
                markdown_content = f.read()
            
            # 清理临时文件
            os.remove(pptx_path)
            os.remove(output_path)
            
            return jsonify({
                'success': True,
                'content': markdown_content
            })
        else:
            return jsonify({'error': '处理文件时出错'}), 500
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)