from flask import Flask, render_template, request, jsonify
import pandas as pd
import numpy as np
import sqlite3
from werkzeug.utils import secure_filename
import os
import logging
import traceback
from datetime import datetime

# 配置日志
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

app = Flask(__name__)

# 配置文件上传
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# 确保上传文件夹存在
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 限制文件大小为16MB


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def init_db():
    try:
        conn = sqlite3.connect('excel_data.db')
        conn.close()
        logger.info("Database initialized successfully")
    except Exception as e:
        logger.error(f"Database initialization error: {str(e)}")
        raise


def store_excel_data(df, table_name):
    try:
        conn = sqlite3.connect('excel_data.db')
        # 将DataFrame中的NaN值替换为None，再存入数据库
        df = df.replace({np.nan: None, np.inf: None, -np.inf: None})
        df.to_sql(table_name, conn, if_exists='replace', index=False)
        conn.close()
        logger.info(f"Data stored successfully in table: {table_name}")
    except Exception as e:
        logger.error(f"Error storing data in database: {str(e)}")
        raise


def clean_value(value):
    """清理数据值，处理特殊情况"""
    if pd.isna(value) or value is None:
        return ""
    elif isinstance(value, (int, float)):
        if np.isnan(value) or np.isinf(value):
            return ""
        return str(value)
    elif isinstance(value, datetime):
        return value.strftime('%Y-%m-%d %H:%M:%S')
    else:
        return str(value)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        if 'file' not in request.files:
            logger.warning("No file part in request")
            return jsonify({'error': 'No file part'}), 400

        file = request.files['file']
        if file.filename == '':
            logger.warning("No selected file")
            return jsonify({'error': 'No selected file'}), 400

        if not allowed_file(file.filename):
            logger.warning(f"Invalid file type: {file.filename}")
            return jsonify({'error': 'Invalid file type. Please upload .xlsx or .xls file'}), 400

        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)

        logger.info(f"Saving file to: {filepath}")
        file.save(filepath)

        try:
            # 读取Excel文件
            logger.info(f"Reading Excel file: {filepath}")
            df = pd.read_excel(filepath, engine='openpyxl')

            # 检查数据框是否为空
            if df.empty:
                logger.warning("Uploaded Excel file is empty")
                return jsonify({'error': 'The uploaded Excel file is empty'}), 400

            # 存储到数据库
            table_name = 'excel_data_' + filename.rsplit('.', 1)[0]
            store_excel_data(df, table_name)

            # 准备返回数据
            columns = df.columns.tolist()
            data = []
            for _, row in df.iterrows():
                cleaned_row = [clean_value(value) for value in row]
                data.append(cleaned_row)

            # 删除临时文件
            os.remove(filepath)
            logger.info("File processed successfully")

            return jsonify({
                'columns': columns,
                'data': data
            })

        except Exception as e:
            logger.error(f"Error processing file: {str(e)}\n{traceback.format_exc()}")
            if os.path.exists(filepath):
                os.remove(filepath)
            return jsonify({'error': f'Error processing file: {str(e)}'}), 500

    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}\n{traceback.format_exc()}")
        return jsonify({'error': f'Server error: {str(e)}'}), 500


@app.errorhandler(413)
def too_large(e):
    return jsonify({'error': 'File is too large. Maximum size is 16MB'}), 413


@app.errorhandler(500)
def internal_error(e):
    logger.error(f"Internal server error: {str(e)}\n{traceback.format_exc()}")
    return jsonify({'error': 'Internal server error occurred'}), 500


if __name__ == '__main__':
    init_db()
    app.run(debug=True)