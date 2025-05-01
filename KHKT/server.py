from flask import Flask, request, jsonify, make_response, send_file
from flask_cors import CORS
import sqlite3
import pandas as pd
import logging
import io
import json

# Khởi tạo ứng dụng Flask và cấu hình CORS
app = Flask(__name__)
CORS(app)

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

DATABASE_NAME = 'violations.db'  # Đường dẫn đến cơ sở dữ liệu

# Hàm khởi tạo cơ sở dữ liệu nếu chưa tồn tại
def init_db():
    conn = sqlite3.connect(DATABASE_NAME)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS violations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            class_name TEXT NOT NULL,
            day TEXT NOT NULL,
            violation_data TEXT
        )
    ''')
    conn.commit()
    conn.close()
    logging.info("Database initialized/checked.")

# Gọi hàm khởi tạo DB ngay khi server chạy
with app.app_context():
    init_db()

# Trang chủ để kiểm tra server hoạt động
@app.route('/')
def home():
    return "Server Flask đang hoạt động!"

# API lưu dữ liệu
@app.route('/save', methods=['POST'])
def save_data():
    try:
        data = request.get_json()
        if not data:
            return make_response(jsonify({'error': 'Invalid JSON data'}), 400)

        conn = sqlite3.connect(DATABASE_NAME)
        c = conn.cursor()

        # Duyệt qua từng class_name trong dữ liệu
        for class_name, rows in data.items():
            if not isinstance(rows, list):
                logging.error(f"Invalid data format for class: {class_name}")
                return make_response(jsonify({'error': f'Invalid data format for class: {class_name}'}), 400)

            for row in rows:
                if not isinstance(row, dict) or 'day' not in row or 'violations' not in row:
                    logging.error(f"Invalid data format for row: {row}")
                    return make_response(jsonify({'error': f'Invalid data format for row: {row}'}), 400)

                day = row['day']  # Chấp nhận mọi giá trị, không kiểm tra định dạng
                violations = str(row['violations'])

                # Lưu dữ liệu vào cơ sở dữ liệu
                c.execute('INSERT INTO violations (class_name, day, violation_data) VALUES (?, ?, ?)',
                          (class_name, day, violations))

        conn.commit()
        conn.close()

        logging.info("Data saved successfully.")
        return jsonify({'status': 'success'}), 201

    except sqlite3.Error as db_err:
        logging.error(f"Database error: {db_err}")
        return make_response(jsonify({'error': f'Database error: {str(db_err)}'}), 500)

    except json.JSONDecodeError as json_err:
        logging.error(f"JSON Decode Error: {json_err}")
        return make_response(jsonify({'error': 'Invalid JSON format'}), 400)

    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")
        return make_response(jsonify({'error': f'An unexpected error occurred: {str(e)}'}), 500)

# API xuất dữ liệu ra file Excel
@app.route('/export_excel', methods=['GET'])
def export_excel():
    try:
        conn = sqlite3.connect(DATABASE_NAME)
        df = pd.read_sql_query("SELECT * FROM violations", conn)
        conn.close()

        # Tạo buffer để chứa file Excel
        output = io.BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)

        # Trả về file Excel dưới dạng tệp đính kèm
        return send_file(output, as_attachment=True, download_name="violations_export.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        logging.error(f"Error exporting Excel: {e}")
        return jsonify({'status': 'error', 'message': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)

