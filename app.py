from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for, send_from_directory
from flask_login import LoginManager, login_required, current_user, login_user, logout_user
from datetime import datetime, time
import os
import sqlite3
from werkzeug.security import generate_password_hash, check_password_hash
from utils.baidu_api import BaiduVoiceAPI
from utils.excel_handler import ExcelHandler
from utils.auth import User, init_db

app = Flask(__name__)
app.config.from_object('config.Config')

# 确保上传文件夹存在
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

# 初始化 LoginManager
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

baidu_api = BaiduVoiceAPI()
excel_handler = ExcelHandler()


@app.route('/')
@login_required
def dashboard():
    return render_template('dashboard.html')


@app.route('/upload', methods=['POST'])
@login_required
def upload_excel():
    if 'file' not in request.files:
        return jsonify({'error': '没有提供文件'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '没有选择文件'}), 400

    # 确保文件是 Excel 格式
    if not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': '无效的文件格式'}), 400

    # 获取原始文件名和后缀
    file_name, file_extension = os.path.splitext(file.filename)
    secure_file_name = file_name + file_extension  # 重新组合文件名和后缀

    # 创建用户目录
    user_directory = os.path.join(app.config['UPLOAD_FOLDER'], current_user.username)
    if not os.path.exists(user_directory):
        os.makedirs(user_directory)

    # 保存文件，保留原始文件名和后缀
    filepath = os.path.join(user_directory, secure_file_name)
    file.save(filepath)

    print(f"Original filename: {file_name}")
    print(f"File will be saved to: {filepath}")

    # 读取Excel数据
    data = excel_handler.read_excel(filepath)
    print("Excel data:", data)  # 调试输出

    if data is None or not data.get('columns'):
        return jsonify({'error': 'Excel file must have headers'}), 400

    # Flatten merged headers if necessary
    merged_headers = []
    for i, col in enumerate(data['columns']):
        if isinstance(col, str) and col.startswith("Unnamed:"):
            merged_headers.append(f"Column {i+1}")
        else:
            merged_headers.append(col)

    data['columns'] = merged_headers
    print("Merged headers:", merged_headers)

    # Convert NaN values to None
    data['data'] = [[None if isinstance(value, float) and (value != value) else value for value in row] for row in data['data']]

    # Convert time objects to string
    for row in data['data']:
        for i, value in enumerate(row):
            if isinstance(value, time):
                row[i] = value.strftime('%H:%M:%S')  # Convert to string format

    # 返回数据时包含文件信息
    return jsonify({
        'columns': data['columns'],
        'data': data.get('data', []),  # 如果没有数据，返回空列表
        'filepath': filepath,
        'filename': secure_file_name
    })


@app.route('/recognize', methods=['POST'])
@login_required
def recognize_voice():
    if 'audio' not in request.files:
        return jsonify({'error': '没有上传音频文件'}), 400
    audio_data = request.files['audio'].read()
    text = baidu_api.recognize(audio_data)
    return jsonify({'text': text})


@app.route('/save', methods=['POST'])
@login_required
def save_excel():
    try:
        data = request.json.get('data')
        filepath = request.json.get('filepath')
        original_filename = request.json.get('filename')

        if not data:
            return jsonify({'error': 'Missing data'}), 400

        # 生成新的保存路径
        timestamp = datetime.now().strftime('%Y%m%d%H%M')
        if original_filename:
            name_without_ext = os.path.splitext(original_filename)[0]
            extension = '.xlsx'  # 统一使用 xlsx 格式
            new_filename = f"{name_without_ext}_{timestamp}{extension}"
        else:
            new_filename = f"excel_{timestamp}.xlsx"

        new_filepath = os.path.join(app.config['UPLOAD_FOLDER'], new_filename)

        # 保存Excel文件
        if excel_handler.save_excel(data, new_filepath):
            return jsonify({
                'success': True,
                'filepath': new_filepath,
                'filename': new_filename
            })
        else:
            return jsonify({'error': 'Failed to save file'}), 500

    except Exception as e:
        print(f"Save error: {str(e)}")  # 错误日志
        return jsonify({'error': str(e)}), 500


@app.route('/download')
@login_required
def download_excel():
    try:
        filepath = request.args.get('filepath')
        if not filepath or not os.path.exists(filepath):
            return jsonify({'error': 'File not found'}), 404

        # 获取文件名
        filename = os.path.basename(filepath)

        return send_file(
            filepath,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        print(f"Download error: {str(e)}")  # 错误日志
        return jsonify({'error': str(e)}), 500


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        user = User.authenticate(username, password)
        if user:
            login_user(user)
            return redirect(url_for('dashboard'))
        return render_template('login.html', error='用户名或密码错误')
    return render_template('login.html')


@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        confirm_password = request.form.get('confirm_password')

        if password != confirm_password:
            return render_template('register.html', error='两次输入的密码不一致')

        # 检查用户名是否唯一
        if User.get_by_username(username):
            return jsonify({'error': '用户名已存在'}), 400

        db = sqlite3.connect('users.db')
        cursor = db.cursor()
        try:
            cursor.execute(
                'INSERT INTO users (username, password_hash) VALUES (?, ?)',
                (username, generate_password_hash(password))
            )
            db.commit()
            db.close()
            return redirect(url_for('login'))
        except sqlite3.IntegrityError:
            db.close()
            return render_template('register.html', error='用户名已存在')
    return render_template('register.html')


@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))


@login_manager.user_loader
def load_user(user_id):
    return User.get(int(user_id))


@app.route('/delete_files', methods=['DELETE'])
@login_required
def delete_files():
    user_directory = os.path.join(app.config['UPLOAD_FOLDER'], current_user.username)

    # 删除用户目录下的所有文件，保留最后一次保存的文件
    if os.path.exists(user_directory):
        files = os.listdir(user_directory)
        if files:
            # 获取最后一次保存的文件
            last_saved_file = max([os.path.join(user_directory, f) for f in files], key=os.path.getctime)
            # 删除所有文件
            for f in files:
                file_path = os.path.join(user_directory, f)
                if file_path != last_saved_file:  # 保留最后一次保存的文件
                    os.remove(file_path)

    return jsonify({'message': '用户目录下的所有文件已删除，保留最后一次保存的文件'}), 200


if __name__ == '__main__':
    init_db()
    app.run(host='0.0.0.0', port=5000, ssl_context=('cert.pem', 'key.pem'))