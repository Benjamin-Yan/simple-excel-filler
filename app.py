import os, signal, threading, time, ipaddress
import atexit
from flask import Flask, render_template, render_template_string, request, session, abort
import pandas as pd
from utils import fill_excel_default, get_column_name, fill_excel_select
from shared import created_temp_files

app = Flask(__name__)
app.secret_key = os.urandom(24) # for session management
apppid = os.getpid()

allowed_ips = os.getenv("ALLOWED_IPS", "127.0.0.1").split(",")

def check_ip():
    client_ip = request.remote_addr
    ip_obj = ipaddress.ip_address(client_ip)

    for net in allowed_ips:
        net = net.strip()
        if ip_obj in ipaddress.ip_network(net, strict=False):
            return  # 允許
    abort(403, description="IP not allowed")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/select')
def about():
    return render_template('select.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    """ 說明
    Handles file upload requests.
    若為預設模式，則回傳狀態資訊；
    若為進階模式，則回傳 get_column_name() 函式的結果。
    """
    check_ip()  # 拒絕非法 IP 消耗資源

    # 確認有沒有傳成功
    if 'input_file' not in request.files or 'output_file' not in request.files:
        return 'No file part'

    input_file = request.files['input_file']
    output_file = request.files['output_file']
    if input_file.filename == '' or output_file.filename == '':
        return 'No selected file'

    source = request.form.get('source')
    result = {}
    if source == 'default':
        result = fill_excel_default(input_file, output_file)
    else:
        return get_column_name(input_file, output_file)

    if result["status"] == "success":
        # Processed file: {input_file.filename}<br>
        # Rows: {result['rows']}, Columns: {result['columns']}<br>
        return render_template_string("""
            <link rel="stylesheet" href="{{ url_for('static', filename='css/style.css') }}">
            The task was completed successfully.<br>
            <a href="/">回首頁</a>
        """)
    else:
        return f"Error: {result['message']}<br><a href='/'>Try again</a>"

@app.route('/select/process', methods=['POST'])
def process_select():
    """ 說明
    取得選定的欄位和資料列數
    回傳表單讓使用者填寫對應的位置
    """
    selected_cols = request.form.getlist('CB') # get column index
    selected_rows = request.form.get('quantity')

    temp_fileI_path = session['temp_fileI_path']
    if not temp_fileI_path or not os.path.exists(temp_fileI_path):
        return 'No file to process', 404

    df = pd.read_excel(temp_fileI_path, sheet_name=0, engine='openpyxl')
    columns = df.columns.tolist()

    selected_cols = [int(i) for i in selected_cols]
    return render_template_string('''
    <meta charset="UTF-8">
    <link rel="stylesheet" href="{{ url_for('static', filename='css/style.css') }}">
    <h2>請填上對應的欄位</h2>
    <p>e.g., "3F"</p>
    <form action="/select/location" method="POST">
        {% for i in selected_cols %}
            <label for="loc{{ i }}">{{ columns[i] }}</label>
            <input type="text" name="location" id="outloc{{ i }}" required><br>
        {% endfor %}

        {% for i in selected_cols %}
            <input type="hidden" name="selected_cols" value="{{ i }}">
        {% endfor %}
        <input type="hidden" name="selrow" value="{{ selected_rows }}">
        <input type="submit" value="確認執行">
    </form>
    ''', selected_cols=selected_cols, columns=columns, selected_rows=selected_rows)

@app.route('/select/location', methods=['POST'])
def process_location():
    """ 說明
    取得使用者指定的位置資料
    回傳執行 fill_excel_select() 的結果。
    """
    location_data = request.form.getlist('location') # ['3F', '27D', '4I']
    selected_cols = request.form.getlist('selected_cols')
    selected_rows = request.form.get('selrow')

    result = {}
    result = fill_excel_select(selected_cols, selected_rows, location_data)
    if result["status"] == "success":
        return render_template_string("""
            <link rel="stylesheet" href="{{ url_for('static', filename='css/style.css') }}">
            The task was completed successfully.<br>
            <a href="/">回首頁</a>
        """)
    else:
        return f"Error: {result['message']}<br><a href='/select'>Try again</a>"

@atexit.register
def cleanup_temp_files():
    for file_path in created_temp_files:
        try:
            if os.path.exists(file_path):
                os.unlink(file_path)
                print(f"Deleted temp file: {file_path}")
        except Exception as e:
            print(f"Error deleting temp file {file_path}: {e}")

@app.route('/endapp')
def kill_backend():
    def delayed_shutdown():
        cleanup_temp_files()
        time.sleep(1)
        os.kill(apppid, signal.SIGTERM)
    threading.Thread(target=delayed_shutdown).start()
    return "任務已完成，可以放心關閉此頁面。"

if __name__ == '__main__':
    app.run(debug=False)

