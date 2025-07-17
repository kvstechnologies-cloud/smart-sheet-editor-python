from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import pandas as pd
import os
import tempfile

app = Flask(__name__, static_folder='public', static_url_path='')
CORS(app)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return send_from_directory(app.static_folder, 'index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    if not file.filename.lower().endswith('.xlsx'):
        return jsonify({'error': 'Invalid file type'}), 400

    # Save to a temporary file within uploads folder
    fd, path = tempfile.mkstemp(dir=UPLOAD_FOLDER, suffix='.xlsx')
    try:
        with os.fdopen(fd, 'wb') as tmp:
            file.save(tmp)
        # Read all sheets
        xl = pd.read_excel(path, sheet_name=None, engine='openpyxl')
        sheets = [{'name': name, 'data': data.fillna('').values.tolist()} for name, data in xl.items()]
        return jsonify({'sheets': sheets})
    finally:
        os.remove(path)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
