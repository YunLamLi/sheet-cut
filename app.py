
from flask import Flask, request, render_template, send_from_directory
import os
from layout_engine import generate_layout_and_summary

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['csvfile']
        if file and file.filename.endswith('.csv'):
            csv_path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(csv_path)

            generate_layout_and_summary(csv_path, OUTPUT_FOLDER)

            output_files = [f for f in os.listdir(OUTPUT_FOLDER) if f.endswith('.png') or f.endswith('.xlsx')]
            return render_template('index.html', files=output_files)

    return render_template('index.html', files=[])

@app.route('/output/<filename>')
def download_file(filename):
    return send_from_directory(OUTPUT_FOLDER, filename)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)

