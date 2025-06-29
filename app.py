from flask import Flask, request, render_template, send_from_directory, jsonify
import os
from layout_engine import generate_layout_and_summary
import traceback

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def clean_output_folder():
    for f in os.listdir(OUTPUT_FOLDER):
        file_path = os.path.join(OUTPUT_FOLDER, f)
        if os.path.isfile(file_path):
            os.remove(file_path)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/process', methods=['POST'])
def process_api():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    if not file:
        return jsonify({'error': 'File is empty'}), 400

    clean_output_folder()

    csv_path = os.path.join(OUTPUT_FOLDER, 'input.csv')
    file.save(csv_path)

    try:
        print("Starting layout generation...")
        generate_layout_and_summary(csv_path, OUTPUT_FOLDER)
        print("Layout generation completed.")
    except Exception as e:
        print("Layout generation failed:")
        traceback.print_exc()
        return jsonify({'error': 'Layout generation failed', 'details': str(e)}), 500

    png_files = sorted([f for f in os.listdir(OUTPUT_FOLDER) if f.endswith('.png')])
    png_links = [f"/output/{f}" for f in png_files]

    xlsx_files = sorted([f for f in os.listdir(OUTPUT_FOLDER) if f.endswith('.xlsx')])
    excel_link = f"/output/{xlsx_files[0]}" if xlsx_files else ""

    return jsonify({
        'layout_pngs': png_links,
        'excel_summary': excel_link
    })

@app.route('/output/<filename>')
def download_file(filename):
    return send_from_directory(OUTPUT_FOLDER, filename)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
