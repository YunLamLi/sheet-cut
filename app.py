from flask import Flask, request, render_template, send_from_directory, jsonify
import os
from layout_engine import generate_layout_and_summary

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Serve the main HTML page
@app.route('/')
def index():
    return render_template('index.html')

# API route to handle AJAX POST from JavaScript
@app.route('/api/process', methods=['POST'])
def process_api():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']
    if not file.filename.endswith('.csv'):
        return jsonify({"error": "Invalid file type"}), 400

    # Save uploaded CSV
    csv_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(csv_path)

    # Generate layout and summary files
    generate_layout_and_summary(csv_path, OUTPUT_FOLDER)

    # Find output files
    layout_file = next((f for f in os.listdir(OUTPUT_FOLDER) if f.endswith('.png')), None)
    summary_file = next((f for f in os.listdir(OUTPUT_FOLDER) if f.endswith('.xlsx')), None)

    if layout_file and summary_file:
        return jsonify({
            "layout_png": f"/output/{layout_file}",
            "excel_summary": f"/output/{summary_file}"
        })
    else:
        return jsonify({"error": "Output files not found"}), 500

# Serve output files (PNG/XLSX)
@app.route('/output/<filename>')
def download_file(filename):
    return send_from_directory(OUTPUT_FOLDER, filename)

# Run app on Render or locally
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 10000))
    app.run(host='0.0.0.0', port=port)
