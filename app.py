
from flask import Flask, request, jsonify, send_from_directory
import pandas as pd
import os
from layout_engine import generate_layout_and_summary

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route("/api/process", methods=["POST"])
def process_csv():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded."}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "Empty filename."}), 400

    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)

    layout_path, summary_path = generate_layout_and_summary(filepath, OUTPUT_FOLDER)

    return jsonify({
        "layout_png": f"/download/{os.path.basename(layout_path)}",
        "excel_summary": f"/download/{os.path.basename(summary_path)}"
    })

@app.route("/download/<filename>")
def download_file(filename):
    return send_from_directory(OUTPUT_FOLDER, filename)

if __name__ == "__main__":
    app.run(debug=True)
