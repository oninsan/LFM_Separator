import os
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
import pdfplumber
import pandas as pd
import pytesseract
from pdf2image import convert_from_path
from PIL import Image
from io import BytesIO

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/api/lfm', methods=['POST'])
def pdf_to_excel():
    if 'file' not in request.files:
        return jsonify({"error": "No file part in the request"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    if not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Only PDF files are allowed."}), 400

    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    try:
        names = []

        with pdfplumber.open(filepath) as pdf:
            has_text = any(page.extract_text() for page in pdf.pages)
            print(f"🔍 PDF has text: {has_text}")

            if has_text:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if table and len(table) > 1:
                            headers = table[0]
                            name_idx = None
                            # Try to find the index of the 'Name' column
                            for i, col in enumerate(headers):
                                if col and 'name' in col.lower():
                                    name_idx = i
                                    break

                            if name_idx is not None:
                                for row in table[1:]:
                                    if len(row) > name_idx and row[name_idx]:
                                        names.append([row[name_idx].strip()])

                if names:
                    df_result = pd.DataFrame(names, columns=["Name"])
                    output = BytesIO()
                    df_result.to_excel(output, index=False)
                    output.seek(0)
                    print("✅ Name column extracted from table.")
                    return send_file(output, as_attachment=True, download_name='names.xlsx')

        # Fallback to OCR using pytesseract
        print("⚠️ No table text found. Using Tesseract OCR fallback.")
        images = convert_from_path(filepath, dpi=300)
        ocr_names = []

        for img in images:
            text = pytesseract.image_to_string(img)
            for line in text.split('\n'):
                line = line.strip()
                if ',' in line and any(char.isalpha() for char in line):
                    ocr_names.append([line])

        if not ocr_names:
            return jsonify({"error": "No names could be extracted from the PDF using OCR."}), 400

        df = pd.DataFrame(ocr_names, columns=["Name"])
        output = BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)
        print("✅ OCR-based name list created using Tesseract.")
        return send_file(output, as_attachment=True, download_name='names_ocr.xlsx')

    except Exception as e:
        print(f"❌ Exception: {e}")
        return jsonify({"error": str(e)}), 500

@app.route('/')
def home():
    return "PDF to Excel Name Extractor API (Tesseract version) is running!"

if __name__ == '__main__':
    app.run(debug=True, port=5000)
