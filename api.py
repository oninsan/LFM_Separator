import os
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
import pdfplumber
import pandas as pd
from paddleocr import PaddleOCR
from pdf2image import convert_from_path
from PIL import Image
from io import BytesIO

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

ocr = PaddleOCR(use_angle_cls=True, lang='en')  # OCR fallback

@app.route('/api/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    if 'file' not in request.files:
        return jsonify({"error": "No file part in the request"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    if not file.filename.lower().endswith('.pdf'):
        return jsonify({"error": "Only PDF files are allowed."}), 400

    # Save PDF file
    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    try:
        all_tables = []

        with pdfplumber.open(filepath) as pdf:
            has_text = any(page.extract_text() for page in pdf.pages)
            print(f"🔍 PDF has text: {has_text}")

            if has_text:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if table and len(table) > 1:
                            df = pd.DataFrame(table[1:], columns=table[0])
                            all_tables.append(df)

                if all_tables:
                    df_result = pd.concat(all_tables, ignore_index=True)
                    output = BytesIO()
                    df_result.to_excel(output, index=False)
                    output.seek(0)
                    print("✅ Extracted tables from text PDF.")
                    return send_file(output, as_attachment=True, download_name='output.xlsx')

        # Fallback to OCR
        print("⚠️ No tables found. Falling back to OCR.")
        images = convert_from_path(filepath, dpi=300)
        lines = []

        for i, img in enumerate(images):
            temp_path = os.path.join(UPLOAD_FOLDER, f"page_{i}.png")
            img.save(temp_path, "PNG")
            result = ocr.predict(temp_path)
            for line in result[0]:
                text = line[1][0]
                if text.strip():
                    lines.append([text])

        if not lines:
            return jsonify({"error": "No text could be extracted from PDF."}), 400

        df = pd.DataFrame(lines)
        output = BytesIO()
        df.to_excel(output, index=False, header=False)
        output.seek(0)

        print("✅ OCR-based Excel created.")
        return send_file(output, as_attachment=True, download_name='output_ocr.xlsx')

    except Exception as e:
        print(f"❌ Exception: {e}")
        return jsonify({"error": str(e)}), 500

@app.route('/')
def home():
    return "PDF to Excel API is running!"

if __name__ == '__main__':
    app.run(debug=True, port=5000)
