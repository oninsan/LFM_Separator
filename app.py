import re
import gc
import pdfplumber
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from io import BytesIO
import xlsxwriter

app = Flask(__name__)
CORS(app)

ALLOWED_EXTENSIONS = {'pdf'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/api/lfm', methods=['POST'])
def pdf_text_to_excel():
    if 'file' not in request.files:
        return jsonify({"error": "No file part in the request"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": "Only PDF files are allowed."}), 400

    try:
        print("📄 Reading PDF for text...")
        ocr_names = []

        with pdfplumber.open(file.stream) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue

                for line in text.splitlines():
                    line = line.strip()
                    if not line:
                        continue

                    # Remove "Name:" or similar prefixes
                    line = re.sub(r'^(name\s*[:\-]*)', '', line, flags=re.IGNORECASE).strip()

                    # Match "Lastname, Firstname Middlename"
                    match = re.search(r"([A-ZÑñ][A-ZÑñ\s.\-']+),\s*([A-ZÑñ\s.\-']+)", line, re.IGNORECASE)
                    if match:
                        last = match.group(1).strip()
                        right = match.group(2).strip().split()
                        if right:
                            first = ' '.join(right[:-1]) if len(right) > 1 else right[0]
                            middle = right[-1][0] + '.' if len(right) > 1 else ''
                            ocr_names.append([last, first, middle])
                        continue

                    # Fallback: "Lastname Firstname Middlename"
                    parts = line.split()
                    if len(parts) >= 2 and all(re.match(r"^[A-ZÑñ.\-']+$", w, re.IGNORECASE) for w in parts):
                        last = parts[0]
                        first = ' '.join(parts[1:-1]) if len(parts) > 2 else parts[1]
                        middle = parts[-1][0] + '.' if len(parts) > 2 else ''
                        ocr_names.append([last, first, middle])

        if not ocr_names:
            return jsonify({"error": "No names could be extracted from the PDF."}), 400

        ocr_names.sort(key=lambda x: x[0].lower())

        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True, 'constant_memory': True})
        worksheet = workbook.add_worksheet("Names")

        worksheet.write_row(0, 0, ['Last Name', 'First Name', 'Middle Initial'])
        for row_num, row_data in enumerate(ocr_names, start=1):
            worksheet.write_row(row_num, 0, row_data)

        workbook.close()
        output.seek(0)

        del ocr_names, text, workbook, worksheet
        gc.collect()

        print("✅ Final structured Excel generated from PDF.")
        return send_file(output, as_attachment=True, download_name='structured_names.xlsx')

    except Exception as e:
        print(f"❌ Exception: {e}")
        return jsonify({"error": str(e)}), 500
