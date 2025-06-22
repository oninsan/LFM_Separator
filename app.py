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
stopwords = {
    "section", "name", "namelist", "grade", "list", "student", "students",
    "cebu", "roosevelt", "memorial", "colleges", "bogo", "city", "prof", "nino", "abao"
}

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

                    if any(sw in line.lower() for sw in stopwords):
                        continue

                    if re.fullmatch(r'[\d\W_]+', line):
                        continue

                    line = re.sub(r'\d+', '', line)
                    line = re.sub(r"[^A-Za-zÑñ\s,\.\-']", '', line)

                    print(f"🔍 Cleaned Line: '{line}'")

                    if ',' not in line:
                        continue

                    line = re.sub(r'^(name\s*[:\-]*)', '', line, flags=re.IGNORECASE).strip()

                    suffixes = {"BSIT", "BSCS", "BSCPE", "BSCE", "BSME", "BSEE", "BSBA", "BSN", "BS", "AB", "JR", "SR", "III", "IV", "II"}

                    # ...inside your match block...
                    match = re.search(r"([^,]+),\s+(.+)", line)
                    if match:
                        last = match.group(1).lstrip('-').strip()
                        first_middle = match.group(2).strip()

                        # Split by whitespace, keep only alphabetic words
                        name_parts = [word for word in first_middle.split() if word.isalpha()]
                        # Remove suffixes/course codes at the end
                        while name_parts and name_parts[-1].upper() in suffixes:
                            name_parts.pop()

                        if len(name_parts) == 0:
                            continue
                        elif len(name_parts) == 1:
                            first = name_parts[0]
                            middle = ''
                        else:
                            first = ' '.join(name_parts[:-1])
                            middle = name_parts[-1][0].upper() + '.'

                        print(f"✅ Parsed: Last='{last}', First='{first}', MI='{middle}'")
                        ocr_names.append([last, first, middle])


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
