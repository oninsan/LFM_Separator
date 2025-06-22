
import re
import gc
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
import pytesseract
from PIL import Image
from io import BytesIO
import xlsxwriter

app = Flask(__name__)
CORS(app)

ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'bmp'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/api/lfm', methods=['POST'])
def image_to_excel():
    if 'file' not in request.files:
        return jsonify({"error": "No file part in the request"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": "Only image files (png, jpg, jpeg, bmp) are allowed."}), 400

    try:
        print("🖼️ Running lightweight OCR...")
        with Image.open(file.stream).convert("L") as image:
            image = image.crop(image.getbbox())

            text = pytesseract.image_to_string(image, lang='eng')

        ocr_names = []
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
            return jsonify({"error": "No names could be extracted from the image."}), 400

        # Sort names by Last Name
        ocr_names.sort(key=lambda x: x[0].lower())

        # Use xlsxwriter directly with constant_memory
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True, 'constant_memory': True})
        worksheet = workbook.add_worksheet("Names")

        # Write headers
        worksheet.write_row(0, 0, ['Last Name', 'First Name', 'Middle Initial'])

        # Write each row
        for row_num, row_data in enumerate(ocr_names, start=1):
            worksheet.write_row(row_num, 0, row_data)

        workbook.close()
        output.seek(0)

        # Force garbage collection to free memory
        del ocr_names, text, image, workbook, worksheet
        gc.collect()

        print("✅ Final structured Excel generated.")
        return send_file(output, as_attachment=True, download_name='structured_names.xlsx')

    except Exception as e:
        print(f"❌ Exception: {e}")
        return jsonify({"error": str(e)}), 500