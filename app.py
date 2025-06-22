import re
import gc
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
import pandas as pd
import pytesseract
from PIL import Image
from io import BytesIO

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

        # Open as grayscale and crop whitespace
        with Image.open(file.stream).convert("L") as image:
            bbox = image.getbbox()
            if bbox:
                image = image.crop(bbox)

            # OCR with optimized settings
            text = pytesseract.image_to_string(image, lang='eng', config='--psm 6')

        del image
        gc.collect()

        # Process text lines
        ocr_names = []
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue

            line = re.sub(r'^(name\s*[:\-]*)', '', line, flags=re.IGNORECASE).strip()

            # "Lastname, Firstname Middlename"
            match = re.search(r"([A-ZÑñ][A-ZÑñ\s\-\.]+),\s*([A-ZÑñ\s\-\.]+)", line, re.IGNORECASE)
            if match:
                last = match.group(1).strip()
                right = match.group(2).strip().split()
                if right:
                    first = ' '.join(right[:-1]) if len(right) > 1 else right[0]
                    middle = right[-1][0] + '.' if len(right) > 1 else ''
                    ocr_names.append((last, first, middle))
                continue

            # Fallback: "Lastname Firstname Middlename"
            parts = line.split()
            if len(parts) >= 2 and all(re.match(r"^[A-ZÑñ\-\.]+$", w, re.IGNORECASE) for w in parts):
                last = parts[0]
                first = ' '.join(parts[1:-1]) if len(parts) > 2 else parts[1]
                middle = parts[-1][0] + '.' if len(parts) > 2 else ''
                ocr_names.append((last, first, middle))

        if not ocr_names:
            return jsonify({"error": "No names could be extracted from the image."}), 400

        # Directly create DataFrame
        df = pd.DataFrame(ocr_names, columns=["Last Name", "First Name", "Middle Initial"])
        df.sort_values("Last Name", inplace=True)

        # Generate Excel in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Names', index=False)

        output.seek(0)
        del df
        gc.collect()

        print("✅ Excel generated.")
        return send_file(output, as_attachment=True, download_name='structured_names.xlsx')

    except Exception as e:
        print(f"❌ Exception: {e}")
        return jsonify({"error": str(e)}), 500
