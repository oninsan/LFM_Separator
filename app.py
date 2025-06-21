import re
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
import pandas as pd
import pytesseract
from PIL import Image, ImageOps
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
        print("🖼️ Processing image with Tesseract OCR...")
        image = Image.open(file.stream)
        image = ImageOps.expand(image, border=(0, 10, 0, 0), fill='white')

        text = pytesseract.image_to_string(image, lang='eng')

        ocr_names = []

        for line in text.split('\n'):
            line = line.strip()
            if not line:
                continue

            # Remove "Name:" or similar prefixes
            line = re.sub(r'^(name\s*[:\-]*)', '', line, flags=re.IGNORECASE).strip()

            # Try to find names like "Lastname, Firstname Middlename"
            match = re.search(r"([A-ZÑñ][A-ZÑñ\s\-\.]+),\s*([A-ZÑñ\s\-\.]+)", line, re.IGNORECASE)
            if match:
                last = match.group(1).strip()
                right = match.group(2).strip()
                first = ' '.join(right.split()[:-1]) if len(right.split()) > 1 else right
                middle = right.split()[-1][0] + '.' if len(right.split()) > 1 else ''
                ocr_names.append([last, first, middle])
                continue

            # Optional fallback if comma is missing but format is valid (Lastname Firstname Middlename)
            words = line.split()
            if len(words) >= 2 and all(re.match(r"^[A-ZÑñ\-\.]+$", w, re.IGNORECASE) for w in words):
                last = words[0]
                first = ' '.join(words[1:-1]) if len(words) > 2 else words[1]
                middle = words[-1][0] + '.' if len(words) > 2 else ''
                ocr_names.append([last, first, middle])

        if not ocr_names:
            return jsonify({"error": "No names could be extracted from the image."}), 400

        # Build DataFrame
        df = pd.DataFrame(ocr_names, columns=["Last Name", "First Name", "Middle Initial"])
        df = df.sort_values(by="Last Name")

        # Write to Excel in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Names', index=False)

        output.seek(0)
        print("✅ Final structured Excel generated.")
        return send_file(output, as_attachment=True, download_name='structured_names.xlsx')

    except Exception as e:
        print(f"❌ Exception: {e}")
        return jsonify({"error": str(e)}), 500

# @app.route('/')
# def home():
#     return "Image to Structured Excel Name Extractor API is running!"

# if __name__ == '__main__':
#     app.run(debug=True, port=5000)
