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
  "cebu", "roosevelt", "memorial", "colleges", "bogo", "city", "prof", "nino, abao",
  "marjorie, reso", "joel, lim", "windel, pelayo", "jonel, gelig","leonard, balabat",
  "carl joshua, cosep", "hilarion, raganas"
}

def allowed_file(filename):
  return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/api/lfm', methods=['POST'])
def pdf_text_to_excel():
  files = request.files.getlist('file')
  if not files or all(f.filename == '' for f in files):
    return jsonify({"error": "No selected file(s)"}), 400

  output = BytesIO()
  workbook = xlsxwriter.Workbook(output, {'in_memory': True, 'constant_memory': True})

  for file in files:
    if not allowed_file(file.filename):
      continue

    print(f"📄 Reading PDF for text: {file.filename}")
    ocr_names = []
    try:
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

            if ',' not in line:
              continue

            line = re.sub(r'^(name\s*[:\-]*)', '', line, flags=re.IGNORECASE).strip()

            suffixes = {"BSIT","BSCRIM", "BSCS", "BSCPE", "BSCE", "BSME", "BSEE", "BSBA", "BSN", "BS", "AB", "JR", "SR","JR.", "SR.", "III", "IV", "II"}

            match = re.search(r"( [^,]+),\s+(.+)", line)
            if match:
              last = match.group(1).lstrip('-').strip()
              first_middle = match.group(2).strip()

              name_parts = [word for word in first_middle.split() if word.isalpha()]
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

              ocr_names.append([last, first, middle])
    except Exception as e:
      print(f"❌ Exception processing {file.filename}: {e}")
      continue

    ocr_names.sort(key=lambda x: x[0].lower())
    # Use filename (without extension) as sheet name, max 31 chars for Excel
    sheet_name = file.filename.rsplit('.', 1)[0][:31] or "Sheet"
    worksheet = workbook.add_worksheet(sheet_name)

    # Define a bordered cell format
    border_fmt = workbook.add_format({'border': 1})

    # Define a bold + bordered cell format for headers
    header_fmt = workbook.add_format({'bold': True, 'border': 1})

    # Write header with bold and border
    headers = ['Last Name', 'First Name', 'Middle Initial']
    worksheet.write_row(0, 0, headers, header_fmt)

    # Write data rows with border
    for row_num, row_data in enumerate(ocr_names, start=1):
      worksheet.write_row(row_num, 0, row_data, border_fmt)

    # Auto-fit column widths
    for col_num, header in enumerate(headers):
      # Get max length of data in this column (including header)
      max_len = max(
        [len(str(row_data[col_num])) for row_data in ocr_names] + [len(header)]
      )
      worksheet.set_column(col_num, col_num, max_len + 2)  # +2 for padding

  workbook.close()
  output.seek(0)
  gc.collect()

  print("✅ Final structured Excel with multiple sheets generated from PDFs.")
  return send_file(output, as_attachment=True, download_name='structured_names.xlsx')