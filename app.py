import pandas as pd
from flask import Flask, request, jsonify
from flask_cors import CORS
import io
import base64

api = Flask(__name__)
CORS(api)
@api.route('/api/lfm', methods=['POST'])
def LFM_SEP():
  workbook = request.files['workbook']
  xls = pd.ExcelFile(workbook)
  output = io.BytesIO()

  with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    for sheet_name in xls.sheet_names:
      raw = pd.read_excel(xls, sheet_name=sheet_name)
      for col in raw.columns:
        # get the Last Name
        raw['Last Name'] = raw[col].str.split(',').str[0].str.strip()
        # get the First Name
        first_middle = raw[col].str.split(',').str[1].str.strip()
        raw['First Name'] = first_middle.str.split().apply(lambda x: ' '.join(x[:-1]) if len(x) > 1 else x[0])
        # get the Middle Initial
        raw['Middle Initial'] = raw[col].str.split().apply(lambda x: x[-1][0]+'.' if len(x) > 1 else '')
        raw = raw.drop(columns=[col])
        raw = raw.sort_values('Last Name', ascending=True)
        break  # Only process the first column

      raw.to_excel(writer, sheet_name=sheet_name, index=False)

  output.seek(0)
  return (
    output.read(),
    200,
    {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': 'attachment; filename=processed.xlsx'
    }
  )

# if __name__=='__main__':
#   api.run(port=5000)