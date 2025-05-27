import os
import re
import io
from flask import Flask, request, render_template, send_file, flash, redirect, url_for
import pdfplumber
import pandas as pd

app = Flask(__name__)
app.secret_key = 'your-secret-key'

EXCEL_FIELDS = [
    '物件名','号室','㎡数','所有者名','状況','自宅番号','勤務先','勤務先番号','メモ','本人携帯',
    '所有者住所','地番','最寄駅','築年数','総戸数','規模','構造','URL1','URL2','管積計',
    'セパか３点か','共有名義人1','共有名義人1番号','共有名義人1住所',
    '共有名義人2','共有名義人2番号','共有名義人2住所'
]

def extract_data_from_pdf(pdf_path):
    owners = []
    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text() or ""
        m = re.search(r'大阪府.+?区.+?丁目[0-9０-９]+[‐－ー―][0-9０-９]+[‐－ー―][0-9０-９]+', text)
        prop_addr = m.group(0) if m else ''
        r = re.search(r'([0-9０-９]{1,4})$', prop_addr)
        room_no = r.group(1) if r else ''
        for line in text.splitlines():
            if '│' in line:
                parts = [p.strip() for p in line.split('│') if p.strip()]
                if len(parts) == 2:
                    row = {f: '' for f in EXCEL_FIELDS}
                    row['物件名']      = ''
                    row['号室']        = room_no
                    row['所有者名']    = parts[1]
                    row['所有者住所']  = parts[0]
                    owners.append(row)
    return owners

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'pdfs' not in request.files:
        flash('PDF が選択されていません')
        return redirect(url_for('index'))

    files = request.files.getlist('pdfs')
    all_rows = []
    for f in files:
        if f.filename.lower().endswith('.pdf'):
            data = f.read()
            with open('tmp.pdf', 'wb') as tmp:
                tmp.write(data)
            all_rows += extract_data_from_pdf('tmp.pdf')
            os.remove('tmp.pdf')

    df = pd.DataFrame(all_rows, columns=EXCEL_FIELDS)   
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
        
    return send_file(
        output,
        as_attachment=True,
        download_name='extracted_owners.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    app.run(debug=True)
