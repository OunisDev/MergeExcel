from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from io import BytesIO

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/merge', methods=['POST'])
def merge_excel():
    files = request.files.getlist('excel_files')
    combined_df = pd.DataFrame()

    for file in files:
        # Read all sheets in the Excel file
        xls = pd.read_excel(file, sheet_name=None)
        for sheet_name, df in xls.items():
            df['Source Sheet'] = sheet_name
            df['Source File'] = file.filename
            combined_df = pd.concat([combined_df, df], ignore_index=True)

    # Save merged file to memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        combined_df.to_excel(writer, index=False, sheet_name='Merged')
    output.seek(0)

    return send_file(output, download_name='merged.xlsx', as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
